import os
import re
from datetime import datetime, timezone, timedelta
from pathlib import Path

import time

import feedparser
import requests
import yfinance as yf  # Lấy giá vàng, dầu, tỷ giá, S&P 500, Bitcoin MIỄN PHÍ từ Yahoo Finance
from bs4 import BeautifulSoup
import trafilatura
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google import genai
from google.genai import types

TELEGRAM_BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
TELEGRAM_CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]
GEMINI_API_KEY = os.environ["GEMINI_API_KEY"]

VN_TZ = timezone(timedelta(hours=7))

TARGET_PAGES = [
    ("Investing.com VN", "https://vn.investing.com/news", "https://vn.investing.com/rss/news.rss"),
    ("Vietstock Thế giới", "https://vietstock.vn/the-gioi.htm", "https://vietstock.vn/736/the-gioi.rss"),
    ("CafeF Tài chính quốc tế", "https://cafef.vn/tai-chinh-quoc-te.chn", "https://cafef.vn/tai-chinh-quoc-te.rss"),
]

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "vi-VN,vi;q=0.9,en-US;q=0.8,en;q=0.7",
}

client = genai.Client(api_key=GEMINI_API_KEY)


def fetch_rss_items(rss_url: str, source_name: str, limit: int = 8) -> list[dict]:
    items = []
    try:
        resp = requests.get(rss_url, headers=HEADERS, timeout=15)
        if resp.status_code == 200 and resp.text.strip():
            parsed = feedparser.parse(resp.text)
            for entry in parsed.entries[:limit]:
                title = (entry.get("title") or "").strip()
                link = (entry.get("link") or "").strip()
                summary = re.sub(r"<[^>]+>", " ", entry.get("summary") or "").strip()
                if title and link:
                    items.append({
                        "source": source_name,
                        "title": title,
                        "link": link,
                        "snippet": re.sub(r"\s+", " ", summary)[:300],
                    })
    except Exception as e:
        print(f"[RSS warning] {source_name}: {e}")
    return items


def fetch_html_links(page_url: str, source_name: str, limit: int = 10) -> list[dict]:
    items = []
    seen_urls = set()
    try:
        resp = requests.get(page_url, headers=HEADERS, timeout=20)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, "html.parser")

        for a in soup.find_all("a", href=True):
            title = re.sub(r"\s+", " ", a.get_text(" ", strip=True))
            href = a["href"].strip()
            if len(title) < 25:
                continue
            if href.startswith("/"):
                base = "/".join(page_url.split("/")[:3])
                href = base + href
            if not href.startswith("http"):
                continue
            if href in seen_urls:
                continue
            # Lọc bớt link điều hướng / chuyên mục
            if any(skip in href for skip in ["/rss", "/login", "/register", "/pro/", "#"]):
                continue

            seen_urls.add(href)
            items.append({
                "source": source_name,
                "title": title,
                "link": href,
                "snippet": "",
            })
            if len(items) >= limit:
                break
    except Exception as e:
        print(f"[HTML warning] {source_name}: {e}")
    return items


def enrich_snippets(items: list[dict], max_articles: int = 18) -> list[dict]:
    for item in items[:max_articles]:
        if item["snippet"]:
            continue
        try:
            downloaded = trafilatura.fetch_url(item["link"])
            if downloaded:
                text = trafilatura.extract(downloaded, include_comments=False, include_tables=False)
                if text:
                    item["snippet"] = re.sub(r"\s+", " ", text)[:320]
        except Exception:
            pass
    return items


def collect_news() -> list[dict]:
    all_items = []
    for source_name, page_url, rss_url in TARGET_PAGES:
        rss_items = fetch_rss_items(rss_url, source_name, limit=8)
        html_items = fetch_html_links(page_url, source_name, limit=8)

        combined = []
        seen_links = set()
        for it in rss_items + html_items:
            if it["link"] not in seen_links:
                seen_links.add(it["link"])
                combined.append(it)
            if len(combined) >= 10:
                break

        all_items.extend(combined)

    return enrich_snippets(all_items, max_articles=18)


# ============================================================
# PHẦN 1: LẤY GIÁ MIỄN PHÍ (thay cho Google Search của Gemini)
# ============================================================

# Mã (ticker) trên Yahoo Finance của từng sản phẩm
YF_TICKERS = [
    ("Vàng thế giới (USD/oz)", "GC=F", "{:,.1f} USD"),
    ("Dầu thô WTI", "CL=F", "{:,.2f} USD/thùng"),
    ("Dầu Brent", "BZ=F", "{:,.2f} USD/thùng"),
    ("Tỷ giá USD/VND", "VND=X", "{:,.0f} VND"),
    ("S&P 500", "^GSPC", "{:,.2f} điểm"),
    ("Bitcoin (BTC/USD)", "BTC-USD", "{:,.0f} USD"),
]


def fetch_sjc_price() -> dict:
    """Thử lấy giá vàng SJC từ web SJC. Lỗi thì ghi 'Đang cập nhật', không làm hỏng cả script."""
    row = {"product": "Vàng SJC (VN)", "price": "Đang cập nhật", "note": "Xem sjc.com.vn"}
    try:
        resp = requests.post(
            "https://sjc.com.vn/GoldPrice/Services/PriceService.ashx",
            headers=HEADERS, timeout=15,
        )
        data = resp.json().get("data", [])
        for item in data:
            name = item.get("TypeName", "")
            if "SJC" in name and ("1L" in name or "1 lượng" in name.lower() or "10L" in name):
                buy = item.get("BuyValue")
                sell = item.get("SellValue")
                if buy and sell:
                    row["price"] = f"Mua {buy:,.0f} / Bán {sell:,.0f} VND/lượng"
                    row["note"] = item.get("BranchName", "SJC")
                break
    except Exception as e:
        print(f"[SJC warning] {e}")
    return row


def fetch_prices() -> list[dict]:
    """Lấy giá các sản phẩm tài chính, kèm % thay đổi so với phiên trước."""
    rows = [fetch_sjc_price()]
    for name, ticker, fmt in YF_TICKERS:
        row = {"product": name, "price": "Đang cập nhật", "note": ""}
        try:
            hist = yf.Ticker(ticker).history(period="5d")["Close"].dropna()
            if len(hist) >= 1:
                last = float(hist.iloc[-1])
                row["price"] = fmt.format(last)
            if len(hist) >= 2:
                prev = float(hist.iloc[-2])
                pct = (last - prev) / prev * 100
                arrow = "▲" if pct > 0 else ("▼" if pct < 0 else "■")
                row["note"] = f"{arrow} {pct:+.2f}% so với phiên trước"
        except Exception as e:
            print(f"[Price warning] {name}: {e}")
        rows.append(row)
    return rows


# ============================================================
# PHẦN 2: GỌI GEMINI (free) — có thử lại + model dự phòng
# ============================================================

# Thử lần lượt: model nào hết quota thì chuyển sang model sau
# (Google đã ngừng 2.5 cho tài khoản mới -> dùng 3.5)
GEMINI_MODELS = ["gemini-3.5-flash", "gemini-3.5-flash-lite"]


def get_model_list() -> list[str]:
    """Danh sách model để thử: model cố định ở trên + các model 'flash' mà API key
    của bạn đang được phép dùng (tự dò, để lần sau Google đổi tên model không bị lỗi nữa)."""
    models = list(GEMINI_MODELS)
    try:
        for m in client.models.list():
            name = (m.name or "").replace("models/", "")
            actions = getattr(m, "supported_actions", None) or []
            if "flash" in name and "generateContent" in actions \
                    and not any(x in name for x in ["image", "tts", "audio", "live", "embedding"]) \
                    and name not in models:
                models.append(name)
    except Exception as e:
        print(f"[Gemini] Không lấy được danh sách model: {e}")
    print(f"[Gemini] Sẽ thử các model: {models[:6]}")
    return models[:6]


def call_gemini(prompt: str) -> str:
    for model in get_model_list():
        for attempt in range(3):  # mỗi model thử tối đa 3 lần
            try:
                resp = client.models.generate_content(
                    model=model,
                    contents=prompt,
                    # KHÔNG dùng tools google_search nữa -> không tốn quota Search
                    config=types.GenerateContentConfig(temperature=0.2),
                )
                print(f"[Gemini] OK với model {model}")
                return resp.text or ""
            except Exception as e:
                msg = str(e)
                if "429" in msg or "RESOURCE_EXHAUSTED" in msg or "503" in msg or "UNAVAILABLE" in msg:
                    wait = 30 * (attempt + 1)  # chờ 30s, 60s, 90s
                    print(f"[Gemini] {model} quá tải/hết quota, chờ {wait}s... ({msg[:120]})")
                    time.sleep(wait)
                else:
                    print(f"[Gemini] {model} lỗi khác: {msg[:200]}")
                    break  # lỗi khác -> chuyển model luôn
        print(f"[Gemini] Bỏ qua model {model}")
    return ""  # tất cả đều lỗi -> trả rỗng, script vẫn gửi bản tin dự phòng


def fallback_digest(articles: list[dict]) -> str:
    """Khi Gemini không dùng được: liệt kê tiêu đề tin để vẫn có bản tin gửi đi."""
    lines = ["Không gọi được AI hôm nay (hết quota free). Danh sách tin nổi bật:"]
    for a in articles[:15]:
        lines.append(f"- [{a['source']}] {a['title']} — {a['link']}")
    return "\n".join(lines)


def generate_digest_and_prices(articles: list[dict], now_vn: datetime) -> tuple[str, list[dict], str]:
    news_lines = "\n".join(
        f"- [{a['source']}] {a['title']} | Link: {a['link']} | Tóm tắt: {a['snippet']}"
        for a in articles
    )

    # Giá lấy bằng Python, rồi đưa cho Gemini làm dữ liệu
    prices_rows = fetch_prices()
    price_lines = "\n".join(f"- {r['product']}: {r['price']} ({r['note']})" for r in prices_rows)

    prompt = f"""
Thời điểm hiện tại: {now_vn.strftime("%d/%m/%Y %H:%M")} (Giờ Việt Nam, UTC+7).

Bạn là chuyên viên phân tích thị trường tài chính. Hãy tổng hợp bản tin buổi sáng dựa trên BẢNG GIÁ và dữ liệu tin tức được cung cấp bên dưới (từ Investing.com VN, Vietstock Thế giới, CafeF Tài chính quốc tế). Chỉ dùng số liệu có trong dữ liệu, không tự bịa giá.

Bảng giá mới nhất:
{price_lines}

Hãy trả về đúng cấu trúc sau (không dùng markdown code block bọc toàn bộ):

===DIGEST===
Viết bản tin sáng bằng tiếng Việt, văn phong báo tài chính, theo đúng hình thức sau:
- Đoạn mở đầu: 1-2 câu nêu tổng quan phiên Phố Wall / thị trường gần nhất, có xu hướng chính.
- Đoạn số liệu: nêu rõ mức điểm và % thay đổi của S&P 500, Nasdaq, Dow Jones (lấy từ BẢNG GIÁ), viết thành câu văn liền mạch, không gạch đầu dòng.
- Các đoạn tiếp theo: mỗi đoạn một chủ đề (công nghệ/AI, chính sách tiền tệ, hàng hóa, tiền tệ, sự kiện doanh nghiệp). Mỗi đoạn 3-5 câu, nêu thông tin chính rồi tác động.
- Đoạn cuối: một câu nêu điểm cần theo dõi trong ngày.
- Các cụm từ quan trọng viết trong **dấu sao kép** để nhấn mạnh.
- Không dùng tiêu đề mục, không dùng bảng. Kèm link nguồn bằng cách ghi tên báo trong ngoặc đơn nếu có.
Không bịa số liệu; nếu thiếu số liệu nào thì ghi "Đang cập nhật".

===TELEGRAM_CAPTION===
Tóm tắt ngắn gọn nhất (dưới 800 ký tự) cho tin nhắn thông báo gửi kèm file Word lúc 7h sáng, cùng văn phong trên.

Dữ liệu tin tức thu thập được:
{news_lines}
"""

    raw = call_gemini(prompt)

    if not raw.strip():
        # Gemini lỗi hết -> vẫn gửi bản tin dự phòng, không crash
        digest_text = fallback_digest(articles)
        caption_text = f"Bản tin sáng {now_vn.strftime('%d/%m/%Y')}: AI hết quota, gửi bảng giá + tiêu đề tin."
        return digest_text, prices_rows, caption_text

    digest_match = re.search(r"===DIGEST===\s*(.*?)\s*===TELEGRAM_CAPTION===", raw, re.DOTALL)
    digest_text = digest_match.group(1).strip() if digest_match else raw.strip()

    caption_match = re.search(r"===TELEGRAM_CAPTION===\s*(.*)$", raw, re.DOTALL)
    caption_text = caption_match.group(1).strip() if caption_match else (
        f"Bản tin tài chính sáng {now_vn.strftime('%d/%m/%Y')} đã sẵn sàng."
    )

    return digest_text, prices_rows, caption_text


def add_formatted_markdown_lines(doc: Document, text: str) -> None:
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            continue

        # Loại bỏ ký tự markdown heading/bold đơn giản để file Word sạch đẹp
        clean = stripped.replace("**", "")
        if clean.startswith("### "):
            p = doc.add_heading(clean[4:], level=3)
        elif clean.startswith("## "):
            p = doc.add_heading(clean[3:], level=2)
        elif clean.startswith("# "):
            p = doc.add_heading(clean[2:], level=1)
        elif clean.startswith(("- ", "* ")):
            p = doc.add_paragraph(clean[2:], style="List Bullet")
            for run in p.runs:
                run.font.size = Pt(10.5)
        else:
            p = doc.add_paragraph(clean)
            for run in p.runs:
                run.font.size = Pt(10.5)


def build_docx(date_str: str, time_str: str, prices_rows: list[dict], digest_text: str) -> Path:
    out_dir = Path("output")
    out_dir.mkdir(exist_ok=True)
    file_path = out_dir / f"Ban_Tin_Tai_Chinh_{date_str}.docx"

    doc = Document()

    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(f"BẢN TIN TÀI CHÍNH BUỔI SÁNG — {date_str}")
    title_run.bold = True
    title_run.font.size = Pt(16)
    title_run.font.color.rgb = RGBColor(0x1B, 0x36, 0x5D)

    meta_p = doc.add_paragraph()
    meta_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta_run = meta_p.add_run(
        f"Cập nhật lúc: {time_str} (GMT+7) | Nguồn: Investing.com VN, Vietstock Thế giới, CafeF Tài chính quốc tế"
    )
    meta_run.italic = True
    meta_run.font.size = Pt(9.5)
    meta_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)

    doc.add_heading("1. Bảng cập nhật giá sản phẩm tài chính", level=2)

    table = doc.add_table(rows=1, cols=3)
    table.style = "Table Grid"
    hdr_cells = table.rows[0].cells
    hdr_titles = ["Sản phẩm tài chính", "Giá cập nhật", "Biến động / Ghi chú"]
    for idx, h_text in enumerate(hdr_titles):
        hdr_cells[idx].text = h_text
        for run in hdr_cells[idx].paragraphs[0].runs:
            run.bold = True
            run.font.size = Pt(10)

    if prices_rows:
        for row in prices_rows:
            cells = table.add_row().cells
            cells[0].text = row["product"]
            cells[1].text = row["price"]
            cells[2].text = row["note"]
            for c in cells:
                for run in c.paragraphs[0].runs:
                    run.font.size = Pt(10)
    else:
        cells = table.add_row().cells
        cells[0].text = "Tổng hợp giá"
        cells[1].text = "Xem chi tiết trong nội dung bên dưới"
        cells[2].text = ""

    doc.add_paragraph("")
    doc.add_heading("2. Tổng hợp tin tức nổi bật", level=2)
    add_formatted_markdown_lines(doc, digest_text)

    doc.save(file_path)
    return file_path


def check_telegram_response(resp: requests.Response, action: str) -> None:
    """In rõ LÝ DO Telegram từ chối (thay vì chỉ báo '400 Bad Request')."""
    if not resp.ok:
        try:
            reason = resp.json().get("description", resp.text)
        except Exception:
            reason = resp.text
        raise RuntimeError(f"Telegram {action} lỗi {resp.status_code}: {reason}")


def send_telegram_document(file_path: Path, caption: str) -> None:
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendDocument"
    with file_path.open("rb") as f:
        resp = requests.post(
            url,
            data={
                "chat_id": TELEGRAM_CHAT_ID.strip(),
                "caption": caption[:950],  # Telegram giới hạn 1024 ký tự, chừa dư cho emoji
            },
            files={
                "document": (file_path.name, f, "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            },
            timeout=60,
        )
    check_telegram_response(resp, "gửi file")


def send_telegram_message(text: str) -> None:
    url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    for i in range(0, len(text), 3800):
        chunk = text[i : i + 3800]
        resp = requests.post(
            url,
            json={
                "chat_id": TELEGRAM_CHAT_ID.strip(),
                "text": chunk,
                "disable_web_page_preview": True,
            },
            timeout=30,
        )
        check_telegram_response(resp, "gửi tin nhắn")


def run() -> None:
    now_vn = datetime.now(VN_TZ)
    date_str = now_vn.strftime("%Y-%m-%d")
    time_str = now_vn.strftime("%H:%M %d/%m/%Y")

    articles = collect_news()
    digest_text, prices_rows, caption_text = generate_digest_and_prices(articles, now_vn)
    docx_path = build_docx(date_str, time_str, prices_rows, digest_text)

    header_caption = f"📊 Bản tin tài chính sáng {now_vn.strftime('%d/%m/%Y')} (7:00)\n\n{caption_text}"
    send_telegram_document(docx_path, header_caption)

    # Nếu muốn đọc nhanh cả phần tóm tắt ngay trên chat Telegram bên cạnh file docx:
    send_telegram_message(f"📰 TÓM TẮT NHANH ({now_vn.strftime('%d/%m/%Y')}):\n\n{digest_text}")


if __name__ == "__main__":
    run()
