"""
Script: Báo cáo doanh thu THEO NGÀNH HÀNG (Top ngành hàng + Top hãng) từ file Excel
"Doanh Thu Theo Ngành Hàng" và gửi dạng THẺ (Flex Message) vào group LINE.

CÁCH DÙNG:
    python send_category_report.py --file "DoanhThuTheoNgànhHàng_xxx.xlsx"
    python send_category_report.py --file "..." --date 2026-11-01 --top 6 --dry-run

    --date     bỏ trống -> lấy NGÀY MỚI NHẤT trong file.
    --top      số ngành hàng hiển thị (mặc định 5), phần còn lại gộp vào "Ngành hàng khác".
    --dry-run  chỉ in ra màn hình, KHÔNG gửi LINE.

Nếu file có nhiều siêu thị -> mỗi siêu thị 1 thẻ, gửi dạng carousel (tối đa 12 thẻ).

CẤU HÌNH (biến môi trường):
    set LINE_CHANNEL_ACCESS_TOKEN=xxxx
    set LINE_GROUP_ID=xxxx
"""

import argparse
import os

import pandas as pd
import requests

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
LINE_GROUP_ID = os.environ.get("LINE_GROUP_ID", "")
LINE_PUSH_URL = "https://api.line.me/v2/bot/message/push"

COLOR_HEADER_BG = "#E67E22"
COLOR_TOTAL = "#D35400"
COLOR_LABEL = "#555555"
COLOR_MUTED = "#888888"
COLOR_WHITE = "#FFFFFF"

WEEKDAY_VN = ["THỨ HAI", "THỨ BA", "THỨ TƯ", "THỨ NĂM", "THỨ SÁU", "THỨ BẢY", "CHỦ NHẬT"]

COL_DATE = "Ngày xuất"
COL_STORE = "Mã siêu thị"
COL_CATEGORY = "Ngành hàng"
COL_BRAND = "Tên Hãng"
COL_REVENUE = "Doanh thu"  # chưa VAT, cùng cơ sở với cột "Doanh thu offline/online" của báo cáo siêu thị


def format_money(value: float) -> str:
    return f"{round(value):,}".replace(",", ".") + "đ"


def strip_code(text: str) -> str:
    """'1054 - Bánh kẹo - Trà' -> 'Bánh kẹo - Trà' (bỏ mã số đứng đầu)."""
    text = str(text)
    head, sep, tail = text.partition(" - ")
    return tail if sep and head.strip().isdigit() else text


def shorten(text: str, limit: int = 34) -> str:
    return text if len(text) <= limit else text[: limit - 1].rstrip() + "…"


def load_day(file_path: str, target_date) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    missing = [c for c in (COL_DATE, COL_STORE, COL_CATEGORY, COL_BRAND, COL_REVENUE) if c not in df.columns]
    if missing:
        raise ValueError(f"File thiếu cột: {', '.join(missing)}")
    df[COL_DATE] = pd.to_datetime(df[COL_DATE])
    wanted = pd.to_datetime(target_date) if target_date else df[COL_DATE].max()
    day = df[df[COL_DATE].dt.date == wanted.date()]
    if day.empty:
        available = ", ".join(d.strftime("%d/%m/%Y") for d in sorted(df[COL_DATE].dt.normalize().unique()))
        raise ValueError(f"Không có dữ liệu ngày {wanted.strftime('%d/%m/%Y')}. Các ngày có trong file: {available}")
    return day


def _row_box(label: str, value: str, pct: str = "", bold: bool = False) -> dict:
    contents = [
        {"type": "text", "text": label, "size": "sm", "color": COLOR_LABEL, "flex": 5, "wrap": True},
        {"type": "text", "text": value, "size": "sm", "color": "#111111", "align": "end", "flex": 3,
         "weight": "bold" if bold else "regular"},
    ]
    if pct:
        contents.append({"type": "text", "text": pct, "size": "xs", "color": COLOR_MUTED, "align": "end", "flex": 1})
    return {"type": "box", "layout": "horizontal", "contents": contents}


def build_bubble(store_df: pd.DataFrame, top_n: int):
    date = store_df[COL_DATE].iloc[0]
    weekday = WEEKDAY_VN[date.weekday()]
    store = str(store_df[COL_STORE].iloc[0])
    total = store_df[COL_REVENUE].sum()

    cats = (store_df.groupby(COL_CATEGORY)[COL_REVENUE].sum()
            .sort_values(ascending=False))
    top_cats, other = cats.head(top_n), cats.iloc[top_n:].sum()

    brands = (store_df.groupby(COL_BRAND)[COL_REVENUE].sum()
              .sort_values(ascending=False).head(3))

    def pct(v):
        return f"{v / total * 100:.0f}%" if total else "0%"

    cat_rows = [_row_box(shorten(strip_code(name)), format_money(v), pct(v)) for name, v in top_cats.items()]
    if other > 0:
        cat_rows.append(_row_box("Ngành hàng khác", format_money(other), pct(other)))
    brand_rows = [_row_box(str(name), format_money(v), pct(v)) for name, v in brands.items()]

    bubble = {
        "type": "bubble",
        "size": "mega",
        "header": {
            "type": "box", "layout": "vertical", "backgroundColor": COLOR_HEADER_BG,
            "paddingAll": "20px", "spacing": "sm",
            "contents": [
                {"type": "text", "text": "🛒 DOANH THU THEO NGÀNH HÀNG", "color": COLOR_WHITE, "weight": "bold", "size": "md"},
                {"type": "text", "text": f"📅 {weekday}, {date.strftime('%d/%m/%Y')}", "color": COLOR_WHITE, "size": "sm"},
            ],
        },
        "body": {
            "type": "box", "layout": "vertical", "paddingAll": "20px", "spacing": "md",
            "contents": [
                {"type": "text", "text": f"🏪 {store}".upper(), "weight": "bold", "size": "md", "wrap": True},
                {"type": "separator", "margin": "md"},
                {"type": "text", "text": f"TOP {len(top_cats)} NGÀNH HÀNG", "size": "xs", "weight": "bold", "color": COLOR_TOTAL},
                *cat_rows,
                {"type": "separator", "margin": "md"},
                {"type": "text", "text": "TOP 3 HÃNG", "size": "xs", "weight": "bold", "color": COLOR_TOTAL},
                *brand_rows,
                {"type": "separator", "margin": "md"},
                {
                    "type": "box", "layout": "horizontal", "margin": "md",
                    "contents": [
                        {"type": "text", "text": "💰 TỔNG DOANH THU", "weight": "bold", "size": "md", "flex": 3},
                        {"type": "text", "text": format_money(total), "weight": "bold", "size": "lg",
                         "color": COLOR_TOTAL, "align": "end", "flex": 2},
                    ],
                },
            ],
        },
    }
    top1 = strip_code(top_cats.index[0]) if len(top_cats) else "-"
    alt_text = (f"DOANH THU NGÀNH HÀNG {date.strftime('%d/%m/%Y')} - Tổng: {format_money(total)} "
                f"- Dẫn đầu: {shorten(top1, 30)}")
    return bubble, alt_text


def send_flex_to_line(contents: dict, alt_text: str) -> None:
    if not LINE_CHANNEL_ACCESS_TOKEN or not LINE_GROUP_ID:
        raise RuntimeError("Thiếu LINE_CHANNEL_ACCESS_TOKEN hoặc LINE_GROUP_ID. Set biến môi trường trước khi chạy.")
    resp = requests.post(
        LINE_PUSH_URL,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}"},
        json={"to": LINE_GROUP_ID, "messages": [{"type": "flex", "altText": alt_text[:400], "contents": contents}]},
        timeout=15,
    )
    if resp.status_code != 200:
        raise RuntimeError(f"Gửi LINE thất bại ({resp.status_code}): {resp.text}")
    print("Đã gửi báo cáo ngành hàng vào LINE thành công.")


def main():
    parser = argparse.ArgumentParser(description="Báo cáo doanh thu theo ngành hàng -> LINE (dạng thẻ)")
    parser.add_argument("--file", required=True, help="Đường dẫn file Excel 'Doanh Thu Theo Ngành Hàng'")
    parser.add_argument("--date", default=None, help="Ngày cần báo cáo YYYY-MM-DD (bỏ trống = ngày mới nhất)")
    parser.add_argument("--top", type=int, default=5, help="Số ngành hàng hiển thị (mặc định 5)")
    parser.add_argument("--dry-run", action="store_true", help="Chỉ in ra, không gửi LINE")
    args = parser.parse_args()

    day = load_day(args.file, args.date)
    results = [build_bubble(g, max(1, args.top)) for _, g in day.groupby(COL_STORE)]
    bubbles = [b for b, _ in results][:12]
    alt_text = results[0][1] if len(results) == 1 else f"DOANH THU NGÀNH HÀNG {day[COL_DATE].iloc[0].strftime('%d/%m/%Y')} ({len(bubbles)} siêu thị)"
    contents = bubbles[0] if len(bubbles) == 1 else {"type": "carousel", "contents": bubbles}

    print("----- ALT TEXT -----")
    print(alt_text)
    for g_store, g in day.groupby(COL_STORE):
        top = g.groupby(COL_CATEGORY)[COL_REVENUE].sum().sort_values(ascending=False).head(args.top)
        print(f"\n{g_store}  |  Tổng: {format_money(g[COL_REVENUE].sum())}")
        for name, v in top.items():
            print(f"  {shorten(strip_code(name), 45):<46}{format_money(v):>14}  {v / g[COL_REVENUE].sum() * 100:4.1f}%")
    print("--------------------")

    if args.dry_run:
        print("(dry-run: chưa gửi LINE)")
        return
    send_flex_to_line(contents, alt_text)


if __name__ == "__main__":
    main()
