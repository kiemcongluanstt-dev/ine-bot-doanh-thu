"""
CHẠY NỀN LIÊN TỤC TRÊN MÁY TÍNH - PHIÊN BẢN 4
================================================================
- Lệnh DT  -> Doanh thu toàn siêu thị + chi tiết ngành hàng, so sánh cùng thứ tuần trước.
- Lệnh NAM -> Doanh thu ngành hàng Nấm, LIỆT KÊ TỪNG SẢN PHẨM (vd: Nấm hải sản TQ,
              Nấm đùi gà lớn tươi TQ...), so sánh mỗi sản phẩm với cùng thứ tuần trước.

3 FILE CẦN CÓ TRONG CÙNG THƯ MỤC (đặt tên như export gốc từ hệ thống, không đổi tên):
  1. "_Doanh_Thu_Theo_Sieu_Thi*.xlsx"  -> tổng doanh thu toàn siêu thị (cho lệnh DT)
  2. "Doanh_Thu_Theo_Nganh_Hang*.xlsx" -> chi tiết ngành hàng (cho lệnh DT)
  3. "_Doanh_Thu_Chi_Tiet*.xlsx"       -> chi tiết từng sản phẩm (cho lệnh NAM)

Không cần gõ lệnh gì - chỉ cần xuất file Excel như bình thường (tải về ĐÚNG
THƯ MỤC này), script tự phát hiện file mới/thay đổi và đẩy báo cáo lên server.

CÁCH DÙNG:
    set CLOUD_URL=https://dt-line-bot.onrender.com
    set PUSH_SECRET=chuoi_bi_mat_da_dat_tren_render
    python watch_and_push.py

    -> Để CỬA SỔ NÀY MỞ LIÊN TỤC (không đóng). Ctrl+C để dừng khi cần.
"""

import glob
import os
import time
from datetime import timedelta

import pandas as pd
import requests

CLOUD_URL = os.environ.get("CLOUD_URL", "")
PUSH_SECRET = os.environ.get("PUSH_SECRET", "")

STORE_CODE = "14978"  # mã siêu thị đang theo dõi
MUC_TIEU_NGAY_NAM = 164000  # mục tiêu doanh thu ngành hàng Nấm mỗi ngày (đ) - sửa số này nếu mục tiêu thay đổi

FRESH_CATEGORIES = [
    "Rau Củ Các Loại",
    "Trái Cây Các Loại",
    "Thịt gia cầm gia súc các loại",
    "Thủy Hải Sản Các Loại",
]

FILE_PATTERN_SIEU_THI = "_Doanh_Thu_Theo_Sieu_Thi*.xlsx"
FILE_PATTERN_NGANH_HANG = "Doanh_Thu_Theo_Nganh_Hang*.xlsx"
FILE_PATTERN_CHI_TIET = "_Doanh_Thu_Chi_Tiet*.xlsx"
CHECK_INTERVAL_SECONDS = 15

COLOR_HEADER_BG = "#2E8B57"
COLOR_HEADER_BG_NAM = "#8D6E63"
COLOR_HEADER_BG_FRESH = "#00897B"
COLOR_TOTAL = "#1DB954"
COLOR_LABEL = "#555555"
COLOR_WHITE = "#FFFFFF"
COLOR_UP = "#1DB954"
COLOR_DOWN = "#E53935"
COLOR_NEUTRAL = "#999999"

WEEKDAY_VN = [
    "THỨ HAI", "THỨ BA", "THỨ TƯ", "THỨ NĂM",
    "THỨ SÁU", "THỨ BẢY", "CHỦ NHẬT",
]

ST_COL_DATE = "Ngày"
ST_COL_STORE_CODE = "Mã siêu thị"
ST_COL_STORE_NAME = "Tên siêu thị"
ST_COL_OFFLINE = "Doanh thu offline"
ST_COL_ONLINE = "Doanh thu Online"
ST_COL_BILL = "Tổng số bill"

NH_COL_DATE = "Ngày xuất"
NH_COL_STORE = "Mã siêu thị"
NH_COL_NGANH_HANG = "Ngành hàng"
NH_COL_REVENUE = "Doanh thu"

CT_COL_DATE = "Ngày xuất"
CT_COL_STORE = "Mã siêu thị"
CT_COL_STORE_NAME = "Tên siêu thị"
CT_COL_PRODUCT = "Tên sản phẩm"
CT_COL_NHOM_HANG = "Nhóm hàng"
CT_COL_UNIT = "Đơn vị"
CT_COL_QTY = "Tổng SL bán"
CT_COL_REVENUE = "Thành tiền phải thu khách hàng (chưa VAT)"


def log(msg: str):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")


def format_money(value: float) -> str:
    return f"{round(value):,}".replace(",", ".") + "đ"


def find_latest_file(pattern: str):
    files = glob.glob(pattern)
    if not files:
        return None
    return max(files, key=os.path.getmtime)


def short_name(raw: str, max_len: int = 24) -> str:
    parts = str(raw).split(" - ", 1)
    name = parts[1] if len(parts) > 1 else str(raw)
    if len(name) > max_len:
        name = name[: max_len - 1].rstrip() + "…"
    return name


def short_product_name(raw: str, max_len: int = 30) -> str:
    name = str(raw).strip()
    if len(name) > max_len:
        name = name[: max_len - 1].rstrip() + "…"
    return name


def change_text_and_color(curr: float, prev):
    if prev is None:
        return "(chưa có dữ liệu tuần trước)", COLOR_NEUTRAL
    if prev == 0:
        return "(tuần trước = 0)", COLOR_NEUTRAL
    pct = (curr - prev) / prev * 100
    if pct >= 0:
        return f"▲ {pct:.0f}% so với tuần trước", COLOR_UP
    return f"▼ {abs(pct):.0f}% so với tuần trước", COLOR_DOWN


def _row_box(label: str, value: str) -> dict:
    return {
        "type": "box",
        "layout": "horizontal",
        "contents": [
            {"type": "text", "text": label, "size": "sm", "color": COLOR_LABEL, "flex": 3},
            {"type": "text", "text": value, "size": "sm", "color": "#111111", "align": "end", "flex": 2, "weight": "bold"},
        ],
    }


def _target_box(current: float, target: float) -> dict:
    pct = (current / target * 100) if target else 0
    if pct >= 100:
        color = COLOR_UP
        bg = "#E8F5E9"
    elif pct >= 70:
        color = "#F9A825"
        bg = "#FFF8E1"
    else:
        color = COLOR_DOWN
        bg = "#FFEBEE"

    return {
        "type": "box",
        "layout": "vertical",
        "margin": "md",
        "paddingAll": "10px",
        "cornerRadius": "8px",
        "backgroundColor": bg,
        "contents": [
            {"type": "text", "text": "🎯 MỤC TIÊU NGÀY", "weight": "bold", "size": "sm", "color": color},
            {
                "type": "box",
                "layout": "horizontal",
                "margin": "xs",
                "contents": [
                    {"type": "text", "text": f"{format_money(current)} / {format_money(target)}", "size": "sm", "color": "#333333", "flex": 3, "wrap": True},
                    {"type": "text", "text": f"{pct:.0f}%", "size": "lg", "weight": "bold", "color": color, "align": "end", "flex": 1},
                ],
            },
        ],
    }


def _card_row(label: str, value: str, change_text: str, change_color: str) -> dict:
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "sm",
        "paddingAll": "8px",
        "borderWidth": "1px",
        "borderColor": "#E0E0E0",
        "cornerRadius": "8px",
        "backgroundColor": "#FAFAFA",
        "contents": [
            {
                "type": "box",
                "layout": "horizontal",
                "contents": [
                    {"type": "text", "text": label, "size": "xs", "color": COLOR_LABEL, "flex": 3, "wrap": True},
                    {"type": "text", "text": value, "size": "xs", "color": "#111111", "align": "end", "flex": 2, "weight": "bold"},
                ],
            },
            {"type": "text", "text": change_text, "size": "xxs", "color": change_color, "align": "end", "margin": "xs"},
        ],
    }


def load_sieu_thi_df(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df[ST_COL_DATE] = pd.to_datetime(df[ST_COL_DATE])
    return df


def get_sieu_thi_row_for_date(df: pd.DataFrame, date):
    matches = df[df[ST_COL_DATE].dt.date == date]
    if matches.empty:
        return None
    return matches.iloc[0]


def load_nganh_hang_df(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df[NH_COL_DATE] = pd.to_datetime(df[NH_COL_DATE])
    return df


def get_category_breakdown_for_date(df: pd.DataFrame, store_code: str, date):
    mask_store = df[NH_COL_STORE].astype(str).str.startswith(str(store_code))
    mask_date = df[NH_COL_DATE].dt.date == date
    day_df = df[mask_store & mask_date]
    if day_df.empty:
        return None
    return day_df.groupby(NH_COL_NGANH_HANG)[NH_COL_REVENUE].sum()


def build_full_dt_bubble(sieu_thi_path: str, nganh_hang_path):
    st_df = load_sieu_thi_df(sieu_thi_path)
    target_date = st_df[ST_COL_DATE].max().date()
    row = get_sieu_thi_row_for_date(st_df, target_date)
    weekday = WEEKDAY_VN[row[ST_COL_DATE].weekday()]

    store_code = row[ST_COL_STORE_CODE]
    store_name = str(row[ST_COL_STORE_NAME]).upper()
    offline = row[ST_COL_OFFLINE]
    online = row[ST_COL_ONLINE]
    bill_count = row[ST_COL_BILL]
    total = offline + online
    bill_value = total / bill_count if bill_count else 0

    last_week_date = target_date - timedelta(days=7)
    last_week_row = get_sieu_thi_row_for_date(st_df, last_week_date)
    last_week_total = None
    if last_week_row is not None:
        last_week_total = last_week_row[ST_COL_OFFLINE] + last_week_row[ST_COL_ONLINE]
    change_text, change_color = change_text_and_color(total, last_week_total)

    body_contents = [
        {"type": "text", "text": f"🏪 {store_code} - {store_name}", "weight": "bold", "size": "md", "wrap": True},
        {"type": "separator", "margin": "md"},
        _row_box("🏬 Doanh thu offline", format_money(offline)),
        _row_box("🛍️ Doanh thu online", format_money(online)),
        _row_box("🧾 Số lượng bill", str(int(bill_count))),
        _row_box("📈 Giá trị bill", format_money(bill_value)),
        {"type": "separator", "margin": "md"},
        {
            "type": "box",
            "layout": "horizontal",
            "margin": "md",
            "contents": [
                {"type": "text", "text": "💰 TỔNG DOANH THU", "weight": "bold", "size": "md", "flex": 3},
                {"type": "text", "text": format_money(total), "weight": "bold", "size": "lg", "color": COLOR_TOTAL, "align": "end", "flex": 2},
            ],
        },
        {"type": "text", "text": change_text, "size": "xs", "color": change_color, "align": "end"},
    ]

    if nganh_hang_path:
        nh_df = load_nganh_hang_df(nganh_hang_path)
        today_cats = get_category_breakdown_for_date(nh_df, store_code, target_date)
        last_week_cats = get_category_breakdown_for_date(nh_df, store_code, last_week_date)

        if today_cats is not None:
            body_contents.append({"type": "separator", "margin": "lg"})
            body_contents.append({"type": "text", "text": "📦 CHI TIẾT NGÀNH HÀNG", "weight": "bold", "size": "sm", "margin": "md"})
            for cat_name, revenue in today_cats.sort_values(ascending=False).items():
                prev_value = None
                if last_week_cats is not None and cat_name in last_week_cats.index:
                    prev_value = last_week_cats[cat_name]
                c_text, c_color = change_text_and_color(revenue, prev_value)
                body_contents.append(_card_row(short_name(cat_name), format_money(revenue), c_text, c_color))
        else:
            body_contents.append({"type": "separator", "margin": "lg"})
            body_contents.append({"type": "text", "text": "📦 Chưa có dữ liệu chi tiết ngành hàng cho ngày này.", "size": "xs", "color": COLOR_NEUTRAL, "margin": "md", "wrap": True})

    bubble = {
        "type": "bubble",
        "size": "mega",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": COLOR_HEADER_BG,
            "paddingAll": "20px",
            "spacing": "sm",
            "contents": [
                {"type": "text", "text": "📊 BÁO CÁO DOANH THU", "color": COLOR_WHITE, "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"📅 {weekday}, {row[ST_COL_DATE].strftime('%d/%m/%Y')}", "color": COLOR_WHITE, "size": "sm"},
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "20px",
            "spacing": "sm",
            "contents": body_contents,
        },
    }
    alt_text = f"BÁO CÁO DOANH THU {target_date.strftime('%d/%m/%Y')} - Tổng: {format_money(total)}"
    return bubble, alt_text, store_code


def push_dt(sieu_thi_path: str, nganh_hang_path):
    bubble, alt_text, store_code = build_full_dt_bubble(sieu_thi_path, nganh_hang_path)
    resp = requests.post(
        f"{CLOUD_URL.rstrip('/')}/update",
        headers={"X-Push-Secret": PUSH_SECRET},
        json={"store_code": str(store_code), "bubble": bubble, "alt_text": alt_text},
        timeout=20,
    )
    if resp.status_code == 200:
        log(f"[DT] Đã đẩy báo cáo (kèm chi tiết ngành hàng) thành công. {alt_text}")
        return True
    log(f"[DT] LỖI đẩy dữ liệu ({resp.status_code}): {resp.text}")
    return False


def build_fresh_bubble(nganh_hang_path: str, store_code: str):
    nh_df = load_nganh_hang_df(nganh_hang_path)

    mask_store = nh_df[NH_COL_STORE].astype(str).str.startswith(str(store_code))
    if not mask_store.any():
        return None
    target_date = nh_df[mask_store][NH_COL_DATE].max().date()
    last_week_date = target_date - timedelta(days=7)

    today_cats = get_category_breakdown_for_date(nh_df, store_code, target_date)
    last_week_cats = get_category_breakdown_for_date(nh_df, store_code, last_week_date)
    if today_cats is None:
        return None

    def match_fresh(series):
        if series is None:
            return {}
        result = {}
        for name, value in series.items():
            for fresh_name in FRESH_CATEGORIES:
                if fresh_name in str(name):
                    result[fresh_name] = value
                    break
        return result

    today_fresh = match_fresh(today_cats)
    last_week_fresh = match_fresh(last_week_cats)

    total_today = sum(today_fresh.values())
    total_last_week = sum(last_week_fresh.values()) if last_week_fresh else None
    change_text, change_color = change_text_and_color(total_today, total_last_week)

    weekday = WEEKDAY_VN[pd.Timestamp(target_date).weekday()]

    body_contents = [
        {"type": "text", "text": f"🏪 Mã siêu thị {store_code}", "weight": "bold", "size": "md", "wrap": True},
        {"type": "separator", "margin": "md"},
        {
            "type": "box",
            "layout": "horizontal",
            "margin": "md",
            "contents": [
                {"type": "text", "text": "🥦 TỔNG DOANH THU FRESH", "weight": "bold", "size": "md", "flex": 3, "wrap": True},
                {"type": "text", "text": format_money(total_today), "weight": "bold", "size": "lg", "color": COLOR_TOTAL, "align": "end", "flex": 2},
            ],
        },
        {"type": "text", "text": change_text, "size": "xs", "color": change_color, "align": "end"},
        {"type": "separator", "margin": "lg"},
        {"type": "text", "text": "📦 CHI TIẾT THEO NHÓM", "weight": "bold", "size": "sm", "margin": "md"},
    ]

    for fresh_name in sorted(today_fresh, key=lambda k: today_fresh[k], reverse=True):
        revenue = today_fresh[fresh_name]
        prev_value = last_week_fresh.get(fresh_name)
        c_text, c_color = change_text_and_color(revenue, prev_value)
        body_contents.append(_card_row(fresh_name, format_money(revenue), c_text, c_color))

    bubble = {
        "type": "bubble",
        "size": "mega",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": COLOR_HEADER_BG_FRESH,
            "paddingAll": "20px",
            "spacing": "sm",
            "contents": [
                {"type": "text", "text": "🥦 BÁO CÁO NGÀNH HÀNG FRESH", "color": COLOR_WHITE, "weight": "bold", "size": "lg", "wrap": True},
                {"type": "text", "text": f"📅 {weekday}, {target_date.strftime('%d/%m/%Y')}", "color": COLOR_WHITE, "size": "sm"},
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "20px",
            "spacing": "sm",
            "contents": body_contents,
        },
    }
    alt_text = f"DOANH THU NGÀNH HÀNG FRESH {target_date.strftime('%d/%m/%Y')} - {format_money(total_today)}"
    return bubble, alt_text


def push_fresh(nganh_hang_path: str):
    result = build_fresh_bubble(nganh_hang_path, STORE_CODE)
    if result is None:
        log(f"[FRESH] Không tìm thấy dữ liệu ngành hàng Fresh cho mã {STORE_CODE} trong file.")
        return False
    bubble, alt_text = result

    resp = requests.post(
        f"{CLOUD_URL.rstrip('/')}/update",
        headers={"X-Push-Secret": PUSH_SECRET},
        json={"store_code": STORE_CODE, "category": "FRESH", "bubble": bubble, "alt_text": alt_text},
        timeout=20,
    )
    if resp.status_code == 200:
        log(f"[FRESH] Đã đẩy báo cáo ngành hàng Fresh thành công. {alt_text}")
        return True
    log(f"[FRESH] LỖI đẩy dữ liệu ({resp.status_code}): {resp.text}")
    return False


def load_chi_tiet_df(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df[CT_COL_DATE] = pd.to_datetime(df[CT_COL_DATE])
    return df


def get_nam_products_for_date(df: pd.DataFrame, store_code: str, date):
    mask_store = df[CT_COL_STORE].astype(str) == str(store_code)
    mask_nam = df[CT_COL_NHOM_HANG].astype(str).str.contains("Nấm", case=False, na=False)
    mask_date = df[CT_COL_DATE].dt.date == date
    day_df = df[mask_store & mask_nam & mask_date]
    if day_df.empty:
        return None
    grouped = (
        day_df.groupby([CT_COL_PRODUCT, CT_COL_UNIT])[[CT_COL_QTY, CT_COL_REVENUE]]
        .sum()
        .reset_index()
    )
    return grouped


def build_nam_bubble(chi_tiet_path: str, store_code: str):
    df = load_chi_tiet_df(chi_tiet_path)

    mask_store = df[CT_COL_STORE].astype(str) == str(store_code)
    mask_nam = df[CT_COL_NHOM_HANG].astype(str).str.contains("Nấm", case=False, na=False)
    filtered = df[mask_store & mask_nam]
    if filtered.empty:
        return None

    target_date = filtered[CT_COL_DATE].max().date()
    today_products = get_nam_products_for_date(df, store_code, target_date)
    last_week_date = target_date - timedelta(days=7)
    last_week_products = get_nam_products_for_date(df, store_code, last_week_date)

    total_revenue = today_products[CT_COL_REVENUE].sum()
    last_week_total = last_week_products[CT_COL_REVENUE].sum() if last_week_products is not None else None

    weekday = WEEKDAY_VN[pd.Timestamp(target_date).weekday()]
    change_text, change_color = change_text_and_color(total_revenue, last_week_total)

    day_rows = filtered[filtered[CT_COL_DATE].dt.date == target_date]
    store_label = f"{store_code} - {day_rows[CT_COL_STORE_NAME].iloc[0]}"

    last_week_lookup = {}
    if last_week_products is not None:
        for _, r in last_week_products.iterrows():
            last_week_lookup[r[CT_COL_PRODUCT]] = r[CT_COL_REVENUE]

    body_contents = [
        {"type": "text", "text": f"🏪 {store_label}", "weight": "bold", "size": "md", "wrap": True},
        {"type": "separator", "margin": "md"},
        {
            "type": "box",
            "layout": "horizontal",
            "margin": "md",
            "contents": [
                {"type": "text", "text": "🍄 TỔNG DOANH THU NẤM", "weight": "bold", "size": "md", "flex": 3, "wrap": True},
                {"type": "text", "text": format_money(total_revenue), "weight": "bold", "size": "lg", "color": COLOR_TOTAL, "align": "end", "flex": 2},
            ],
        },
        {"type": "text", "text": change_text, "size": "xs", "color": change_color, "align": "end"},
        _target_box(total_revenue, MUC_TIEU_NGAY_NAM),
        {"type": "separator", "margin": "lg"},
        {"type": "text", "text": "🍄 CHI TIẾT SẢN PHẨM", "weight": "bold", "size": "sm", "margin": "md"},
    ]

    for _, r in today_products.sort_values(CT_COL_REVENUE, ascending=False).iterrows():
        product_name = r[CT_COL_PRODUCT]
        unit = r[CT_COL_UNIT]
        qty = r[CT_COL_QTY]
        revenue = r[CT_COL_REVENUE]
        prev_value = last_week_lookup.get(product_name)
        p_text, p_color = change_text_and_color(revenue, prev_value)
        label = short_product_name(product_name)
        value = f"{qty:.0f} {unit} - {format_money(revenue)}"
        body_contents.append(_card_row(label, value, p_text, p_color))

    bubble = {
        "type": "bubble",
        "size": "mega",
        "header": {
            "type": "box",
            "layout": "vertical",
            "backgroundColor": COLOR_HEADER_BG_NAM,
            "paddingAll": "20px",
            "spacing": "sm",
            "contents": [
                {"type": "text", "text": "🍄 BÁO CÁO NGÀNH HÀNG NẤM", "color": COLOR_WHITE, "weight": "bold", "size": "lg", "wrap": True},
                {"type": "text", "text": f"📅 {weekday}, {target_date.strftime('%d/%m/%Y')}", "color": COLOR_WHITE, "size": "sm"},
            ],
        },
        "body": {
            "type": "box",
            "layout": "vertical",
            "paddingAll": "20px",
            "spacing": "sm",
            "contents": body_contents,
        },
    }
    alt_text = f"DOANH THU NGÀNH HÀNG NẤM {target_date.strftime('%d/%m/%Y')} - {format_money(total_revenue)}"
    return bubble, alt_text


def push_nam(chi_tiet_path: str):
    result = build_nam_bubble(chi_tiet_path, STORE_CODE)
    if result is None:
        log(f"[NAM] Không tìm thấy dữ liệu ngành hàng Nấm cho mã {STORE_CODE} trong file.")
        return False
    bubble, alt_text = result

    resp = requests.post(
        f"{CLOUD_URL.rstrip('/')}/update",
        headers={"X-Push-Secret": PUSH_SECRET},
        json={"store_code": STORE_CODE, "category": "NAM", "bubble": bubble, "alt_text": alt_text},
        timeout=20,
    )
    if resp.status_code == 200:
        log(f"[NAM] Đã đẩy báo cáo chi tiết sản phẩm Nấm thành công. {alt_text}")
        return True
    log(f"[NAM] LỖI đẩy dữ liệu ({resp.status_code}): {resp.text}")
    return False


def main():
    if not CLOUD_URL or not PUSH_SECRET:
        log("!!! Thiếu CLOUD_URL hoặc PUSH_SECRET. Set biến môi trường trước khi chạy. !!!")
        return

    log(f"Đang theo dõi thư mục hiện tại, kiểm tra mỗi {CHECK_INTERVAL_SECONDS} giây... (Ctrl+C để dừng)")

    last_mtimes = {"sieu_thi": None, "nganh_hang": None, "chi_tiet": None}

    while True:
        try:
            sieu_thi_path = find_latest_file(FILE_PATTERN_SIEU_THI)
            nganh_hang_path = find_latest_file(FILE_PATTERN_NGANH_HANG)
            chi_tiet_path = find_latest_file(FILE_PATTERN_CHI_TIET)

            st_mtime = os.path.getmtime(sieu_thi_path) if sieu_thi_path else None
            nh_mtime = os.path.getmtime(nganh_hang_path) if nganh_hang_path else None
            ct_mtime = os.path.getmtime(chi_tiet_path) if chi_tiet_path else None

            dt_changed = (st_mtime != last_mtimes["sieu_thi"]) or (nh_mtime != last_mtimes["nganh_hang"])
            nam_changed = ct_mtime != last_mtimes["chi_tiet"]

            if dt_changed and sieu_thi_path:
                log("Phát hiện thay đổi dữ liệu doanh thu - đang tính lại báo cáo DT...")
                if push_dt(sieu_thi_path, nganh_hang_path):
                    last_mtimes["sieu_thi"] = st_mtime
                    last_mtimes["nganh_hang"] = nh_mtime
                    if nganh_hang_path:
                        push_fresh(nganh_hang_path)

            if nam_changed and chi_tiet_path:
                log("Phát hiện thay đổi dữ liệu chi tiết sản phẩm - đang tính lại báo cáo NAM...")
                if push_nam(chi_tiet_path):
                    last_mtimes["chi_tiet"] = ct_mtime

            time.sleep(CHECK_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            log("Đã dừng theo dõi.")
            break
        except Exception as e:
            log(f"Lỗi không mong muốn: {e} - thử lại sau {CHECK_INTERVAL_SECONDS} giây")
            time.sleep(CHECK_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
