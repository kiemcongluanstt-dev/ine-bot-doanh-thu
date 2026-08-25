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

FILE_PATTERN_SIEU_THI = ["_Doanh_Thu_Theo_Sieu_Thi*.xlsx", "Doanh_Thu_Theo_Sieu_Thi*.xlsx"]
FILE_PATTERN_NGANH_HANG = ["Doanh_Thu_Theo_Nganh_Hang*.xlsx", "_Doanh_Thu_Theo_Nganh_Hang*.xlsx"]
FILE_PATTERN_CHI_TIET = ["_Doanh_Thu_Chi_Tiet*.xlsx", "Doanh_Thu_Chi_Tiet*.xlsx"]
FILE_PATTERN_PHIEU_NHAP = ["Chi_Tiet_Phieu_Nhap*.xlsx", "_Chi_Tiet_Phieu_Nhap*.xlsx"]
FILE_PATTERN_CHENH_LECH = ["Bao_Cao_Chenh_Lech*.xlsx", "_Bao_Cao_Chenh_Lech*.xlsx"]
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
NH_COL_NHOM_HANG = "Nhóm hàng"
NH_COL_REVENUE = "Doanh thu"

# Phân loại ngành hàng -> nhóm KB (theo sheet "KHAI BÁO" trong file gốc)
CATEGORY_GROUP_MAP = {
    "Sản Phẩm Từ Sữa - Bảo Quản Mát": "ĐÔNG MÁT",
    "Kem các loại": "ĐÔNG MÁT",
    "Thực phẩm đông lạnh - Hàng mát các loại": "ĐÔNG MÁT",
    "Hóa phẩm các loại": "FMCG",
    "Dụng cụ nhà bếp": "FMCG",
    "Bánh kẹo - Trà - Cà phê - Bột Dinh Dưỡng các loại": "FMCG",
    "Thực phẩm - Gia vị các loại": "FMCG",
    "Sữa - Thức uống bổ dưỡng các loại": "FMCG",
    "Thức uống giải khát các loại": "FMCG",
    "Bia Các Loại": "FMCG",
    "Làm Đẹp": "FMCG",
    "Mỹ phẩm các loại": "FMCG",
    "Dụng cụ gia đình": "FMCG",
    "Rau Củ Các Loại": "FRESH",
    "Trái Cây Các Loại": "FRESH",
    "Thịt gia cầm gia súc các loại": "FRESH",
    "Thủy Hải Sản Các Loại": "FRESH",
    "BHX - Hàng khuyến mãi": "FMCG",
}


def classify_category(nganh_hang_name: str) -> str:
    for key, group in CATEGORY_GROUP_MAP.items():
        if key in str(nganh_hang_name):
            return group
    return "FMCG"


def compute_group_totals(cats_series):
    totals = {"FMCG": 0, "FRESH": 0, "ĐÔNG MÁT": 0}
    if cats_series is None:
        return totals
    for name, val in cats_series.items():
        grp = classify_category(name)
        totals[grp] = totals.get(grp, 0) + val
    return totals

CT_COL_DATE = "Ngày xuất"
CT_COL_STORE = "Mã siêu thị"
CT_COL_STORE_NAME = "Tên siêu thị"
CT_COL_PRODUCT = "Tên sản phẩm"
CT_COL_NHOM_HANG = "Nhóm hàng"
CT_COL_UNIT = "Đơn vị"
CT_COL_QTY = "Tổng SL bán"
CT_COL_REVENUE = "Thành tiền phải thu khách hàng (chưa VAT)"

PN_COL_DATE = "Ngày nhập"
PN_COL_STORE = "Mã siêu thị"
PN_COL_NGANH_HANG = "Ngành hàng"
PN_COL_QTY = "Số lượng"
PN_COL_GIA_NHAP = "Giá nhập"

CL_COL_DATE = "Ngày tạo"
CL_COL_STORE = "Mã siêu thị"
CL_COL_NGANH_HANG = "Ngành hàng"
CL_COL_VALUE = "Giá trị chênh lệch"
CL_COL_GHI_CHU = "Ghi chú"
CL_GHI_CHU_HUY = "Clear tồn hằng ngày"
CL_GHI_CHU_MAT_MAT = "Clear tồn 14 ngày k nhập/k xuất"


def log(msg: str):
    print(f"[{time.strftime('%H:%M:%S')}] {msg}")


def format_money(value: float) -> str:
    return f"{round(value):,}".replace(",", ".") + "đ"


def format_money_short(value: float) -> str:
    if abs(value) >= 1_000_000:
        return f"{value / 1_000_000:.1f}tr"
    if abs(value) >= 1_000:
        return f"{value / 1_000:.0f}k"
    return f"{value:.0f}đ"


def find_latest_file(patterns):
    if isinstance(patterns, str):
        patterns = [patterns]
    files = []
    for p in patterns:
        files.extend(glob.glob(p))
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


def _grid_cell(text: str, header: bool = False) -> dict:
    return {
        "type": "box",
        "layout": "vertical",
        "flex": 1,
        "borderWidth": "1px",
        "borderColor": "#CCCCCC",
        "backgroundColor": COLOR_HEADER_BG if header else "#FFFFFF",
        "paddingAll": "6px",
        "contents": [
            {
                "type": "text",
                "text": text,
                "size": "xxs",
                "align": "center",
                "wrap": True,
                "weight": "bold",
                "color": COLOR_WHITE if header else "#111111",
            }
        ],
    }


def _grid_row(cells: list, header: bool = False) -> dict:
    return {
        "type": "box",
        "layout": "horizontal",
        "spacing": "none",
        "contents": [_grid_cell(c, header) for c in cells],
    }


def _summary_grid_table(fmcg: float, fresh: float, dong_mat: float, total: float, bill_count, bill_value: float) -> dict:
    return {
        "type": "box",
        "layout": "vertical",
        "margin": "md",
        "contents": [
            _grid_row(["FMCG", "FRESH", "ĐÔNG MÁT"], header=True),
            _grid_row([format_money(fmcg), format_money(fresh), format_money(dong_mat)]),
            _grid_row(["TỔNG", "LƯỢT BILL", "GIÁ TRỊ BILL"], header=True),
            _grid_row([format_money(total), str(int(bill_count)), format_money(bill_value)]),
        ],
    }


def _row_box(label: str, value: str) -> dict:
    return {
        "type": "box",
        "layout": "horizontal",
        "contents": [
            {"type": "text", "text": label, "size": "sm", "color": COLOR_LABEL, "flex": 3},
            {"type": "text", "text": value, "size": "sm", "color": "#111111", "align": "end", "flex": 2, "weight": "bold"},
        ],
    }


def _progress_box(current: float, compare_value: float, title: str) -> dict:
    pct = (current / compare_value * 100) if compare_value else 0
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
            {"type": "text", "text": title, "weight": "bold", "size": "sm", "color": color},
            {
                "type": "box",
                "layout": "horizontal",
                "margin": "xs",
                "contents": [
                    {"type": "text", "text": f"{format_money(current)} / {format_money(compare_value)}", "size": "sm", "color": "#333333", "flex": 3, "wrap": True},
                    {"type": "text", "text": f"{pct:.0f}%", "size": "lg", "weight": "bold", "color": color, "align": "end", "flex": 1},
                ],
            },
        ],
    }


def _week_compare_box(current: float, prev) -> dict:
    if prev is None:
        return {"type": "text", "text": "📊 Chưa có dữ liệu tuần trước để so sánh.", "size": "xs", "color": COLOR_NEUTRAL, "margin": "md"}
    if prev == 0:
        return {"type": "text", "text": "📊 Doanh thu tuần trước = 0đ, không thể tính tỷ lệ.", "size": "xs", "color": COLOR_NEUTRAL, "margin": "md"}
    return _progress_box(current, prev, "📊 SO VỚI TUẦN TRƯỚC")


def _card_row(label: str, value: str, prev_text, change_text: str, change_color: str) -> dict:
    second_line = []
    if prev_text:
        second_line = [
            {"type": "text", "text": f"Tuần trước: {prev_text}", "size": "xxs", "color": COLOR_NEUTRAL, "flex": 3, "wrap": True},
            {"type": "text", "text": change_text, "size": "xxs", "color": change_color, "align": "end", "flex": 2},
        ]
    else:
        second_line = [
            {"type": "text", "text": change_text, "size": "xxs", "color": change_color, "align": "end", "flex": 1},
        ]

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
            {"type": "box", "layout": "horizontal", "margin": "xs", "contents": second_line},
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


def build_full_dt_bubble(sieu_thi_path: str, nganh_hang_path, target_date=None):
    st_df = load_sieu_thi_df(sieu_thi_path)
    if target_date is None:
        target_date = st_df[ST_COL_DATE].max().date()
    row = get_sieu_thi_row_for_date(st_df, target_date)
    if row is None:
        return None
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

    today_cats = None
    last_week_cats = None
    if nganh_hang_path:
        nh_df = load_nganh_hang_df(nganh_hang_path)
        today_cats = get_category_breakdown_for_date(nh_df, store_code, target_date)
        last_week_cats = get_category_breakdown_for_date(nh_df, store_code, last_week_date)

    group_totals = compute_group_totals(today_cats)

    body_contents = [
        {"type": "text", "text": f"🏪 {store_code} - {store_name}", "weight": "bold", "size": "md", "wrap": True},
        {"type": "separator", "margin": "md"},
        _summary_grid_table(group_totals["FMCG"], group_totals["FRESH"], group_totals["ĐÔNG MÁT"], total, bill_count, bill_value),
        _week_compare_box(total, last_week_total),
    ]

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


HISTORY_DAYS = 8  # so ngay gan day duoc luu lai de co the tra cuu (vd hoi lai hom qua)


def _post_update(store_code: str, category, bubble: dict, alt_text: str, date_str=None):
    payload = {"store_code": store_code, "bubble": bubble, "alt_text": alt_text}
    if category:
        payload["category"] = category
    if date_str:
        payload["date"] = date_str
    return requests.post(
        f"{CLOUD_URL.rstrip('/')}/update",
        headers={"X-Push-Secret": PUSH_SECRET},
        json=payload,
        timeout=20,
    )


def push_dt(sieu_thi_path: str, nganh_hang_path):
    st_df = load_sieu_thi_df(sieu_thi_path)
    all_dates = sorted(st_df[ST_COL_DATE].dt.date.unique())
    if not all_dates:
        return False
    latest_date = all_dates[-1]
    recent_dates = [d for d in all_dates if (latest_date - d).days < HISTORY_DAYS]

    any_ok = False
    for d in recent_dates:
        result = build_full_dt_bubble(sieu_thi_path, nganh_hang_path, target_date=d)
        if result is None:
            continue
        bubble, alt_text, store_code = result
        date_str = d.strftime("%Y-%m-%d")
        resp = _post_update(str(store_code), None, bubble, alt_text, date_str=date_str)
        if resp.status_code == 200:
            any_ok = True
        else:
            log(f"[DT] LỖI đẩy ngày {d.strftime('%d/%m')} ({resp.status_code}): {resp.text}")

    result = build_full_dt_bubble(sieu_thi_path, nganh_hang_path, target_date=latest_date)
    if result is None:
        return any_ok
    bubble, alt_text, store_code = result
    resp = _post_update(str(store_code), None, bubble, alt_text, date_str=None)
    if resp.status_code == 200:
        log(f"[DT] Đã đẩy báo cáo (kèm chi tiết ngành hàng, {len(recent_dates)} ngày gần nhất) thành công. {alt_text}")
        return True
    log(f"[DT] LỖI đẩy dữ liệu ({resp.status_code}): {resp.text}")
    return any_ok


def get_nhom_hang_breakdown_for_date(df: pd.DataFrame, store_code: str, ngan_hang_filter: str, date):
    mask_store = df[NH_COL_STORE].astype(str).str.startswith(str(store_code))
    mask_cat = df[NH_COL_NGANH_HANG].astype(str).str.contains(ngan_hang_filter, case=False, na=False)
    mask_date = df[NH_COL_DATE].dt.date == date
    day_df = df[mask_store & mask_cat & mask_date]
    if day_df.empty:
        return None
    return day_df.groupby(NH_COL_NHOM_HANG)[NH_COL_REVENUE].sum()


def _meat_grid_table(today_meat: dict, last_week_meat: dict) -> dict:
    labels = ["Thịt Heo", "Thịt Gia Cầm", "Thịt Bò"]
    keys = ["Thịt Heo Các Loại", "Thịt Gia Cầm Các Loại", "Thịt Bò Các Loại"]

    today_row = [format_money(today_meat.get(k, 0)) for k in keys]
    last_week_row = []
    for k in keys:
        prev = last_week_meat.get(k)
        last_week_row.append(format_money(prev) if prev is not None else "—")
    pct_row = []
    for k in keys:
        cur = today_meat.get(k, 0)
        prev = last_week_meat.get(k)
        if prev is None or prev == 0:
            pct_row.append("—")
        else:
            pct = (cur - prev) / prev * 100
            pct_row.append(f"{'▲' if pct >= 0 else '▼'}{abs(pct):.0f}%")

    return {
        "type": "box",
        "layout": "vertical",
        "margin": "md",
        "contents": [
            _grid_row(labels, header=True),
            _grid_row(today_row),
            _grid_row([f"TT: {v}" for v in last_week_row]),
            _grid_row(pct_row),
        ],
    }


def match_fresh_categories(series):
    if series is None:
        return {}
    result = {}
    for name, value in series.items():
        for fresh_name in FRESH_CATEGORIES:
            if fresh_name in str(name):
                result[fresh_name] = result.get(fresh_name, 0) + value
                break
    return result


def load_phieu_nhap_df(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df[PN_COL_DATE] = pd.to_datetime(df[PN_COL_DATE])
    df["_gia_tri_nhap"] = df[PN_COL_QTY] * df[PN_COL_GIA_NHAP]
    return df


def get_import_value_by_fresh_for_date(df: pd.DataFrame, store_code: str, date):
    mask_store = df[PN_COL_STORE].astype(str) == str(store_code)
    mask_date = df[PN_COL_DATE].dt.date == date
    day_df = df[mask_store & mask_date]
    if day_df.empty:
        return None
    by_cat = day_df.groupby(PN_COL_NGANH_HANG)["_gia_tri_nhap"].sum()
    return match_fresh_categories(by_cat)


def load_chenh_lech_df(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df[CL_COL_DATE] = pd.to_datetime(df[CL_COL_DATE])
    return df


def get_loss_breakdown_for_date(df: pd.DataFrame, store_code: str, date):
    mask_store = df[CL_COL_STORE].astype(str) == str(store_code)
    mask_date = df[CL_COL_DATE].dt.date == date
    day_df = df[mask_store & mask_date]
    if day_df.empty:
        return {}

    result = {}
    for ghi_chu, key in [(CL_GHI_CHU_HUY, "huy"), (CL_GHI_CHU_MAT_MAT, "mat_mat")]:
        sub = day_df[day_df[CL_COL_GHI_CHU] == ghi_chu]
        by_cat = sub.groupby(CL_COL_NGANH_HANG)[CL_COL_VALUE].sum()
        matched = match_fresh_categories(by_cat)
        for cat, val in matched.items():
            result.setdefault(cat, {"huy": 0, "mat_mat": 0})[key] = val
    return result


def build_fresh_chitiet_bubble(nganh_hang_path: str, phieu_nhap_path, chenh_lech_path, store_code: str, target_date=None):
    nh_df = load_nganh_hang_df(nganh_hang_path)
    mask_store = nh_df[NH_COL_STORE].astype(str).str.startswith(str(store_code))
    if not mask_store.any():
        return None
    if target_date is None:
        target_date = nh_df[mask_store][NH_COL_DATE].max().date()

    today_cats = get_category_breakdown_for_date(nh_df, store_code, target_date)
    if today_cats is None:
        return None
    today_ban = match_fresh_categories(today_cats)

    today_nhap = {}
    if phieu_nhap_path:
        pn_df = load_phieu_nhap_df(phieu_nhap_path)
        today_nhap = get_import_value_by_fresh_for_date(pn_df, store_code, target_date) or {}

    today_loss = {}
    if chenh_lech_path:
        cl_df = load_chenh_lech_df(chenh_lech_path)
        today_loss = get_loss_breakdown_for_date(cl_df, store_code, target_date)

    weekday = WEEKDAY_VN[pd.Timestamp(target_date).weekday()]

    header = ["NHÓM HÀNG", "NHẬP", "BÁN", "BÁN/NHẬP"]
    rows = [header]
    total_ban = 0
    total_nhap = 0
    for cat in FRESH_CATEGORIES:
        ban = today_ban.get(cat, 0)
        nhap = today_nhap.get(cat, 0)
        total_ban += ban
        total_nhap += nhap
        ty_le = f"{(ban / nhap * 100):.0f}%" if nhap else "—"
        rows.append([short_name(cat, 14), format_money_short(nhap), format_money_short(ban), ty_le])

    total_ty_le = f"{(total_ban / total_nhap * 100):.0f}%" if total_nhap else "—"
    rows.append(["TỔNG", format_money_short(total_nhap), format_money_short(total_ban), total_ty_le])

    grid_contents = []
    for i, r in enumerate(rows):
        grid_contents.append(_grid_row(r, header=(i == 0)))

    body_contents = [
        {"type": "text", "text": f"🏪 Mã siêu thị {store_code}", "weight": "bold", "size": "md", "wrap": True},
        {"type": "separator", "margin": "md"},
        {"type": "box", "layout": "vertical", "margin": "md", "contents": grid_contents},
    ]

    if not phieu_nhap_path:
        body_contents.append({"type": "text", "text": "⚠️ Chưa có dữ liệu Nhập hàng.", "size": "xxs", "color": COLOR_NEUTRAL, "margin": "md", "wrap": True})

    body_contents.append({"type": "separator", "margin": "lg"})
    body_contents.append({"type": "text", "text": "🗑️ HỦY / MẤT MÁT", "weight": "bold", "size": "sm", "margin": "md"})

    if chenh_lech_path:
        loss_header = ["NHÓM HÀNG", "HỦY", "MẤT MÁT", "% HAO HỤT"]
        loss_rows = [loss_header]
        total_huy = 0
        total_mat_mat = 0
        for cat in FRESH_CATEGORIES:
            info = today_loss.get(cat, {"huy": 0, "mat_mat": 0})
            huy = info.get("huy", 0)
            mat_mat = info.get("mat_mat", 0)
            total_huy += huy
            total_mat_mat += mat_mat
            nhap_cat = today_nhap.get(cat, 0)
            pct = f"{((huy + mat_mat) / nhap_cat * 100):.0f}%" if nhap_cat else "—"
            loss_rows.append([short_name(cat, 14), format_money_short(huy), format_money_short(mat_mat), pct])

        total_pct = f"{((total_huy + total_mat_mat) / total_nhap * 100):.0f}%" if total_nhap else "—"
        loss_rows.append(["TỔNG", format_money_short(total_huy), format_money_short(total_mat_mat), total_pct])

        loss_grid_contents = [_grid_row(r, header=(i == 0)) for i, r in enumerate(loss_rows)]
        body_contents.append({"type": "box", "layout": "vertical", "margin": "md", "contents": loss_grid_contents})
    else:
        body_contents.append({
            "type": "text",
            "text": "⚠️ Chưa có báo cáo Hủy/Mất mát từ hệ thống.",
            "size": "xxs",
            "color": COLOR_NEUTRAL,
            "margin": "md",
            "wrap": True,
        })

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
                {"type": "text", "text": "📋 CHI TIẾT NHẬP/BÁN FRESH", "color": COLOR_WHITE, "weight": "bold", "size": "lg", "wrap": True},
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
    alt_text = f"CHI TIẾT NHẬP/BÁN FRESH {target_date.strftime('%d/%m/%Y')} - Bán: {format_money(total_ban)}"
    return bubble, alt_text


def push_fresh_chitiet(nganh_hang_path: str, phieu_nhap_path, chenh_lech_path):
    nh_df = load_nganh_hang_df(nganh_hang_path)
    mask_store = nh_df[NH_COL_STORE].astype(str).str.startswith(str(STORE_CODE))
    if not mask_store.any():
        return False
    all_dates = sorted(nh_df[mask_store][NH_COL_DATE].dt.date.unique())
    if not all_dates:
        return False
    latest_date = all_dates[-1]
    recent_dates = [d for d in all_dates if (latest_date - d).days < HISTORY_DAYS]

    any_ok = False
    for d in recent_dates:
        result = build_fresh_chitiet_bubble(nganh_hang_path, phieu_nhap_path, chenh_lech_path, STORE_CODE, target_date=d)
        if result is None:
            continue
        bubble, alt_text = result
        date_str = d.strftime("%Y-%m-%d")
        resp = _post_update(STORE_CODE, "FRESHCT", bubble, alt_text, date_str=date_str)
        if resp.status_code == 200:
            any_ok = True
        else:
            log(f"[FRESHCT] LỖI đẩy ngày {d.strftime('%d/%m')} ({resp.status_code}): {resp.text}")

    result = build_fresh_chitiet_bubble(nganh_hang_path, phieu_nhap_path, chenh_lech_path, STORE_CODE, target_date=latest_date)
    if result is None:
        return any_ok
    bubble, alt_text = result
    resp = _post_update(STORE_CODE, "FRESHCT", bubble, alt_text, date_str=None)
    if resp.status_code == 200:
        log(f"[FRESHCT] Đã đẩy báo cáo chi tiết Nhập/Bán Fresh ({len(recent_dates)} ngày gần nhất) thành công. {alt_text}")
        return True
    log(f"[FRESHCT] LỖI đẩy dữ liệu ({resp.status_code}): {resp.text}")
    return any_ok


def build_fresh_bubble(nganh_hang_path: str, store_code: str, target_date=None):
    nh_df = load_nganh_hang_df(nganh_hang_path)

    mask_store = nh_df[NH_COL_STORE].astype(str).str.startswith(str(store_code))
    if not mask_store.any():
        return None
    if target_date is None:
        target_date = nh_df[mask_store][NH_COL_DATE].max().date()
    last_week_date = target_date - timedelta(days=7)

    today_cats = get_category_breakdown_for_date(nh_df, store_code, target_date)
    last_week_cats = get_category_breakdown_for_date(nh_df, store_code, last_week_date)
    if today_cats is None:
        return None

    today_fresh = match_fresh_categories(today_cats)
    last_week_fresh = match_fresh_categories(last_week_cats)

    total_today = sum(today_fresh.values())
    total_last_week = sum(last_week_fresh.values()) if last_week_fresh else None

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
        _week_compare_box(total_today, total_last_week),
        {"type": "separator", "margin": "lg"},
        {"type": "text", "text": "📦 CHI TIẾT THEO NHÓM", "weight": "bold", "size": "sm", "margin": "md"},
    ]

    for fresh_name in sorted(today_fresh, key=lambda k: today_fresh[k], reverse=True):
        revenue = today_fresh[fresh_name]
        prev_value = last_week_fresh.get(fresh_name)
        c_text, c_color = change_text_and_color(revenue, prev_value)
        prev_text = format_money(prev_value) if prev_value is not None else None
        body_contents.append(_card_row(fresh_name, format_money(revenue), prev_text, c_text, c_color))

    today_meat = get_nhom_hang_breakdown_for_date(nh_df, store_code, "Thịt gia cầm gia súc", target_date)
    last_week_meat = get_nhom_hang_breakdown_for_date(nh_df, store_code, "Thịt gia cầm gia súc", last_week_date)
    if today_meat is not None:
        today_meat_dict = {short_name(k, 40): v for k, v in today_meat.items()}
        last_week_meat_dict = {short_name(k, 40): v for k, v in last_week_meat.items()} if last_week_meat is not None else {}
        body_contents.append({"type": "separator", "margin": "lg"})
        body_contents.append({"type": "text", "text": "🥩 CHI TIẾT THỊT", "weight": "bold", "size": "sm", "margin": "md"})
        body_contents.append(_meat_grid_table(today_meat_dict, last_week_meat_dict))

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
    nh_df = load_nganh_hang_df(nganh_hang_path)
    mask_store = nh_df[NH_COL_STORE].astype(str).str.startswith(str(STORE_CODE))
    if not mask_store.any():
        return False
    all_dates = sorted(nh_df[mask_store][NH_COL_DATE].dt.date.unique())
    if not all_dates:
        return False
    latest_date = all_dates[-1]
    recent_dates = [d for d in all_dates if (latest_date - d).days < HISTORY_DAYS]

    any_ok = False
    for d in recent_dates:
        result = build_fresh_bubble(nganh_hang_path, STORE_CODE, target_date=d)
        if result is None:
            continue
        bubble, alt_text = result
        date_str = d.strftime("%Y-%m-%d")
        resp = _post_update(STORE_CODE, "FRESH", bubble, alt_text, date_str=date_str)
        if resp.status_code == 200:
            any_ok = True
        else:
            log(f"[FRESH] LỖI đẩy ngày {d.strftime('%d/%m')} ({resp.status_code}): {resp.text}")

    result = build_fresh_bubble(nganh_hang_path, STORE_CODE, target_date=latest_date)
    if result is None:
        return any_ok
    bubble, alt_text = result
    resp = _post_update(STORE_CODE, "FRESH", bubble, alt_text, date_str=None)
    if resp.status_code == 200:
        log(f"[FRESH] Đã đẩy báo cáo ngành hàng Fresh ({len(recent_dates)} ngày gần nhất) thành công. {alt_text}")
        return True
    log(f"[FRESH] LỖI đẩy dữ liệu ({resp.status_code}): {resp.text}")
    return any_ok


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


def build_nam_bubble(chi_tiet_path: str, store_code: str, target_date=None):
    df = load_chi_tiet_df(chi_tiet_path)

    mask_store = df[CT_COL_STORE].astype(str) == str(store_code)
    mask_nam = df[CT_COL_NHOM_HANG].astype(str).str.contains("Nấm", case=False, na=False)
    filtered = df[mask_store & mask_nam]
    if filtered.empty:
        return None

    is_stale = False
    if target_date is None:
        target_date = filtered[CT_COL_DATE].max().date()
        file_max_date = df[df[CT_COL_STORE].astype(str) == str(store_code)][CT_COL_DATE].max()
        file_max_date = file_max_date.date() if pd.notna(file_max_date) else None
        is_stale = file_max_date is not None and target_date < file_max_date
    else:
        if not (filtered[CT_COL_DATE].dt.date == target_date).any():
            return None

    today_products = get_nam_products_for_date(df, store_code, target_date)
    last_week_date = target_date - timedelta(days=7)
    last_week_products = get_nam_products_for_date(df, store_code, last_week_date)

    total_revenue = today_products[CT_COL_REVENUE].sum()
    last_week_total = last_week_products[CT_COL_REVENUE].sum() if last_week_products is not None else None

    weekday = WEEKDAY_VN[pd.Timestamp(target_date).weekday()]

    day_rows = filtered[filtered[CT_COL_DATE].dt.date == target_date]
    store_label = f"{store_code} - {day_rows[CT_COL_STORE_NAME].iloc[0]}"

    last_week_lookup = {}
    if last_week_products is not None:
        for _, r in last_week_products.iterrows():
            last_week_lookup[r[CT_COL_PRODUCT]] = r[CT_COL_REVENUE]

    body_contents = [
        {"type": "text", "text": f"🏪 {store_label}", "weight": "bold", "size": "md", "wrap": True},
    ]
    if is_stale:
        body_contents.append({
            "type": "text",
            "text": f"⚠️ Chưa ghi nhận đơn Nấm nào cho ngày {file_max_date.strftime('%d/%m/%Y')} tính đến lúc xuất file. Đang hiện số liệu ngày gần nhất có bán ({target_date.strftime('%d/%m/%Y')}).",
            "size": "xxs",
            "color": COLOR_DOWN,
            "wrap": True,
            "margin": "sm",
        })
    body_contents += [
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
        _week_compare_box(total_revenue, last_week_total),
        _progress_box(total_revenue, MUC_TIEU_NGAY_NAM, "🎯 MỤC TIÊU NGÀY"),
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
        prev_text = format_money(prev_value) if prev_value is not None else None
        body_contents.append(_card_row(label, value, prev_text, p_text, p_color))

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
    df = load_chi_tiet_df(chi_tiet_path)
    mask_store = df[CT_COL_STORE].astype(str) == str(STORE_CODE)
    mask_nam = df[CT_COL_NHOM_HANG].astype(str).str.contains("Nấm", case=False, na=False)
    filtered = df[mask_store & mask_nam]
    if filtered.empty:
        log(f"[NAM] Không tìm thấy dữ liệu ngành hàng Nấm cho mã {STORE_CODE} trong file.")
        return False

    all_dates = sorted(filtered[CT_COL_DATE].dt.date.unique())
    latest_date = all_dates[-1]
    recent_dates = [d for d in all_dates if (latest_date - d).days < HISTORY_DAYS]

    any_ok = False
    for d in recent_dates:
        result = build_nam_bubble(chi_tiet_path, STORE_CODE, target_date=d)
        if result is None:
            continue
        bubble, alt_text = result
        date_str = d.strftime("%Y-%m-%d")
        resp = _post_update(STORE_CODE, "NAM", bubble, alt_text, date_str=date_str)
        if resp.status_code == 200:
            any_ok = True
        else:
            log(f"[NAM] LỖI đẩy ngày {d.strftime('%d/%m')} ({resp.status_code}): {resp.text}")

    result = build_nam_bubble(chi_tiet_path, STORE_CODE)
    if result is None:
        return any_ok
    bubble, alt_text = result
    resp = _post_update(STORE_CODE, "NAM", bubble, alt_text, date_str=None)
    if resp.status_code == 200:
        log(f"[NAM] Đã đẩy báo cáo chi tiết sản phẩm Nấm ({len(recent_dates)} ngày gần nhất) thành công. {alt_text}")
        return True
    log(f"[NAM] LỖI đẩy dữ liệu ({resp.status_code}): {resp.text}")
    return any_ok


def main():
    if not CLOUD_URL or not PUSH_SECRET:
        log("!!! Thiếu CLOUD_URL hoặc PUSH_SECRET. Set biến môi trường trước khi chạy. !!!")
        return

    log(f"Đang theo dõi thư mục hiện tại, kiểm tra mỗi {CHECK_INTERVAL_SECONDS} giây... (Ctrl+C để dừng)")

    last_mtimes = {"sieu_thi": None, "nganh_hang": None, "chi_tiet": None, "phieu_nhap": None, "chenh_lech": None}

    while True:
        try:
            sieu_thi_path = find_latest_file(FILE_PATTERN_SIEU_THI)
            nganh_hang_path = find_latest_file(FILE_PATTERN_NGANH_HANG)
            chi_tiet_path = find_latest_file(FILE_PATTERN_CHI_TIET)
            phieu_nhap_path = find_latest_file(FILE_PATTERN_PHIEU_NHAP)
            chenh_lech_path = find_latest_file(FILE_PATTERN_CHENH_LECH)

            st_mtime = os.path.getmtime(sieu_thi_path) if sieu_thi_path else None
            nh_mtime = os.path.getmtime(nganh_hang_path) if nganh_hang_path else None
            ct_mtime = os.path.getmtime(chi_tiet_path) if chi_tiet_path else None
            pn_mtime = os.path.getmtime(phieu_nhap_path) if phieu_nhap_path else None
            cl_mtime = os.path.getmtime(chenh_lech_path) if chenh_lech_path else None

            dt_changed = (st_mtime != last_mtimes["sieu_thi"]) or (nh_mtime != last_mtimes["nganh_hang"])
            nam_changed = ct_mtime != last_mtimes["chi_tiet"]
            freshct_changed = (
                (nh_mtime != last_mtimes["nganh_hang"])
                or (pn_mtime != last_mtimes["phieu_nhap"])
                or (cl_mtime != last_mtimes["chenh_lech"])
            )

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

            if freshct_changed and nganh_hang_path:
                log("Phát hiện thay đổi dữ liệu Nhập/Bán/Hủy - đang tính lại báo cáo FRESHCT...")
                if push_fresh_chitiet(nganh_hang_path, phieu_nhap_path, chenh_lech_path):
                    last_mtimes["nganh_hang"] = nh_mtime
                    last_mtimes["phieu_nhap"] = pn_mtime
                    last_mtimes["chenh_lech"] = cl_mtime

            time.sleep(CHECK_INTERVAL_SECONDS)
        except KeyboardInterrupt:
            log("Đã dừng theo dõi.")
            break
        except Exception as e:
            log(f"Lỗi không mong muốn: {e} - thử lại sau {CHECK_INTERVAL_SECONDS} giây")
            time.sleep(CHECK_INTERVAL_SECONDS)


if __name__ == "__main__":
    main()
