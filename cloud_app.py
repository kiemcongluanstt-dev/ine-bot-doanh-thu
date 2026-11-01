"""
Theo dõi thư mục: khi có file Excel "DoanhThuTheoNgànhHàng*.xlsx" MỚI xuất hiện
(hoặc được ghi đè bằng bản mới) -> tự gửi báo cáo ngành hàng vào LINE.

CÁCH DÙNG (chạy trong thư mục chứa file Excel, cùng thư mục với send_category_report.py):
    python watch_category.py
    python watch_category.py --dry-run          # chỉ in, không gửi LINE (để test)
    python watch_category.py --send-existing    # gửi luôn file đang có sẵn khi khởi động

Cần set LINE_CHANNEL_ACCESS_TOKEN và LINE_GROUP_ID trước khi chạy (file .bat đã set sẵn).
Dừng bằng Ctrl+C.
"""

import argparse
import glob
import os
import time

import send_category_report as rpt

FILE_PATTERN = "DoanhThuTheoNgànhHàng*.xlsx"
POLL_SECONDS = 10


def file_key(path: str):
    st = os.stat(path)
    return (os.path.basename(path), st.st_mtime_ns, st.st_size)


def is_stable(path: str, wait: float = 2.0) -> bool:
    """File đã tải/ghi xong chưa (kích thước không đổi và mở đọc được)."""
    try:
        size1 = os.path.getsize(path)
        time.sleep(wait)
        if size1 != os.path.getsize(path) or size1 == 0:
            return False
        with open(path, "rb"):
            return True
    except OSError:
        return False


def process(path: str, top_n: int, dry_run: bool) -> None:
    day = rpt.load_day(path, None)
    results = [rpt.build_bubble(g, top_n) for _, g in day.groupby(rpt.COL_STORE)]
    bubbles = [b for b, _ in results][:12]
    date_txt = day[rpt.COL_DATE].iloc[0].strftime("%d/%m/%Y")
    alt = results[0][1] if len(results) == 1 else f"DOANH THU NGÀNH HÀNG {date_txt} ({len(bubbles)} siêu thị)"
    contents = bubbles[0] if len(bubbles) == 1 else {"type": "carousel", "contents": bubbles}
    total = day[rpt.COL_REVENUE].sum()
    print(f"   Ngày {date_txt} | {len(bubbles)} siêu thị | Tổng: {rpt.format_money(total)}")
    if dry_run:
        print("   (dry-run: chưa gửi LINE)")
    else:
        rpt.send_flex_to_line(contents, alt)


def main():
    ap = argparse.ArgumentParser(description="Tự gửi báo cáo ngành hàng khi có file Excel mới")
    ap.add_argument("--dry-run", action="store_true", help="Chỉ in, không gửi LINE")
    ap.add_argument("--send-existing", action="store_true", help="Gửi cả file đã có sẵn lúc khởi động")
    ap.add_argument("--top", type=int, default=5, help="Số ngành hàng hiển thị (mặc định 5)")
    args = ap.parse_args()

    seen = set()
    if not args.send_existing:
        seen = {file_key(p) for p in glob.glob(FILE_PATTERN)}
    print(f"Đang theo dõi '{FILE_PATTERN}' trong {os.getcwd()} (kiểm tra mỗi {POLL_SECONDS}s). Ctrl+C để dừng.")
    if seen:
        print(f"Bỏ qua {len(seen)} file đã có sẵn (dùng --send-existing nếu muốn gửi luôn).")

    while True:
        try:
            for path in sorted(glob.glob(FILE_PATTERN), key=os.path.getmtime):
                key = file_key(path)
                if key in seen or not is_stable(path):
                    continue
                print(f"[{time.strftime('%H:%M:%S')}] Phát hiện file mới: {path}")
                try:
                    process(path, max(1, args.top), args.dry_run)
                    seen.add(file_key(path))  # chỉ đánh dấu đã xử lý khi thành công
                except Exception as e:  # lỗi -> thử lại ở vòng sau
                    print(f"   !!! Lỗi: {e}  (sẽ thử lại sau {POLL_SECONDS}s)")
        except FileNotFoundError:
            pass
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nĐã dừng.")
