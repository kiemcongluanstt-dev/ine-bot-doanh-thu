"""
BOT LINE TRẢ LỜI DOANH THU — chạy trên Render (gunicorn cloud_app:app)
=====================================================================
Chức năng:
  1) POST /update    : máy tính ở cửa hàng đẩy báo cáo lên (push_report.py / watch_and_push.py).
                       Bắt buộc có header  X-Push-Secret  khớp biến môi trường PUSH_SECRET.
  2) POST /callback  : LINE gọi vào khi có tin nhắn trong group/chat (webhook).
                       Gõ  "DT 30400"  -> bot trả thẻ báo cáo của siêu thị mã 30400.
                       Gõ  "DT"  hoặc  "DTST"  -> bot trả thẻ của tất cả siêu thị đang có.
  3) GET  /          : kiểm tra bot còn sống (health check).

Biến môi trường cần đặt trên Render (mục Environment):
  LINE_CHANNEL_ACCESS_TOKEN   (bắt buộc)  token của kênh Messaging API
  PUSH_SECRET                 (bắt buộc)  chuỗi bí mật, phải giống PUSH_SECRET ở máy tính
  LINE_CHANNEL_SECRET         (khuyến nghị) để xác thực tin nhắn thật sự do LINE gửi

Start Command trên Render:  gunicorn cloud_app:app
"""

import base64
import hashlib
import hmac
import json
import os
import re
import tempfile
import threading
import time

import requests
from flask import Flask, abort, jsonify, request

app = Flask(__name__)

LINE_REPLY_URL = "https://api.line.me/v2/bot/message/reply"
STORE_FILE = os.path.join(tempfile.gettempdir(), "dt_reports.json")
_lock = threading.Lock()

# "DT 30400", "dt30400", "DT" (một mình), "DTST"
CMD_RE = re.compile(r"^\s*(?:DT|DTST)\s*(\d+)?\s*$", re.IGNORECASE)


def env(name: str) -> str:
    return os.environ.get(name, "").strip()


# ---------------- Lưu / đọc báo cáo (đọc file mỗi lần -> an toàn nếu có nhiều worker) ----------------
def load_reports() -> dict:
    try:
        with open(STORE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def save_reports(data: dict) -> None:
    tmp = STORE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False)
    os.replace(tmp, STORE_FILE)


# ---------------- 1) Nhận báo cáo từ máy tính ----------------
@app.route("/update", methods=["POST"])
def update():
    secret = env("PUSH_SECRET")
    if not secret:
        return jsonify(error="Server chưa cấu hình PUSH_SECRET"), 500
    if not hmac.compare_digest(request.headers.get("X-Push-Secret", ""), secret):
        return jsonify(error="Sai X-Push-Secret"), 401

    body = request.get_json(silent=True) or {}
    store_code = str(body.get("store_code", "")).strip()
    bubble = body.get("bubble")
    alt_text = str(body.get("alt_text", "BÁO CÁO DOANH THU"))[:400]
    if not store_code or not isinstance(bubble, dict):
        return jsonify(error="Thiếu store_code hoặc bubble"), 400

    with _lock:
        data = load_reports()
        data[store_code] = {"bubble": bubble, "alt_text": alt_text, "updated_at": time.time()}
        save_reports(data)
    print(f"[update] Đã lưu báo cáo mã {store_code}: {alt_text}", flush=True)
    return jsonify(ok=True, store_code=store_code, stores=sorted(data))


# ---------------- 2) Webhook LINE ----------------
def verify_signature(raw_body: bytes) -> bool:
    channel_secret = env("LINE_CHANNEL_SECRET")
    if not channel_secret:
        print("[cảnh báo] Chưa đặt LINE_CHANNEL_SECRET -> bỏ qua xác thực chữ ký", flush=True)
        return True
    digest = hmac.new(channel_secret.encode("utf-8"), raw_body, hashlib.sha256).digest()
    expected = base64.b64encode(digest).decode("utf-8")
    return hmac.compare_digest(expected, request.headers.get("X-Line-Signature", ""))


def reply(reply_token: str, messages: list) -> None:
    token = env("LINE_CHANNEL_ACCESS_TOKEN")
    if not token:
        print("[lỗi] Thiếu LINE_CHANNEL_ACCESS_TOKEN", flush=True)
        return
    resp = requests.post(
        LINE_REPLY_URL,
        headers={"Content-Type": "application/json", "Authorization": f"Bearer {token}"},
        json={"replyToken": reply_token, "messages": messages},
        timeout=10,
    )
    if resp.status_code != 200:
        print(f"[lỗi] Reply LINE {resp.status_code}: {resp.text}", flush=True)


def text_msg(text: str) -> dict:
    return {"type": "text", "text": text}


def build_reply(store_code: str | None) -> list:
    data = load_reports()
    if not data:
        return [text_msg("Chưa có dữ liệu. Hãy chạy đẩy báo cáo từ máy tính trước.")]

    if store_code:
        item = data.get(store_code)
        if not item:
            codes = ", ".join(sorted(data))
            return [text_msg(f"Không có báo cáo cho mã {store_code}. Các mã hiện có: {codes}")]
        return [{"type": "flex", "altText": item["alt_text"], "contents": item["bubble"]}]

    items = [data[k] for k in sorted(data)][:12]
    if len(items) == 1:
        return [{"type": "flex", "altText": items[0]["alt_text"], "contents": items[0]["bubble"]}]
    return [{
        "type": "flex",
        "altText": f"BÁO CÁO DOANH THU ({len(items)} siêu thị)",
        "contents": {"type": "carousel", "contents": [i["bubble"] for i in items]},
    }]


@app.route("/callback", methods=["POST"])
def callback():
    raw = request.get_data()
    if not verify_signature(raw):
        abort(400, "Sai chữ ký LINE")

    payload = request.get_json(silent=True) or {}
    for event in payload.get("events", []):
        source = event.get("source", {})
        # In ra log để anh lấy được LINE_GROUP_ID (dạng C....) khi có người nhắn trong group
        if source.get("groupId"):
            print(f"[event] groupId = {source['groupId']}", flush=True)

        if event.get("type") != "message" or event.get("message", {}).get("type") != "text":
            continue
        m = CMD_RE.match(event["message"].get("text", ""))
        if not m:
            continue  # tin nhắn thường -> bot im lặng
        try:
            reply(event["replyToken"], build_reply(m.group(1)))
        except Exception as e:
            print(f"[lỗi] Xử lý lệnh thất bại: {e}", flush=True)
    return "OK"  # LINE cần nhận 200 nhanh


# ---------------- 3) Health check ----------------
@app.route("/", methods=["GET"])
def health():
    return jsonify(status="ok", stores=sorted(load_reports()))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
