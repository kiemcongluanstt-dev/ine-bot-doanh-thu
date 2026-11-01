"""
SERVER CHẠY TRÊN CLOUD (deploy lên Render.com)
================================================
- Nhận webhook từ LINE, khi có ai gõ "DT <mã siêu thị>" trong group -> trả lời
  báo cáo doanh thu MỚI NHẤT đã được máy tính đẩy lên (qua endpoint /update).
- KHÔNG tự đọc file Excel (server ở xa, không có file đó) - chỉ lưu lại số liệu
  mà máy tính gửi lên qua push_report.py.

BIẾN MÔI TRƯỜNG CẦN SET TRÊN RENDER (mục Environment):
    LINE_CHANNEL_ACCESS_TOKEN
    LINE_CHANNEL_SECRET
    PUSH_SECRET          (tự đặt 1 chuỗi bí mật bất kỳ, dùng để máy tính xác thực khi đẩy dữ liệu lên)
"""

import base64
import hashlib
import hmac
import json
import os

from flask import Flask, request, abort
import requests

LINE_CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN", "")
LINE_CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET", "")
PUSH_SECRET = os.environ.get("PUSH_SECRET", "")
LINE_REPLY_URL = "https://api.line.me/v2/bot/message/reply"

app = Flask(__name__)

# Lưu báo cáo mới nhất theo mã siêu thị, trong bộ nhớ (mất khi server khởi động lại -
# vì vậy mỗi khi server restart, cần đẩy lại dữ liệu 1 lần từ máy tính).
latest_reports = {}


def verify_signature(body: bytes, signature: str) -> bool:
    if not LINE_CHANNEL_SECRET:
        return False
    mac = hmac.new(LINE_CHANNEL_SECRET.encode("utf-8"), body, hashlib.sha256).digest()
    expected = base64.b64encode(mac).decode("utf-8")
    return hmac.compare_digest(expected, signature)


def _reply(reply_token: str, messages: list):
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {LINE_CHANNEL_ACCESS_TOKEN}",
    }
    payload = {"replyToken": reply_token, "messages": messages}
    resp = requests.post(LINE_REPLY_URL, headers=headers, json=payload, timeout=15)
    if resp.status_code != 200:
        print(f"[LOI GUI REPLY] {resp.status_code}: {resp.text}")


def reply_text(reply_token: str, text: str):
    _reply(reply_token, [{"type": "text", "text": text}])


def reply_flex(reply_token: str, bubble: dict, alt_text: str):
    _reply(reply_token, [{"type": "flex", "altText": alt_text, "contents": bubble}])


@app.route("/webhook", methods=["POST"])
def webhook():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data()

    if not verify_signature(body, signature):
        print("Chu ky khong hop le - tu choi request.")
        abort(400)

    data = json.loads(body.decode("utf-8"))
    for event in data.get("events", []):
        if event.get("type") != "message":
            continue
        message = event.get("message", {})
        if message.get("type") != "text":
            continue

        text = message.get("text", "").strip().upper()
        reply_token = event.get("replyToken")

        if text.startswith("DT "):
            store_code = text[3:].strip()
            report = latest_reports.get(store_code)
            if report:
                reply_flex(reply_token, report["bubble"], report["alt_text"])
            else:
                reply_text(
                    reply_token,
                    f"Chưa có dữ liệu doanh thu cho mã {store_code}. "
                    f"Hãy đợi máy tính đẩy dữ liệu lên (chạy push_report.py).",
                )

    return "OK"


@app.route("/update", methods=["POST"])
def update():
    """Máy tính gọi endpoint này để đẩy báo cáo mới nhất lên."""
    secret = request.headers.get("X-Push-Secret", "")
    if not PUSH_SECRET or secret != PUSH_SECRET:
        abort(403)

    data = request.get_json(force=True)
    store_code = str(data.get("store_code", "")).strip()
    if not store_code:
        return {"status": "error", "message": "thieu store_code"}, 400

    latest_reports[store_code] = {
        "bubble": data["bubble"],
        "alt_text": data["alt_text"],
    }
    print(f"Da nhan du lieu moi cho ma sieu thi {store_code}")
    return {"status": "ok", "store_code": store_code}


@app.route("/", methods=["GET"])
def health_check():
    return "Bot cloud dang chay OK"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
