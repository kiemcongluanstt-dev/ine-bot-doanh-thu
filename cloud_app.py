"""
SERVER CHẠY TRÊN CLOUD (deploy lên Render.com) - PHIÊN BẢN 2
================================================================
- Lệnh "DT <mã siêu thị>"  -> trả lời báo cáo DOANH THU TOÀN SIÊU THỊ
- Lệnh "NAM <mã siêu thị>" -> trả lời báo cáo DOANH THU NGÀNH HÀNG NẤM
- Không tự đọc file Excel (server ở xa) - chỉ lưu số liệu máy tính đẩy lên qua /update.

BIẾN MÔI TRƯỜNG CẦN SET TRÊN RENDER (giữ nguyên như cũ, không đổi):
    LINE_CHANNEL_ACCESS_TOKEN
    LINE_CHANNEL_SECRET
    PUSH_SECRET
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


def _lookup_and_reply(reply_token: str, key: str, not_found_label: str):
    report = latest_reports.get(key)
    if report:
        reply_flex(reply_token, report["bubble"], report["alt_text"])
    else:
        reply_text(
            reply_token,
            f"Chưa có dữ liệu {not_found_label}. Hãy đợi máy tính đẩy dữ liệu lên.",
        )


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
        source = event.get("source", {})
        print(f"[TIN NHAN] '{text}' tu source: {source}")

        if text.startswith("NAM "):
            store_code = text[4:].strip()
            _lookup_and_reply(reply_token, f"NAM_{store_code}", f"ngành hàng Nấm cho mã {store_code}")
        elif text.startswith("DT "):
            store_code = text[3:].strip()
            _lookup_and_reply(reply_token, store_code, f"doanh thu cho mã {store_code}")

    return "OK"


@app.route("/update", methods=["POST"])
def update():
    """Máy tính gọi endpoint này để đẩy báo cáo mới nhất lên."""
    secret = request.headers.get("X-Push-Secret", "")
    if not PUSH_SECRET or secret != PUSH_SECRET:
        abort(403)

    data = request.get_json(force=True)
    store_code = str(data.get("store_code", "")).strip()
    category = str(data.get("category", "")).strip().upper()
    if not store_code:
        return {"status": "error", "message": "thieu store_code"}, 400

    key = f"{category}_{store_code}" if category else store_code
    latest_reports[key] = {
        "bubble": data["bubble"],
        "alt_text": data["alt_text"],
    }
    print(f"Da nhan du lieu moi cho key {key}")
    return {"status": "ok", "key": key}


@app.route("/", methods=["GET"])
def health_check():
    return "Bot cloud dang chay OK"


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
