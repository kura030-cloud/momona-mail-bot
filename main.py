import imaplib
import email
from email.header import decode_header
from datetime import datetime, timedelta
import os
import requests
import json
from bs4 import BeautifulSoup
from openai import OpenAI

# ====== 读取环境变量 ======
EMAIL_USER = os.environ["EMAIL_USER"]
EMAIL_PASS = os.environ["EMAIL_PASS"]
OPENAI_KEY = os.environ["OPENAI_KEY"]
FEISHU_WEBHOOK = os.environ["FEISHU_WEBHOOK"]
SENDER_EMAIL = os.environ["SENDER_EMAIL"]

# ====== 连接邮箱 (Hotmail/Outlook) ======
mail = imaplib.IMAP4_SSL("outlook.office365.com")
mail.login(EMAIL_USER, EMAIL_PASS)
mail.select("inbox")

yesterday = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")

status, messages = mail.search(
    None,
    f'(FROM "{SENDER_EMAIL}" SINCE "{yesterday}")'
)

if not messages[0]:
    print("没有找到邮件")
    exit()

latest_email_id = messages[0].split()[-1]
status, msg_data = mail.fetch(latest_email_id, "(RFC822)")
msg = email.message_from_bytes(msg_data[0][1])

# ====== 提取正文 ======
content = ""

if msg.is_multipart():
    for part in msg.walk():
        if part.get_content_type() == "text/html":
            html = part.get_payload(decode=True)
            soup = BeautifulSoup(html, "html.parser")
            content = soup.get_text()
else:
    content = msg.get_payload(decode=True).decode()

if not content:
    print("未找到正文")
    exit()

# ====== 翻译 ======
client = OpenAI(api_key=OPENAI_KEY)

response = client.chat.completions.create(
    model="gpt-4o-mini",
    messages=[
        {"role": "system", "content": "把下面的日文翻译成自然流畅的中文"},
        {"role": "user", "content": content}
    ]
)

translated = response.choices[0].message.content

# ====== 发送到飞书 ======
data = {
    "msg_type": "text",
    "content": {
        "text": translated
    }
}

requests.post(FEISHU_WEBHOOK, data=json.dumps(data))

print("已发送到飞书")
