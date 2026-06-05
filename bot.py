from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import openpyxl
import os
import json
import re
import requests
import pytesseract
from PIL import Image
from io import BytesIO
from datetime import datetime

app = Flask(__name__)
EXCEL_FILE = "orders.xlsx"
CONTACTS_FILE = "contacts.json"

pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def setup_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Orders"
        ws.append(["Order No", "Date", "Time", "Sender Name", "Phone", "Item", "Quantity"])
        wb.save(EXCEL_FILE)

def load_contacts():
    try:
        if os.path.exists(CONTACTS_FILE):
            with open(CONTACTS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except:
        pass
    return {}

def get_name(phone):
    contacts = load_contacts()
    phone_clean = phone.strip().replace(" ", "")
    for key in contacts:
        if key.strip().replace(" ", "") == phone_clean:
            return contacts[key]
    return phone

def log_order(sender_name, phone, item, quantity):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    order_no = ws.max_row
    date = datetime.now().strftime("%d-%m-%Y")
    time_now = datetime.now().strftime("%H:%M:%S")
    ws.append([order_no, date, time_now, sender_name, phone, item, quantity])
    wb.save(EXCEL_FILE)
    print(f"✅ New Order! | {sender_name} | {item} | Qty: {quantity}")

def parse_line(line):
    line = line.strip()
    if not line:
        return None

    line_lower = line.lower()

    # Format: "item ke liye order: qty"
    if "order:" in line_lower:
        parts = line_lower.split("order:")
        item = parts[0].replace("ke liye", "").strip()
        quantity = parts[1].strip().split()[0]
        if item and quantity:
            return (item, quantity)

    # Format: "item order qty"
    if "order" in line_lower:
        parts = line_lower.replace("order", "").strip()
        tokens = parts.split()
        if len(tokens) >= 2:
            if tokens[0].isdigit():
                return (" ".join(tokens[1:]), tokens[0])
            else:
                return (" ".join(tokens[:-1]), tokens[-1])

    # Format: "item: qty" or "item = qty" or "item - qty"
    match = re.match(r'^([a-zA-Z\u0900-\u097F ]+)[:\-=]+\s*(\d+)', line)
    if match:
        return (match.group(1).strip(), match.group(2).strip())

    # Format: "qty item" (number pehle, no space) like "20tamatar"
    match = re.match(r'^(\d+)\s*([a-zA-Z\u0900-\u097F ]+)', line)
    if match:
        return (match.group(2).strip(),  match.group(1).strip())

    # Format: "item qty" or "itemqty" (number last, no space) like "Bengan10"
    match = re.match(r'^([a-zA-Z\u0900-\u097F ]+?)\s*(\d+)$', line)
    if match:
        return (match.group(1).strip(), match.group(2).strip())

    return None

def parse_order(message):
    orders = []
    lines = message.strip().split("\n")
    for line in lines:
        result = parse_line(line)
        if result:
            orders.append(result)
    return orders

def extract_text_from_image(image_url, account_sid, auth_token):
    try:
        response = requests.get(image_url, auth=(account_sid, auth_token))
        img = Image.open(BytesIO(response.content))
        text = pytesseract.image_to_string(img)
        print(f"📷 Image se text mila: {text}")
        return text
    except Exception as e:
        print(f"Image error: {e}")
        return ""

@app.route("/webhook", methods=["POST"])
def webhook():
    incoming_msg = request.values.get("Body", "").strip()
    sender = request.values.get("From", "").replace("whatsapp:", "")
    media_url = request.values.get("MediaUrl0", "")
    account_sid = request.values.get("AccountSid", "")
    auth_token = "3Ed3yW0pPmQelT9EFyTUYAyKGyH_74NvhABHDF11Z8PFRb9Vt"

    print(f"📩 Message from {sender}:\n{incoming_msg}")

    resp = MessagingResponse()
    msg = resp.message()

    name = get_name(sender)
    print(f"👤 Name found: {name}")

    if media_url:
        print(f"📷 Image mili: {media_url}")
        extracted_text = extract_text_from_image(media_url, account_sid, auth_token)
        if extracted_text:
            incoming_msg = extracted_text

    orders = parse_order(incoming_msg)

    if orders:
        reply_lines = [f"✅ Orders Received! ({name})"]
        for item, quantity in orders:
            log_order(name, sender, item, quantity)
            reply_lines.append(f"• {item.title()} — Qty: {quantity}")
        reply_lines.append("Thank you! 🙏")
        msg.body("\n".join(reply_lines))
    else:
        if media_url:
            msg.body("❌ Image se order detect nahi hua!\nPlease text mein order likho.")
        else:
            msg.body(f"❌ Format samajh nahi aaya!\nExample:\nCement 50\nSand: 100\nBrick - 200")

    return str(resp)

if __name__ == "__main__":
    setup_excel()
    print("🤖 Bot server shuru ho gaya!")
    print("🌐 Webhook: http://localhost:5000/webhook")
    app.run(debug=False, port=5000)