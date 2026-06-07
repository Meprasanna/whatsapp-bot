from flask import Flask, request
from twilio.twiml.messaging_response import MessagingResponse
import openpyxl
import os
import json
import re
from datetime import datetime

app = Flask(__name__)
EXCEL_FILE = "orders.xlsx"
CONTACTS_FILE = "contacts.json"

def setup_excel():
    if not os.path.exists(EXCEL_FILE):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Orders"
        ws.append(["Order No", "Date", "Time", "Sender Name", "Phone", "Item", "Qty (kg)", "Qty (pieces/nag)", "Qty (bundles/gaddi)"])
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

def parse_quantity(qty_str):
    qty_str = qty_str.strip().lower()
    
    kg = ""
    pieces = ""
    bundles = ""
    
    # gm/gram → kg convert
    gm_match = re.search(r'(\d+\.?\d*)\s*(gm|gram|grm|g)\b', qty_str)
    if gm_match:
        grams = float(gm_match.group(1))
        kg = round(grams / 1000, 3)
        return str(kg), "", ""
    
    # kg
    kg_match = re.search(r'(\d+\.?\d*)\s*(kg|kilo|kilogram)\b', qty_str)
    if kg_match:
        kg = kg_match.group(1)
        return kg, "", ""
    
    # nag/nug/piece/pcs
    nag_match = re.search(r'(\d+\.?\d*)\s*(nag|nug|piece|pcs|pc|nos|no)\b', qty_str)
    if nag_match:
        pieces = nag_match.group(1)
        return "", pieces, ""
    
    # gaddi/bundle/gadi
    gaddi_match = re.search(r'(\d+\.?\d*)\s*(gaddi|gadi|bundle|bunch|bundi)\b', qty_str)
    if gaddi_match:
        bundles = gaddi_match.group(1)
        return "", "", bundles
    
    # sirf number — default kg
    num_match = re.search(r'(\d+\.?\d*)', qty_str)
    if num_match:
        kg = num_match.group(1)
        return kg, "", ""
    
    return "", "", ""

def log_order(sender_name, phone, item, kg, pieces, bundles):
    wb = openpyxl.load_workbook(EXCEL_FILE)
    ws = wb.active
    order_no = ws.max_row
    date = datetime.now().strftime("%d-%m-%Y")
    time_now = datetime.now().strftime("%H:%M:%S")
    ws.append([order_no, date, time_now, sender_name, phone, item, kg, pieces, bundles])
    wb.save(EXCEL_FILE)
    print(f"✅ Order! | {sender_name} | {item} | kg:{kg} | pcs:{pieces} | bundle:{bundles}")

def parse_line(line):
    line = line.strip()
    if not line:
        return None
    
    # Remove serial numbers like "1)", "2.", "3-"
    line = re.sub(r'^\d+[\)\.:\-]\s*', '', line)
    
    line_lower = line.lower()

    # Format: "item ke liye order: qty"
    if "order:" in line_lower:
        parts = line_lower.split("order:")
        item = parts[0].replace("ke liye", "").strip()
        quantity = parts[1].strip()
        if item and quantity:
            return (item, quantity)

    # Format: "item - qty" or "item : qty" or "item = qty"
    match = re.match(r'^([\w\s\u0900-\u097F]+?)\s*[\-:=]+\s*([\d\w\s\.]+)$', line)
    if match:
        item = match.group(1).strip()
        quantity = match.group(2).strip()
        if item and quantity:
            return (item, quantity)

    # Format: "item qty unit" like "Tamatar 10 kg"
    match = re.match(r'^([\w\s\u0900-\u097F]+?)\s+(\d+\.?\d*\s*(?:kg|gm|gram|nag|gaddi|bundle|pcs|pc|nos)?)\s*$', line, re.IGNORECASE)
    if match:
        item = match.group(1).strip()
        quantity = match.group(2).strip()
        if item and quantity:
            return (item, quantity)

    return None

def parse_order(message):
    orders = []
    lines = message.strip().split("\n")
    for line in lines:
        result = parse_line(line)
        if result:
            orders.append(result)
    return orders

def format_qty_reply(kg, pieces, bundles):
    parts = []
    if kg:
        parts.append(f"{kg} kg")
    if pieces:
        parts.append(f"{pieces} pcs")
    if bundles:
        parts.append(f"{bundles} bundle")
    return " | ".join(parts) if parts else "?"

@app.route("/webhook", methods=["POST"])
def webhook():
    incoming_msg = request.values.get("Body", "").strip()
    sender = request.values.get("From", "").replace("whatsapp:", "")
    print(f"📩 Message from {sender}:\n{incoming_msg}")

    resp = MessagingResponse()
    msg = resp.message()

    name = get_name(sender)
    orders = parse_order(incoming_msg)

    if orders:
        reply_lines = [f"✅ Orders Received! ({name})"]
        for item, quantity in orders:
            kg, pieces, bundles = parse_quantity(quantity)
            log_order(name, sender, item, kg, pieces, bundles)
            qty_display = format_qty_reply(kg, pieces, bundles)
            reply_lines.append(f"• {item.title()} — {qty_display}")
        reply_lines.append("Thank you! 🙏")
        msg.body("\n".join(reply_lines))
    else:
        msg.body("❌ Format samajh nahi aaya!\nExample:\nTamatar - 10kg\nDhaniya - 500gm\nLemon - 2 nag\nPudina - 1 gaddi")

    return str(resp)

if __name__ == "__main__":
    setup_excel()
    print("🤖 Bot server shuru ho gaya!")
    print("🌐 Webhook: http://localhost:5000/webhook")
    app.run(debug=False, port=5000)