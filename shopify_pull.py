import requests
import pandas as pd
from datetime import datetime
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

# ========== CONFIG ==========
ACCESS_TOKEN = os.getenv("SHOPIFY_TOKEN")
SHOP = "o2otestv2.myshopify.com"

EMAIL_FROM = "tesprojek2025@gmail.com"
EMAIL_TO = "bimapgusti@gmail.com"
EMAIL_PASS = os.getenv("EMAIL_PASS")
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

DEBUG = True  # <--- aktifkan untuk print detail order

# ========== RANGE TANGGAL (MANUAL) ==========
start_date = datetime(2024, 8, 1)   # <-- ubah sesuai kebutuhan
end_date   = datetime(2024, 8, 31)   # <-- ubah sesuai kebutuhan

# ========== REQUEST DATA DARI SHOPIFY ==========
current_url = f"https://{SHOP}/admin/api/2024-07/orders.json"
params = {
    "status": "any",
    "created_at_min": start_date.strftime("%Y-%m-%dT00:00:00Z"),
    "created_at_max": end_date.strftime("%Y-%m-%dT23:59:59Z"),
    "limit": 250
}

headers = {
    "X-Shopify-Access-Token": ACCESS_TOKEN,
    "Content-Type": "application/json"
}

orders = []
while current_url:
    response = requests.get(current_url, headers=headers, params=params)
    if response.status_code != 200:
        print("Error getting orders:", response.status_code, response.text)
        break

    data = response.json()
    orders.extend(data.get("orders", []))

    # pagination
    next_link = response.links.get("next", {}).get("url")
    current_url = next_link if next_link else None
    params = None

print(f"Jumlah order mentah diambil: {len(orders)}")

# ========== PROSES DATA ==========
summary = {}
seen_ids = set()

for order in orders:
    oid = order.get("id")
    if oid in seen_ids:
        continue
    seen_ids.add(oid)

    created = order["created_at"][:10]  # YYYY-MM-DD

    # Ambil sales
    if "current_total_price" in order:
        total_price = float(order["current_total_price"])
    elif "total_price_set" in order and order["total_price_set"].get("shop_money"):
        total_price = float(order["total_price_set"]["shop_money"].get("amount", 0))
    else:
        total_price = float(order.get("total_price", 0) or 0)

    # total qty
    total_qty = sum(int(i.get("quantity", 0)) for i in order.get("line_items", []))

    # refunded qty
    refunded_qty = 0
    for refund in order.get("refunds", []):
        for rli in refund.get("refund_line_items", []):
            refunded_qty += int(rli.get("quantity", 0))

    net_items = max(total_qty - refunded_qty, 0)

    # DEBUG PRINT
    if DEBUG:
        print(f"OrderID {oid} | Date {created} | Sales {total_price} | "
              f"Items {net_items} (total:{total_qty}, refunded:{refunded_qty})")

    if created not in summary:
        summary[created] = {"sales": 0, "orders": 0, "items": 0}

    summary[created]["sales"] += total_price
    summary[created]["orders"] += 1
    summary[created]["items"] += net_items

# ========== DATAFRAME ==========
all_days = pd.date_range(start=start_date, end=end_date).strftime("%Y-%m-%d")
rows = []

for day in all_days:
    rows.append({
        "Day": datetime.strptime(day, "%Y-%m-%d").strftime("%b %d, %Y").replace(" 0", " "),
        "Total sales": summary.get(day, {}).get("sales", 0),
        "Net items sold": summary.get(day, {}).get("items", 0),
        "Orders": summary.get(day, {}).get("orders", 0)
    })

df = pd.DataFrame(rows)

# hitung total keseluruhan
total_sales_all = df["Total sales"].sum()
total_orders_all = df["Orders"].sum()
total_items_all = df["Net items sold"].sum()

# tambahkan baris 'TOTAL' di df
df.loc[len(df)] = {
    "Day": "TOTAL",
    "Total sales": total_sales_all,
    "Net items sold": total_items_all,
    "Orders": total_orders_all
}

print(df)

# Simpan ke Excel
filename = f"report_{start_date.strftime('%Y%m%d')}_to_{end_date.strftime('%Y%m%d')}.xlsx"
df.to_excel(filename, index=False)

# ========== KIRIM EMAIL ==========
msg = MIMEMultipart()
msg["From"] = EMAIL_FROM
msg["To"] = EMAIL_TO
msg["Subject"] = f"Laporan Shopify {start_date.strftime('%Y-%m-%d')} s/d {end_date.strftime('%Y-%m-%d')}"

body = f"""
Halo,

Berikut laporan penjualan Shopify untuk periode {start_date.strftime('%Y-%m-%d')} sampai {end_date.strftime('%Y-%m-%d')}.

📊 Ringkasan:
- Total Sales   : Rp {total_sales_all:,.2f}
- Total Orders  : {total_orders_all}
- Total Items   : {total_items_all}

Detail harian ada di file Excel terlampir.

Salam,
Bot Shopify
"""
msg.attach(MIMEText(body, "plain"))

with open(filename, "rb") as f:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", f"attachment; filename= {filename}")
    msg.attach(part)

server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
server.starttls()
server.login(EMAIL_FROM, EMAIL_PASS)
server.send_message(msg)
server.quit()

print(f"✅ Laporan berhasil dikirim ke {EMAIL_TO}")
