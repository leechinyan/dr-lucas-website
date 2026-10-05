"""五金鹿卡斯 - small server: serves the page and records visits/bookings into Excel.

Setup:   pip install flask openpyxl requests
Run:     python server.py
Open:    http://localhost:5001
Edit EXCEL_PATH below to point at your Excel file in the Downloads folder.
Close the Excel file before submitting (Excel locks open files).
"""
import os, threading, datetime as dt
from pathlib import Path
import requests
from flask import Flask, request, jsonify, send_from_directory
from openpyxl import Workbook, load_workbook

# >>> change this to your real file name <<<
EXCEL_PATH = https://github.com/leechinyan/dr-lucas-website/blob/0674a6c87f141020df5ebb1e9345c9ab17fa3cd3/logtemp.xlsx
HEADERS = ["Timestamp", "IP", "Country", "City", "Contact", "Purpose", "Appointment Date"]

app = Flask(__name__)
lock = threading.Lock()


def client_ip():
    fwd = request.headers.get("X-Forwarded-For", "")
    return fwd.split(",")[0].strip() if fwd else request.remote_addr


def locate(ip):
    if ip.startswith(("127.", "10.", "192.168.")) or ip == "::1":
        return "Local", "Local"
    try:
        r = requests.get(f"http://ip-api.com/json/{ip}?fields=status,country,city", timeout=4).json()
        if r.get("status") == "success":
            return r.get("country", ""), r.get("city", "")
    except Exception:
        pass
    return "Unknown", "Unknown"


def append_row(row):
    with lock:
        if EXCEL_PATH.exists():
            wb = load_workbook(EXCEL_PATH)
            ws = wb.active
            if ws.max_row == 1 and ws["A1"].value is None:
                ws.append(HEADERS)
        else:
            wb = Workbook()
            ws = wb.active
            ws.append(HEADERS)
        ws.append(row)
        wb.save(EXCEL_PATH)


def base_row():
    ip = client_ip()
    country, city = locate(ip)
    return [dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), ip, country, city]


@app.get("/")
def index():
    return send_from_directory(Path(__file__).parent, "index.html")


@app.post("/api/visit")
def visit():
    try:
        append_row(base_row() + ["", "", ""])
    except PermissionError:
        pass  # Excel file is open; skip this visit
    return jsonify(ok=True)


@app.post("/api/book")
def book():
    data = request.get_json(force=True)
    try:
        d = dt.date.fromisoformat(data["date"])
    except Exception:
        return jsonify(message="Invalid Date"), 400
    today = dt.date.today()
    # keep rule in sync with the page: today..end of this month is fully booked
    next_month = (today.replace(day=28) + dt.timedelta(days=4)).replace(day=1)
    if d < next_month:
        return jsonify(message="Selected Date Is Full, Please Choose Another Date"), 400
    try:
        append_row(base_row() + [data["contact"], data["purpose"], d.isoformat()])
    except PermissionError:
        return jsonify(message="Excel 文件正被打开，请关闭后重试"), 500
    return jsonify(message="Reservation Successful! We Will Contact You Soon!")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5001)))
