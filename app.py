from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import sqlite3
from datetime import datetime
import qrcode
import os
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "sunflag_secret_key"

DATABASE = "visitors.db"


# ----------------------------
# Database Connection
# ----------------------------
def get_db_connection():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


# ----------------------------
# Create Table If Not Exists
# ----------------------------
def init_db():
    conn = get_db_connection()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS visitors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sn INTEGER,
            visitor_name TEXT NOT NULL,
            full_address TEXT,
            tag_no TEXT,
            phone_no TEXT,
            whom_to_see TEXT,
            purpose TEXT,
            time_in TEXT,
            time_out TEXT,
            acknowledged INTEGER
        )
    """)
    conn.commit()
    conn.close()


init_db()


# ----------------------------
# Generate QR Code Automatically
# ----------------------------


def generate_qr():
    base_url = os.environ.get("RENDER_EXTERNAL_URL", "http://127.0.0.1:5000")
    qr = qrcode.make(f"{base_url}/register")
    qr.save("static/qr.png")
    if not os.path.exists("static"):
        os.makedirs("static")

    qr.save("static/qr.png")

generate_qr()


# ----------------------------
# Home Page
# ----------------------------
@app.route("/")
def home():
    return render_template("home.html")


# ----------------------------
# Registration Page
# ----------------------------
@app.route("/register")
def register():
    return render_template("register.html")


@app.route("/add", methods=["POST"])
def add_visitor():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM visitors")
    sn = cur.fetchone()[0] + 1

    acknowledged = 1 if request.form.get("acknowledged") else 0

    time_in = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    cur.execute("""
        INSERT INTO visitors 
        (sn, visitor_name, full_address, tag_no, phone_no, whom_to_see, purpose, time_in, acknowledged)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        sn,
        request.form["visitor_name"],
        request.form["full_address"],
        request.form["tag_no"],
        request.form["phone_no"],
        request.form["whom_to_see"],
        request.form["purpose"],
        time_in,
        acknowledged
    ))

    conn.commit()
    conn.close()

    flash("Thank you. Your visit has been successfully registered.")
    return redirect(url_for("register"))


# ----------------------------
# Dashboard
# ----------------------------
@app.route("/dashboard")
def dashboard():
    conn = get_db_connection()
    visitors = conn.execute("SELECT * FROM visitors").fetchall()
    conn.close()
    return render_template("dashboard.html", visitors=visitors)


@app.route("/timeout/<int:id>")
def timeout(id):
    conn = get_db_connection()
    time_out = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn.execute("UPDATE visitors SET time_out = ? WHERE id = ?", (time_out, id))
    conn.commit()
    conn.close()
    return redirect(url_for("dashboard"))


# ----------------------------
# Export Today's Visitors To Excel
# ----------------------------
@app.route("/export-today")
def export_today():
    conn = get_db_connection()
    today = datetime.now().strftime("%Y-%m-%d")

    visitors = conn.execute(
        "SELECT * FROM visitors WHERE date(time_in) = ?", (today,)
    ).fetchall()

    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Report"

    headers = [
        "SN", "Visitor Name", "Address", "Tag No",
        "Phone", "Whom To See", "Purpose",
        "Time In", "Time Out", "Acknowledged"
    ]

    ws.append(headers)

    for v in visitors:
        ws.append([
            v["sn"],
            v["visitor_name"],
            v["full_address"],
            v["tag_no"],
            v["phone_no"],
            v["whom_to_see"],
            v["purpose"],
            v["time_in"],
            v["time_out"],
            "Yes" if v["acknowledged"] else "No"
        ])

        filename = f"sunflag_daily_report_{today}.xlsx"
        wb.save(filename)

        return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)