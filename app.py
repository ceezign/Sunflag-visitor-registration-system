from flask import Flask, render_template, request, redirect, url_for, send_file, flash
import os
import psycopg2
from psycopg2.extras import RealDictCursor
import qrcode
from openpyxl import Workbook
from datetime import datetime
import pytz

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "fallback-secret")


# ----------------------------
# Nigeria Time
# ----------------------------
def get_nigeria_time():
    lagos = pytz.timezone("Africa/Lagos")
    return datetime.now(lagos)


# ----------------------------
# Database Connection (Postgres)
# ----------------------------
def get_db_connection():
    return psycopg2.connect(
        os.environ.get("postgresql://sunflag_visitor_user:48p3fqYm3SG3Ko21s8kOeiFSR0MGC76x@dpg-d7kesd67r5hc738a2oug-a.frankfurt-postgres.render.com/sunflag_visitor"),
        cursor_factory=RealDictCursor
    )


# ----------------------------
# Create Table
# ----------------------------
def init_db():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS visitors (
            id SERIAL PRIMARY KEY,
            sn INTEGER,
            visitor_name TEXT,
            full_address TEXT,
            tag_no TEXT,
            phone_no TEXT,
            whom_to_see TEXT,
            purpose TEXT,
            time_in TIMESTAMP,
            time_out TIMESTAMP,
            acknowledged INTEGER
        )
    """)

    conn.commit()
    cur.close()
    conn.close()


init_db()


# ----------------------------
# Generate QR Code
# ----------------------------
def generate_qr():
    base_url = os.environ.get("RENDER_EXTERNAL_URL", "http://127.0.0.1:5000")

    qr = qrcode.make(f"{base_url}/register")

    if not os.path.exists("static"):
        os.makedirs("static")

    qr.save("static/qr.png")


generate_qr()


# ----------------------------
# Routes
# ----------------------------
@app.route("/")
def home():
    return render_template("home.html")


@app.route("/register")
def register():
    return render_template("register.html")


@app.route("/add", methods=["POST"])
def add_visitor():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM visitors")
    sn = cur.fetchone()['count'] + 1

    acknowledged = 1 if request.form.get("acknowledged") else 0

    time_in = get_nigeria_time()

    cur.execute("""
        INSERT INTO visitors 
        (sn, visitor_name, full_address, tag_no, phone_no, whom_to_see, purpose, time_in, acknowledged)
        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)
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
    cur.close()
    conn.close()

    flash("Thank you. Your visit has been successfully registered.")
    return redirect(url_for("register"))


@app.route("/dashboard")
def dashboard():
    conn = get_db_connection()
    cur = conn.cursor()

    cur.execute("SELECT * FROM visitors ORDER BY id DESC")
    visitors = cur.fetchall()

    cur.close()
    conn.close()

    return render_template("dashboard.html", visitors=visitors)


@app.route("/timeout/<int:id>")
def timeout(id):
    conn = get_db_connection()
    cur = conn.cursor()

    time_out = get_nigeria_time()

    cur.execute(
        "UPDATE visitors SET time_out = %s WHERE id = %s",
        (time_out, id)
    )

    conn.commit()
    cur.close()
    conn.close()

    return redirect(url_for("dashboard"))


# ----------------------------
# Export Excel (Today)
# ----------------------------
@app.route("/export-today")
def export_today():
    conn = get_db_connection()
    cur = conn.cursor()

    today = get_nigeria_time().date()

    cur.execute("""
        SELECT * FROM visitors 
        WHERE DATE(time_in) = %s
    """, (today,))

    visitors = cur.fetchall()

    cur.close()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Daily Report"

    headers = [
        "SN", "Name", "Phone", "Whom To See",
        "Purpose", "Time In", "Time Out"
    ]

    ws.append(headers)

    for v in visitors:
        ws.append([
            v["sn"],
            v["visitor_name"],
            v["phone_no"],
            v["whom_to_see"],
            v["purpose"],
            str(v["time_in"]),
            str(v["time_out"])
        ])

    filename = "daily_report.xlsx"
    wb.save(filename)

    return send_file(filename, as_attachment=True)


# ----------------------------
# Run App
# ----------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)