from flask import Flask, send_from_directory, jsonify, request
import sqlite3
import os

app = Flask(__name__)
DB_PATH = "attendance.db"

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

# 근무표 조회
@app.route("/schedule")
def get_schedule():
    dates = request.args.get("dates")
    month = request.args.get("month")

    conn = get_db()
    cur = conn.cursor()

    if dates:
        date_list = dates.split(",")
        placeholders = ",".join("?"*len(date_list))
        cur.execute(
            f"SELECT name, date, status FROM work_schedule WHERE date IN ({placeholders})",
            date_list
        )
    elif month:
        cur.execute(
            "SELECT name, date, status FROM work_schedule WHERE date LIKE ?",
            (f"{month}%",)
        )
    else:
        return jsonify({})

    rows = cur.fetchall()
    conn.close()

    result = {}
    for r in rows:
        if r["name"] not in result:
            result[r["name"]] = {}
        result[r["name"]][r["date"]] = r["status"]

    return jsonify(result)

# 비상근무 조회
@app.route("/special")
def get_special():
    month = request.args.get("month")

    conn = get_db()
    cur = conn.cursor()

    cur.execute(
        "SELECT duty, name, date FROM special_duty WHERE date LIKE ?",
        (f"{month}%",)
    )

    rows = cur.fetchall()
    conn.close()

    result = {}
    for r in rows:
        result.setdefault(r["date"], {}).setdefault(r["duty"], []).append(r["name"])

    return jsonify(result)

# 근무 추가
@app.route("/add_schedule", methods=["POST"])
def add_schedule():
    data = request.json
    conn = get_db()
    cur = conn.cursor()

    cur.execute(
        "INSERT INTO work_schedule (name, date, status) VALUES (?, ?, ?)",
        (data["name"], data["date"], data["status"])
    )

    conn.commit()
    conn.close()
    return jsonify({"result": "ok"})

# ⭐ 근무 일괄 추가 (엑셀 복붙)
@app.route("/add_schedule_bulk", methods=["POST"])
def add_schedule_bulk():
    data = request.json["data"]

    conn = get_db()
    cur = conn.cursor()

    for line in data:
        if "\t" in line:
            parts = line.split("\t")
        else:
            parts = line.split(",")

        if len(parts) != 3:
            continue

        name, date, status = [p.strip() for p in parts]

        cur.execute(
            "INSERT INTO work_schedule (name, date, status) VALUES (?, ?, ?)",
            (name, date, status)
        )

    conn.commit()
    conn.close()
    return jsonify({"result": "ok"})

# ⭐ 비상근무 추가
@app.route("/add_special", methods=["POST"])
def add_special():
    data = request.json
    conn = get_db()
    cur = conn.cursor()

    cur.execute(
        "INSERT INTO special_duty (duty, name, date) VALUES (?, ?, ?)",
        (data["duty"], data["name"], data["date"])
    )

    conn.commit()
    conn.close()
    return jsonify({"result": "ok"})

@app.route("/")
def index():
    return send_from_directory(os.getcwd(), "test.html")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
