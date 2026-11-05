"""Consistency Tracker - Flask app that stores everything in one Excel file."""
import hashlib
import hmac
import os
import threading
import time
import uuid
from datetime import date, datetime, timedelta

from flask import Flask, jsonify, redirect, render_template, request, session, url_for
from openpyxl import Workbook, load_workbook

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", BASE_DIR)
os.makedirs(DATA_DIR, exist_ok=True)
XLSX_PATH = os.path.join(DATA_DIR, "tracker.xlsx")

# ---- Single user. Only salted PBKDF2 hashes live in code; the browser never sees them.
SALT = b"c7f1a94e2b6d4f08"
USER_HASH = "8ad9f91c0eafc0ac1432959d3b3d7140ae6b30c5b0f60aa4e564a3d86414bf27"
PASS_HASH = "1181da1f8e503fa1c16790d11aad15b215575e731fac53a54facd154d7cf3901"


def _h(value: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", value.encode(), SALT, 200_000).hex()


app = Flask(__name__)
# Set SECRET_KEY in your host's environment variables for best security.
app.secret_key = os.environ.get("SECRET_KEY") or hashlib.sha256((PASS_HASH + "session").encode()).hexdigest()
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=timedelta(days=14),
)

LOCK = threading.RLock()
FAILS = {}  # ip -> [count, first_fail_time]

# ------------------------------------------------------------------ Excel layer
SHEETS = {
    "Tasks": ["ID", "Task", "Start Date", "Target Date", "Created"],
    "Checkins": ["Task ID", "Date", "Done"],
    "Notes": ["Note ID", "Task ID", "Added On", "Note"],
}


def _open():
    if not os.path.exists(XLSX_PATH):
        wb = Workbook()
        wb.remove(wb.active)
        for name, cols in SHEETS.items():
            wb.create_sheet(name).append(cols)
        wb.save(XLSX_PATH)
    return load_workbook(XLSX_PATH)


def _save(wb):
    tmp = XLSX_PATH + ".tmp"
    wb.save(tmp)
    os.replace(tmp, XLSX_PATH)


def _append(ws, values):
    ws.append([None] * len(values))
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=ws.max_row, column=i)
        c.value = v
        if isinstance(v, str):
            c.data_type = "s"  # stops text like "=1+1" being treated as a formula


def _rows(ws):
    return [r for r in ws.iter_rows(min_row=2, values_only=True) if r and r[0] is not None]


def _d(s):
    return datetime.strptime(str(s), "%Y-%m-%d").date()


# ------------------------------------------------------------------ Analytics
def _streaks(done_days, start, end, today):
    longest = run = 0
    d = start
    while d <= end:
        run = run + 1 if d in done_days else 0
        longest = max(longest, run)
        d += timedelta(days=1)
    # current streak: today is still open, so don't break it until the day ends
    cur_end = min(today, end)
    d = cur_end if cur_end in done_days else cur_end - timedelta(days=1)
    cur = 0
    while d >= start and d in done_days:
        cur += 1
        d -= timedelta(days=1)
    return cur, longest


def build_data():
    today = date.today()
    wb = _open()
    tasks_raw = _rows(wb["Tasks"])
    checks = {}
    for tid, d, done in _rows(wb["Checkins"]):
        if str(done).lower() == "yes":
            checks.setdefault(tid, set()).add(_d(d))
    notes = {}
    for nid, tid, added, text in _rows(wb["Notes"]):
        notes.setdefault(tid, []).append({"id": nid, "added": added, "text": text})

    tasks, day_sched, day_done = [], {}, {}
    for tid, name, s, t, _created in tasks_raw:
        start, target = _d(s), _d(t)
        done_days = checks.get(tid, set())
        total = (target - start).days + 1
        last = min(today, target)
        elapsed = max(0, (last - start).days + 1)
        done_total = len([d for d in done_days if start <= d <= target])
        done_past = len([d for d in done_days if start <= d <= min(today - timedelta(days=1), target)])
        past_days = max(0, (min(today - timedelta(days=1), target) - start).days + 1)
        cur, longest = _streaks(done_days, start, target, today)
        days = []
        d = start
        while d <= target:
            days.append({"date": d.isoformat(), "done": d in done_days, "future": d > today})
            if d <= today:
                day_sched[d] = day_sched.get(d, 0) + 1
                day_done[d] = day_done.get(d, 0) + (1 if d in done_days else 0)
            d += timedelta(days=1)
        status = "upcoming" if start > today else ("completed" if target < today else "active")
        tasks.append({
            "id": tid, "name": name, "start": s, "target": t, "status": status,
            "total_days": total, "elapsed": elapsed, "done": done_total,
            "missed": past_days - done_past,
            "progress_pct": round(100 * done_total / total) if total else 0,
            "consistency_pct": round(100 * done_total / elapsed) if elapsed else 0,
            "days_left": max(0, (target - today).days),
            "streak": cur, "longest": longest,
            "days": days, "notes": list(reversed(notes.get(tid, []))),
        })

    sched, done = sum(day_sched.values()), sum(day_done.values())
    monday = today - timedelta(days=today.weekday())
    weeks = []
    for i in range(7, -1, -1):
        ws_ = monday - timedelta(weeks=i)
        sc = sum(day_sched.get(ws_ + timedelta(days=k), 0) for k in range(7))
        dn = sum(day_done.get(ws_ + timedelta(days=k), 0) for k in range(7))
        weeks.append({"label": ws_.strftime("%d %b"), "pct": round(100 * dn / sc) if sc else 0, "has": sc > 0})
    today_tasks = [t for t in tasks if t["status"] == "active"]
    today_done = sum(1 for t in today_tasks if any(x["date"] == today.isoformat() and x["done"] for x in t["days"]))
    return {
        "today": today.isoformat(),
        "overall_pct": round(100 * done / sched) if sched else 0,
        "tasks_total": len(tasks),
        "active": len(today_tasks),
        "completed": sum(1 for t in tasks if t["status"] == "completed"),
        "today_done": today_done,
        "best_streak": max([t["longest"] for t in tasks], default=0),
        "weeks": weeks,
        "tasks": tasks,
    }


# ------------------------------------------------------------------ Auth
def login_required(fn):
    from functools import wraps

    @wraps(fn)
    def wrap(*a, **k):
        if not session.get("ok"):
            if request.path.startswith("/api/"):
                return jsonify(error="Please sign in again."), 401
            return redirect(url_for("login"))
        return fn(*a, **k)

    return wrap


@app.after_request
def headers(resp):
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["X-Frame-Options"] = "DENY"
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("ok"):
        return redirect(url_for("home"))
    error = None
    if request.method == "POST":
        ip = request.headers.get("X-Forwarded-For", request.remote_addr or "?").split(",")[0].strip()
        count, first = FAILS.get(ip, [0, time.time()])
        if time.time() - first > 300:
            count, first = 0, time.time()
        if count >= 5:
            error = "Too many attempts. Try again in a few minutes."
        else:
            u_ok = hmac.compare_digest(_h(request.form.get("username", "").strip()), USER_HASH)
            p_ok = hmac.compare_digest(_h(request.form.get("password", "")), PASS_HASH)
            if u_ok and p_ok:
                FAILS.pop(ip, None)
                session.clear()
                session["ok"] = True
                session.permanent = True
                return redirect(url_for("home"))
            FAILS[ip] = [count + 1, first]
            error = "Username or password is incorrect."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ------------------------------------------------------------------ Pages & API
@app.route("/")
@login_required
def home():
    return render_template("index.html")


@app.route("/api/data")
@login_required
def api_data():
    with LOCK:
        return jsonify(build_data())


@app.route("/api/tasks", methods=["POST"])
@login_required
def add_task():
    j = request.get_json(silent=True) or {}
    name = str(j.get("name", "")).strip()[:120]
    note = str(j.get("note", "")).strip()[:500]
    try:
        start = _d(j.get("start") or date.today().isoformat())
        target = _d(j.get("target"))
    except (ValueError, TypeError):
        return jsonify(error="Pick a valid target date."), 400
    if not name:
        return jsonify(error="Give the task a name."), 400
    if target < start:
        return jsonify(error="Target date must be on or after the start date."), 400
    if (target - start).days > 730:
        return jsonify(error="Keep the range under 2 years."), 400
    with LOCK:
        wb = _open()
        tid = uuid.uuid4().hex[:8]
        _append(wb["Tasks"], [tid, name, start.isoformat(), target.isoformat(), datetime.now().strftime("%Y-%m-%d %H:%M")])
        if note:
            _append(wb["Notes"], [uuid.uuid4().hex[:8], tid, datetime.now().strftime("%d %b %Y, %H:%M"), note])
        _save(wb)
    return jsonify(ok=True)


@app.route("/api/tasks/<tid>/toggle", methods=["POST"])
@login_required
def toggle(tid):
    j = request.get_json(silent=True) or {}
    try:
        d = _d(j.get("date"))
    except (ValueError, TypeError):
        return jsonify(error="Bad date."), 400
    with LOCK:
        wb = _open()
        task = next((r for r in _rows(wb["Tasks"]) if r[0] == tid), None)
        if not task:
            return jsonify(error="Task not found."), 404
        if not (_d(task[2]) <= d <= _d(task[3])) or d > date.today():
            return jsonify(error="That day can't be ticked."), 400
        ws = wb["Checkins"]
        hit = next((i for i, r in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1)
                    if i > 1 and r[0] == tid and r[1] == d.isoformat()), None)
        if hit:
            ws.delete_rows(hit)
        else:
            _append(ws, [tid, d.isoformat(), "Yes"])
        _save(wb)
    return jsonify(ok=True)


@app.route("/api/tasks/<tid>/notes", methods=["POST"])
@login_required
def add_note(tid):
    text = str((request.get_json(silent=True) or {}).get("note", "")).strip()[:500]
    if not text:
        return jsonify(error="Write something first."), 400
    with LOCK:
        wb = _open()
        if not any(r[0] == tid for r in _rows(wb["Tasks"])):
            return jsonify(error="Task not found."), 404
        _append(wb["Notes"], [uuid.uuid4().hex[:8], tid, datetime.now().strftime("%d %b %Y, %H:%M"), text])
        _save(wb)
    return jsonify(ok=True)


@app.route("/api/tasks/<tid>", methods=["DELETE"])
@login_required
def delete_task(tid):
    with LOCK:
        wb = _open()
        for sheet, col in (("Tasks", 0), ("Checkins", 0), ("Notes", 1)):
            ws = wb[sheet]
            for i in range(ws.max_row, 1, -1):
                if ws.cell(row=i, column=col + 1).value == tid:
                    ws.delete_rows(i)
        _save(wb)
    return jsonify(ok=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
