import re
import os
import sqlite3
import uuid
import io
import hashlib
import functools
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, jsonify, request, redirect, flash, url_for, session
from werkzeug.utils import secure_filename
import openpyxl
import requests as http_requests

app = Flask(__name__)
app.secret_key = "bni_scorecard_secret_2026"

EXCEL_DIR   = Path(__file__).parent
EXCEL_FILE  = EXCEL_DIR / "Copy of MTL Recommendations - 2.0.xlsx"
DATABASE    = EXCEL_DIR / "bni.db"
CHAPTERS_DIR = EXCEL_DIR / "chapters"
CHAPTERS_DIR.mkdir(exist_ok=True)


def get_db():
    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn


def hash_password(pw):
    return hashlib.sha256(pw.encode()).hexdigest()


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS chapters (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                name         TEXT NOT NULL,
                slug         TEXT NOT NULL UNIQUE,
                filename     TEXT,
                member_count INTEGER DEFAULT 0,
                uploaded_at  TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT NOT NULL,
                email      TEXT NOT NULL UNIQUE,
                phone      TEXT,
                password   TEXT NOT NULL,
                chapter_id INTEGER,
                role       TEXT NOT NULL DEFAULT 'vp',
                FOREIGN KEY (chapter_id) REFERENCES chapters(id)
            )
        """)
        # Add columns if upgrading from older schema
        for col, defn in [("member_count", "INTEGER DEFAULT 0"), ("uploaded_at", "TEXT")]:
            try:
                conn.execute(f"ALTER TABLE chapters ADD COLUMN {col} {defn}")
            except Exception:
                pass
        conn.execute("""
            CREATE TABLE IF NOT EXISTS log (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                ts           TEXT NOT NULL,
                user_name    TEXT NOT NULL,
                chapter_name TEXT,
                activity     TEXT NOT NULL
            )
        """)
        # Seed super admin if not exists
        existing = conn.execute("SELECT id FROM users WHERE email=?", ("info@keylines.net",)).fetchone()
        if not existing:
            conn.execute(
                "INSERT INTO users (name, email, phone, password, chapter_id, role) VALUES (?,?,?,?,?,?)",
                ("Subrata Kundu", "info@keylines.net", "9330109091", hash_password("pass1234"), None, "superadmin")
            )


# ── Auth helpers ──────────────────────────────────────────────────────────────

def write_log(user_name, activity, chapter_name=None):
    ts = datetime.now().strftime("%d %b %Y, %I:%M:%S %p")
    with get_db() as conn:
        conn.execute(
            "INSERT INTO log (ts, user_name, chapter_name, activity) VALUES (?,?,?,?)",
            (ts, user_name, chapter_name, activity)
        )


def login_required(f):
    @functools.wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user_id"):
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated


def current_user():
    if not session.get("user_id"):
        return None
    with get_db() as conn:
        return conn.execute("SELECT * FROM users WHERE id=?", (session["user_id"],)).fetchone()


def is_superadmin():
    u = current_user()
    return u and u["role"] == "superadmin"


def extract_sheet_id(url):
    """Extract Google Sheets file ID from various URL formats."""
    m = re.search(r'/spreadsheets/d/([a-zA-Z0-9_-]+)', url)
    return m.group(1) if m else None


def validate_excel_bytes(data):
    """
    Validate that bytes are a valid .xlsx with a Recommendations sheet.
    Returns (ok: bool, error_message: str, member_count: int)
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        return False, "Not a valid Excel (.xlsx) file.", 0
    if "Recommendations" not in wb.sheetnames:
        return False, f"Sheet 'Recommendations' not found. Sheets found: {', '.join(wb.sheetnames)}", 0
    # Quick member count
    ws = wb["Recommendations"]
    count = sum(
        1 for row in ws.iter_rows(min_row=1, values_only=True)
        if len(row) > 2 and isinstance(row[2], str) and "score" in row[2].lower()
        and row[1] and str(row[1]).strip()
    )
    return True, "", count


def save_chapter_file(chapter_id, chapter_name, data):
    """Save bytes as chapter Excel, delete old file, update DB. Returns member_count."""
    with get_db() as conn:
        row = conn.execute("SELECT filename FROM chapters WHERE id=?", (chapter_id,)).fetchone()
    if row and row["filename"]:
        old = CHAPTERS_DIR / row["filename"]
        if old.exists():
            old.unlink()
    safe_name = re.sub(r'[^a-z0-9]+', '_', chapter_name.lower()).strip('_')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"{safe_name}_{timestamp}.xlsx"
    (CHAPTERS_DIR / filename).write_bytes(data)
    try:
        _, members = parse_recommendations(filepath=CHAPTERS_DIR / filename)
        member_count = len(members)
    except Exception:
        member_count = 0
    uploaded_at = datetime.now().strftime("%d %b %Y, %I:%M %p")
    with get_db() as conn:
        conn.execute(
            "UPDATE chapters SET filename=?, member_count=?, uploaded_at=? WHERE id=?",
            (filename, member_count, uploaded_at, chapter_id)
        )
    return member_count


def slugify(text):
    text = text.lower().strip()
    text = re.sub(r'[^a-z0-9\s-]', '', text)
    text = re.sub(r'\s+', '-', text)
    return text or str(uuid.uuid4())[:8]

# ── Excel parser ─────────────────────────────────────────────────────────────

def parse_recommendations(filepath=None):
    """
    Read the Recommendations sheet and return:
      title   : str   — sheet title (e.g. "March 2026")
      members : list of {
          name, total_score, max_score, pct, traffic_light,
          metrics: [ {metric, your_score, max_score, pct, status, rec_text, tiers} ]
      }
    """
    path = Path(filepath) if filepath else EXCEL_FILE
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Recommendations"]

    title = ""
    members = []
    current = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        b = row[1] if len(row) > 1 else None
        c = row[2] if len(row) > 2 else None
        d = row[3] if len(row) > 3 else None
        e = row[4] if len(row) > 4 else None

        # Sheet title row
        if row[0] and b is None:
            title = str(row[0]).strip()
            continue

        if b is None:
            continue

        b_str = str(b).strip()

        # Member header row: c == "Your Score"
        if isinstance(c, str) and "score" in c.lower():
            if current and current["name"]:
                members.append(current)
            current = {"name": b_str, "metrics": []} if b_str else None
            continue

        if current is None:
            continue

        # Metric row
        try:
            your  = int(float(c)) if c is not None else 0
            max_s = int(float(d)) if d is not None else 0
        except (ValueError, TypeError):
            your, max_s = 0, 0

        rec_raw = str(e).strip() if e is not None else ""
        if rec_raw in ("None", "nan", ""):
            rec_raw = ""

        pct    = round(your / max_s * 100) if max_s else 0
        status = "full" if pct == 100 else ("partial" if pct > 0 else "none")

        tiers  = _parse_tiers(rec_raw, your, b_str)

        current["metrics"].append({
            "metric":     b_str,
            "your_score": your,
            "max_score":  max_s,
            "pct":        pct,
            "status":     status,
            "rec_text":   rec_raw,
            "tiers":      tiers,
        })

    if current and current["name"]:
        members.append(current)

    # Attach total_score & traffic_light to each member
    for m in members:
        total_row = next((x for x in m["metrics"] if x["metric"].upper() == "TOTAL"), None)
        m["total_score"] = total_row["your_score"] if total_row else 0
        m["max_score"]   = total_row["max_score"]  if total_row else 100
        m["pct"]         = total_row["pct"]         if total_row else 0
        m["traffic_light"] = _traffic_light(m["total_score"])

    members.sort(key=lambda x: x["total_score"], reverse=True)
    return title, members


def _traffic_light(score):
    """
    Black  : 25 or below  — Severe under-performance
    Red    : 26 – 49      — Critical / at risk
    Amber  : 50 – 69      — Warning / needs improvement
    Green  : 70 – 100     — Good standing
    Gray   : 0 / no data  — Exempt / grace period
    """
    if score <= 25:  return "gray"
    if score <= 49:  return "red"
    if score <= 69:  return "amber"
    return "green"


def _parse_tiers(rec_text, your_score, metric):
    """Parse 'X1/X2/.../XN for P1/P2/.../PN points' into tier list."""
    m = re.match(r"^([\d/]+)\s+for\s+([\d/]+)\s+points?", rec_text.strip(), re.I)
    if not m:
        return []
    try:
        thresholds = [int(x) for x in m.group(1).split("/") if x.strip()]
        pts_vals   = [int(x) for x in m.group(2).split("/") if x.strip()]
    except ValueError:
        return []
    if len(thresholds) != len(pts_vals):
        return []

    tiers = []
    found_next = False
    for threshold, pts in zip(thresholds, pts_vals):
        if pts <= your_score:
            status = "done"
        elif not found_next:
            status = "next"
            found_next = True
        else:
            status = "future"

        # Format large numbers (TYFCB amounts) in Indian style
        fmt = _indian(threshold) if threshold >= 1000 else str(threshold)
        tiers.append({"pts": pts, "threshold": threshold, "fmt": fmt, "status": status})
    return tiers


def _indian(n):
    """Format number in Indian lakh/crore style."""
    s = str(abs(int(n)))
    if len(s) <= 3:
        return s
    result = s[-3:]
    s = s[:-3]
    while s:
        result = s[-2:] + "," + result
        s = s[:-2]
    return result.lstrip(",")


# ── Jinja template helpers ───────────────────────────────────────────────────

def _act_key(activity):
    a = activity.lower()
    if "login" in a:        return "login"
    if "logout" in a:       return "logout"
    if "created" in a or "create" in a: return "create"
    if "deleted" in a or "delete" in a: return "delete"
    if "uploaded" in a or "upload" in a: return "upload"
    if "google sheet" in a or "imported" in a: return "sheet"
    if "password" in a:     return "profile"
    if "profile" in a:      return "profile"
    if "renamed" in a or "rename" in a: return "rename"
    return "default"

def activity_icon(activity):
    return "icon-" + _act_key(activity)

def activity_emoji(activity):
    return {"login":"🔓","logout":"🔒","create":"➕","delete":"🗑",
            "upload":"📤","sheet":"📊","profile":"✏️","rename":"🏷","default":"📌"}.get(_act_key(activity),"📌")

app.jinja_env.globals.update(activity_icon=activity_icon, activity_emoji=activity_emoji)


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    with get_db() as conn:
        chapters = conn.execute("SELECT * FROM chapters ORDER BY name ASC").fetchall()
    return render_template("home.html", chapters=chapters)


@app.route("/api/member/<path:name>")
def member_api(name):
    _, members = parse_recommendations()
    m = next((x for x in members if x["name"] == name), None)
    if not m:
        return jsonify({"error": "not found"}), 404
    return jsonify(m)


@app.route("/upload", methods=["POST"])
def upload():
    f = request.files.get("excel")
    if f and f.filename.endswith(".xlsx"):
        f.save(EXCEL_FILE)
    return redirect("/")


# ── Auth routes ───────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("user_id"):
        return redirect("/admin")
    error = None
    if request.method == "POST":
        email = request.form.get("email", "").strip().lower()
        password = request.form.get("password", "")
        with get_db() as conn:
            user = conn.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if user and user["password"] == hash_password(password):
            session["user_id"] = user["id"]
            ch_name = None
            if user["role"] != "superadmin" and user["chapter_id"]:
                with get_db() as c2:
                    ch = c2.execute("SELECT name FROM chapters WHERE id=?", (user["chapter_id"],)).fetchone()
                    ch_name = ch["name"] if ch else None
            write_log(user["name"], "Login", "Super Admin" if user["role"] == "superadmin" else ch_name)
            return redirect("/admin")
        error = "Invalid email or password."
    return render_template("login.html", error=error)


@app.route("/logout")
def logout():
    u = current_user()
    if u:
        ch_name = None
        if u["role"] != "superadmin" and u["chapter_id"]:
            with get_db() as conn:
                ch = conn.execute("SELECT name FROM chapters WHERE id=?", (u["chapter_id"],)).fetchone()
                ch_name = ch["name"] if ch else None
        write_log(u["name"], "Logout", "Super Admin" if u["role"] == "superadmin" else ch_name)
    session.clear()
    return redirect("/login")


# ── Admin routes ─────────────────────────────────────────────────────────────

@app.route("/admin")
@login_required
def admin():
    user = current_user()
    with get_db() as conn:
        if user["role"] == "superadmin":
            chapters = conn.execute("SELECT * FROM chapters ORDER BY name ASC").fetchall()
        else:
            chapters = conn.execute(
                "SELECT * FROM chapters WHERE id=?", (user["chapter_id"],)
            ).fetchall()
        # Attach VP user to each chapter
        chapter_list = []
        for ch in chapters:
            vp = conn.execute(
                "SELECT * FROM users WHERE chapter_id=? AND role='vp'", (ch["id"],)
            ).fetchone()
            tl = {"green": 0, "amber": 0, "red": 0, "gray": 0}
            if ch["filename"]:
                try:
                    _, members = parse_recommendations(filepath=CHAPTERS_DIR / ch["filename"])
                    for m in members:
                        tl[m["traffic_light"]] += 1
                except Exception:
                    pass
            chapter_list.append({"chapter": ch, "vp": vp, "tl": tl})
    create_form = session.pop('create_form', {})
    return render_template("admin.html", chapter_list=chapter_list, user=user, create_form=create_form)


@app.route("/admin/edit-user/<int:user_id>", methods=["POST"])
@login_required
def admin_edit_user(user_id):
    u = current_user()
    # Superadmin can edit anyone; VP can only edit themselves
    if u["role"] != "superadmin" and u["id"] != user_id:
        return redirect("/admin")
    name   = request.form.get("name", "").strip()
    email  = request.form.get("email", "").strip().lower()
    phone  = request.form.get("phone", "").strip()
    new_pw = request.form.get("password", "").strip()

    if not name or not email:
        flash("Name and email are required.", "error")
        return redirect("/admin")

    with get_db() as conn:
        # Check email not taken by a different user
        conflict = conn.execute(
            "SELECT id FROM users WHERE email=? AND id!=?", (email, user_id)
        ).fetchone()
        if conflict:
            flash(f"Email '{email}' is already in use by another user.", "error")
            return redirect("/admin")

        target = conn.execute("SELECT name, chapter_id FROM users WHERE id=?", (user_id,)).fetchone()
        target_name = target["name"] if target else "?"
        target_ch_id = target["chapter_id"] if target else None
        ch_row = conn.execute("SELECT name FROM chapters WHERE id=?", (target_ch_id,)).fetchone() if target_ch_id else None
        target_ch = ch_row["name"] if ch_row else None
        if new_pw:
            conn.execute("UPDATE users SET name=?, email=?, phone=?, password=? WHERE id=?",
                         (name, email, phone, hash_password(new_pw), user_id))
        else:
            conn.execute("UPDATE users SET name=?, email=?, phone=? WHERE id=?",
                         (name, email, phone, user_id))
    actor = current_user()
    activity = f"Password & profile updated for user '{target_name}' (name, email, phone, password)" if new_pw \
               else f"Profile updated for user '{target_name}' (name={name}, email={email}, phone={phone})"
    write_log(actor["name"], activity, "Super Admin" if actor["role"] == "superadmin" else target_ch)
    flash("User updated successfully.", "success")
    return redirect("/admin")


@app.route("/admin/create", methods=["POST"])
@login_required
def admin_create():
    if not is_superadmin():
        return redirect("/admin")
    name        = request.form.get("name", "").strip()
    vp_name     = request.form.get("vp_name", "").strip()
    vp_email    = request.form.get("vp_email", "").strip().lower()
    vp_phone    = request.form.get("vp_phone", "").strip()
    vp_password = request.form.get("vp_password", "").strip()

    # Store form data in session so it can be repopulated on error
    session['create_form'] = {
        "name": name, "vp_name": vp_name,
        "vp_email": vp_email, "vp_phone": vp_phone
    }

    # Validate chapter name
    if not name:
        flash("Chapter name is required.", "error")
        return redirect("/admin")

    # Validate VP fields — all required
    if not vp_name:
        flash("VP full name is required.", "error")
        return redirect("/admin")
    if not vp_email:
        flash("VP email is required.", "error")
        return redirect("/admin")
    if not vp_password:
        flash("VP password is required.", "error")
        return redirect("/admin")

    with get_db() as conn:
        # Check email not already used
        existing_user = conn.execute("SELECT id FROM users WHERE email=?", (vp_email,)).fetchone()
        if existing_user:
            flash(f"Email '{vp_email}' is already assigned to another user.", "error")
            return redirect("/admin")

        # Check VP not already assigned to another chapter
        existing_vp = conn.execute(
            "SELECT u.id, c.name FROM users u JOIN chapters c ON u.chapter_id=c.id WHERE u.email=?",
            (vp_email,)
        ).fetchone()
        if existing_vp:
            flash(f"This user is already assigned to chapter '{existing_vp['name']}'.", "error")
            return redirect("/admin")

        slug = slugify(name)
        if conn.execute("SELECT id FROM chapters WHERE slug=?", (slug,)).fetchone():
            slug = slug + "-" + str(uuid.uuid4())[:6]

        cur = conn.execute("INSERT INTO chapters (name, slug) VALUES (?, ?)", (name, slug))
        chapter_id = cur.lastrowid
        conn.execute(
            "INSERT INTO users (name, email, phone, password, chapter_id, role) VALUES (?,?,?,?,?,?)",
            (vp_name, vp_email, vp_phone, hash_password(vp_password), chapter_id, "vp")
        )

    session.pop('create_form', None)
    u = current_user()
    write_log(u["name"], f"Created new chapter '{name}' with VP {vp_email}", "Super Admin")
    flash(f"Chapter '{name}' created with VP login for {vp_email}.", "success")
    return redirect("/admin")


@app.route("/admin/rename/<int:chapter_id>", methods=["POST"])
@login_required
def admin_rename(chapter_id):
    if not is_superadmin():
        return redirect("/admin")
    new_name = request.form.get("name", "").strip()
    if not new_name:
        flash("Chapter name cannot be empty.", "error")
        return redirect("/admin")
    new_slug = slugify(new_name)
    with get_db() as conn:
        conflict = conn.execute(
            "SELECT id FROM chapters WHERE name=? AND id!=?", (new_name, chapter_id)
        ).fetchone()
        if conflict:
            flash(f"A chapter named '{new_name}' already exists. Please choose a unique name.", "error")
            return redirect("/admin")
        # Ensure slug uniqueness
        slug_conflict = conn.execute(
            "SELECT id FROM chapters WHERE slug=? AND id!=?", (new_slug, chapter_id)
        ).fetchone()
        if slug_conflict:
            new_slug = new_slug + "-" + str(uuid.uuid4())[:6]
        old_row = conn.execute("SELECT name FROM chapters WHERE id=?", (chapter_id,)).fetchone()
        old_name = old_row["name"] if old_row else "?"
        conn.execute("UPDATE chapters SET name=?, slug=? WHERE id=?", (new_name, new_slug, chapter_id))
    u = current_user()
    write_log(u["name"], f"Renamed chapter '{old_name}' → '{new_name}'", "Super Admin")
    flash(f"Chapter renamed to '{new_name}'.", "success")
    return redirect("/admin")


@app.route("/admin/upload/<int:chapter_id>", methods=["POST"])
@login_required
def admin_upload(chapter_id):
    user = current_user()
    if user["role"] != "superadmin" and user["chapter_id"] != chapter_id:
        return redirect("/admin")
    f = request.files.get("excel")
    if not f or not f.filename.endswith(".xlsx"):
        flash("Please select a valid .xlsx file.", "error")
        return redirect("/admin")
    data = f.read()
    ok, err, _ = validate_excel_bytes(data)
    if not ok:
        flash(f"Invalid file: {err}", "error")
        return redirect("/admin")
    with get_db() as conn:
        row = conn.execute("SELECT name FROM chapters WHERE id=?", (chapter_id,)).fetchone()
    save_chapter_file(chapter_id, row["name"], data)
    u = current_user()
    write_log(u["name"], f"Uploaded new file for chapter", row["name"])
    flash("File uploaded successfully.", "success")
    return redirect("/admin")


@app.route("/admin/fetch/<int:chapter_id>", methods=["POST"])
@login_required
def admin_fetch(chapter_id):
    user = current_user()
    if user["role"] != "superadmin" and user["chapter_id"] != chapter_id:
        return redirect("/admin")
    url = request.form.get("sheet_url", "").strip()
    sheet_id = extract_sheet_id(url)
    if not sheet_id:
        flash("Invalid Google Sheets URL. Please paste the full spreadsheet link.", "error")
        return redirect("/admin")
    export_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    try:
        resp = http_requests.get(export_url, timeout=20)
        if resp.status_code != 200:
            flash(f"Could not fetch the spreadsheet (HTTP {resp.status_code}). Make sure it is shared publicly.", "error")
            return redirect("/admin")
        data = resp.content
    except Exception as e:
        flash(f"Network error: {e}", "error")
        return redirect("/admin")
    ok, err, _ = validate_excel_bytes(data)
    if not ok:
        flash(f"Validation failed: {err}", "error")
        return redirect("/admin")
    with get_db() as conn:
        row = conn.execute("SELECT name FROM chapters WHERE id=?", (chapter_id,)).fetchone()
    save_chapter_file(chapter_id, row["name"], data)
    u = current_user()
    write_log(u["name"], f"Imported Google Sheet for chapter", row["name"])
    flash("Google Sheet imported successfully.", "success")
    return redirect("/admin")


@app.route("/admin/delete/<int:chapter_id>", methods=["POST"])
@login_required
def admin_delete(chapter_id):
    if not is_superadmin():
        return redirect("/admin")
    with get_db() as conn:
        row = conn.execute("SELECT name, filename FROM chapters WHERE id=?", (chapter_id,)).fetchone()
        ch_name = row["name"] if row else "?"
        if row and row["filename"]:
            f = CHAPTERS_DIR / row["filename"]
            if f.exists():
                f.unlink()
        conn.execute("DELETE FROM chapters WHERE id=?", (chapter_id,))
        conn.execute("DELETE FROM users WHERE chapter_id=?", (chapter_id,))
    u = current_user()
    write_log(u["name"], f"Deleted chapter '{ch_name}'", "Super Admin")
    flash("Chapter deleted.", "success")
    return redirect("/admin")


@app.route("/admin/logs")
@login_required
def admin_logs():
    if not is_superadmin():
        return redirect("/admin")
    with get_db() as conn:
        logs = conn.execute(
            "SELECT * FROM log ORDER BY id DESC LIMIT 500"
        ).fetchall()
    return render_template("logs.html", logs=logs, user=current_user())


@app.route("/c/<slug>")
def chapter_dashboard(slug):
    with get_db() as conn:
        chapter = conn.execute("SELECT * FROM chapters WHERE slug=?", (slug,)).fetchone()
    if not chapter:
        return "Chapter not found", 404
    if not chapter["filename"]:
        return render_template("index.html",
            title="No file uploaded yet", chapter_name=chapter["name"],
            members=[], counts={"green": 0, "amber": 0, "red": 0, "gray": 0})
    filepath = CHAPTERS_DIR / chapter["filename"]
    title, members = parse_recommendations(filepath=filepath)
    counts = {"green": 0, "amber": 0, "red": 0, "gray": 0}
    for m in members:
        counts[m["traffic_light"]] += 1
    return render_template("index.html", title=title, chapter_name=chapter["name"],
                           members=members, counts=counts)


# ── Startup ───────────────────────────────────────────────────────────────────

init_db()

# Backfill member_count for chapters uploaded before this column existed
with get_db() as _conn:
    _rows = _conn.execute(
        "SELECT id, filename FROM chapters WHERE filename IS NOT NULL AND member_count = 0"
    ).fetchall()
for _row in _rows:
    try:
        _, _members = parse_recommendations(filepath=CHAPTERS_DIR / _row["filename"])
        with get_db() as _conn:
            _conn.execute("UPDATE chapters SET member_count=? WHERE id=?",
                          (len(_members), _row["id"]))
    except Exception:
        pass

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=True)
