import re
import os
from pathlib import Path
from flask import Flask, render_template, jsonify, request, redirect
import openpyxl

app = Flask(__name__)

EXCEL_DIR  = Path(__file__).parent
EXCEL_FILE = EXCEL_DIR / "Copy of MTL Recommendations - 2.0.xlsx"

# ── Excel parser ─────────────────────────────────────────────────────────────

def parse_recommendations():
    """
    Read the Recommendations sheet and return:
      title   : str   — sheet title (e.g. "March 2026")
      members : list of {
          name, total_score, max_score, pct, traffic_light,
          metrics: [ {metric, your_score, max_score, pct, status, rec_text, tiers} ]
      }
    """
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
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


# ── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    title, members = parse_recommendations()
    counts = {"green": 0, "amber": 0, "red": 0, "gray": 0}
    for m in members:
        counts[m["traffic_light"]] += 1
    return render_template("index.html", title=title, members=members, counts=counts)


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


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5050, debug=True)
