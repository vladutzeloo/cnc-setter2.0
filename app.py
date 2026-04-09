"""
CNC Setter — Web Application
Flask-based interface for the CMM/NC offset recommendation tool.
"""

import hashlib
import importlib.util
import json
import os
import shutil
import time
from functools import wraps
from pathlib import Path

from flask import (Flask, abort, flash, jsonify, redirect, render_template,
                   request, send_file, session, url_for)
from werkzeug.utils import secure_filename

# ── Load parser module (filename has a space) ────────────────────────────────
_PARSER_PATH = Path(__file__).parent / "generate_report (4).py"
_spec = importlib.util.spec_from_file_location("cnc_parser", _PARSER_PATH)
parser = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(parser)

# ── Flask app setup ───────────────────────────────────────────────────────────
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "cncsetter-change-in-prod-xyz987")

BASE_DIR  = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
DB_PATH   = BASE_DIR / "cncsetter.db"
UPLOAD_DIR.mkdir(exist_ok=True)

ALLOWED_NC  = {".nc", ".cnc", ".mpf", ".h", ".prg"}
ALLOWED_PDF = {".pdf"}

# ── Database ──────────────────────────────────────────────────────────────────

def get_db():
    import sqlite3
    db = sqlite3.connect(DB_PATH)
    db.row_factory = sqlite3.Row
    return db


def init_db():
    db = get_db()
    db.executescript("""
        CREATE TABLE IF NOT EXISTS users (
            id            INTEGER PRIMARY KEY,
            username      TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            created_at    TEXT DEFAULT CURRENT_TIMESTAMP
        );
        CREATE TABLE IF NOT EXISTS reports (
            id               INTEGER PRIMARY KEY,
            user_id          INTEGER,
            part_name        TEXT,
            cmm_filename     TEXT,
            nc_filenames     TEXT,
            drawing_filename TEXT,
            report_json      TEXT,
            created_at       TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (user_id) REFERENCES users(id)
        );
        CREATE TABLE IF NOT EXISTS tool_overrides (
            id            INTEGER PRIMARY KEY,
            report_id     INTEGER,
            feature_key   TEXT,
            tool          TEXT,
            tool_desc     TEXT,
            tool_type     TEXT,
            FOREIGN KEY (report_id) REFERENCES reports(id)
        );
    """)
    db.commit()
    db.close()


def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()

# ── Auth helpers ──────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if "user_id" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated


# ── Auth routes ───────────────────────────────────────────────────────────────

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute(
            "SELECT * FROM users WHERE username = ?", (username,)
        ).fetchone()
        db.close()
        if user and user["password_hash"] == hash_pw(password):
            session["user_id"]  = user["id"]
            session["username"] = user["username"]
            return redirect(url_for("dashboard"))
        flash("Invalid username or password.", "danger")
    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    db = get_db()
    user_count = db.execute("SELECT COUNT(*) FROM users").fetchone()[0]
    db.close()
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        confirm  = request.form.get("confirm",  "")
        if not username or not password:
            flash("Username and password are required.", "danger")
        elif password != confirm:
            flash("Passwords do not match.", "danger")
        elif len(password) < 6:
            flash("Password must be at least 6 characters.", "danger")
        else:
            try:
                db = get_db()
                db.execute(
                    "INSERT INTO users (username, password_hash) VALUES (?, ?)",
                    (username, hash_pw(password))
                )
                db.commit()
                db.close()
                flash("Account created — please log in.", "success")
                return redirect(url_for("login"))
            except Exception:
                flash("Username already taken.", "danger")
    return render_template("register.html", first_user=(user_count == 0))


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


# ── Dashboard ─────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def dashboard():
    db = get_db()
    reports = db.execute(
        "SELECT id, part_name, cmm_filename, nc_filenames, created_at, "
        "drawing_filename, "
        "json_extract(report_json,'$.summary.total_features') AS total, "
        "json_extract(report_json,'$.summary.oot_count')      AS oot, "
        "json_extract(report_json,'$.summary.high_count')     AS high, "
        "json_extract(report_json,'$.summary.medium_count')   AS medium, "
        "json_extract(report_json,'$.summary.low_count')      AS low "
        "FROM reports WHERE user_id = ? ORDER BY id DESC",
        (session["user_id"],)
    ).fetchall()
    db.close()
    return render_template("dashboard.html", reports=reports)


# ── Upload & process ──────────────────────────────────────────────────────────

@app.route("/upload", methods=["GET", "POST"])
@login_required
def upload():
    if request.method == "GET":
        return render_template("upload.html")

    part_name   = request.form.get("part_name", "").strip() or "Unknown Part"
    cmm_file    = request.files.get("cmm_pdf")
    nc_files    = request.files.getlist("nc_files")
    drawing_file = request.files.get("drawing_pdf")

    # Basic validation
    if not cmm_file or not cmm_file.filename:
        flash("CMM PDF is required.", "danger")
        return render_template("upload.html")
    if not nc_files or not any(f.filename for f in nc_files):
        flash("At least one NC file is required.", "danger")
        return render_template("upload.html")

    # Create a temp directory for this upload
    ts  = int(time.time())
    job_dir = UPLOAD_DIR / f"job_{ts}"
    job_dir.mkdir(parents=True)

    try:
        # Save CMM PDF
        cmm_name = secure_filename(cmm_file.filename)
        cmm_path = job_dir / cmm_name
        cmm_file.save(str(cmm_path))

        # Save NC files
        nc_paths = []
        nc_names = []
        for f in nc_files:
            if not f.filename:
                continue
            fn  = secure_filename(f.filename)
            fp  = job_dir / fn
            f.save(str(fp))
            nc_paths.append(str(fp))
            nc_names.append(fn)

        # Save drawing PDF (optional)
        drawing_name = None
        if drawing_file and drawing_file.filename:
            drawing_name = secure_filename(drawing_file.filename)
            drawing_file.save(str(job_dir / drawing_name))

        # ── Run parser ───────────────────────────────────────────────────────
        t0 = time.time()
        cmm_features = parser.parse_cmm_pdf(str(cmm_path), verbose=False)

        all_holes, all_ops = [], []
        for nc_path in nc_paths:
            holes, ops = parser.parse_nc_file(nc_path)
            all_holes.extend(holes)
            all_ops.extend(ops)

        recs = parser.correlate(cmm_features, all_holes)
        elapsed = round(time.time() - t0, 1)

        # ── Build report JSON ────────────────────────────────────────────────
        oot_count     = sum(1 for f in cmm_features if f.out_of_tol and not f.suspect)
        suspect_count = sum(1 for f in cmm_features if f.suspect)
        high_count    = sum(1 for r in recs if r.confidence == "HIGH"   and r.feature.out_of_tol)
        medium_count  = sum(1 for r in recs if r.confidence == "MEDIUM" and r.feature.out_of_tol)
        low_count     = sum(1 for r in recs if r.confidence == "LOW"    and r.feature.out_of_tol)

        # Build unique tool list (for override dropdowns)
        seen_tools = {}
        for h in all_holes:
            if h.tool and h.tool not in seen_tools:
                seen_tools[h.tool] = {
                    "tool": h.tool, "desc": h.tool_desc or "",
                    "type": h.tool_type, "dia": h.tool_dia,
                }
        tools_list = sorted(seen_tools.values(), key=lambda t: (t["type"], t["dia"]))

        report_data = {
            "part_name":  part_name,
            "cmm_filename":  cmm_name,
            "nc_filenames":  nc_names,
            "drawing_filename": drawing_name,
            "generated_at":  time.strftime("%Y-%m-%d %H:%M"),
            "elapsed_s":     elapsed,
            "summary": {
                "total_features": len(cmm_features),
                "oot_count":     oot_count,
                "suspect_count": suspect_count,
                "high_count":    high_count,
                "medium_count":  medium_count,
                "low_count":     low_count,
            },
            "recommendations": [r.to_dict() for r in recs],
            "all_features":    [f.to_dict() for f in cmm_features
                                if abs(f.dev) < 0.0001 and not f.suspect],
            "tools": tools_list,
        }

        report_json = json.dumps(report_data)

        # ── Save to DB ───────────────────────────────────────────────────────
        db = get_db()
        cur = db.execute(
            "INSERT INTO reports (user_id, part_name, cmm_filename, nc_filenames, "
            "drawing_filename, report_json) VALUES (?, ?, ?, ?, ?, ?)",
            (session["user_id"], part_name, cmm_name,
             json.dumps(nc_names), drawing_name, report_json)
        )
        report_id = cur.lastrowid
        db.commit()
        db.close()

        # Move job dir to named folder
        final_dir = UPLOAD_DIR / f"report_{report_id}"
        shutil.move(str(job_dir), str(final_dir))

        flash(f"Report generated in {elapsed}s — {len(recs)} recommendations.", "success")
        return redirect(url_for("view_report", report_id=report_id))

    except Exception as e:
        shutil.rmtree(str(job_dir), ignore_errors=True)
        flash(f"Processing error: {e}", "danger")
        return render_template("upload.html")


# ── View report ───────────────────────────────────────────────────────────────

@app.route("/report/<int:report_id>")
@login_required
def view_report(report_id):
    db = get_db()
    row = db.execute(
        "SELECT * FROM reports WHERE id = ? AND user_id = ?",
        (report_id, session["user_id"])
    ).fetchone()
    if not row:
        abort(404)

    overrides = db.execute(
        "SELECT feature_key, tool, tool_desc, tool_type FROM tool_overrides WHERE report_id = ?",
        (report_id,)
    ).fetchall()
    db.close()

    report = json.loads(row["report_json"])

    # Apply overrides
    override_map = {o["feature_key"]: dict(o) for o in overrides}
    for rec in report["recommendations"]:
        key = _feature_key(rec["feature"])
        if key in override_map:
            ov = override_map[key]
            rec["tool"]        = ov["tool"]
            rec["tool_desc"]   = ov["tool_desc"]
            rec["confidence"]  = "MEDIUM"
            rec["match_reason"] = f"Manual override → T{ov['tool']}"

    return render_template("report.html",
                           report=report,
                           report_id=report_id,
                           part_name=row["part_name"],
                           has_drawing=bool(row["drawing_filename"]))


def _feature_key(feat_dict):
    return f"{feat_dict['name']}:{feat_dict['axis']}:{feat_dict['nominal']}"


# ── API: tool override ────────────────────────────────────────────────────────

@app.route("/api/report/<int:report_id>/override", methods=["POST"])
@login_required
def api_override(report_id):
    db = get_db()
    row = db.execute(
        "SELECT id FROM reports WHERE id = ? AND user_id = ?",
        (report_id, session["user_id"])
    ).fetchone()
    if not row:
        db.close()
        return jsonify({"error": "not found"}), 404

    data        = request.get_json()
    feature_key = data.get("feature_key")
    tool_num    = data.get("tool")
    tool_desc   = data.get("tool_desc", "")
    tool_type   = data.get("tool_type", "")

    if not feature_key or not tool_num:
        db.close()
        return jsonify({"error": "missing fields"}), 400

    # Upsert override
    db.execute(
        "DELETE FROM tool_overrides WHERE report_id = ? AND feature_key = ?",
        (report_id, feature_key)
    )
    db.execute(
        "INSERT INTO tool_overrides (report_id, feature_key, tool, tool_desc, tool_type) "
        "VALUES (?, ?, ?, ?, ?)",
        (report_id, feature_key, tool_num, tool_desc, tool_type)
    )
    db.commit()
    db.close()
    return jsonify({"ok": True})


# ── API: update part name ─────────────────────────────────────────────────────

@app.route("/api/report/<int:report_id>/name", methods=["POST"])
@login_required
def api_update_name(report_id):
    db = get_db()
    row = db.execute(
        "SELECT id FROM reports WHERE id = ? AND user_id = ?",
        (report_id, session["user_id"])
    ).fetchone()
    if not row:
        db.close()
        return jsonify({"error": "not found"}), 404

    new_name = (request.get_json() or {}).get("name", "").strip()
    if not new_name:
        db.close()
        return jsonify({"error": "empty name"}), 400

    db.execute("UPDATE reports SET part_name = ? WHERE id = ?", (new_name, report_id))
    db.commit()
    db.close()
    return jsonify({"ok": True, "name": new_name})


# ── Serve drawing PDF ─────────────────────────────────────────────────────────

@app.route("/report/<int:report_id>/drawing")
@login_required
def serve_drawing(report_id):
    db = get_db()
    row = db.execute(
        "SELECT drawing_filename FROM reports WHERE id = ? AND user_id = ?",
        (report_id, session["user_id"])
    ).fetchone()
    db.close()
    if not row or not row["drawing_filename"]:
        abort(404)
    path = UPLOAD_DIR / f"report_{report_id}" / row["drawing_filename"]
    if not path.exists():
        abort(404)
    return send_file(str(path), mimetype="application/pdf")


# ── Delete report ─────────────────────────────────────────────────────────────

@app.route("/report/<int:report_id>/delete", methods=["POST"])
@login_required
def delete_report(report_id):
    db = get_db()
    row = db.execute(
        "SELECT id FROM reports WHERE id = ? AND user_id = ?",
        (report_id, session["user_id"])
    ).fetchone()
    if row:
        db.execute("DELETE FROM tool_overrides WHERE report_id = ?", (report_id,))
        db.execute("DELETE FROM reports WHERE id = ?",               (report_id,))
        db.commit()
        shutil.rmtree(str(UPLOAD_DIR / f"report_{report_id}"), ignore_errors=True)
    db.close()
    flash("Report deleted.", "info")
    return redirect(url_for("dashboard"))


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    init_db()
    print("\n  CNC Setter Web — http://127.0.0.1:5000\n")
    app.run(debug=True, port=5000, host="0.0.0.0")
