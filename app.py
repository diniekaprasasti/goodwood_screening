# -*- coding: utf-8 -*-
"""Goodwood Screening System — screening awal manuskrip dengan login dan riwayat."""
import hmac
import os
import sqlite3
import uuid
from datetime import datetime
from functools import wraps
from zoneinfo import ZoneInfo

from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_from_directory,
    session,
    url_for,
)
from werkzeug.utils import secure_filename

from annotator import annotate_docx
from i18n import LANG_LABELS, normalize_lang


BASE = os.path.dirname(os.path.abspath(__file__))

# Untuk Railway, set DATA_DIR=/app/data dan pasang Railway Volume ke /app/data.
DATA_DIR = os.getenv("DATA_DIR", os.path.join(BASE, "data"))
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
EXPORT_DIR = os.path.join(DATA_DIR, "exports")
DB_PATH = os.path.join(DATA_DIR, "screening_history.db")

os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

VERSION = "2.0"
JAKARTA_TZ = ZoneInfo("Asia/Jakarta")

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "change-this-secret-key-for-production")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB
app.config["SESSION_COOKIE_HTTPONLY"] = True
app.config["SESSION_COOKIE_SAMESITE"] = "Lax"
app.config["SESSION_COOKIE_SECURE"] = bool(os.getenv("RAILWAY_ENVIRONMENT"))

LOGIN_USERNAME = os.getenv("LOGIN_USERNAME", "admin")
LOGIN_PASSWORD = os.getenv("LOGIN_PASSWORD", "admin123")


def get_db():
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    with get_db() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS screening_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                created_at TEXT NOT NULL,
                username TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                saved_filename TEXT NOT NULL UNIQUE,
                comments INTEGER NOT NULL DEFAULT 0,
                verdict TEXT,
                passed INTEGER,
                total INTEGER,
                lang TEXT NOT NULL DEFAULT 'id'
            )
            """
        )
        # Migrasi untuk database lama yang dibuat sebelum fitur dwibahasa.
        existing = {row["name"] for row in connection.execute(
            "PRAGMA table_info(screening_history)")}
        if "lang" not in existing:
            connection.execute(
                "ALTER TABLE screening_history ADD COLUMN lang TEXT NOT NULL DEFAULT 'id'")


init_db()


def login_required(view_function):
    @wraps(view_function)
    def wrapped_view(*args, **kwargs):
        if not session.get("logged_in"):
            flash("Silakan login untuk mengakses Goodwood Screening System.")
            return redirect(url_for("login", next=request.path))
        return view_function(*args, **kwargs)

    return wrapped_view


def safe_next_url(value):
    """Hanya izinkan redirect internal yang diawali satu garis miring."""
    if value and value.startswith("/") and not value.startswith("//"):
        return value
    return url_for("index")


def make_result_filename(original_filename):
    safe = secure_filename(original_filename)
    stem, extension = os.path.splitext(safe)
    stem = stem or "manuscript"
    extension = extension.lower() if extension else ".docx"
    timestamp = datetime.now(JAKARTA_TZ).strftime("%Y%m%d_%H%M%S")
    short_id = uuid.uuid4().hex[:6].upper()
    return f"{stem}_SCREENED_{timestamp}_{short_id}{extension}"


@app.route("/login", methods=["GET", "POST"])
def login():
    if session.get("logged_in"):
        return redirect(url_for("index"))

    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        username_ok = hmac.compare_digest(username, LOGIN_USERNAME)
        password_ok = hmac.compare_digest(password, LOGIN_PASSWORD)

        if username_ok and password_ok:
            session.clear()
            session["logged_in"] = True
            session["username"] = username
            session.permanent = True
            return redirect(safe_next_url(request.form.get("next")))

        flash("Username atau password salah.")

    return render_template("login.html", version=VERSION, next=request.args.get("next", ""))


@app.route("/logout")
def logout():
    session.clear()
    flash("Anda telah logout.")
    return redirect(url_for("login"))


@app.route("/", methods=["GET"])
@login_required
def index():
    return render_template(
        "index.html",
        result=None,
        filename=None,
        version=VERSION,
        lang="id",
        lang_labels=LANG_LABELS,
    )


@app.route("/screen", methods=["POST"])
@login_required
def screen():
    uploaded_file = request.files.get("manuscript")
    lang = normalize_lang(request.form.get("lang"))

    if not uploaded_file or not uploaded_file.filename:
        flash("Silakan pilih file Word (.docx) terlebih dahulu.")
        return redirect(url_for("index"))

    if not uploaded_file.filename.lower().endswith(".docx"):
        flash(
            "Format tidak didukung. Sistem hanya menerima file .docx "
            "(jika file Anda .doc, simpan ulang sebagai .docx di Microsoft Word)."
        )
        return redirect(url_for("index"))

    original_filename = secure_filename(uploaded_file.filename)
    upload_name = f"{uuid.uuid4().hex[:8]}_{original_filename}"
    upload_path = os.path.join(UPLOAD_DIR, upload_name)
    uploaded_file.save(upload_path)

    saved_filename = make_result_filename(original_filename)
    output_path = os.path.join(EXPORT_DIR, saved_filename)

    try:
        result, n_comments = annotate_docx(upload_path, output_path, lang=lang)
    except Exception as exc:
        # Hapus file hasil yang mungkin terbentuk sebagian.
        try:
            os.remove(output_path)
        except OSError:
            pass
        flash(
            f"Gagal membaca dokumen: {exc}. "
            "Coba buka file ini di Microsoft Word lalu simpan ulang melalui "
            "File \u2192 Save As \u2192 Word Document (.docx), kemudian unggah kembali."
        )
        return redirect(url_for("index"))
    finally:
        try:
            os.remove(upload_path)
        except OSError:
            pass

    repaired = (result.get("repaired") if isinstance(result, dict) else None) or []
    if repaired:
        flash(
            "Arsip dokumen ini cacat pada {} berkas internal (biasanya gambar) dan sudah "
            "diperbaiki otomatis agar tetap dapat di-screening. Isi teks manuskrip tidak "
            "terpengaruh.".format(len(repaired))
        )

    verdict = str(getattr(result, "verdict", ""))
    passed = int(getattr(result, "passed", 0) or 0)
    total = int(getattr(result, "total", 0) or 0)
    created_at = datetime.now(JAKARTA_TZ).strftime("%d %b %Y, %H:%M WIB")

    with get_db() as connection:
        connection.execute(
            """
            INSERT INTO screening_history (
                created_at, username, original_filename, saved_filename,
                comments, verdict, passed, total, lang
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                created_at,
                session.get("username", "unknown"),
                uploaded_file.filename,
                saved_filename,
                int(n_comments or 0),
                verdict,
                passed,
                total,
                lang,
            ),
        )

    return render_template(
        "index.html",
        result=result,
        filename=uploaded_file.filename,
        download=saved_filename,
        n_comments=n_comments,
        version=VERSION,
        lang=lang,
        lang_labels=LANG_LABELS,
    )


@app.route("/history")
@login_required
def history():
    with get_db() as connection:
        rows = connection.execute(
            """
            SELECT id, created_at, username, original_filename, saved_filename,
                   comments, verdict, passed, total,
                   COALESCE(lang, 'id') AS lang
            FROM screening_history
            ORDER BY id DESC
            """
        ).fetchall()

    return render_template(
        "history.html", records=rows, version=VERSION, lang_labels=LANG_LABELS)


@app.route("/download/<path:fname>")
@login_required
def download(fname):
    safe = secure_filename(fname)
    path = os.path.join(EXPORT_DIR, safe)

    if not os.path.isfile(path):
        flash("File hasil screening tidak ditemukan. Pastikan Railway Volume sudah terpasang.")
        return redirect(url_for("history"))

    return send_from_directory(
        EXPORT_DIR,
        safe,
        as_attachment=True,
        download_name=safe,
    )


if __name__ == "__main__":
    print(f"\n  Goodwood Screening System v{VERSION}")
    print("  Buka browser: http://localhost:5001\n")
    app.run(host="0.0.0.0", port=5001, debug=False)
