# -*- coding: utf-8 -*-
"""Goodwood Screening System — aplikasi lokal untuk screening awal manuskrip."""
import os
import uuid
from flask import Flask, render_template, request, redirect, url_for, flash
from werkzeug.utils import secure_filename
from flask import send_from_directory
from annotator import annotate_docx

BASE = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE, "uploads")
EXPORT_DIR = os.path.join(BASE, "exports")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

VERSION = "1.4"

app = Flask(__name__)
app.secret_key = "goodwood-screening-local"
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", result=None, filename=None, version=VERSION)


@app.route("/screen", methods=["POST"])
def screen():
    f = request.files.get("manuscript")
    if not f or not f.filename:
        flash("Silakan pilih file Word (.docx) terlebih dahulu.")
        return redirect(url_for("index"))
    if not f.filename.lower().endswith(".docx"):
        flash("Format tidak didukung. Sistem hanya menerima file .docx "
              "(jika file Anda .doc, simpan ulang sebagai .docx di Microsoft Word).")
        return redirect(url_for("index"))

    safe = secure_filename(f.filename)
    path = os.path.join(UPLOAD_DIR, "{}_{}".format(uuid.uuid4().hex[:8], safe))
    f.save(path)

    out_name = "SCREENING_{}_{}".format(uuid.uuid4().hex[:6], safe)
    out_path = os.path.join(EXPORT_DIR, out_name)
    try:
        result, n_comments = annotate_docx(path, out_path)
    except Exception as e:
        flash("Gagal membaca dokumen: {}".format(e))
        return redirect(url_for("index"))
    finally:
        try:
            os.remove(path)
        except OSError:
            pass

    return render_template("index.html", result=result, filename=f.filename,
                           download=out_name, n_comments=n_comments, version=VERSION)


@app.route("/download/<path:fname>")
def download(fname):
    safe = secure_filename(fname)
    return send_from_directory(EXPORT_DIR, safe, as_attachment=True,
                               download_name=safe.split("_", 2)[-1].replace(".docx", "") 
                                             + "_SCREENING.docx")


if __name__ == "__main__":
    print("\n  Goodwood Screening System v" + VERSION)
    print("  Buka browser: http://localhost:5001\n")
    app.run(host="0.0.0.0", port=5001, debug=False)
