# -*- coding: utf-8 -*-
"""
Goodwood Screening — Annotator
Menghasilkan salinan manuskrip (.docx) dengan comment Word di lokasi temuan.

Setiap comment ditempatkan SEDEKAT MUNGKIN dengan lokasi masalah yang
sebenarnya di dalam dokumen (judul, abstrak, keywords, bab tertentu,
referensi tertentu, atau heading tertentu) — bukan ditumpuk semua di awal
dokumen. Comment ringkasan tetap ditaruh di awal sebagai overview, tapi
rincian tiap temuan selalu dianchor ke paragraf terkait.
"""
from docx import Document
from screener import screen_document

AUTHOR = "Goodwood Screening"
INITIALS = "GS"


def _runs_at(paragraphs, idx):
    """Ambil runs pada paragraf ke-idx (dari list `paragraphs`, hasil get_all_paragraphs);
    fallback ke paragraf non-kosong terdekat."""
    if idx is not None and 0 <= idx < len(paragraphs) and paragraphs[idx].runs:
        return paragraphs[idx].runs
    return None


def _first_runs(paragraphs):
    for p in paragraphs:
        if p.runs and p.text.strip():
            return p.runs
    return None


def annotate_docx(src_path, out_path):
    """Screening + sisipkan comment. Mengembalikan (result, jumlah_comment)."""
    doc = Document(src_path)
    result, parts, all_paragraphs = screen_document(doc)
    checks = {c["name"].split(" (")[0]: c for c in result["checks"]}
    n_comments = 0

    def comment(runs, text):
        nonlocal n_comments
        if runs:
            doc.add_comment(runs, text=text, author=AUTHOR, initials=INITIALS)
            n_comments += 1

    # ---------- 2. Judul ----------
    c = checks.get("Judul")
    if c and not c["passed"]:
        comment(_runs_at(all_paragraphs, parts["title_idx"]) or _first_runs(all_paragraphs),
                "JUDUL: {} Mohon dipersingkat menjadi maksimal 15 kata "
                "(termasuk kata sambung).".format(c["detail"]))

    # ---------- 3. Keywords ----------
    c = checks.get("Keywords")
    if c and not c["passed"]:
        anchor = _runs_at(all_paragraphs, parts["keywords_idx"]) or \
                 _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, "KEYWORDS: {} Syarat: 3-5 keywords, dipisahkan "
                "titik koma (;).".format(c["detail"]))

    # ---------- 3b. Format Keywords (Title Case) ----------
    c = checks.get("Format Keywords")
    if c and not c["passed"]:
        anchor = _runs_at(all_paragraphs, parts["keywords_idx"]) or \
                 _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, "FORMAT KEYWORDS: {} Setiap keyword harus ditulis menggunakan format "
                "Capital Each Word (Title Case), mis. \u201cService Quality\u201d, bukan "
                "\u201cservice quality\u201d.".format(c["detail"]))

    # ---------- 4. Abstrak ----------
    c = checks.get("Abstrak")
    if c and not c["passed"]:
        anchor = _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, "ABSTRAK: {}".format(c["detail"]))

    # ---------- 5 & 6. Bab utama + bab khusus (yang hilang) ----------
    missing = []
    for key in ("Struktur 5 Bab Utama", "Bab Khusus"):
        c = checks.get(key)
        if c and not c["passed"]:
            missing += [ch["label"] for ch in c["chapters"] if not ch["found"]]
    if missing:
        # anchor di heading pertama yang terdeteksi agar dekat dengan struktur bab
        anchor = None
        if parts["heading_idxs"]:
            anchor = _runs_at(all_paragraphs, parts["heading_idxs"][0])
        comment(anchor or _first_runs(all_paragraphs),
                "STRUKTUR BAB: Bab berikut tidak ditemukan dalam manuskrip: "
                + "; ".join(missing) + ". Mohon lengkapi sesuai template jurnal "
                "(Introduction, Literature Review & Hypothesis Development, "
                "Research Methodology, Result and Discussion, Conclusion, "
                "Acknowledgement, Author Contribution).")

    # ---------- 7. Referensi: jumlah + persentase + minimal jurnal Goodwood ----------
    # Digabung jadi SATU comment (bernomor) di heading "References", supaya
    # semua kekurangan pada tingkat keseluruhan daftar pustaka terlihat
    # sekaligus dalam satu tempat, bukan tersebar di beberapa comment balon
    # terpisah yang membuat salah satu temuan mudah terlewat.
    c_count = checks.get("Jumlah Referensi")
    c_pct = checks.get("Jurnal Internasional")
    c_goodwood = checks.get("Referensi Jurnal Goodwood")
    ref_anchor = _runs_at(all_paragraphs, parts["ref_heading_idx"])
    msgs = []
    if c_count and not c_count["passed"]:
        msgs.append("JUMLAH REFERENSI: {} Syarat minimal 30 referensi.".format(c_count["detail"]))
    if c_pct and not c_pct["passed"]:
        msgs.append("KOMPOSISI REFERENSI: {}".format(c_pct["detail"].split(" Klasifikasi")[0]) +
                    " Syarat: minimal 80% berupa artikel jurnal internasional. "
                    "Mohon tinjau dan sesuaikan komposisi referensi pada daftar Referensi.")
    if c_goodwood and not c_goodwood["passed"]:
        msgs.append("REFERENSI JURNAL GOODWOOD: {} Syarat minimal {} referensi dari jurnal "
                    "terbitan Goodwood Publishing.".format(
                        c_goodwood["detail"], c_goodwood.get("goodwood_min", 5)))
    if msgs:
        if len(msgs) > 1:
            lines = ["DAFTAR PUSTAKA \u2014 {} kekurangan terdeteksi:".format(len(msgs))]
            lines += ["{}. {}".format(n, m) for n, m in enumerate(msgs, 1)]
            comment(ref_anchor or _first_runs(all_paragraphs), "\n".join(lines))
        else:
            comment(ref_anchor or _first_runs(all_paragraphs), msgs[0])

    # ---------- 8. Comment per-referensi: SELURUH kekurangan format APA ----------
    # Setiap referensi yang punya kekurangan diberi comment tersendiri, tepat
    # di paragraf referensi tersebut, memuat SEMUA kekurangan yang terdeteksi
    # (bukan hanya yang pertama) beserta rekomendasi perbaikannya — sehingga
    # penulis/editor tidak perlu mencari sendiri letak kesalahannya.
    c_apa = checks.get("Validasi Format APA per Referensi")
    if c_apa and c_apa.get("ref_apa_rows"):
        for row, idx in zip(c_apa["ref_apa_rows"], parts["ref_idxs"]):
            if row["issues"]:
                lines = ["REFERENSI No. {} — {} kekurangan terdeteksi:".format(
                    row["no"], len(row["issues"]))]
                for n, issue in enumerate(row["issues"], 1):
                    lines.append("{}. {}".format(n, issue["message"]))
                comment(_runs_at(all_paragraphs, idx), "\n".join(lines))

    # ---------- 10. Jumlah kata total artikel ----------
    c_words = checks.get("Jumlah Kata Artikel") or next(
        (c for c in result["checks"] if c["name"].startswith("Jumlah Kata Artikel")), None)
    if c_words and not c_words["passed"]:
        anchor = _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, "JUMLAH KATA: {}".format(c_words["detail"]))

    # ---------- 11. Struktur penomoran Section/Subsection/Sub-subsection ----------
    c_struct = checks.get("Struktur Penomoran Section/Subsection")
    if c_struct and not c_struct["passed"]:
        for issue in c_struct.get("hierarchy_issues", []):
            comment(_runs_at(all_paragraphs, issue["idx"]) or _first_runs(all_paragraphs),
                    issue["message"])

    # ---------- 12. Subsection wajib Bab 5 (Conclusion) ----------
    c_concl = checks.get("Struktur Subsection Bab 5")
    if c_concl and not c_concl["passed"]:
        for issue in c_concl.get("conclusion_issues", []):
            comment(_runs_at(all_paragraphs, issue["idx"]) or _first_runs(all_paragraphs),
                    issue["message"])

    # ---------- 13. Cross-reference Table & Figure ----------
    c_xref = checks.get("Cross-reference Table & Figure")
    if c_xref and not c_xref["passed"]:
        for issue in c_xref.get("crossref_issues", []):
            comment(_runs_at(all_paragraphs, issue["idx"]) or _first_runs(all_paragraphs),
                    issue["message"])

    # ---------- 9. Sitasi pada bagian Methodology ----------
    c_method = checks.get("Sitasi pada Bagian Methodology")
    if c_method and not c_method["passed"] and c_method.get("methodology_result"):
        mr = c_method["methodology_result"]
        comment(_runs_at(all_paragraphs, mr["idx"]) or _first_runs(all_paragraphs),
                "METHODOLOGY: Bagian ini belum memiliki satu pun sitasi pendukung. Umumnya "
                "bagian Methodology perlu menyitasi sumber yang mendukung metode, teknik "
                "analisis, instrumen, atau pendekatan penelitian yang digunakan. Mohon "
                "tambahkan sitasi yang relevan.")

    # ---------- 9b. Kesesuaian sitasi & daftar pustaka (dua arah) ----------
    c_citmatch = checks.get("Kesesuaian Sitasi & Daftar Pustaka")
    if c_citmatch and not c_citmatch["passed"]:
        for row in c_citmatch.get("uncited_refs", []):
            idx = parts["ref_idxs"][row["no"] - 1] if row["no"] - 1 < len(parts["ref_idxs"]) else None
            comment(_runs_at(all_paragraphs, idx) or _first_runs(all_paragraphs),
                    "REFERENSI TIDAK DISITASI: Referensi No. {} ini tidak ditemukan disitasi "
                    "di badan teks manuskrip. Mohon pastikan referensi ini disitasi minimal "
                    "satu kali, atau hapus jika memang tidak relevan/tidak dipakai.".format(row["no"]))
        for orphan in c_citmatch.get("orphan_citations", []):
            comment(_runs_at(all_paragraphs, orphan["idx"]) or _first_runs(all_paragraphs),
                    "SITASI TANPA PASANGAN: Sitasi \u201c{} ({})\u201d di badan teks ini tidak "
                    "memiliki entri yang sesuai pada daftar pustaka. Mohon tambahkan entri "
                    "referensi yang sesuai pada daftar pustaka, atau periksa kembali penulisan "
                    "nama penulis/tahun pada sitasi ini.".format(
                        orphan["surname"].capitalize(), orphan["year"]))

    doc.save(out_path)
    return result, n_comments
