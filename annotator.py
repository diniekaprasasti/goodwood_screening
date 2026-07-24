# -*- coding: utf-8 -*-
"""
Goodwood Screening — Annotator
Menghasilkan salinan manuskrip (.docx) dengan comment Word di lokasi temuan.

Setiap comment ditempatkan SEDEKAT MUNGKIN dengan lokasi masalah yang
sebenarnya di dalam dokumen (judul, abstrak, keywords, bab tertentu,
referensi tertentu, atau heading tertentu) — bukan ditumpuk semua di awal
dokumen.

Bahasa comment dapat dipilih saat screening ("id" atau "en") melalui
parameter ``lang``. Bahasa Indonesia tetap menjadi perilaku bawaan sehingga
hasil screening lama tetap identik.
"""
from docx import Document

from screener import (
    ABSTRACT_MAX,
    ABSTRACT_MIN,
    DOC_WORDS_MAX,
    DOC_WORDS_MIN,
    GOODWOOD_JOURNAL_MIN,
    KEYWORDS_MAX,
    KEYWORDS_MIN,
    REF_INTL_PCT,
    REF_MIN,
    TITLE_MAX_WORDS,
    screen_document,
)
from i18n import detail_for, normalize_lang, pick, t

AUTHOR = "Goodwood Screening"
INITIALS = "GS"

# Batas ambang yang perlu diketahui i18n saat menyusun detail versi Inggris.
LIMITS = {
    "abs_min": ABSTRACT_MIN,
    "abs_max": ABSTRACT_MAX,
    "doc_min": DOC_WORDS_MIN,
    "doc_max": DOC_WORDS_MAX,
}


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


def annotate_docx(src_path, out_path, lang="id"):
    """Screening + sisipkan comment. Mengembalikan (result, jumlah_comment).

    ``lang`` menentukan bahasa comment: "id" (default) atau "en".
    """
    lang = normalize_lang(lang)

    doc = Document(src_path)
    result, parts, all_paragraphs = screen_document(doc)
    checks = {c["name"].split(" (")[0]: c for c in result["checks"]}
    n_comments = 0

    def comment(runs, text):
        nonlocal n_comments
        if runs:
            doc.add_comment(runs, text=text, author=AUTHOR, initials=INITIALS)
            n_comments += 1

    def detail(base_name, check):
        return detail_for(base_name, check, lang, **LIMITS)

    # ---------- 2. Judul ----------
    c = checks.get("Judul")
    if c and not c["passed"]:
        comment(_runs_at(all_paragraphs, parts["title_idx"]) or _first_runs(all_paragraphs),
                t("comment.title", lang,
                  detail=detail("Judul", c), max_words=TITLE_MAX_WORDS))

    # ---------- 3. Keywords ----------
    c = checks.get("Keywords")
    if c and not c["passed"]:
        anchor = _runs_at(all_paragraphs, parts["keywords_idx"]) or \
                 _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, t("comment.keywords", lang,
                          detail=detail("Keywords", c),
                          kw_min=KEYWORDS_MIN, kw_max=KEYWORDS_MAX))

    # ---------- 3b. Format Keywords (Title Case) ----------
    c = checks.get("Format Keywords")
    if c and not c["passed"]:
        anchor = _runs_at(all_paragraphs, parts["keywords_idx"]) or \
                 _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, t("comment.keywords_case", lang,
                          detail=detail("Format Keywords", c)))

    # ---------- 4. Abstrak ----------
    c = checks.get("Abstrak")
    if c and not c["passed"]:
        anchor = _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, t("comment.abstract", lang, detail=detail("Abstrak", c)))

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
                t("comment.chapters", lang, missing="; ".join(missing)))

    # ---------- 7. Referensi: jumlah + persentase + minimal jurnal Goodwood ----------
    # Digabung jadi SATU comment (bernomor) di heading "References", supaya
    # semua kekurangan pada tingkat keseluruhan daftar pustaka terlihat
    # sekaligus dalam satu tempat.
    c_count = checks.get("Jumlah Referensi")
    c_pct = checks.get("Jurnal Internasional")
    c_goodwood = checks.get("Referensi Jurnal Goodwood")
    ref_anchor = _runs_at(all_paragraphs, parts["ref_heading_idx"])
    msgs = []
    if c_count and not c_count["passed"]:
        msgs.append(t("comment.ref_count", lang,
                      detail=detail("Jumlah Referensi", c_count), ref_min=REF_MIN))
    if c_pct and not c_pct["passed"]:
        # Versi Indonesia memuat kalimat "Klasifikasi ini bersifat estimasi..."
        # yang tidak perlu diulang di dalam comment, jadi dipotong.
        pct_detail = detail("Jurnal Internasional", c_pct)
        if lang == "id":
            pct_detail = pct_detail.split(" Klasifikasi")[0]
        msgs.append(t("comment.ref_composition", lang,
                      detail=pct_detail, pct_min=REF_INTL_PCT))
    if c_goodwood and not c_goodwood["passed"]:
        msgs.append(t("comment.ref_goodwood", lang,
                      detail=detail("Referensi Jurnal Goodwood", c_goodwood),
                      gw_min=c_goodwood.get("goodwood_min", GOODWOOD_JOURNAL_MIN)))
    if msgs:
        if len(msgs) > 1:
            lines = [t("comment.ref_group_header", lang, n=len(msgs))]
            lines += ["{}. {}".format(n, m) for n, m in enumerate(msgs, 1)]
            comment(ref_anchor or _first_runs(all_paragraphs), "\n".join(lines))
        else:
            comment(ref_anchor or _first_runs(all_paragraphs), msgs[0])

    # ---------- 8. Comment per-referensi: SELURUH kekurangan format APA ----------
    c_apa = checks.get("Validasi Format APA per Referensi")
    if c_apa and c_apa.get("ref_apa_rows"):
        for row, idx in zip(c_apa["ref_apa_rows"], parts["ref_idxs"]):
            if row["issues"]:
                lines = [t("comment.ref_item_header", lang,
                           no=row["no"], n=len(row["issues"]))]
                for n, issue in enumerate(row["issues"], 1):
                    lines.append("{}. {}".format(n, pick(issue, lang)))
                comment(_runs_at(all_paragraphs, idx), "\n".join(lines))

    # ---------- 10. Jumlah kata total artikel ----------
    c_words = checks.get("Jumlah Kata Artikel") or next(
        (c for c in result["checks"] if c["name"].startswith("Jumlah Kata Artikel")), None)
    if c_words and not c_words["passed"]:
        anchor = _runs_at(all_paragraphs, parts["abstract_idx"]) or _first_runs(all_paragraphs)
        comment(anchor, t("comment.word_count", lang,
                          detail=detail("Jumlah Kata Artikel", c_words)))

    # ---------- 11. Struktur penomoran Section/Subsection/Sub-subsection ----------
    c_struct = checks.get("Struktur Penomoran Section/Subsection")
    if c_struct and not c_struct["passed"]:
        for issue in c_struct.get("hierarchy_issues", []):
            comment(_runs_at(all_paragraphs, issue["idx"]) or _first_runs(all_paragraphs),
                    pick(issue, lang))

    # ---------- 12. Subsection wajib Bab 5 (Conclusion) ----------
    c_concl = checks.get("Struktur Subsection Bab 5")
    if c_concl and not c_concl["passed"]:
        for issue in c_concl.get("conclusion_issues", []):
            comment(_runs_at(all_paragraphs, issue["idx"]) or _first_runs(all_paragraphs),
                    pick(issue, lang))

    # ---------- 13. Cross-reference Table & Figure ----------
    c_xref = checks.get("Cross-reference Table & Figure")
    if c_xref and not c_xref["passed"]:
        for issue in c_xref.get("crossref_issues", []):
            comment(_runs_at(all_paragraphs, issue["idx"]) or _first_runs(all_paragraphs),
                    pick(issue, lang))

    # ---------- 9. Sitasi pada bagian Methodology ----------
    c_method = checks.get("Sitasi pada Bagian Methodology")
    if c_method and not c_method["passed"] and c_method.get("methodology_result"):
        mr = c_method["methodology_result"]
        comment(_runs_at(all_paragraphs, mr["idx"]) or _first_runs(all_paragraphs),
                t("comment.methodology", lang))

    # ---------- 9b. Kesesuaian sitasi & daftar pustaka (dua arah) ----------
    c_citmatch = checks.get("Kesesuaian Sitasi & Daftar Pustaka")
    if c_citmatch and not c_citmatch["passed"]:
        for row in c_citmatch.get("uncited_refs", []):
            idx = parts["ref_idxs"][row["no"] - 1] if row["no"] - 1 < len(parts["ref_idxs"]) else None
            comment(_runs_at(all_paragraphs, idx) or _first_runs(all_paragraphs),
                    t("comment.uncited_ref", lang, no=row["no"]))
        for orphan in c_citmatch.get("orphan_citations", []):
            comment(_runs_at(all_paragraphs, orphan["idx"]) or _first_runs(all_paragraphs),
                    t("comment.orphan_citation", lang,
                      surname=orphan["surname"].capitalize(), year=orphan["year"]))

    doc.save(out_path)
    return result, n_comments
