# -*- coding: utf-8 -*-
"""
Goodwood Screening — katalog bahasa untuk comment Word.

Bahasa Indonesia tetap menjadi sumber utama (nilai ``detail`` dan ``message``
di screener.py tidak diubah), sehingga tampilan web tidak terpengaruh sama
sekali. Modul ini hanya menyediakan padanan bahasa Inggris untuk teks yang
benar-benar masuk ke dalam balon comment pada file .docx hasil screening.

Cara pakai:
    from i18n import normalize_lang, t, detail_for
    lang = normalize_lang("en")            # -> "en"
    t("comment.title", lang, detail="...", max_words=15)
"""

SUPPORTED = ("id", "en")

LANG_LABELS = {
    "id": "Bahasa Indonesia",
    "en": "English",
}


def normalize_lang(value):
    """Terima input apa pun dari form, kembalikan 'id' atau 'en'."""
    return "en" if str(value or "").strip().lower().startswith("en") else "id"


# ---------------------------------------------------------------
# Teks comment yang ditulis oleh annotator
# ---------------------------------------------------------------
MESSAGES = {
    "comment.title": {
        "id": ("JUDUL: {detail} Mohon dipersingkat menjadi maksimal {max_words} kata "
               "(termasuk kata sambung)."),
        "en": ("TITLE: {detail} Please shorten the title to a maximum of {max_words} words "
               "(including conjunctions and articles)."),
    },
    "comment.keywords": {
        "id": ("KEYWORDS: {detail} Syarat: {kw_min}-{kw_max} keywords, dipisahkan "
               "titik koma (;)."),
        "en": ("KEYWORDS: {detail} Requirement: {kw_min}-{kw_max} keywords, separated by "
               "semicolons (;)."),
    },
    "comment.keywords_case": {
        "id": ("FORMAT KEYWORDS: {detail} Setiap keyword harus ditulis menggunakan format "
               "Capital Each Word (Title Case), mis. \u201cService Quality\u201d, bukan "
               "\u201cservice quality\u201d."),
        "en": ("KEYWORD FORMAT: {detail} Every keyword must be written in Capital Each Word "
               "(Title Case) format, for example \u201cService Quality\u201d rather than "
               "\u201cservice quality\u201d."),
    },
    "comment.abstract": {
        "id": "ABSTRAK: {detail}",
        "en": "ABSTRACT: {detail}",
    },
    "comment.chapters": {
        "id": ("STRUKTUR BAB: Bab berikut tidak ditemukan dalam manuskrip: {missing}. "
               "Mohon lengkapi sesuai template jurnal (Introduction, Literature Review & "
               "Hypothesis Development, Research Methodology, Result and Discussion, "
               "Conclusion, Acknowledgement, Author Contribution)."),
        "en": ("CHAPTER STRUCTURE: The following chapters were not found in the manuscript: "
               "{missing}. Please complete them in line with the journal template "
               "(Introduction, Literature Review & Hypothesis Development, Research "
               "Methodology, Result and Discussion, Conclusion, Acknowledgement, Author "
               "Contribution)."),
    },
    "comment.ref_count": {
        "id": "JUMLAH REFERENSI: {detail} Syarat minimal {ref_min} referensi.",
        "en": "NUMBER OF REFERENCES: {detail} A minimum of {ref_min} references is required.",
    },
    "comment.ref_composition": {
        "id": ("KOMPOSISI REFERENSI: {detail} Syarat: minimal {pct_min:.0f}% berupa artikel "
               "jurnal internasional. Mohon tinjau dan sesuaikan komposisi referensi pada "
               "daftar Referensi."),
        "en": ("REFERENCE COMPOSITION: {detail} Requirement: at least {pct_min:.0f}% of the "
               "references must be international journal articles. Please review and adjust "
               "the composition of the reference list."),
    },
    "comment.ref_goodwood": {
        "id": ("REFERENSI JURNAL GOODWOOD: {detail} Syarat minimal {gw_min} referensi dari "
               "jurnal terbitan Goodwood Publishing."),
        "en": ("GOODWOOD JOURNAL REFERENCES: {detail} A minimum of {gw_min} references from "
               "journals published by Goodwood Publishing is required."),
    },
    "comment.ref_group_header": {
        "id": "DAFTAR PUSTAKA \u2014 {n} kekurangan terdeteksi:",
        "en": "REFERENCE LIST \u2014 {n} issues detected:",
    },
    "comment.ref_item_header": {
        "id": "REFERENSI No. {no} \u2014 {n} kekurangan terdeteksi:",
        "en": "REFERENCE No. {no} \u2014 {n} issues detected:",
    },
    "comment.word_count": {
        "id": "JUMLAH KATA: {detail}",
        "en": "WORD COUNT: {detail}",
    },
    "comment.methodology": {
        "id": ("METHODOLOGY: Bagian ini belum memiliki satu pun sitasi pendukung. Umumnya "
               "bagian Methodology perlu menyitasi sumber yang mendukung metode, teknik "
               "analisis, instrumen, atau pendekatan penelitian yang digunakan. Mohon "
               "tambahkan sitasi yang relevan."),
        "en": ("METHODOLOGY: This section does not contain any supporting citation. A "
               "Methodology section normally needs to cite sources that support the methods, "
               "analytical techniques, instruments, or research approach being used. Please "
               "add the relevant citations."),
    },
    "comment.uncited_ref": {
        "id": ("REFERENSI TIDAK DISITASI: Referensi No. {no} ini tidak ditemukan disitasi di "
               "badan teks manuskrip. Mohon pastikan referensi ini disitasi minimal satu "
               "kali, atau hapus jika memang tidak relevan/tidak dipakai."),
        "en": ("UNCITED REFERENCE: Reference No. {no} was not found cited anywhere in the body "
               "text of the manuscript. Please make sure this reference is cited at least "
               "once, or remove it if it is not relevant or not actually used."),
    },
    "comment.orphan_citation": {
        "id": ("SITASI TANPA PASANGAN: Sitasi \u201c{surname} ({year})\u201d di badan teks ini "
               "tidak memiliki entri yang sesuai pada daftar pustaka. Mohon tambahkan entri "
               "referensi yang sesuai pada daftar pustaka, atau periksa kembali penulisan "
               "nama penulis/tahun pada sitasi ini."),
        "en": ("UNMATCHED CITATION: The citation \u201c{surname} ({year})\u201d in the body "
               "text has no corresponding entry in the reference list. Please add the "
               "matching reference entry, or re-check the author name and year used in this "
               "citation."),
    },
}


def t(key, lang="id", **kwargs):
    """Ambil teks untuk ``key`` pada bahasa ``lang`` dan isi placeholder-nya."""
    entry = MESSAGES.get(key)
    if entry is None:
        return ""
    template = entry.get(lang) or entry["id"]
    try:
        return template.format(**kwargs)
    except (KeyError, IndexError, ValueError):
        return template


# ---------------------------------------------------------------
# Padanan bahasa Inggris untuk nilai ``detail`` yang ikut masuk comment
# ---------------------------------------------------------------
def _en_title(check):
    title = check.get("title") or ""
    if not title:
        return "Title not detected."
    return '"{}" \u2014 {} words'.format(title, check.get("n_words", 0))


def _en_keywords(check):
    kws = check.get("keywords")
    if kws is None:
        return "The Keywords section was not found."
    return "{} keywords detected: {}".format(len(kws), "; ".join(kws))


def _en_keywords_case(check):
    bad = check.get("bad_keywords") or []
    if not bad:
        return "All keywords already use Capital Each Word (Title Case) format."
    return ("The following keywords do not yet use Capital Each Word (Title Case) format: "
            "{}.".format("; ".join('"{}"'.format(k) for k in bad)))


def _en_abstract(check, abs_min, abs_max):
    n = check.get("word_count")
    if not n:
        return "The Abstract section was not found."
    text = "{} words (requirement: {}-{} words).".format(n, abs_min, abs_max)
    if n < abs_min:
        text += " {} words short.".format(abs_min - n)
    elif n > abs_max:
        text += " {} words over the limit.".format(n - abs_max)
    return text


def _en_ref_count(check):
    n = check.get("n_refs", 0)
    if not n:
        return "The References section was not found or is empty."
    return "{} references detected.".format(n)


def _en_ref_intl(check):
    n_refs = check.get("n_refs", 0)
    if not n_refs:
        return "There are no references to analyse."
    return "{} of {} references ({:.1f}%) were detected as international journal articles.".format(
        check.get("n_intl", 0), n_refs, check.get("pct", 0.0))


def _en_goodwood(check):
    n_refs = check.get("n_refs", 0)
    count = check.get("goodwood_count", 0)
    minimum = check.get("goodwood_min", 5)
    if not n_refs:
        return "There are no references to check."
    if count >= minimum:
        return "{} references from Goodwood Publishing journals were detected (minimum {}).".format(
            count, minimum)
    return ("The article does not yet meet the minimum requirement for references from Goodwood "
            "Publishing journals: only {} of the required {} were detected. Please add "
            "references from journals published by Goodwood Publishing.".format(count, minimum))


def _en_word_count(check, doc_min, doc_max):
    n = check.get("word_count", 0)
    if not n:
        return "Word count could not be computed (Abstract/References not detected)."
    text = "{:,} words (counted from Abstract through References; requirement: {:,}-{:,} words).".format(
        n, doc_min, doc_max)
    if n < doc_min:
        text += " {:,} words short.".format(doc_min - n)
    elif n > doc_max:
        text += " {:,} words over the limit.".format(n - doc_max)
    return text


def detail_for(base_name, check, lang, **limits):
    """
    Kembalikan teks ``detail`` yang siap ditempel ke dalam comment.

    Untuk bahasa Indonesia dipakai apa adanya dari screener (tanpa perubahan
    perilaku sama sekali). Untuk bahasa Inggris teks dibangun ulang dari
    field terstruktur pada dict ``check``.
    """
    if lang != "en":
        return check.get("detail", "")

    if base_name == "Judul":
        return _en_title(check)
    if base_name == "Keywords":
        return _en_keywords(check)
    if base_name == "Format Keywords":
        return _en_keywords_case(check)
    if base_name == "Abstrak":
        return _en_abstract(check, limits.get("abs_min", 180), limits.get("abs_max", 250))
    if base_name == "Jumlah Referensi":
        return _en_ref_count(check)
    if base_name == "Jurnal Internasional":
        return _en_ref_intl(check)
    if base_name == "Referensi Jurnal Goodwood":
        return _en_goodwood(check)
    if base_name.startswith("Jumlah Kata Artikel"):
        return _en_word_count(check, limits.get("doc_min", 5000), limits.get("doc_max", 10000))

    return check.get("detail", "")


def pick(issue, lang):
    """Ambil ``message`` (id) atau ``message_en`` (en) dari sebuah issue dict."""
    if lang == "en":
        return issue.get("message_en") or issue.get("message", "")
    return issue.get("message", "")
