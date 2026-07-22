# Goodwood Screening System

Sistem lokal untuk screening awal manuskrip (.docx) sebelum masuk proses review.

## Kriteria yang Diperiksa
1. Judul maksimal 15 kata (termasuk kata sambung)
2. Keywords 3–5 buah, setiap keyword menggunakan format Capital Each Word (Title Case)
3. Abstrak 180–250 kata
4. 5 bab utama: Introduction, Literature Review & Hypothesis Development,
   Research Methodology, Result and Discussion, Conclusion
5. Bab khusus: Acknowledgement dan Author Contribution
6. Referensi minimal 30, dengan minimal 80% jurnal internasional
   (klasifikasi otomatis berbasis heuristik — verifikasi akhir tetap oleh editor)
7. Setiap referensi diperiksa satu per satu terhadap format APA 6th Edition:
   nama penulis, tahun publikasi, format judul artikel (sentence case), nama
   jurnal (harus italic), volume/issue, nomor halaman, DOI (ada/valid/URL
   dipakai sebagai pengganti DOI). Jika satu referensi punya lebih dari satu
   kekurangan, SEMUA kekurangan ditampilkan sekaligus (bukan hanya yang
   pertama ditemukan), lengkap dengan rekomendasi perbaikannya.

## Cara Install (sekali saja)
**Windows** : klik dua kali `install_windows.bat`
**Mac/Linux**: `bash install_mac_linux.sh`

## Cara Menjalankan
**Windows** : klik dua kali `start_windows.bat`
**Mac/Linux**: `bash start_mac_linux.sh`

Browser akan terbuka otomatis di http://localhost:5001
Upload file .docx → klik "Jalankan Screening" → hasil muncul langsung.

## Mengubah Kriteria
Semua angka kriteria ada di bagian atas file `screener.py`
(TITLE_MAX_WORDS, ABSTRACT_MIN/MAX, REF_MIN, REF_INTL_PCT, dst.)
sehingga mudah disesuaikan jika kebijakan jurnal berubah.

## Catatan
- Hanya menerima .docx (file .doc harus disimpan ulang sebagai .docx di Word).
- File yang di-upload dihapus otomatis setelah screening (tidak disimpan).

## Output Word + Comment (fitur baru)
Setelah screening, sistem otomatis membuat salinan manuskrip (.docx) yang
sudah berisi comment Word di lokasi temuan:
- Comment ringkasan hasil di awal dokumen (lolos/tidak + daftar kriteria gagal)
- Comment di judul, keywords, dan abstrak jika bermasalah
- Comment daftar bab yang hilang
- Comment di bagian References (jumlah & persentase jurnal internasional)
- Comment per-referensi yang memuat SEMUA kekurangan format APA 6th Edition
  yang terdeteksi pada referensi tersebut (nama penulis, tahun, judul, nama
  jurnal & italic, volume/issue, halaman, DOI), beserta rekomendasi
  perbaikannya — ditempatkan tepat di paragraf referensi yang bersangkutan

Author comment: "Goodwood Screening". Unduh lewat tombol
"Unduh Word + Comment" di halaman hasil — file siap dikirim ke penulis.
File hasil tersimpan di folder `exports/`.
