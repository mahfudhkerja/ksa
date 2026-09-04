"""
import_val.py — lihat catatan di import_printing_2.py.
Import data dari sheet "VAL" (Validasi Check) ke tab VAL_1 di spreadsheet
tujuan. Konfigurasi (link + sheet yang dicentang) dibaca dari config.json,
diisi lewat halaman Input Data, card "Validasi (VAL)".

BEDA dari script import lain: script ini TIDAK dipanggil dari run_all.py.
Dipanggil khusus dari tombol "Refresh" di sub-page "Data Validasi"
(lihat /api/validasi/refresh-import di app.py), bersama import_form_st.py.

ID spreadsheet sumber (referensi, sudah diisi lewat halaman Input Data):
16iUkDJ_XZs866OhxL-iDKjrVM5KudTYUn6zb8j70IRo
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "val"
TARGET_SHEET_NAME = "VAL_1"

TARGET_HEADERS = [
    "NO", "TANGGAL", "SHIFT", "CHECK", "AREA", "JO_DIGIT", "JO", "NAMA_PRODUK",
    "JUMLAH", "JUMLAH_MASUK_REWIND", "KETERANGAN"
]

HEADER_KEYWORDS = ["JUMLAH_MASUK_REWIND", "TANGGAL"]

# Sheet VAL (Validasi Check) bukan sheet mesin produksi (tidak ada baris
# ringkasan "TOTAL"/"JAM SETTING" dsb di tengah data) -- data tiap baris
# murni transaksi produk. DEFAULT_JUNK_KEYWORDS (KG, JAM, TOTAL, JUMLAH,
# dst) malah salah tangkap baris valid yang NAMA_PRODUK-nya kebetulan
# mengandung kata itu sebagai bagian kata lain (mis. "JAMUR" kena "JAM").
# Sama kasusnya kayak import_form_st.py -- filter junk dimatikan total.
JUNK_KEYWORDS = []

if __name__ == "__main__":
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS,
                       junk_keywords=JUNK_KEYWORDS)
