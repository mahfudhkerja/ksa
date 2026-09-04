"""
import_form_st.py — lihat catatan di import_printing_2.py.
Import data dari sheet "Form Serah Terima" (ST) ke tab FORM_ST_1 di
spreadsheet tujuan. Konfigurasi (link + sheet yang dicentang) dibaca dari
config.json, diisi lewat halaman Input Data, card "Form Serah Terima (ST)".

BEDA dari script import lain: script ini TIDAK dipanggil dari run_all.py.
Dipanggil khusus dari tombol "Refresh" di sub-page "Data Validasi"
(lihat /api/validasi/refresh-import di app.py), bersama import_val.py.

Catatan header: di kepala tabel sumber ada beberapa kolom kosong di antara
"BERAT/KG" -> "STATUS" dan setelah "DARI_SLITTING" — diasumsikan itu
kolom kosong/pemisah di sheet aslinya, jadi tidak dimasukkan ke
TARGET_HEADERS. Kalau ternyata ada nama kolom di situ, tambahkan ke
TARGET_HEADERS sesuai urutan aslinya.

ID spreadsheet sumber (referensi, sudah diisi lewat halaman Input Data):
1a-s1YkDObIXIP2a2D6QJxRp7oOwTRDzcej-NChjiHfA
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "form_st"
TARGET_SHEET_NAME = "FORM_ST_1"

TARGET_HEADERS = [
    "TANGGAL", "SHIFT", "JO_DIGIT", "JO", "NAMA_PRODUK", "JUMLAH_MASUK_GBJ",
    "BERAT/KG", "STATUS", "MASUK_REWIND", "HASIL_RIWEN", "DARI_SLITTING"
]

HEADER_KEYWORDS = ["JUMLAH_MASUK_GBJ", "TANGGAL"]

# Form Serah Terima bukan sheet mesin produksi (tidak ada baris ringkasan
# "TOTAL"/"JAM SETTING" dsb di tengah data) -- data tiap baris murni
# transaksi produk. DEFAULT_JUNK_KEYWORDS (KG, JAM, TOTAL, JUMLAH, dst)
# malah salah tangkap baris valid yang NAMA_PRODUK atau kolom lainnya
# kebetulan mengandung kata itu sebagai bagian kata lain (mis. "JAMUR"
# kena "JAM"). Makanya filter junk dimatikan total di sini.
JUNK_KEYWORDS = []

if __name__ == "__main__":
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS,
                       junk_keywords=JUNK_KEYWORDS)
