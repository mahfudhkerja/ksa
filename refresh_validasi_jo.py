"""
refresh_validasi_jo.py

Dipanggil sebagai refresh ke-3 (bareng import_val.py & import_form_st.py)
dari tombol Refresh di halaman Data Validasi.

BEDA dari script import_*.py lain: ini TIDAK baca config.json "sources"
(bukan import dari spreadsheet luar). Yang dibaca cuma 2 tab yang sudah
ada di spreadsheet TUJUAN sendiri -- SL_1 (kolom TANGGAL & JO) dan JO_1
(buat cari NAMA produk) -- lalu ditulis ke tab "Validasi" (kolom
TANGGAL, JO, NAMA saja; kolom lain tidak disentuh).

Lihat sync_validasi_header() di import_engine.py untuk detail alurnya.
"""

from import_engine import sync_validasi_header

if __name__ == "__main__":
    sync_validasi_header()
