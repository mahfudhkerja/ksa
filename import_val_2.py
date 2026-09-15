"""
import_val_2.py — versi KEDUA/tambahan dari import_val.py.
Import data dari file Excel yang DIUPLOAD (bukan link Google Sheet) lewat
kartu "Validasi 2" di halaman Input Data Produksi, ke tab VAL_2 di
spreadsheet Monitor Bahan Baku (beda dari VAL_1 yang ada di spreadsheet
utama -- lihat SOURCE_KEY "val_2" di config.json, field
target_id/target_sheet).

Header & filter SAMA PERSIS dengan import_val.py (TARGET_HEADERS,
HEADER_KEYWORDS, JUNK_KEYWORDS) -- sumbernya diasumsikan format tabel yang
sama, cuma dari file Excel hasil download/export, bukan link spreadsheet
langsung. Kalau ternyata beda, cukup ubah konstanta di sini saja (logic
di import_engine.py -- run_local_excel_import() -- tetap generik).

BEDA dari import_val.py:
  - Sumbernya file upload lokal (python_calamine), bukan link Google
    Sheet -- lihat run_local_excel_import() / VARIAN 5 di import_engine.py.
  - TIDAK dipanggil dari halaman Data Validasi / /api/validasi/refresh-import.
    Dipanggil otomatis lewat /api/gudang/refresh (bareng import Gudang &
    import_form_st_2.py) tiap kali tombol Refresh di halaman Data Gudang
    BJB/BJL diklik, atau manual: `python import_val_2.py`.
"""

from import_engine import run_local_excel_import

SOURCE_KEY = "val_2"
TARGET_SHEET_NAME = "VAL_2"

TARGET_HEADERS = [
    "NO", "TANGGAL", "SHIFT", "CHECK", "AREA", "JO_DIGIT", "JO", "NAMA_PRODUK",
    "JUMLAH", "JUMLAH_MASUK_REWIND", "KETERANGAN"
]

HEADER_KEYWORDS = ["JUMLAH_MASUK_REWIND", "TANGGAL"]

# Sama alasannya dengan import_val.py: bukan sheet mesin produksi, filter
# junk (KG/JAM/TOTAL/JUMLAH dst) malah salah tangkap baris valid --
# dimatikan total.
JUNK_KEYWORDS = []

if __name__ == "__main__":
    print(f"\n=== Import: {SOURCE_KEY} ===")
    try:
        rows_written = run_local_excel_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS,
                                               HEADER_KEYWORDS, junk_keywords=JUNK_KEYWORDS)
        print(f"✅ Selesai: {rows_written} baris ditulis.")
    except Exception as e:
        print(f"❌ Gagal: {e}")
        raise
