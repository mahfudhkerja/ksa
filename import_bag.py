"""
import_bag.py — lihat catatan di import_printing_2.py & import_dry_1.py.
Import data dari file Excel "Bag Making" di Google Drive (bukan Google
Sheet langsung — sama seperti Dry 1-5) ke tab BAG_1 di spreadsheet
tujuan. Konfigurasi (link Drive + sheet/tab yang dicentang) dibaca dari
config.json, diisi lewat halaman Input Data, card "Bag Making".

Dipanggil dari run_all.py (Refresh Semua), sama seperti Printing/Rw/Sl/
Dry/Sf/Ex/JO.
"""

from import_engine import run_excel_import

SOURCE_KEY = "bag"
TARGET_SHEET_NAME = "BAG_1"

TARGET_HEADERS = [
    "TANGGAL", "MESIN", "SHIFT", "SPK", "SPK/JO", "PRODUK",
    "UKURAN_BAG_1", "UKURAN_BAG_2",
    "AWAL_ROLL", "AWAL_METER", "AWAL_KG",
    "AKHIR_BAIK", "AKHIR_KW", "AKHIR_JELEK", "AKHIR_TOTAL", "AKHIR_METER", "AKHIR_AREA",
    "BERAT/PACK", "WASTE_KG", "KETERANGAN_WASTE"
]

HEADER_KEYWORDS = ["TANGGAL", "SPK/JO"]

if __name__ == "__main__":
    run_excel_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS)
