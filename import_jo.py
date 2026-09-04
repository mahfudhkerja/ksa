"""
import_jo.py — lihat catatan di import_printing_2.py.
Import data dari sheet "JO" (Cek JO Akhir) ke tab JO_1 di spreadsheet
tujuan. Konfigurasi (link + sheet yang dicentang) dibaca dari config.json,
diisi lewat halaman Input Data, card "JO".

Dipanggil dari run_all.py (Refresh Semua), sama seperti Printing/Rw/Sl/
Dry/Sf/Ex/Bag.
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "jo"
TARGET_SHEET_NAME = "JO_1"

TARGET_HEADERS = [
    "CEK_JO_AKHIR", "STATUS_JO", "MANUAL", "CUSTOMER", "NO_PO", "JO", "KEMASAN",
    "ORDER", "UK", "UP", "POTONGAN", "BHN", "METER", "ROLL", "EST", "OPP",
    "LAP_1", "LAP_2", "KETERANGAN"
]

HEADER_KEYWORDS = ["CEK_JO_AKHIR", "STATUS_JO"]

if __name__ == "__main__":
    # JO: matikan filter kata "sampah" (KG/JAM/TOTAL/JUMLAH/dst) — kolom
    # BHN/KETERANGAN di sheet JO wajar memuat kata-kata itu sebagai data
    # asli, bukan baris ringkasan, jadi tidak boleh ikut ke-skip.
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS, junk_keywords=[])
