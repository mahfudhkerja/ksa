"""
import_sj.py — Import data Surat Jalan (History Pengiriman).

Sumber: spreadsheet multi-sheet (bulan), di-load & pilih sheet dari
halaman Input Data Produksi (config.json key "surat_jalan").

Tujuan: sheet SURAT_JALAN di spreadsheet
1h720fO5xiBdTCCwmsxmImSlEa4472HU-0rNlztPI_8E.

Alur sama persis dengan import_rw.py / import_sl.py:
  run_gsheet_import() → baca source → hapus sheet lama → tulis ulang.
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "surat_jalan"
TARGET_SHEET_NAME = "SURAT_JALAN"

# Spreadsheet tujuan BUKAN warehouse utama (target_sheet_id global),
# tapi spreadsheet khusus History Pengiriman.
TARGET_SPREADSHEET_ID = "1h720fO5xiBdTCCwmsxmImSlEa4472HU-0rNlztPI_8E"

TARGET_HEADERS = [
    "No_SJ",
    "Tanggal",
    "Tanggal_SJ",
    "Nama_Customer",
    "Product",
    "No_JO",
    "JO_Master",
    "Nomor_kendaraan",
    "Ukuran",
    "Unit",
    "Kg/Unit",
    "Status",
    "Keterangan_Internal",
    "Keterangan",
]

HEADER_KEYWORDS = ["No_SJ", "Tanggal", "Nama_Customer", "Product", "No_JO"]

if __name__ == "__main__":
    run_gsheet_import(
        SOURCE_KEY,
        TARGET_SHEET_NAME,
        TARGET_HEADERS,
        HEADER_KEYWORDS,
        target_id=TARGET_SPREADSHEET_ID,
    )
