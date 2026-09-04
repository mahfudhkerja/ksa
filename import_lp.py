"""
import_lp.py
Import untuk "Laporan Produksi" (LP). Sama seperti import_printing_2.py
dkk: SOURCE_SHEET_ID dan SHEETS_TO_IMPORT tidak ditulis manual di sini,
diambil otomatis dari config.json (diisi lewat halaman Input Data ->
kartu "Laporan Produksi" -> Load + Pilih Sheet).

Beda dari card lain: sheet sumber LP TIDAK pakai filter baris "sampah"
(junk_keywords) sama sekali, karena format tabelnya beda (per-proses,
bukan per-mesin harian) dan bisa wajar memuat kata seperti "TOTAL"/"KG"/
"JAM" di kolom teksnya. Baris cuma dianggap kosong/berhenti kalau
kolom A-F beneran kosong (lihat get_data_rows di import_engine.py).

Fitur lain tetap sama seperti card lain:
  - Sheet diurutkan otomatis per bulan (JAN -> DES), walau urutan tab di
    sumber / urutan centang user tidak berurutan.
  - Cell yang kosong ditulis "-" di sheet target.
"""

from import_engine import run_gsheet_import

SOURCE_KEY = "lp"
TARGET_SHEET_NAME = "LP_1"

TARGET_HEADERS = [
    "SPK/JO1", "Tanggal", "SPK/JO", "Produk", "Customer", "Meter_Order",
    "Urutan_Proses", "Proses", "Mesin", "Tanggal_Proses", "Meter_Awal",
    "Meter_Hasil", "Meter_Waste", "Presentase_Waste", "Indikator",
    "Meter_waste2", "Faktor_Penyebab_Waste", "KLASIFIKASI"
]

HEADER_KEYWORDS = ["SPK/JO", "TANGGAL", "PRODUK", "PROSES"]

if __name__ == "__main__":
    # junk_keywords=[] -> matikan total filter baris "sampah" untuk source ini.
    # header_min_matches=len(HEADER_KEYWORDS) -> baris baru dianggap baris
    # header kalau SEMUA kata kunci di atas ketemu BARENGAN dalam 1 baris
    # yang sama (boleh di kolom mana saja / urutan berapa saja di baris
    # itu, tidak harus posisi tetap) -- bukan cukup 1 kata kunci nyasar
    # sendirian di suatu baris data.
    run_gsheet_import(SOURCE_KEY, TARGET_SHEET_NAME, TARGET_HEADERS, HEADER_KEYWORDS,
                       junk_keywords=[], header_min_matches=len(HEADER_KEYWORDS))
