"""
import_update_stock.py
=======================
Runner untuk source "Update Stock" (lihat card di halaman Input Data
Produksi). Beda dari import_*.py lain: bukan menyalin tabel dengan
header standar, tapi mengekstrak blok-blok kolom lebar 10 (A-J, L-U,
W-AF, dst -- tanpa header, baris 1 langsung data) dari tiap sheet yang
dicentang, menumpuknya, lalu menulis ke kolom B..K di satu tab tujuan
pada spreadsheet Monitor Bahan Baku (terpisah dari target_sheet_id
utama). Kolom A di tujuan (rumus manual) tidak pernah disentuh.

Dipanggil otomatis oleh run_all.py (tambahkan "import_update_stock.py"
ke SCRIPTS_ORDER di sana), atau manual: `python import_update_stock.py`.
"""

import import_engine as ie

SOURCE_KEY = "update_stock"


def main():
    print(f"\n=== Import: {SOURCE_KEY} ===")
    try:
        rows_written = ie.run_update_stock_import(SOURCE_KEY)
        ie.set_import_result(SOURCE_KEY, "OK", rows_written=rows_written, error=None)
        print(f"✅ Selesai: {rows_written} baris ditulis.")
    except Exception as e:
        ie.set_import_result(SOURCE_KEY, "ERROR", rows_written=None, error=str(e))
        print(f"❌ Gagal: {e}")
        raise


if __name__ == "__main__":
    main()
