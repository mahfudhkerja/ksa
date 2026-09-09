"""
classify_gudang_sheets.py
==========================
Baca kolom KETERANGAN & STATUS dari sheet "BJB" dan "BJL" (Barang Jadi
Baru / Barang Jadi Lama), klasifikasi tiap baris pakai classify_status(),
lalu tulis hasilnya ke sheet baru "BJB_KATEGORI" / "BJL_KATEGORI" yang
diposisikan TEPAT SETELAH sheet sumbernya.

Kebijakan run ulang: sheet target di-HAPUS TOTAL isinya lalu ditulis ulang
dari nol (bukan cuma di-update sebagian) -- supaya kalau data baru lebih
sedikit barisnya dari data lama, tidak ada baris sisa dari run sebelumnya
yang nyangkut jadi data basi.

Kebutuhan:
  pip install gspread google-auth
  file service account JSON (path diisi di SERVICE_ACCOUNT_FILE di bawah,
  atau lewat environment variable GOOGLE_SERVICE_ACCOUNT_FILE)
"""

import os
import gspread

from classify_status import classify_status

# --------------------------------------------------------------------------
# KONFIGURASI -- sesuaikan bagian ini
# --------------------------------------------------------------------------

SPREADSHEET_ID = "1-ZyKSwXLzZaA6uNYRcpJNQZWX_ssYzvX45Z51xERipI"

SERVICE_ACCOUNT_FILE = os.environ.get(
    "GOOGLE_SERVICE_ACCOUNT_FILE", "credentials.json"
)

# Nama kolom dipakai untuk KLASIFIKASI (dicari berdasarkan HEADER, bukan
# posisi huruf, supaya tetap jalan walau urutan kolom berubah sedikit) --
# fallback ke Q/T kalau header dengan nama ini tidak ketemu persis.
COL_KETERANGAN = "KETERANGAN"  # kolom Q
COL_STATUS = "STATUS"          # kolom T

# Rentang kolom yang DISALIN APA ADANYA ke sheet _KATEGORI (posisi huruf,
# bukan nama header, karena ini soal "kolom A-E dan K-U" secara harfiah).
COPY_RANGES = [("A", "E"), ("K", "U")]

SOURCE_SHEETS = ["BJB", "BJL"]


# --------------------------------------------------------------------------


def _col_letter_to_index(letter):
    """'A' -> 0, 'Q' -> 16, dst. (0-based)"""
    idx = 0
    for ch in letter:
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx - 1


def _find_col_index(header_row, nama_kolom, fallback_letter):
    """Cari index kolom (0-based) berdasarkan nama header persis.
    Kalau tidak ketemu, pakai fallback posisi huruf kolom (mis. 'Q')."""
    for i, h in enumerate(header_row):
        if str(h).strip().upper() == nama_kolom.upper():
            return i
    return _col_letter_to_index(fallback_letter)


def _get_or_create_target_sheet(spreadsheet, source_title):
    """Ambil worksheet '<source>_KATEGORI'; kalau belum ada, buat baru
    tepat di posisi setelah sheet sumbernya."""
    target_title = f"{source_title}_KATEGORI"

    for ws in spreadsheet.worksheets():
        if ws.title == target_title:
            return ws

    all_ws = spreadsheet.worksheets()
    source_index = next(i for i, ws in enumerate(all_ws) if ws.title == source_title)

    return spreadsheet.add_worksheet(
        title=target_title,
        rows=16000,
        cols=len(_copy_indices()) + 1,  # + 1 untuk kolom KATEGORI
        index=source_index + 1,  # tepat setelah sheet sumber
    )


def _copy_indices():
    """Bangun daftar index kolom (0-based) yang perlu disalin, berdasarkan
    COPY_RANGES (mis. A-E dan K-U), urut sesuai urutan asalnya."""
    indices = []
    for start_letter, end_letter in COPY_RANGES:
        start = _col_letter_to_index(start_letter)
        end = _col_letter_to_index(end_letter)
        indices.extend(range(start, end + 1))
    return indices


def classify_sheet(spreadsheet, source_title):
    print(f"--- Memproses sheet '{source_title}' ---")

    source_ws = spreadsheet.worksheet(source_title)
    all_values = source_ws.get_all_values()

    if not all_values:
        print(f"  Sheet '{source_title}' kosong, dilewati.")
        return

    header_row = all_values[0]
    data_rows = all_values[1:]

    idx_ket = _find_col_index(header_row, COL_KETERANGAN, "Q")
    idx_status = _find_col_index(header_row, COL_STATUS, "T")
    copy_idx = _copy_indices()

    print(f"  Kolom disalin (A-E, K-U): {len(copy_idx)} kolom")
    print(f"  Kolom untuk klasifikasi -> KETERANGAN: {idx_ket}, STATUS: {idx_status}")

    def _cell(row, idx):
        return row[idx] if idx < len(row) else ""

    header_out = [_cell(header_row, i) for i in copy_idx] + ["KATEGORI"]
    output_rows = [header_out]

    for row in data_rows:
        keterangan = _cell(row, idx_ket)
        status = _cell(row, idx_status)
        kategori = classify_status(keterangan, status)
        salinan = [_cell(row, i) for i in copy_idx]
        output_rows.append(salinan + [kategori])

    target_ws = _get_or_create_target_sheet(spreadsheet, source_title)

    # Kebijakan run ulang: hapus total isi lama, baru tulis ulang dari nol.
    target_ws.clear()
    target_ws.update(values=output_rows, range_name="A1", value_input_option="RAW")

    print(f"  Selesai: {len(output_rows) - 1} baris ditulis ke '{target_ws.title}'.")


def main():
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    spreadsheet = gc.open_by_key(SPREADSHEET_ID)

    for source_title in SOURCE_SHEETS:
        classify_sheet(spreadsheet, source_title)

    print("Selesai semua.")


if __name__ == "__main__":
    main()
