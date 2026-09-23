"""
KSA System - Backend API
Menggantikan google.script.run (Apps Script) dengan Flask + gspread,
supaya frontend (index.html) bisa baca-tulis ke Google Sheets lewat
REST API biasa.

Struktur Spreadsheet yang diharapkan (1 spreadsheet, beberapa sheet/tab).
Nama TAB harus persis (huruf besar/kecil ikut dicek), nama KOLOM di baris 1 juga harus persis:

  - Tab "Login"
      USERNAME | PASSWORD | NAMA | ROLE

  - Tab "Validasi"
      TANGGAL | JO | NAMA | ORDER | HASIL SLITTING | HASIL SLIT(QTY) |
      HASIL BAG | VALIDASI | FORM SERAH TERIMA | TOTAL | SELISIH | STATUS | POTONGAN

  - Tab "UpdateStock"
      JO | NAMA | ORDER | METER ORDER | METER VALIDASI |
      LAPISAN ORDER | LAPISAN VALIDASI | ACC

  - Tab "StockBahan"
      TANGGAL | USER | JO | NAMA | ORDER | METER ORDER | METER VALIDASI |
      LAPISAN ORDER | LAPISAN VALIDASI

  - Tab "PIC"
      NAMA | NOMOR

Kalau header di sheet kamu beda, cukup ubah nilai di *_COLUMN_MAP di bawah
(bagian kiri = nama kolom asli di sheet, bagian kanan = nama field yang
dipakai kode/frontend, jangan diubah bagian kanannya).

--------------------------------------------------------------------------
BAGIAN "INPUT DATA PRODUKSI" (baru)
--------------------------------------------------------------------------
Endpoint /api/produksi/* di bawah menggantikan cara lama isi
SOURCE_SHEET_ID / SHEETS_TO_IMPORT manual di tiap import_*.py. Sekarang:

  1. User paste link spreadsheet di kartu source (mis. "Printing 2") lalu
     klik Load -> /api/produksi/load -> deteksi ID + nama semua sheet/tab.
  2. User klik "Pilih Sheet" -> centang beberapa sheet dari hasil deteksi
     -> /api/produksi/sheets -> disimpan ke config.json.
  3. User klik satu tombol "Refresh Semua" -> /api/produksi/run-all ->
     menjalankan run_all.py di background thread (semua script import
     baca config.json sendiri-sendiri) -> frontend polling
     /api/produksi/run-status untuk lihat progress live.

Semua penyimpanan konfigurasi ada di config.json (lihat import_engine.py).

--------------------------------------------------------------------------
BAGIAN "FSTL — LAMPIRAN WASTE" (baru)
--------------------------------------------------------------------------
Endpoint /api/fstl/* pakai spreadsheet TERPISAH (FSTL_SPREADSHEET_ID, lihat
env var / default di bawah), bukan SPREADSHEET_ID utama:

  - /api/fstl/cek-jo      : cocokkan JO input ke sheet "JO_1" kolom F
                            (suffix setelah "/" terakhir, huruf nyangkut di
                            belakang angka diabaikan), balikin produk dari
                            kolom G.
  - /api/fstl/keterangan  : buat tiap proses yang dicentang (Printing, Dry
                            Laminasi, Slitting, Rewinding, Extrusi, Bag
                            Making), textjoin semua KETERANGAN yang JO-nya
                            cocok dari sheet sumbernya (lihat
                            FSTL_PROCESS_SOURCES), ditambah hasil dari sheet
                            "LP_1" (difilter kolom KLASIFIKASI) sebagai
                            "... LAPORAN PROD: ...".
  - /api/fstl/save        : simpan catatan waste ke sheet "{USERNAME}_Kitir"
                            (dibuat otomatis kalau belum ada, TANPA hapus
                            sheet lama), sebagai satu "kartu" mulai kolom B
                            baris 2 (kolom A & baris 1 TIDAK disentuh):
                              baris 1 (hijau) : SPK/JO | Produk | Waste Besar
                                                Proses : <daftar proses>
                              baris 2 (biru)  : label kolom (Waste Besar
                                                Proses/Keterangan/Action
                                                Plan/Status)
                              baris 3..N (hijau) : satu baris per proses yang
                                                dicentang
                            Kartu baru selalu disisipkan tepat di baris 2,
                            jadi kartu-kartu lama otomatis ikut turun tanpa
                            baris kosong pemisah.
"""

import math
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime, date, timedelta, timezone
from pathlib import Path

import gspread
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
from google.oauth2.service_account import Credentials

import import_engine
import rewind_qty
import kg_bruto
import konversi_meter_jumbo
import chatbot_engine
import run_all as run_all_module  # dipakai buat daftar script (SCRIPTS_ORDER) & jalankan satu-satu

# Baca file .env (SPREADSHEET_ID, GOOGLE_CREDENTIALS_FILE, PORT,
# DEEPSEEK_API_KEY, dst) dan masukkan ke environment variable proses ini,
# supaya os.environ.get(...) di bawah bisa nemu nilainya. Kalau file .env
# nggak ada, ini nggak error -- cuma dianggap kosong (masih bisa jalan
# kalau env var-nya sudah di-set manual lewat "set" di CMD).
load_dotenv()

# --------------------------------------------------------------------------
# KONFIGURASI
# --------------------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

CREDENTIALS_FILE = os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")  # isi di file .env

BASE_DIR = Path(__file__).resolve().parent
RUN_ALL_PATH = BASE_DIR / "run_all.py"

# Label yang enak dibaca untuk tiap script individual di dropdown "Jalankan
# Satu Script" (halaman Input Data Produksi). Urutan & isi list script-nya
# sendiri TETAP ambil dari run_all.SCRIPTS_ORDER (satu sumber kebenaran),
# ini cuma peta nama file -> label tampilan.
SCRIPT_LABELS = {
    "import_printing_2.py": "Printing 2",
    "import_printing_3.py": "Printing 3",
    "import_printing_4.py": "Printing 4",
    "import_printing_5.py": "Printing 5",
    "import_rw.py": "RW (Rewinding)",
    "import_sl.py": "SL (Slitting)",
    "import_dry_1.py": "Dry 1",
    "import_dry_2.py": "Dry 2",
    "import_dry_3.py": "Dry 3",
    "import_dry_4.py": "Dry 4",
    "import_dry_5.py": "Dry 5",
    "import_sf.py": "SF",
    "import_ex.py": "EX (Extrusi)",
    "import_bag.py": "Bag Making",
    "import_jo.py": "JO",
    "import_lp.py": "LP (Laporan Produksi)",
    "import_rewind_kecil.py": "Rewind Kecil (REWIND_PY_RAW)",
}

app = Flask(__name__)
# PENTING: Flask defaultnya SORT_KEYS alphabetical buat semua jsonify() --
# ini yang bikin urutan kolom di modal Cek Stok (dan tabel lain yang
# ngambil kolom dari Object.keys(row) di frontend) jadi kacau (alfabetis,
# bukan urutan yang kita susun di WRW_STOK_*_COLUMNS dkk). Dimatikan biar
# urutan key di dict Python (row_out = {...}) yang nentuin urutan kolom.
app.json.sort_keys = False
CORS(app)  # izinkan dipanggil dari frontend berbeda origin (mis. Figma / GitHub Pages)


# --------------------------------------------------------------------------
# SERVE FRONTEND (index.html) -- tanpa ini, buka domain Render langsung
# bakal muncul "Not Found" 404 karena app.py aslinya cuma nyediain
# route /api/... saja (index.html sebelumnya dibuka manual dari komputer,
# bukan lewat server, jadi ini nggak ketahuan sampai di-deploy ke Render).
# --------------------------------------------------------------------------
@app.route("/")
def serve_index():
    return send_from_directory(BASE_DIR, "index.html")


@app.route("/<path:filename>")
def serve_static_asset(filename):
    """Buat file pendukung frontend lain (css/js/gambar) kalau ada, yang
    ditaruh sejajar index.html dan direferensikan pakai path relatif."""
    return send_from_directory(BASE_DIR, filename)


import os as _os_diag  # noqa: E402  (cuma buat print PID di bawah, nggak ganggu import 'os' yang di atas)
print(f"=== SERVER STARTED (PID={_os_diag.getpid()}) — kalau baris ini muncul LAGI di tengah-tengah kamu testing, artinya server abis restart otomatis (cache ke-reset) ===", flush=True)


_gspread_client = None
_gspread_client_lock = threading.Lock()


def get_client():
    """Login ke Google (baca credentials.json + otorisasi) itu operasi yang
    lumayan berat kalau diulang tiap request. Sebelumnya dipanggil dari nol
    di SETIAP endpoint yang butuh Sheets -- termasuk /api/fstl/keterangan,
    jadi tiap klik "Ambil Keterangan" (JO sama ATAUPUN beda) selalu kena
    biaya login ulang ini duluan, sebelum sempat manfaatin cache sheet di
    bawah. Client login cuma dibuat SEKALI lalu dipakai ulang terus --
    aman, karena Credentials dari google-auth otomatis refresh token-nya
    sendiri kalau kadaluarsa, tanpa perlu login dari awal lagi."""
    global _gspread_client
    with _gspread_client_lock:
        if _gspread_client is None:
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
            _gspread_client = gspread.authorize(creds)
        return _gspread_client


_spreadsheet_handle = {"sh": None}
_spreadsheet_lock = threading.Lock()


def _is_quota_error(e):
    """True kalau exception ini gspread.exceptions.APIError status 429
    (limit 'requests per minute per user' Google Sheets API kelampauan).
    Beda dari error lain (sheet nggak ada, credential salah, dst) yang
    memang harus langsung gagal -- 429 ini murni soal kebanyakan request
    dalam satu menit, jadi wajar buat dicoba lagi setelah nunggu sebentar."""
    if not isinstance(e, gspread.exceptions.APIError):
        return False
    try:
        status = e.response.status_code
    except AttributeError:
        status = None
    return status == 429 or "RESOURCE_EXHAUSTED" in str(e) or "Quota exceeded" in str(e)


def get_sheet(sheet_name, attempts=3):
    """Sama kayak _fstl_spreadsheet() di bawah -- open_by_key() itu
    panggilan ke Google (fetch metadata spreadsheet), dan ID-nya nggak
    pernah berubah selama app jalan. SEBELUMNYA dipanggil dari nol di
    SETIAP get_sheet(), jadi tool baru search_produk (yang buka hampir
    SEMUA sheet buat cari nama produk lintas grup, bisa belasan sheet
    sekaligus) jadi buka spreadsheet dari nol belasan kali cuma buat satu
    pertanyaan -- ini yang bikin request lambat/timeout ('Failed to
    fetch') waktu user tanya berdasarkan nama produk. Sekarang handle
    spreadsheet-nya dibuka sekali lalu dipakai ulang terus.

    Ditambah retry khusus buat 429 ('Quota exceeded ... per minute') --
    SEBELUMNYA sekali kena 429 (misalnya persis setelah tombol Refresh
    dipencet dan kuota per-menit abis) endpoint manapun yang manggil
    get_sheet() langsung crash jadi 500 tanpa dicoba ulang. Sekarang
    ditunggu sebentar dulu (kuotanya reset per menit) lalu dicoba lagi
    sebelum benar-benar dianggap gagal."""
    if not SPREADSHEET_ID:
        raise RuntimeError("SPREADSHEET_ID belum diset (lihat file .env)")
    with _spreadsheet_lock:
        sh = _spreadsheet_handle["sh"]
    if sh is None:
        client = get_client()
        sh = client.open_by_key(SPREADSHEET_ID)
        with _spreadsheet_lock:
            _spreadsheet_handle["sh"] = sh

    last_exc = None
    for attempt in range(1, attempts + 1):
        try:
            return sh.worksheet(sheet_name)
        except gspread.exceptions.APIError as e:
            if not _is_quota_error(e) or attempt == attempts:
                raise
            last_exc = e
            wait = 15 * attempt
            print(f"   ⚠️ Kena limit kuota Google Sheets (get_sheet '{sheet_name}'): {e}. Coba lagi dalam {wait}s...")
            time.sleep(wait)
    raise last_exc


# --------------------------------------------------------------------------
# 1. LOGIN
# --------------------------------------------------------------------------
LOGIN_COLUMN_MAP = {
    "USERNAME": "username",
    "PASSWORD": "password",
    "NAMA": "nama",
    "ROLE": "role",
}


def normalize_row(row, column_map):
    """Ubah key dari header asli sheet jadi key yang dipakai frontend,
    sekaligus tetap simpan key aslinya kalau-kalau dibutuhkan."""
    out = dict(row)  # simpan versi asli juga
    for original_key, new_key in column_map.items():
        if original_key in row:
            out[new_key] = row[original_key]
    return out


@app.route("/api/login", methods=["POST"])
def login():
    body = request.get_json(force=True) or {}
    username = str(body.get("username", "")).strip()
    password = str(body.get("password", "")).strip()

    ws = get_sheet("Login")
    raw = ws.get_all_records()  # baris pertama dianggap header
    rows = [normalize_row(r, LOGIN_COLUMN_MAP) for r in raw]

    for r in rows:
        if str(r.get("username", "")).strip() == username and str(r.get("password", "")).strip() == password:
            return jsonify({
                "status": "SUKSES",
                "nama": r.get("nama", ""),
                "role": r.get("role", ""),
            })

    return jsonify({"status": "GAGAL", "pesan": "Username atau password salah"}), 401


# --------------------------------------------------------------------------
# 2. DATA VALIDASI
# --------------------------------------------------------------------------
# Nama kolom ASLI di sheet -> nama field yang dipakai frontend.
# Sesuaikan bagian kiri kalau header di sheet kamu berubah.
VALIDASI_COLUMN_MAP = {
    "TANGGAL": "tanggal",
    "JO": "jo",
    "NAMA": "nama",
    "ORDER": "order",
    "HASIL SLITTING": "slitting",
    "HASIL SLIT\n(QTY)": "qtySlit",
    "HASIL BAG": "hasilBag",
    "VALIDASI": "validasi",
    "FORM SERAH TERIMA": "serahTerima",
    "TOTAL": "total",
    "SELISIH": "selisih",
    "STATUS": "status",
    "POTONGAN": "potongan",
}


@app.route("/api/validasi", methods=["GET"])
def get_validasi():
    ws = get_sheet("Validasi")
    # numericise_ignore=['all']: JANGAN biarkan gspread otomatis mengubah
    # cell yang keliatan seperti angka jadi int/float. Kita pakai format
    # angka Indonesia (titik = ribuan, koma = desimal) -- kalau dibiarkan,
    # gspread nganggep titik itu desimal (konvensi US) dan "160.000" jadi
    # kebaca 160.0, ditampilkan "160" di frontend. Ambil apa adanya (string).
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, VALIDASI_COLUMN_MAP) for r in raw]
    return jsonify(data)



# ---- Refresh sumber Data Validasi (VAL + Form Serah Terima + sinkron JO) ----
# Beda dari "Refresh Semua" di halaman Input Data Produksi (run_all.py,
# 15 script Printing/Rw/Sl/Dry/Sf/Ex/Bag/JO): ini cuma 3 script kecil, jadi
# dijalankan langsung (blocking) di request ini, tanpa background thread.
# - import_val.py & import_form_st.py : import dari spreadsheet luar (VAL_1 & FORM_ST_1)
# - refresh_validasi_jo.py            : sinkron internal SL_1 + JO_1 -> kolom
#                                        TANGGAL/JO/NAMA di tab "Validasi"
VALIDASI_IMPORT_SCRIPTS = ["import_val.py", "import_form_st.py", "refresh_validasi_jo.py"]


@app.route("/api/validasi/refresh-import", methods=["POST"])
def refresh_import_validasi():
    """Dipanggil tombol Refresh di halaman Data Validasi: jalankan
    import_val.py, import_form_st.py, lalu refresh_validasi_jo.py,
    supaya tab VAL_1 & FORM_ST_1 ter-update dan tab "Validasi" (kolom
    TANGGAL/JO/NAMA) tersinkron, sebelum data ditarik ulang lewat
    GET /api/validasi."""
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    results = []
    for script_name in VALIDASI_IMPORT_SCRIPTS:
        script_path = BASE_DIR / script_name
        if not script_path.exists():
            results.append({"script": script_name, "status": "NOT FOUND", "log": ""})
            continue
        try:
            proc = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            ok = proc.returncode == 0
            results.append({
                "script": script_name,
                "status": "OK" if ok else f"FAILED (exit {proc.returncode})",
                "log": (proc.stdout or "") + (proc.stderr or ""),
            })
        except Exception as e:
            results.append({"script": script_name, "status": f"FAILED ({e})", "log": ""})

    success = all(r["status"] == "OK" for r in results)
    return jsonify({"success": success, "results": results})


@app.route("/api/validasi/status", methods=["POST"])
def update_status_rekap():
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    status = body.get("status", "")

    ws = get_sheet("Validasi")
    cell = ws.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan"}), 404

    header = ws.row_values(1)
    if "STATUS" not in header:
        return jsonify({"success": False, "message": "Kolom 'STATUS' tidak ada di sheet Validasi"}), 400

    col_status = header.index("STATUS") + 1
    ws.update_cell(cell.row, col_status, status)
    return jsonify({"success": True})


# ---- Tombol "Status OK (Manual)" di halaman Rekap ----
@app.route("/api/validasi/ok-manual", methods=["POST"])
def ok_manual_validasi():
    """Dipanggil tombol 'Status OK (Manual)' di tabel Rekap Data Selisih:
    1. Salin 1 baris JO dari sheet "Validasi" ke sheet "Revisi_Manual"
       (kalau JO itu belum pernah tersimpan di sana, supaya tidak dobel).
       Sheet "Revisi_Manual" inilah yang jadi SUMBER KEBENARAN status OK --
       tiap kali refresh_validasi_jo.py jalan (lihat sync_validasi_header()
       di import_engine.py), kolom STATUS di seluruh sheet Validasi
       dihitung ULANG dari sini (dicocokkan lewat JO, bukan nomor baris),
       jadi tidak akan salah baris walau posisi baris JO berubah antar-refresh.
    2. Set juga kolom STATUS di baris JO ini langsung jadi "OK", supaya
       Rekap langsung update seketika tanpa perlu tunggu tombol Refresh
       (loadRekap() di frontend sudah filter status != "OK")."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    if not jo:
        return jsonify({"success": False, "message": "JO wajib diisi"}), 400

    ws_val = get_sheet("Validasi")
    cell = ws_val.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan di Validasi"}), 404

    header = ws_val.row_values(1)
    row_values = ws_val.row_values(cell.row)

    ws_revisi = get_sheet("Revisi_Manual")
    if not ws_revisi.find(jo):
        ws_revisi.append_row(row_values)

    if "STATUS" in header:
        col_status = header.index("STATUS") + 1
        ws_val.update_cell(cell.row, col_status, "OK")

    return jsonify({"success": True})


# --------------------------------------------------------------------------
# 3. PIC (untuk dropdown kirim WA)
# --------------------------------------------------------------------------
PIC_COLUMN_MAP = {
    "NAMA": "nama",
    "NOMOR": "nomor",
}


@app.route("/api/pic", methods=["GET"])
def get_pic_list():
    ws = get_sheet("PIC")
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, PIC_COLUMN_MAP) for r in raw]
    return jsonify(data)  # [{"nama": ..., "nomor": ...}, ...]


# --------------------------------------------------------------------------
# 4. UPDATE STOCK (monitor bahan baku - butuh ACC)
# --------------------------------------------------------------------------
UPDATESTOCK_COLUMN_MAP = {
    "JO": "jo",
    "NAMA": "nama",
    "ORDER": "order",
    "METER ORDER": "meterOrder",
    "METER VALIDASI": "meterValidasi",
    "LAPISAN ORDER": "lapisanOrder",
    "LAPISAN VALIDASI": "lapisanValidasi",
    "KETERANGAN": "keterangan",
    "ACC": "acc",
}


@app.route("/api/update-stock/refresh-import", methods=["POST"])
def refresh_import_update_stock():
    """Dipanggil tombol Refresh di halaman 'Update Stock Bahan Baku'.
    Menjalankan DUA sinkronisasi terpisah (target spreadsheet beda,
    tidak saling menimpa):
      1. run_update_stock_import() -- baca sheet2 PL/PET/CPPM dst yang
         dicentang di kartu 'Update Stock' -> tulis ulang kolom B-K di
         tab tujuan pada spreadsheet EKSTERNAL Monitor Bahan Baku
         (kolom A di sana tidak disentuh).
      2. sync_update_stock_from_jo() -- baca JO_1 (spreadsheet utama)
         -> tulis kolom A/B/C/F tab 'UpdateStock' (spreadsheet utama)
         yang dipakai halaman ini sendiri lewat GET /api/update-stock.
    Kalau salah satu gagal, tetap coba jalankan yang satunya (supaya
    satu bagian yang error tidak ikut menggagalkan bagian lain), lalu
    laporkan errornya di response."""
    errors = []

    try:
        rows_written = import_engine.run_update_stock_import("update_stock")
        import_engine.set_import_result("update_stock", "OK", rows_written=rows_written, error=None)
    except Exception as e:
        rows_written = None
        import_engine.set_import_result("update_stock", "ERROR", rows_written=None, error=str(e))
        errors.append(f"run_update_stock_import: {e}")

    # Jeda sebelum lanjut ke sync_update_stock_from_jo() -- fungsi itu
    # LANGSUNG buka lagi spreadsheet eksternal Monitor Bahan Baku
    # (1lSj54tQP8QKMR96HiAHBCd1Zm-fOxnkdt9x3gD1tuEM) yang barusan ditulis
    # di atas (sheet UPDATE_STOCK, kolom B-K), buat baca ulang isinya
    # (_build_update_stock_monitor_lookup, kolom E/G). Tulis besar
    # (batch_clear + update) langsung disusul baca lagi ke spreadsheet
    # YANG SAMA dalam hitungan detik itu yang bikin gampang numpuk kena
    # limit "requests per minute" -- kasih jeda dulu di sini biar kuotanya
    # sempat longgar sebelum dipakai lagi.
    time.sleep(20)

    try:
        jo_rows_synced = import_engine.sync_update_stock_from_jo()
    except Exception as e:
        jo_rows_synced = None
        errors.append(f"sync_update_stock_from_jo: {e}")

    if errors:
        return jsonify({
            "success": False,
            "message": " | ".join(errors),
            "rows_written": rows_written,
            "jo_rows_synced": jo_rows_synced,
        }), 400

    return jsonify({
        "success": True,
        "rows_written": rows_written,
        "jo_rows_synced": jo_rows_synced,
    })


@app.route("/api/update-stock", methods=["GET"])
def get_update_stock():
    ws = get_sheet("UpdateStock")
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, UPDATESTOCK_COLUMN_MAP) for r in raw]
    for r in data:
        r["isLocked"] = str(r.get("acc", "0")) == "1"
    return jsonify(data)


@app.route("/api/update-stock/acc", methods=["POST"])
def acc_update_stock():
    """Tombol 'ACC & Kirim': kunci baris + salin data ke sheet StockBahan."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()

    ws_update = get_sheet("UpdateStock")
    cell = ws_update.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan"}), 404

    header = ws_update.row_values(1)
    col_acc = header.index("ACC") + 1 if "ACC" in header else None
    if col_acc:
        ws_update.update_cell(cell.row, col_acc, "1")

    ws_stock = get_sheet("StockBahan")
    ws_stock.append_row([
        datetime.now().strftime("%d-%m-%Y %H:%M"),
        body.get("user", "Tidak Diketahui"),
        body.get("jo", ""),
        body.get("nama", ""),
        body.get("order", ""),
        body.get("meterOrder", ""),
        body.get("meterValidasi", ""),
        body.get("lapisanOrder", ""),
        body.get("lapisanValidasi", ""),
    ])
    return jsonify({"success": True})


@app.route("/api/update-stock/unlock", methods=["POST"])
def unlock_update_stock():
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()

    ws = get_sheet("UpdateStock")
    cell = ws.find(jo)
    if not cell:
        return jsonify({"success": False, "message": f"JO {jo} tidak ditemukan"}), 404

    header = ws.row_values(1)
    if "ACC" in header:
        ws.update_cell(cell.row, header.index("ACC") + 1, "0")
    return jsonify({"success": True})


# --------------------------------------------------------------------------
# 5. STOCK BAHAN BAKU (hasil ACC)
# --------------------------------------------------------------------------
STOCKBAHAN_COLUMN_MAP = {
    "TANGGAL": "tanggal",
    "USER": "user",
    "JO": "jo",
    "NAMA": "nama",
    "ORDER": "order",
    "METER ORDER": "meterOrder",
    "METER VALIDASI": "meterValidasi",
    "LAPISAN ORDER": "lapisanOrder",
    "LAPISAN VALIDASI": "lapisanValidasi",
}


@app.route("/api/stock-bahan", methods=["GET"])
def get_stock_bahan():
    ws = get_sheet("StockBahan")
    raw = ws.get_all_records(numericise_ignore=["all"])
    data = [normalize_row(r, STOCKBAHAN_COLUMN_MAP) for r in raw]
    return jsonify(data)


# --------------------------------------------------------------------------
# 6. INPUT DATA PRODUKSI — Load link / Pilih Sheet / Refresh (Run All)
# --------------------------------------------------------------------------

@app.route("/api/produksi/sources", methods=["GET"])
def produksi_sources():
    """Daftar semua source (Printing 2..5, RW, SL, SF, Dry 1..5) beserta
    status koneksi & sheet yang sudah dicentang — dipakai untuk render kartu
    generik (link + checklist sheet). Source "gudang", "update_stock",
    "form_st_2", & "val_2" SENGAJA DIKECUALIKAN di sini karena alurnya beda
    (upload file, bukan link) dan punya kartu hardcoded sendiri di frontend
    (lihat index.html, kartu "Data Gudang", "Update Stock", "Form Serah
    Terima 2", "Validasi 2")."""
    cfg = import_engine.load_config()
    sources = cfg.get("sources", {})
    sources = {
        k: v for k, v in sources.items()
        if k not in import_engine.GUDANG_SOURCES and k not in import_engine.FILE_UPLOAD_EXTRA_SOURCES
    }
    return jsonify(sources)


@app.route("/api/produksi/load", methods=["POST"])
def produksi_load():
    """Body: {source_key, link}
    Ekstrak ID dari link, coba connect, deteksi nama semua sheet/tab,
    simpan source_id ke config.json. Sheet yang sudah pernah dicentang
    sebelumnya TIDAK dihapus otomatis, biar user bisa cocokkan ulang."""
    body = request.get_json(force=True) or {}
    source_key = str(body.get("source_key", "")).strip()
    link = str(body.get("link", "")).strip()

    if not source_key:
        return jsonify({"success": False, "message": "source_key wajib diisi"}), 400
    if not link:
        return jsonify({"success": False, "message": "Link/ID spreadsheet wajib diisi"}), 400

    try:
        import_engine.get_source(source_key)
    except KeyError as e:
        return jsonify({"success": False, "message": str(e)}), 404

    try:
        source_id = import_engine.extract_id_from_link(link)
    except ValueError as e:
        return jsonify({"success": False, "message": str(e)}), 400

    # Jenis sumber (Google Sheets asli vs file Excel/WPS di Drive) dideteksi
    # OTOMATIS lewat mimeType-nya di Drive API -- user tidak perlu pilih
    # manual lagi (dulu ada dropdown khusus di kartu "Rewind Kecil").
    try:
        source_type = import_engine.detect_source_type(source_id)
        detected_sheets, file_name = import_engine.detect_sheets(source_id, source_type)
    except Exception as e:
        return jsonify({"success": False, "message": f"Gagal connect ke spreadsheet: {e}"}), 400

    updated = import_engine.update_source(
        source_key,
        type=source_type,
        source_id=source_id,
        source_name=file_name,
        last_connected=datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )

    # Simpan juga link mentah yang dipaste user ke sheet "ListMesin" (kolom B),
    # pada baris yang cocok dengan nama mesin source ini (kolom A). Kalau ini
    # gagal (mis. sheet ListMesin belum ada / nama mesin tidak match), jangan
    # sampai menggagalkan proses Load utama — cukup diabaikan.
    try:
        import_engine.update_list_mesin_config(source_key, link=link)
    except Exception:
        pass

    return jsonify({
        "success": True,
        "source_id": source_id,
        "file_name": file_name,
        "detected_sheets": detected_sheets,
        "selected_sheets": updated.get("sheets", []),
    })


@app.route("/api/produksi/sheets", methods=["POST"])
def produksi_sheets():
    """Body: {source_key, sheets: [...]}
    Simpan daftar sheet yang dicentang user untuk source ini -> menggantikan
    SHEETS_TO_IMPORT yang dulu hardcoded di tiap script."""
    body = request.get_json(force=True) or {}
    source_key = str(body.get("source_key", "")).strip()
    sheets = body.get("sheets")

    if not source_key:
        return jsonify({"success": False, "message": "source_key wajib diisi"}), 400
    if not isinstance(sheets, list):
        return jsonify({"success": False, "message": "sheets harus berupa list"}), 400

    try:
        import_engine.update_source(source_key, sheets=sheets)
    except KeyError as e:
        return jsonify({"success": False, "message": str(e)}), 404

    # Simpan juga daftar sheet yang dicentang ke sheet "ListMesin" (kolom
    # C), sama alasannya dengan link di produksi_load() -- ini catatan
    # cadangan yang persisten di Google Sheets, dipakai buat memulihkan
    # config.json otomatis (_recover_source_from_list_mesin) kalau
    # sampai hilang (mis. abis server restart di hosting yang disknya
    # ephemeral). Kalau gagal, jangan sampai menggagalkan proses utama.
    try:
        import_engine.update_list_mesin_config(source_key, sheets=sheets)
    except Exception:
        pass

    return jsonify({"success": True, "sheets": sheets})


# ---- Refresh / Run All (background, supaya 1 tombol tapi tidak nge-block) ----
RUN_STATE_LOCK = threading.Lock()
RUN_STATE = {
    "running": False,
    "log": "",
    "started_at": None,
    "finished_at": None,
    "returncode": None,
}


def _run_all_worker():
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    try:
        proc = subprocess.Popen(
            [sys.executable, str(RUN_ALL_PATH)],
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            bufsize=1,
        )
        for line in proc.stdout:
            with RUN_STATE_LOCK:
                RUN_STATE["log"] += line
        proc.wait()
        returncode = proc.returncode

        # Rewind Kecil (import mentah ke REWIND_PY_RAW) ikut Refresh Semua,
        # supaya tombol Refresh di halaman Waste Rewind TIDAK perlu narik
        # data mentah lagi (hemat kuota API / hindari 429). Kalau
        # run_all.SCRIPTS_ORDER sudah memuat script ini, jangan dobel.
        if REWIND_KECIL_SCRIPT.name not in run_all_module.SCRIPTS_ORDER and REWIND_KECIL_SCRIPT.exists():
            with RUN_STATE_LOCK:
                RUN_STATE["log"] += f"\n=== {REWIND_KECIL_SCRIPT.name} (Rewind Kecil -> REWIND_PY_RAW) ===\n"
            proc2 = subprocess.Popen(
                [sys.executable, str(REWIND_KECIL_SCRIPT)],
                cwd=str(BASE_DIR),
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
                bufsize=1,
            )
            for line in proc2.stdout:
                with RUN_STATE_LOCK:
                    RUN_STATE["log"] += line
            proc2.wait()
            if returncode == 0:
                returncode = proc2.returncode

        with RUN_STATE_LOCK:
            RUN_STATE["returncode"] = returncode
    except Exception as e:
        with RUN_STATE_LOCK:
            RUN_STATE["log"] += f"\n[GAGAL MENJALANKAN run_all.py] {e}\n"
            RUN_STATE["returncode"] = -1
    finally:
        with RUN_STATE_LOCK:
            RUN_STATE["running"] = False
            RUN_STATE["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@app.route("/api/produksi/run-all", methods=["POST"])
def produksi_run_all():
    """Tombol Refresh tunggal: jalankan run_all.py di background thread.
    Frontend lalu polling /api/produksi/run-status untuk lihat progress."""
    with RUN_STATE_LOCK, RUN_ONE_STATE_LOCK:
        if RUN_STATE["running"]:
            return jsonify({"success": False, "message": "Sedang berjalan, tunggu sampai selesai."}), 409
        if RUN_ONE_STATE["running"]:
            return jsonify({"success": False, "message": "Ada script tunggal yang sedang jalan, tunggu sampai selesai."}), 409
        RUN_STATE["running"] = True
        RUN_STATE["log"] = ""
        RUN_STATE["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        RUN_STATE["finished_at"] = None
        RUN_STATE["returncode"] = None

    thread = threading.Thread(target=_run_all_worker, daemon=True)
    thread.start()
    return jsonify({"success": True, "message": "run_all.py mulai dijalankan."})


@app.route("/api/produksi/run-status", methods=["GET"])
def produksi_run_status():
    with RUN_STATE_LOCK:
        return jsonify(dict(RUN_STATE))


# ---- Jalankan SATU script saja (dropdown di sebelah tombol Run All) ----
# Daftar & urutan file-nya ikut run_all.SCRIPTS_ORDER (satu sumber kebenaran,
# tidak ditulis ulang di sini) supaya kalau run_all.py nambah/hapus script,
# dropdown ini otomatis ikut update tanpa perlu ubah app.py.
RUN_ONE_STATE_LOCK = threading.Lock()
RUN_ONE_STATE = {
    "running": False,
    "log": "",
    "started_at": None,
    "finished_at": None,
    "returncode": None,
    "script": None,
}


def _run_one_worker(script_name):
    script_path = BASE_DIR / script_name
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"
    try:
        proc = subprocess.Popen(
            [sys.executable, str(script_path)],
            cwd=str(BASE_DIR),
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
            bufsize=1,
        )
        for line in proc.stdout:
            with RUN_ONE_STATE_LOCK:
                RUN_ONE_STATE["log"] += line
        proc.wait()
        with RUN_ONE_STATE_LOCK:
            RUN_ONE_STATE["returncode"] = proc.returncode
    except Exception as e:
        with RUN_ONE_STATE_LOCK:
            RUN_ONE_STATE["log"] += f"\n[GAGAL MENJALANKAN {script_name}] {e}\n"
            RUN_ONE_STATE["returncode"] = -1
    finally:
        with RUN_ONE_STATE_LOCK:
            RUN_ONE_STATE["running"] = False
            RUN_ONE_STATE["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def _produksi_scripts_order():
    """SCRIPTS_ORDER dari run_all.py + import_rewind_kecil.py (kalau belum
    ada di sana), supaya Rewind Kecil bisa dijalankan lewat "Jalankan Satu
    Script" di halaman Input Data Produksi."""
    items = list(run_all_module.SCRIPTS_ORDER)
    if REWIND_KECIL_SCRIPT.name not in items:
        items.append(REWIND_KECIL_SCRIPT.name)
    return items


@app.route("/api/produksi/scripts", methods=["GET"])
def produksi_scripts():
    """Daftar script individual (file + label) buat isi dropdown di frontend,
    urutannya sama seperti yang dijalankan run_all.py."""
    items = [
        {"file": f, "label": SCRIPT_LABELS.get(f, f)}
        for f in _produksi_scripts_order()
    ]
    return jsonify(items)


@app.route("/api/produksi/run-one", methods=["POST"])
def produksi_run_one():
    """Body: {script: "import_rw.py"}
    Jalankan satu script import saja (bukan run_all.py), di background
    thread, dengan panel log/status terpisah dari Run All."""
    body = request.get_json(force=True) or {}
    script = str(body.get("script", "")).strip()

    if not script:
        return jsonify({"success": False, "message": "Pilih script dulu."}), 400
    if script not in _produksi_scripts_order():
        return jsonify({"success": False, "message": f"Script '{script}' tidak dikenal."}), 400

    script_path = BASE_DIR / script
    if not script_path.exists():
        return jsonify({"success": False, "message": f"{script} tidak ditemukan di server."}), 404

    with RUN_STATE_LOCK, RUN_ONE_STATE_LOCK:
        if RUN_STATE["running"]:
            return jsonify({"success": False, "message": "Refresh Semua sedang berjalan, tunggu sampai selesai."}), 409
        if RUN_ONE_STATE["running"]:
            return jsonify({"success": False, "message": "Sedang ada script lain berjalan, tunggu sampai selesai."}), 409
        RUN_ONE_STATE["running"] = True
        RUN_ONE_STATE["log"] = ""
        RUN_ONE_STATE["script"] = script
        RUN_ONE_STATE["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        RUN_ONE_STATE["finished_at"] = None
        RUN_ONE_STATE["returncode"] = None

    thread = threading.Thread(target=_run_one_worker, args=(script,), daemon=True)
    thread.start()
    label = SCRIPT_LABELS.get(script, script)
    return jsonify({"success": True, "message": f"{label} mulai dijalankan."})


@app.route("/api/produksi/run-one-status", methods=["GET"])
def produksi_run_one_status():
    with RUN_ONE_STATE_LOCK:
        return jsonify(dict(RUN_ONE_STATE))



# --------------------------------------------------------------------------
# 6a.5 REWIND KECIL — import mentah (REWIND_PY_RAW) ikut Refresh Semua & Run One;
#      endpoint /rewind-kecil/run di bawah = refresh Waste Rewind (tanpa import)
# --------------------------------------------------------------------------
REWIND_KECIL_SCRIPT = BASE_DIR / "import_rewind_kecil.py"
REWIND_KECIL_RUN_STATE_LOCK = threading.Lock()
REWIND_KECIL_RUN_STATE = {
    "running": False,
    "log": "",
    "started_at": None,
    "finished_at": None,
    "returncode": None,
    "rows_written": None,
    "spk_jo_added": None,
    "bahan_awal_updated": None,
    "error": None,
}


def _run_rewind_kecil_worker():
    """Tombol Refresh di halaman Waste Rewind.

    TIDAK lagi menjalankan import_rewind_kecil.py / menulis ke
    REWIND_PY_RAW (itu bikin kena limit 429 dari Sheets API). Import data
    mentah sekarang lewat Refresh Semua / Jalankan Satu Script di halaman
    Input Data Produksi. Di sini cuma menghitung ulang REWIND_PY dari data
    yang SUDAH ada di REWIND_PY_RAW + LP/JO/SL/PRINTING."""
    try:
        rows_written = None
        err = None

        spk_jo_added = None
        bahan_awal_updated = None
        hasil_slitting_updated = None
        printing_updated = None
        tanggal_qty_updated = None
        kg_bruto_updated = None
        konversi_updated = None
        meter_hilang_updated = None
        waste_kolom_updated = None
        revisi_applied = None
        _invalidate_waste_rewind_source_cache()  # Refresh selalu baca LP_1/JO_1/SL_1/PRINTING_x terbaru
        try:
            spk_jo_added = _sync_rewind_kecil_spk_jo_into_rewind_py()
        except Exception as sync_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += f"\n[GAGAL SINKRON SPK/NO_JO ke {WASTE_REWIND_SHEET_NAME}] {sync_err}\n"
            err = err or str(sync_err)
        try:
            bahan_awal_updated = _sync_bahan_awal_printing_into_rewind_py()
        except Exception as bahan_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += f"\n[GAGAL ISI Bahan_Awal_Printing_(Meter) di {WASTE_REWIND_SHEET_NAME}] {bahan_err}\n"
            err = err or str(bahan_err)
        try:
            hasil_slitting_updated = _sync_slitting_kolom_into_rewind_py()
        except Exception as slit_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL ISI Hasil_Slitting_(Rol)/UP_Slitting/Hasil_Slitting_(Meter) "
                    f"di {WASTE_REWIND_SHEET_NAME}] {slit_err}\n"
                )
            err = err or str(slit_err)
        try:
            printing_updated = _sync_printing_kolom_into_rewind_py()
        except Exception as printing_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL ISI Printing_1..5/Total_Hasil_Printing "
                    f"di {WASTE_REWIND_SHEET_NAME}] {printing_err}\n"
                )
            err = err or str(printing_err)
        try:
            tanggal_qty_updated = _sync_rewind_kecil_tanggal_qty_into_rewind_py()
        except Exception as tq_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL ISI Tanggal_Rewind/Qty_Awal_Rewind/Qty_Akhir_Rewind "
                    f"di {WASTE_REWIND_SHEET_NAME}] {tq_err}\n"
                )
            err = err or str(tq_err)
        try:
            kg_bruto_updated = _sync_kg_bruto_into_rewind_py()
        except Exception as kg_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL ISI Kg_Bruto di {WASTE_REWIND_SHEET_NAME}] {kg_err}\n"
                )
            err = err or str(kg_err)
        # HARUS setelah Qty_Awal_Rewind, UP_Slitting & Kg_Bruto terisi (ketiganya jadi input)
        try:
            konversi_updated = _sync_konversi_meter_jumbo_into_rewind_py()
        except Exception as kv_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL ISI Konversi_Meter_Jumbo_Qty_Awal/Akhir_Rewind "
                    f"di {WASTE_REWIND_SHEET_NAME}] {kv_err}\n"
                )
            err = err or str(kv_err)
        # HARUS setelah Konversi_Meter_Jumbo_* terisi (jadi input Meter_Jumbo_Hilang_Rewind)
        try:
            meter_hilang_updated = _sync_meter_hilang_into_rewind_py()
        except Exception as mh_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL ISI Meter_Jumbo_Hilang_Rewind/Meter_Hilang_Rewind "
                    f"di {WASTE_REWIND_SHEET_NAME}] {mh_err}\n"
                )
            err = err or str(mh_err)
        # HARUS setelah Bahan_Awal_Printing_(Meter)/Hasil_Slitting_(Meter)/
        # Meter_Jumbo_Hilang_Rewind terisi (jadi input hitung_waste_rewind()).
        # Ini yang bikin tombol Refresh ikut menghitung 4 kolom waste untuk
        # SEMUA JO (bukan cuma tombol "Hitung Waste" per-JO di modal Detail).
        try:
            waste_kolom_updated = _sync_hitung_waste_kolom_into_rewind_py()
        except Exception as wk_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL HITUNG 4 KOLOM WASTE di {WASTE_REWIND_SHEET_NAME}] {wk_err}\n"
                )
            err = err or str(wk_err)
        # PALING AKHIR: terapkan revisi manual (REWIND_PY_REVISI) per sel yang
        # berbeda, supaya tidak ditimpa sync-sync di atas.
        try:
            revisi_applied = _apply_rewind_revisi_into_rewind_py()
        except Exception as rv_err:
            with REWIND_KECIL_RUN_STATE_LOCK:
                REWIND_KECIL_RUN_STATE["log"] += (
                    f"\n[GAGAL TERAPKAN REVISI dari {WASTE_REWIND_REVISI_SHEET}] {rv_err}\n"
                )
            err = err or str(rv_err)

        with REWIND_KECIL_RUN_STATE_LOCK:
            REWIND_KECIL_RUN_STATE["returncode"] = 0
            REWIND_KECIL_RUN_STATE["rows_written"] = rows_written
            REWIND_KECIL_RUN_STATE["spk_jo_added"] = spk_jo_added
            REWIND_KECIL_RUN_STATE["bahan_awal_updated"] = bahan_awal_updated
            REWIND_KECIL_RUN_STATE["hasil_slitting_updated"] = hasil_slitting_updated
            REWIND_KECIL_RUN_STATE["printing_updated"] = printing_updated
            REWIND_KECIL_RUN_STATE["tanggal_qty_updated"] = tanggal_qty_updated
            REWIND_KECIL_RUN_STATE["kg_bruto_updated"] = kg_bruto_updated
            REWIND_KECIL_RUN_STATE["konversi_updated"] = konversi_updated
            REWIND_KECIL_RUN_STATE["meter_hilang_updated"] = meter_hilang_updated
            REWIND_KECIL_RUN_STATE["waste_kolom_updated"] = waste_kolom_updated
            REWIND_KECIL_RUN_STATE["revisi_applied"] = revisi_applied
            REWIND_KECIL_RUN_STATE["error"] = err
    except Exception as e:
        with REWIND_KECIL_RUN_STATE_LOCK:
            REWIND_KECIL_RUN_STATE["log"] += f"\n[GAGAL REFRESH WASTE REWIND] {e}\n"
            REWIND_KECIL_RUN_STATE["returncode"] = -1
            REWIND_KECIL_RUN_STATE["error"] = str(e)
    finally:
        with REWIND_KECIL_RUN_STATE_LOCK:
            REWIND_KECIL_RUN_STATE["running"] = False
            REWIND_KECIL_RUN_STATE["finished_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@app.route("/api/produksi/rewind-kecil/run", methods=["POST"])
def produksi_run_rewind_kecil():
    """Refresh Waste Rewind: hitung ulang REWIND_PY dari data yang sudah ada.
    Tidak import data mentah (lihat _run_rewind_kecil_worker)."""
    with REWIND_KECIL_RUN_STATE_LOCK:
        if REWIND_KECIL_RUN_STATE["running"]:
            return jsonify({
                "success": False,
                "message": "Refresh Rewind Kecil sedang berjalan."
            }), 409

        REWIND_KECIL_RUN_STATE["running"] = True
        REWIND_KECIL_RUN_STATE["log"] = ""
        REWIND_KECIL_RUN_STATE["started_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        REWIND_KECIL_RUN_STATE["finished_at"] = None
        REWIND_KECIL_RUN_STATE["returncode"] = None
        REWIND_KECIL_RUN_STATE["rows_written"] = None
        REWIND_KECIL_RUN_STATE["spk_jo_added"] = None
        REWIND_KECIL_RUN_STATE["bahan_awal_updated"] = None
        REWIND_KECIL_RUN_STATE["hasil_slitting_updated"] = None
        REWIND_KECIL_RUN_STATE["printing_updated"] = None
        REWIND_KECIL_RUN_STATE["konversi_updated"] = None
        REWIND_KECIL_RUN_STATE["meter_hilang_updated"] = None
        REWIND_KECIL_RUN_STATE["waste_kolom_updated"] = None
        REWIND_KECIL_RUN_STATE["revisi_applied"] = None
        REWIND_KECIL_RUN_STATE["error"] = None

    thread = threading.Thread(target=_run_rewind_kecil_worker, daemon=True)
    thread.start()

    return jsonify({
        "success": True,
        "message": "Refresh Waste Rewind mulai dijalankan."
    })


@app.route("/api/produksi/rewind-kecil/status", methods=["GET"])
def produksi_status_rewind_kecil():
    with REWIND_KECIL_RUN_STATE_LOCK:
        return jsonify(dict(REWIND_KECIL_RUN_STATE))


# --------------------------------------------------------------------------
# 6b. DATA GUDANG — file DIUPLOAD & sheet DIPILIH di kartu "Data Gudang"
# pada halaman Input Data Produksi (upload, save-selection), tapi
# DIEKSEKUSI (refresh) dari halaman Data Gudang BJB/BJL, TIDAK ikut
# run_all.py / tombol "Refresh Semua". Lihat import_engine.py bagian
# "VARIAN 4".
# --------------------------------------------------------------------------

@app.route("/api/gudang/sources", methods=["GET"])
def gudang_sources():
    """Status & pilihan folder/file/sheet yang tersimpan untuk source
    "Data Gudang" -- dipakai oleh kartu di Input Data Produksi (buat
    tahu apa yang sudah tersimpan) dan halaman Data Gudang BJB/BJL
    (buat nampilin status terakhir + sumber yang lagi aktif)."""
    cfg = import_engine.load_config()
    sources = cfg.get("sources", {})
    return jsonify({key: sources.get(key, {}) for key in import_engine.GUDANG_SOURCES})


@app.route("/api/gudang/upload", methods=["POST"])
def gudang_upload():
    """Multipart/form-data, field 'file'. Ganti dari cara lama (browse
    folder lokal di komputer user) -- server sekarang TIDAK PERNAH baca
    filesystem komputer user (tidak bisa, apalagi setelah di-deploy
    online), jadi browser yang kirim file-nya langsung lewat upload, baru
    server simpan & baca dari disknya sendiri. Dipanggil dari kartu "Data
    Gudang" di halaman Input Data Produksi begitu user pilih file lewat
    <input type="file">."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_gudang_upload(file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "filename": f"current{Path(file_storage.filename).suffix.lower() or '.xlsx'}", "sheets": sheets})


@app.route("/api/gudang/save-selection", methods=["POST"])
def gudang_save_selection():
    """Body: {filename, sheet}. Dipanggil setelah user selesai pilih
    sheet di kartu "Data Gudang" (Input Data Produksi), setelah file-nya
    diupload lewat /api/gudang/upload. HANYA menyimpan pilihan sheet ke
    config.json -- TIDAK menjalankan import."""
    body = request.get_json(force=True) or {}
    filename = str(body.get("filename", "")).strip()
    sheet = str(body.get("sheet", "")).strip()
    try:
        src = import_engine.save_gudang_selection(filename, sheet)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "source": src})


GUDANG_EXTRA_IMPORT_SCRIPTS = [
    ("form_st_2", "import_form_st_2.py"),
    ("val_2", "import_val_2.py"),
]


@app.route("/api/gudang/refresh", methods=["POST"])
def gudang_refresh():
    """Tombol Refresh di halaman Data Gudang BJB *atau* BJL -- keduanya
    memanggil endpoint yang sama ini. Tidak perlu body: folder/file/sheet
    dibaca dari config.json (hasil save-selection di kartu "Data Gudang",
    dan hasil upload di kartu "Form Serah Terima 2" / "Validasi 2").

    SATU klik di sini sekarang menjalankan TIGA proses (kalau salah satu
    gagal, yang lain TETAP dicoba -- sama polanya dengan
    refresh_import_validasi() di atas):
      1. run_gudang_import() -- tulis tab "API" (Monitor Bahan Baku) +
         klasifikasi BJB/BJL (classify_gudang_sheets.py). Dijalankan
         LANGSUNG (in-process, bukan subprocess) -- kalau ini gagal,
         seluruh refresh dianggap gagal (data Gudang BJB/BJL sendiri
         yang mau ditampilkan tidak ke-update).
      2. import_form_st_2.py -- kartu "Form Serah Terima 2" -> tab
         "FORM_ST_2" di spreadsheet Monitor Bahan Baku.
      3. import_val_2.py -- kartu "Validasi 2" -> tab "VAL_2" di
         spreadsheet yang sama.
    (2) & (3) dijalankan sebagai SUBPROCESS terpisah (python
    import_form_st_2.py / import_val_2.py) -- sama seperti
    VALIDASI_IMPORT_SCRIPTS di atas -- karena keduanya header-based
    (TARGET_HEADERS/HEADER_KEYWORDS ada di script masing2, lihat
    run_local_excel_import() di import_engine.py), bukan generik lewat
    satu fungsi seperti run_gudang_import()/run_update_stock_import().

    Kalau (1) Gudang gagal, response success=False. Kalau (1) berhasil
    tapi (2)/(3) ada yang gagal, response TETAP success=True (supaya
    tampilan Gudang BJB/BJL tetap ke-update) tapi errornya dikirim lewat
    'extra_errors' & detail per-script di 'extra_results'."""
    try:
        result = import_engine.run_gudang_import()
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    extra_results = []
    for source_key, script_name in GUDANG_EXTRA_IMPORT_SCRIPTS:
        script_path = BASE_DIR / script_name
        if not script_path.exists():
            extra_results.append({"source_key": source_key, "script": script_name, "status": "NOT FOUND", "rows_written": None, "log": ""})
            continue
        try:
            proc = subprocess.run(
                [sys.executable, str(script_path)],
                cwd=str(BASE_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
            )
            ok = proc.returncode == 0
        except Exception as e:
            extra_results.append({"source_key": source_key, "script": script_name, "status": f"FAILED ({e})", "rows_written": None, "log": ""})
            continue

        # Baca status/rows_written TERBARU dari config.json (ditulis oleh
        # set_import_result() di dalam run_local_excel_import()) -- lebih
        # akurat daripada parsing stdout script.
        try:
            _, src_after = import_engine.get_source(source_key)
        except KeyError:
            src_after = {}

        extra_results.append({
            "source_key": source_key,
            "script": script_name,
            "status": "OK" if ok else f"FAILED (exit {proc.returncode})",
            "rows_written": src_after.get("last_rows"),
            "last_error": src_after.get("last_error") if not ok else None,
            "log": (proc.stdout or "") + (proc.stderr or ""),
        })

    extra_errors = [
        f"{r['source_key']}: {r.get('last_error') or r['status']}"
        for r in extra_results if r["status"] != "OK"
    ]

    return jsonify({
        "success": True,
        "rows_written": result["rows_written"],
        "classification_errors": result["classification_errors"],
        "extra_results": extra_results,
        "extra_errors": extra_errors,
    })


# --------------------------------------------------------------------------
# 6c. UPDATE STOCK — SUMBER FILE (upload Excel langsung di kartu "Update
# Stock" pada halaman Input Data Produksi, ALTERNATIF dari cara lama
# paste link spreadsheet). Alurnya mirip Data Gudang di atas, bedanya
# sheet yang dicentang bisa lebih dari satu -- makanya SETELAH upload,
# pemilihan sheet-nya lewat modal "Pilih Sheet" GENERIK yang sama dengan
# source link lain (endpoint /api/produksi/sheets, TIDAK butuh endpoint
# "save-selection" terpisah seperti Data Gudang).
# --------------------------------------------------------------------------

@app.route("/api/update-stock-source", methods=["GET"])
def update_stock_source():
    """Status & pilihan file/sheet yang tersimpan untuk source
    'update_stock' -- dipakai kartu "Update Stock" (mode upload file) di
    halaman Input Data Produksi. Dipisah dari /api/produksi/sources
    karena source ini dikecualikan dari daftar situ (lihat komentar di
    produksi_sources())."""
    cfg = import_engine.load_config()
    return jsonify(cfg.get("sources", {}).get("update_stock", {}))


@app.route("/api/update-stock-source/upload", methods=["POST"])
def update_stock_source_upload():
    """Multipart/form-data, field 'file'. Sama alurnya seperti
    /api/gudang/upload, tapi untuk source 'update_stock' -- file
    disimpan ke server lalu dibalikin daftar nama sheet di dalamnya,
    supaya user bisa langsung centang sheet mana yang mau dipakai lewat
    modal "Pilih Sheet" (checkbox multi-select, sama dengan source link
    lain)."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_update_stock_upload(file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "sheets": sheets})


# --------------------------------------------------------------------------
# 6d. FORM SERAH TERIMA 2 & VALIDASI 2 — SUMBER FILE, alurnya PERSIS sama
# dengan "Update Stock" di 6c (upload Excel -> centang sheet lewat modal
# "Pilih Sheet" generik -> run_update_stock_import(source_key) ekstrak
# blok kolom lebar 10 & tumpuk). BEDA dari Update Stock: dua source ini
# TIDAK ikut "Refresh Semua" -- keduanya dijalankan otomatis bareng
# import Gudang tiap kali tombol Refresh di halaman Data Gudang BJB/BJL
# diklik (lihat gudang_refresh() di 6b, 3 proses sekali klik). Target
# tulisnya juga BUKAN spreadsheet Monitor Bahan Baku, tapi spreadsheet
# terpisah (lihat config.json: sources.form_st_2 / sources.val_2 ->
# target_id "1-ZyKSwXLzZaA6uNYRcpJNQZWX_ssYzvX45Z51xERipI", target_sheet
# "FORM_ST_2" / "VAL_2").
# --------------------------------------------------------------------------

@app.route("/api/form-st2-source", methods=["GET"])
def form_st2_source():
    """Status & pilihan file/sheet yang tersimpan untuk source 'form_st_2'
    -- dipakai kartu "Form Serah Terima 2" (mode upload file) di halaman
    Input Data Produksi."""
    cfg = import_engine.load_config()
    return jsonify(cfg.get("sources", {}).get("form_st_2", {}))


@app.route("/api/form-st2-source/upload", methods=["POST"])
def form_st2_source_upload():
    """Multipart/form-data, field 'file'. Sama alurnya seperti
    /api/update-stock-source/upload, tapi untuk source 'form_st_2'."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_local_excel_upload("form_st_2", file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "sheets": sheets})


@app.route("/api/val2-source", methods=["GET"])
def val2_source():
    """Status & pilihan file/sheet yang tersimpan untuk source 'val_2'
    -- dipakai kartu "Validasi 2" (mode upload file) di halaman Input
    Data Produksi."""
    cfg = import_engine.load_config()
    return jsonify(cfg.get("sources", {}).get("val_2", {}))


@app.route("/api/val2-source/upload", methods=["POST"])
def val2_source_upload():
    """Multipart/form-data, field 'file'. Sama alurnya seperti
    /api/update-stock-source/upload, tapi untuk source 'val_2'."""
    file_storage = request.files.get("file")
    if file_storage is None or not file_storage.filename:
        return jsonify({"success": False, "message": "Tidak ada file yang dikirim."}), 400
    try:
        sheets = import_engine.save_local_excel_upload("val_2", file_storage, file_storage.filename)
    except Exception as e:
        return jsonify({"success": False, "message": str(e)}), 400
    return jsonify({"success": True, "sheets": sheets})



# --------------------------------------------------------------------------
# 6d. WASTE REWIND — tampilan data dari spreadsheet REWIND_PY
# --------------------------------------------------------------------------
WASTE_REWIND_SPREADSHEET_ID = "1DnXtcMPkRdoadgO7ML7y7M9s4injmMTsKqEQ4BPrJxc"
WASTE_REWIND_SHEET_NAME = "REWIND_PY"

WASTE_REWIND_VISIBLE_HEADERS = [
    "JO",
    "SPK",
    "NO_JO",
    "Nama_Produk",
    "Planning_Meter",
    "Bahan_Awal_Printing_(Meter)",
    "Hasil_Slitting_(Rol)",
    "Meter_Hilang_Rewind",
    "Persentase_Waste_(%)",
    "Waste_Slitting_After_Rewind_Presentase",
]

_waste_rewind_spreadsheet_handle = {"sh": None}
_waste_rewind_spreadsheet_lock = threading.Lock()
_waste_rewind_cache = {"ts": 0.0, "headers": [], "rows": []}
_waste_rewind_cache_lock = threading.Lock()
_WASTE_REWIND_CACHE_TTL = int(os.environ.get("WASTE_REWIND_CACHE_TTL_SECONDS", "60"))


def _waste_rewind_spreadsheet():
    with _waste_rewind_spreadsheet_lock:
        if _waste_rewind_spreadsheet_handle["sh"] is not None:
            return _waste_rewind_spreadsheet_handle["sh"]

    client = get_client()
    sh = client.open_by_key(WASTE_REWIND_SPREADSHEET_ID)

    with _waste_rewind_spreadsheet_lock:
        _waste_rewind_spreadsheet_handle["sh"] = sh
    return sh


def _read_waste_rewind(force=False):
    now = time.time()
    with _waste_rewind_cache_lock:
        cached = dict(_waste_rewind_cache)
    if (
        not force
        and cached["headers"]
        and now - cached["ts"] < _WASTE_REWIND_CACHE_TTL
    ):
        return cached["headers"], cached["rows"]

    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()

    if not values:
        headers, rows = [], []
    else:
        headers = [str(x).strip() for x in values[0]]
        rows = []
        for row in values[1:]:
            padded = list(row) + [""] * max(0, len(headers) - len(row))
            row_obj = {}
            for i, header in enumerate(headers):
                if not header:
                    continue
                row_obj[header] = padded[i]
            if any(str(v).strip() for v in row_obj.values()):
                rows.append(row_obj)

    revised = _read_revisi_map(sh, headers) if headers else {}
    with _waste_rewind_cache_lock:
        _waste_rewind_cache["ts"] = now
        _waste_rewind_cache["headers"] = headers
        _waste_rewind_cache["rows"] = rows
        _waste_rewind_cache["revised"] = revised

    return headers, rows


@app.route("/api/waste-rewind", methods=["GET"])
def get_waste_rewind():
    """Ambil data REWIND_PY. Main table frontend boleh memakai
    WASTE_REWIND_VISIBLE_HEADERS, modal memakai seluruh headers."""
    force = str(request.args.get("refresh", "")).strip().lower() in ("1", "true", "yes")
    try:
        headers, rows = _read_waste_rewind(force=force)
        return jsonify({
            "success": True,
            "sheet": WASTE_REWIND_SHEET_NAME,
            "headers": headers,
            "visible_headers": WASTE_REWIND_VISIBLE_HEADERS,
            "rows": rows,
            "revised": dict(_waste_rewind_cache.get("revised") or {}),  # {"SPK||NO_JO": [kolom revisi]}
            "count": len(rows),
        })
    except Exception as e:
        return jsonify({
            "success": False,
            "message": f"Gagal membaca {WASTE_REWIND_SHEET_NAME}: {e}"
        }), 500



# --------------------------------------------------------------------------
# 6d.1 WASTE REWIND — aksi di modal Detail: Hitung Waste / Set Status Finish /
#      Revisi. Baris dicari berdasarkan pasangan (SPK, NO_JO) -- pasangan itu
#      unik di REWIND_PY (lihat _sync_rewind_kecil_spk_jo_into_rewind_py).
#
#   Hitung Waste     -> isi 4 kolom di REWIND_PY:
#        Waste_Slitting_Meter                    = Bahan_Awal_Printing_(Meter) - Hasil_Slitting_(Meter)
#        Persentase_Waste_(%)                    = |Waste_Slitting_Meter| / Bahan_Awal_Printing_(Meter)
#        Waste_Slitting_After_Rewind_Meter       = Bahan_Awal_Printing_(Meter)
#                                                  - (Hasil_Slitting_(Meter) - Meter_Jumbo_Hilang_Rewind)
#        Waste_Slitting_After_Rewind_Presentase  = |Waste_Slitting_After_Rewind_Meter| / Bahan_Awal_Printing_(Meter)
#   Set Status Finish -> simpan snapshot baris ke sheet REWIND_PY_FINISH, lalu
#        isi kolom Status di REWIND_PY = "Finish". Harus sudah Hitung Waste dulu.
#        Setelah Finish: tidak bisa Finish / Hitung Waste / Revisi / Hapus Revisi
#        (terkunci; dicek dari Status DAN keberadaan di REWIND_PY_FINISH).
#   Lepas Finish (/unfinish, KHUSUS Admin) -> kosongkan Status + hapus baris
#        SPK+NO_JO dari REWIND_PY_FINISH.
#   Revisi           -> user memilih/mengubah kolom di modal, HANYA kolom yang
#        berubah dikirim ke REWIND_PY_REVISI (append = riwayat) dan langsung
#        ditulis ke REWIND_PY. Di refresh, _apply_rewind_revisi_into_rewind_py()
#        menerapkan lagi sel revisi yang beda dari data hasil refresh.
#
# Header REWIND_PY_FINISH / REWIND_PY_REVISI dicocokkan PER NAMA KOLOM
# (bukan per posisi). "Jam" & "Nama_User" diisi otomatis; header kosong atau
# nama kolom kembar (mis. Waste_Slitting_After_Rewind_Presentase muncul 2x)
# cuma diisi di kemunculan PERTAMA, sisanya dibiarkan kosong.
# --------------------------------------------------------------------------
WASTE_REWIND_FINISH_SHEET = "REWIND_PY_FINISH"
WASTE_REWIND_REVISI_SHEET = "REWIND_PY_REVISI"
_WRW_CALC_COLS = (
    "Waste_Slitting_Meter",
    "Persentase_Waste_(%)",
    "Waste_Slitting_After_Rewind_Meter",
    "Waste_Slitting_After_Rewind_Presentase",
)


def _wib_now_str():
    """Waktu sekarang WIB (server bisa jalan di UTC, mis. di Render)."""
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Asia/Jakarta")
    except Exception:
        tz = timezone(timedelta(hours=7))
    return datetime.now(tz).strftime("%d/%m/%Y %H:%M:%S")


def _fmt_persen(fraction):
    """0.07518 -> '7,52%' (format Indonesia; Sheets membacanya sbg persen)."""
    return f"{fraction * 100:.2f}".replace(".", ",") + "%"


def hitung_waste_rewind(bahan_awal, hasil_meter, jumbo_hilang):
    """Balikin dict 4 kolom waste (string siap tulis ke sheet).
    ValueError kalau Bahan_Awal_Printing_(Meter) / Hasil_Slitting_(Meter)
    kosong atau Bahan Awal = 0. Kalau Meter_Jumbo_Hilang_Rewind kosong, dua
    kolom "After Rewind" dikosongkan."""
    if bahan_awal is None:
        raise ValueError("Bahan_Awal_Printing_(Meter) kosong, tidak bisa menghitung waste.")
    if hasil_meter is None:
        raise ValueError("Hasil_Slitting_(Meter) kosong, tidak bisa menghitung waste.")
    if bahan_awal == 0:
        raise ValueError("Bahan_Awal_Printing_(Meter) = 0, persentase waste tidak bisa dihitung.")
    fmt = import_engine._format_number
    waste = bahan_awal - hasil_meter
    out = {
        "Waste_Slitting_Meter": fmt(round(waste, 2)),
        "Persentase_Waste_(%)": _fmt_persen(abs(waste) / bahan_awal),
        "Waste_Slitting_After_Rewind_Meter": "",
        "Waste_Slitting_After_Rewind_Presentase": "",
    }
    if jumbo_hilang is not None:
        after = bahan_awal - (hasil_meter - jumbo_hilang)
        out["Waste_Slitting_After_Rewind_Meter"] = fmt(round(after, 2))
        out["Waste_Slitting_After_Rewind_Presentase"] = _fmt_persen(abs(after) / bahan_awal)
    return out


def _wrw_pad(row, width):
    return list(row) + [""] * max(0, width - len(row))


def _wrw_row_dict(header, row):
    row = _wrw_pad(row, len(header))
    d = {}
    for i, h in enumerate(header):
        if h and h not in d:
            d[h] = row[i]
    return d


def _wrw_find_pair(values, header, spk, no_jo):
    """Cari baris (nomor baris sheet 1-based, isi baris) dengan SPK & NO_JO
    yang sama. (None, None) kalau tidak ketemu."""
    col_spk = import_engine._find_col_index(header, "SPK")
    col_nojo = import_engine._find_col_index(header, "NO_JO")
    if col_spk is None or col_nojo is None:
        raise RuntimeError("Kolom SPK/NO_JO tidak ketemu di header sheet.")
    spk, no_jo = str(spk).strip(), str(no_jo).strip()
    for i, row in enumerate(values[1:], start=2):
        r = _wrw_pad(row, len(header))
        if str(r[col_spk]).strip() == spk and str(r[col_nojo]).strip() == no_jo:
            return i, r
    return None, None


def _wrw_load_target(spk, no_jo):
    if not str(spk).strip() or not str(no_jo).strip():
        raise ValueError("SPK / NO_JO kosong.")
    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        raise LookupError(f"Sheet {WASTE_REWIND_SHEET_NAME} kosong.")
    header = [str(h).strip() for h in values[0]]
    row_no, row = _wrw_find_pair(values, header, spk, no_jo)
    if row_no is None:
        raise LookupError(
            f"Baris SPK {spk} / NO_JO {no_jo} tidak ditemukan di {WASTE_REWIND_SHEET_NAME} "
            f"(mungkin baru di-refresh). Muat ulang halaman lalu coba lagi.")
    return sh, ws, header, row_no, row


def _wrw_open_sheet(sh, name):
    try:
        return sh.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        raise LookupError(f"Sheet {name} tidak ditemukan di spreadsheet. Buat dulu sheet-nya (lengkap dengan kepala tabel).")


def _wrw_build_snapshot(target_header, src_header, src_row, user, overrides=None):
    """Susun satu baris untuk sheet FINISH/REVISI mengikuti urutan header
    sheet tujuan. overrides = {nama_kolom: nilai} (mis. Status)."""
    norm = import_engine._norm
    src_row = _wrw_pad(src_row, len(src_header))
    src_map = {}
    for i, h in enumerate(src_header):
        k = norm(h)
        if k and k not in src_map:
            src_map[k] = src_row[i]
    over = {norm(k): v for k, v in (overrides or {}).items()}
    fixed = {norm("Jam"): _wib_now_str(), norm("Nama_User"): user}
    out, seen = [], set()
    for h in target_header:
        k = norm(h)
        if not k or k in seen:
            out.append("")
            continue
        seen.add(k)
        if k in fixed:
            out.append(fixed[k])
        elif k in over:
            out.append(over[k])
        else:
            out.append(src_map.get(k, ""))
    return out


def _wrw_write_row(ws, values, header, row_no, row_out):
    width = len(header)
    end_col = gspread.utils.rowcol_to_a1(1, width).rstrip("0123456789")
    if row_no > ws.row_count:
        ws.add_rows(row_no - ws.row_count)
    ws.batch_update(
        [{"range": f"A{row_no}:{end_col}{row_no}", "values": [row_out]}],
        value_input_option="USER_ENTERED",
    )


def _wrw_sheet_and_header(sh, sheet_name):
    ws = _wrw_open_sheet(sh, sheet_name)
    values = ws.get_all_values()
    if not values or not any(str(h).strip() for h in values[0]):
        raise LookupError(f"Kepala tabel sheet {sheet_name} kosong.")
    return ws, values, [str(h).strip() for h in values[0]]


def _wrw_append_finish(sh, src_header, src_row, user):
    """Tambah snapshot ke REWIND_PY_FINISH. Ditolak kalau pasangan (SPK,
    NO_JO) sudah ada di sana (mencegah dobel)."""
    ws, values, header = _wrw_sheet_and_header(sh, WASTE_REWIND_FINISH_SHEET)
    src = _wrw_pad(src_row, len(src_header))
    spk = src[import_engine._find_col_index(src_header, "SPK")]
    no_jo = src[import_engine._find_col_index(src_header, "NO_JO")]
    exist, _ = _wrw_find_pair(values, header, spk, no_jo)
    if exist is not None:
        raise ValueError(f"SPK {spk} / NO JO {no_jo} sudah ada di {WASTE_REWIND_FINISH_SHEET} (baris {exist}).")
    row_out = _wrw_build_snapshot(header, src_header, src_row, user, {"Status": "Finish"})
    row_no = len(values) + 1
    _wrw_write_row(ws, values, header, row_no, row_out)
    return row_no


# Kolom yang TIDAK boleh direvisi (kunci baris / dikunci sistem).
_WRW_REVISI_LOCKED = ("JO", "SPK", "NO_JO", "Status")
# Kolom di REWIND_PY_REVISI yang bukan data revisi (identitas / metadata).
_WRW_REVISI_SKIP = ("Jam", "Nama_User", "JO", "SPK", "NO_JO", "Status")


def _wrw_build_revisi_row(target_header, src_header, src_row, user, changes):
    """Baris untuk REWIND_PY_REVISI: Jam, Nama_User, JO, SPK, NO_JO + HANYA
    kolom yang direvisi (changes = {nama_kolom: nilai_baru}). Kolom lain
    dikosongkan -- saat refresh cuma kolom berisi ini yang dipakai."""
    norm = import_engine._norm
    src = _wrw_pad(src_row, len(src_header))
    src_map = {}
    for i, h in enumerate(src_header):
        k = norm(h)
        if k and k not in src_map:
            src_map[k] = src[i]
    target_keys = {norm(h) for h in target_header if norm(h)}
    absent = [k for k in changes if norm(k) not in target_keys]
    if absent:
        raise ValueError(f"Kolom {', '.join(absent)} tidak ada di header {WASTE_REWIND_REVISI_SHEET}.")
    ch = {norm(k): str(v).strip() for k, v in changes.items()}
    fixed = {
        norm("Jam"): _wib_now_str(),
        norm("Nama_User"): user,
        norm("JO"): src_map.get(norm("JO"), ""),
        norm("SPK"): src_map.get(norm("SPK"), ""),
        norm("NO_JO"): src_map.get(norm("NO_JO"), ""),
    }
    out, seen = [], set()
    for h in target_header:
        k = norm(h)
        if not k or k in seen:
            out.append("")
            continue
        seen.add(k)
        out.append(ch[k] if k in ch else fixed.get(k, ""))
    return out


def _wrw_append_revisi(sh, src_header, src_row, user, changes):
    ws, values, header = _wrw_sheet_and_header(sh, WASTE_REWIND_REVISI_SHEET)
    row_out = _wrw_build_revisi_row(header, src_header, src_row, user, changes)
    row_no = len(values) + 1
    _wrw_write_row(ws, values, header, row_no, row_out)
    return row_no


def _wrw_same_value(a, b):
    """Sama secara numerik (format Indonesia) atau teks persis."""
    a, b = str(a).strip(), str(b).strip()
    if a == b:
        return True
    pa, pb = import_engine._parse_flexible_number(a), import_engine._parse_flexible_number(b)
    return pa is not None and pb is not None and abs(pa - pb) < 0.005


def _wrw_is_finished(header, row):
    c = import_engine._find_col_index(header, "Status")
    return c is not None and str(row[c]).strip().lower() == "finish"


def _wrw_finish_locked(sh, header, row):
    """True kalau baris ini sudah Finish: Status di REWIND_PY = "Finish" ATAU
    pasangan (SPK, NO_JO) ada di REWIND_PY_FINISH. Cek kedua-duanya supaya
    kunci tidak lepas kalau Status di REWIND_PY sempat kosong (mis. di tengah
    refresh). Sheet FINISH belum ada -> cuma andalkan Status."""
    if _wrw_is_finished(header, row):
        return True
    try:
        ws_f = sh.worksheet(WASTE_REWIND_FINISH_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        return False
    values_f = ws_f.get_all_values()
    if not values_f:
        return False
    header_f = [str(h).strip() for h in values_f[0]]
    src = _wrw_pad(row, len(header))
    spk = src[import_engine._find_col_index(header, "SPK")]
    no_jo = src[import_engine._find_col_index(header, "NO_JO")]
    hit, _ = _wrw_find_pair(values_f, header_f, spk, no_jo)
    return hit is not None


def _wrw_is_admin(nama):
    """Cek ke sheet Login: apakah user dengan NAMA ini ber-role Admin.
    (Aplikasi ini belum punya sesi/token login, jadi nama dikirim dari
    frontend -- pengecekan ini menghindari cuma percaya string 'role'.)"""
    nama = str(nama or "").strip().lower()
    if not nama:
        return False
    ws = get_sheet("Login")
    for r in ws.get_all_records():
        low = {str(k).strip().lower(): v for k, v in r.items()}
        if str(low.get("nama", "")).strip().lower() == nama and str(low.get("role", "")).strip().lower() == "admin":
            return True
    return False


def _wrw_error_response(e):
    if isinstance(e, ValueError):
        return jsonify({"success": False, "message": str(e)}), 400
    if isinstance(e, LookupError):
        return jsonify({"success": False, "message": str(e)}), 404
    return jsonify({"success": False, "message": f"Gagal: {e}"}), 500


def _wrw_invalidate_cache():
    with _waste_rewind_cache_lock:
        _waste_rewind_cache["ts"] = 0.0


def _wrw_body():
    body = request.get_json(silent=True) or {}
    return (str(body.get("spk", "")).strip(), str(body.get("no_jo", "")).strip(),
            str(body.get("user", "")).strip() or "Tidak diketahui")


@app.route("/api/waste-rewind/hitung", methods=["POST"])
def waste_rewind_hitung():
    spk, no_jo, _user = _wrw_body()
    try:
        sh, ws, header, row_no, row = _wrw_load_target(spk, no_jo)
        if _wrw_finish_locked(sh, header, row):
            raise ValueError("Sudah Finish: waste terkunci. Minta Admin melepas status Finish kalau perlu diubah.")
        find = import_engine._find_col_index
        parse = import_engine._parse_flexible_number
        need = ("Bahan_Awal_Printing_(Meter)", "Hasil_Slitting_(Meter)", "Meter_Jumbo_Hilang_Rewind") + _WRW_CALC_COLS
        cols = {n: find(header, n) for n in need}
        missing = [n for n, c in cols.items() if c is None]
        if missing:
            raise RuntimeError(f"Kolom {', '.join(missing)} tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")
        res = hitung_waste_rewind(
            parse(row[cols["Bahan_Awal_Printing_(Meter)"]]),
            parse(row[cols["Hasil_Slitting_(Meter)"]]),
            parse(row[cols["Meter_Jumbo_Hilang_Rewind"]]),
        )
        ws.batch_update(
            [{"range": gspread.utils.rowcol_to_a1(row_no, cols[n] + 1), "values": [[res[n]]]} for n in _WRW_CALC_COLS],
            value_input_option="USER_ENTERED",
        )
        _wrw_invalidate_cache()
        new_row = _wrw_row_dict(header, ws.row_values(row_no))
        msg = "Waste dihitung."
        if not res["Waste_Slitting_After_Rewind_Meter"]:
            msg += " (Meter_Jumbo_Hilang_Rewind kosong, kolom After Rewind dikosongkan.)"
        return jsonify({"success": True, "message": msg, "row": new_row})
    except Exception as e:
        return _wrw_error_response(e)


@app.route("/api/waste-rewind/finish", methods=["POST"])
def waste_rewind_finish():
    spk, no_jo, user = _wrw_body()
    try:
        sh, ws, header, row_no, row = _wrw_load_target(spk, no_jo)
        find = import_engine._find_col_index
        col_status = find(header, "Status")
        col_waste = find(header, "Waste_Slitting_Meter")
        if col_status is None or col_waste is None:
            raise RuntimeError(f"Kolom Status/Waste_Slitting_Meter tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")
        if _wrw_is_finished(header, row):
            raise ValueError("Sudah Finish. Status tidak bisa di-Finish-kan dua kali (hanya Admin yang bisa melepasnya).")
        if not str(row[col_waste]).strip():
            raise ValueError("Klik \"Hitung Waste\" dulu sebelum Set Status Finish.")
        # 1) simpan ke REWIND_PY_FINISH (ditolak kalau SPK+NO_JO sudah ada di sana)
        saved_row = _wrw_append_finish(sh, header, row, user)
        # 2) baru set Status di REWIND_PY
        ws.batch_update(
            [{"range": gspread.utils.rowcol_to_a1(row_no, col_status + 1), "values": [["Finish"]]}],
            value_input_option="USER_ENTERED",
        )
        _wrw_invalidate_cache()
        new_row = _wrw_row_dict(header, ws.row_values(row_no))
        return jsonify({
            "success": True,
            "message": f"Status Finish tersimpan & data dikirim ke {WASTE_REWIND_FINISH_SHEET} (baris {saved_row}).",
            "row": new_row,
        })
    except Exception as e:
        return _wrw_error_response(e)


@app.route("/api/waste-rewind/unfinish", methods=["POST"])
def waste_rewind_unfinish():
    """KHUSUS ADMIN: lepas status Finish (kosongkan Status di REWIND_PY) dan
    hapus baris SPK+NO_JO itu dari REWIND_PY_FINISH. Kolom waste/persentase
    di REWIND_PY dibiarkan (akan ikut terhapus di refresh berikutnya)."""
    spk, no_jo, user = _wrw_body()
    try:
        if not _wrw_is_admin(user):
            return jsonify({"success": False, "message": "Hanya Admin yang boleh melepas status Finish."}), 403
        sh, ws, header, row_no, row = _wrw_load_target(spk, no_jo)
        col_status = import_engine._find_col_index(header, "Status")
        if col_status is None:
            raise RuntimeError(f"Kolom Status tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")
        # 1) hapus dari REWIND_PY_FINISH dulu (kalau gagal, Status belum berubah)
        ws_f, values_f, header_f = _wrw_sheet_and_header(sh, WASTE_REWIND_FINISH_SHEET)
        deleted = 0
        while True:
            hit, _ = _wrw_find_pair(values_f, header_f, spk, no_jo)
            if hit is None:
                break
            ws_f.delete_rows(hit)
            del values_f[hit - 1]
            deleted += 1
        # 2) kosongkan Status di REWIND_PY
        ws.batch_update(
            [{"range": gspread.utils.rowcol_to_a1(row_no, col_status + 1), "values": [[""]]}],
            value_input_option="USER_ENTERED",
        )
        _wrw_invalidate_cache()
        new_row = _wrw_row_dict(header, ws.row_values(row_no))
        return jsonify({
            "success": True,
            "message": f"Status Finish dilepas; {deleted} baris dihapus dari {WASTE_REWIND_FINISH_SHEET}.",
            "row": new_row,
        })
    except Exception as e:
        return _wrw_error_response(e)


@app.route("/api/waste-rewind/revisi", methods=["POST"])
def waste_rewind_revisi():
    """Body: {spk, no_jo, user, changes: {nama_kolom: nilai_baru}}.
    Cuma kolom yang benar-benar berubah yang dikirim ke REWIND_PY_REVISI
    (bersama Jam, Nama_User, JO, SPK, NO_JO), lalu nilai barunya juga
    langsung ditulis ke REWIND_PY supaya tampilan konsisten. Saat refresh,
    kolom-kolom ini diterapkan lagi (lihat _apply_rewind_revisi_into_rewind_py).
    Ditolak kalau baris sudah Finish."""
    spk, no_jo, user = _wrw_body()
    body = request.get_json(silent=True) or {}
    changes_in = body.get("changes")
    try:
        if not isinstance(changes_in, dict) or not changes_in:
            raise ValueError("Tidak ada perubahan yang dikirim.")
        sh, ws, header, row_no, row = _wrw_load_target(spk, no_jo)
        if _wrw_finish_locked(sh, header, row):
            raise ValueError("Sudah Finish: data tidak bisa direvisi. Minta Admin melepas status Finish dulu.")
        find = import_engine._find_col_index
        locked = {import_engine._norm(x) for x in _WRW_REVISI_LOCKED}
        changes = {}  # nama header sheet -> nilai baru
        for name, new_val in changes_in.items():
            c = find(header, str(name))
            if c is None:
                raise ValueError(f"Kolom {name} tidak ada di {WASTE_REWIND_SHEET_NAME}.")
            real_name = header[c]
            if import_engine._norm(real_name) in locked:
                raise ValueError(f"Kolom {real_name} dikunci dan tidak bisa direvisi.")
            new_val = str(new_val).strip()
            if _wrw_same_value(row[c], new_val):
                continue  # tidak berubah
            if new_val == "":
                raise ValueError(f"Kolom {real_name} tidak boleh dikosongkan (isi nilai baru atau batalkan perubahannya).")
            changes[real_name] = new_val
        if not changes:
            raise ValueError("Tidak ada nilai yang berubah.")
        # 1) REWIND_PY_REVISI  2) REWIND_PY
        saved_row = _wrw_append_revisi(sh, header, row, user, changes)
        ws.batch_update(
            [{"range": gspread.utils.rowcol_to_a1(row_no, find(header, n) + 1), "values": [[v]]} for n, v in changes.items()],
            value_input_option="USER_ENTERED",
        )
        _wrw_invalidate_cache()
        new_row = _wrw_row_dict(header, ws.row_values(row_no))
        return jsonify({
            "success": True,
            "message": f"{len(changes)} kolom revisi dikirim ke {WASTE_REWIND_REVISI_SHEET} (baris {saved_row}).",
            "row": new_row,
        })
    except Exception as e:
        return _wrw_error_response(e)


def _wrw_revisi_col_map(rev_header, header):
    """Pasangan (indeks kolom di REWIND_PY_REVISI, indeks kolom di REWIND_PY)
    untuk kolom-kolom DATA revisi (bukan Jam/Nama_User/JO/SPK/NO_JO/Status).
    Nama kembar / kosong: hanya kemunculan pertama yang dipakai."""
    norm = import_engine._norm
    find = import_engine._find_col_index
    skip = {norm(x) for x in _WRW_REVISI_SKIP}
    out, seen = [], set()
    for j, h in enumerate(rev_header):
        k = norm(h)
        if not k or k in seen:
            continue
        seen.add(k)
        if k in skip:
            continue
        c = find(header, h)
        if c is not None:
            out.append((j, c))
    return out


def _wrw_pair_key(spk, no_jo):
    return f"{str(spk).strip()}||{str(no_jo).strip()}"


def _read_revisi_map(sh, header):
    """{"SPK||NO_JO": [nama kolom REWIND_PY yang punya revisi]} -- dipakai
    frontend buat menandai sel hasil revisi dengan *. Sheet REVISI tidak ada
    / gagal dibaca -> {} (tampilan tetap jalan)."""
    try:
        rev = sh.worksheet(WASTE_REWIND_REVISI_SHEET).get_all_values()
    except Exception as e:
        print(f"[revisi map] dilewati: {e}")
        return {}
    if len(rev) < 2:
        return {}
    rev_header = [str(h).strip() for h in rev[0]]
    rc_spk = import_engine._find_col_index(rev_header, "SPK")
    rc_nojo = import_engine._find_col_index(rev_header, "NO_JO")
    if rc_spk is None or rc_nojo is None:
        return {}
    col_map = _wrw_revisi_col_map(rev_header, header)
    out = {}
    for row in rev[1:]:
        r = _wrw_pad(row, len(rev_header))
        key = _wrw_pair_key(r[rc_spk], r[rc_nojo])
        for j, c in col_map:
            if str(r[j]).strip():
                out.setdefault(key, set()).add(header[c])
    return {k: sorted(v) for k, v in out.items()}


@app.route("/api/waste-rewind/hapus-revisi", methods=["POST"])
def waste_rewind_hapus_revisi():
    """Hapus SEMUA baris revisi milik pasangan (SPK, NO_JO) dari
    REWIND_PY_REVISI. Nilai default di REWIND_PY dikembalikan oleh refresh
    Waste Rewind berikutnya (frontend menjalankannya otomatis) -- nilai default
    hanya bisa dihitung ulang dari data mentah, jadi tidak ditebak di sini."""
    spk, no_jo, _user = _wrw_body()
    try:
        if not spk or not no_jo:
            raise ValueError("SPK / NO_JO kosong.")
        sh, _ws, py_header, _row_no, py_row = _wrw_load_target(spk, no_jo)
        if _wrw_finish_locked(sh, py_header, py_row):
            raise ValueError("Sudah Finish: revisi tidak bisa dihapus. Minta Admin melepas status Finish dulu.")
        ws_r, values, header = _wrw_sheet_and_header(sh, WASTE_REWIND_REVISI_SHEET)
        deleted = 0
        while True:
            hit, _ = _wrw_find_pair(values, header, spk, no_jo)
            if hit is None:
                break
            ws_r.delete_rows(hit)
            del values[hit - 1]
            deleted += 1
        if deleted == 0:
            raise LookupError(f"Tidak ada revisi untuk SPK {spk} / NO JO {no_jo} di {WASTE_REWIND_REVISI_SHEET}.")
        _wrw_invalidate_cache()
        return jsonify({
            "success": True,
            "deleted": deleted,
            "message": f"{deleted} baris revisi dihapus dari {WASTE_REWIND_REVISI_SHEET}.",
        })
    except Exception as e:
        return _wrw_error_response(e)



# --------------------------------------------------------------------------
# 6d.2 WASTE REWIND — CEK STOK (tombol di modal Detail). Read-only, jalan
#      di baris Finish ATAUPUN belum (tidak dikunci seperti Hitung Waste/
#      Revisi -- lihat catatan "kalau sudah finish masih bisa cek stok").
#
#      DIUBAH: sekarang SEMUA sumber dicocokkan lewat NOMOR JO (suffix
#      angka di belakang, huruf nyangkut diabaikan -- _fstl_suffix_key()),
#      BUKAN lewat Nama_Produk lagi seperti sebelumnya (chatbot_engine.
#      query_stok_gudang() sudah tidak dipakai di endpoint ini karena itu
#      cari berdasar produk). Baris sumber "JO" di sheet REWIND_PY (format
#      lengkap, mis. "JO/26/VIII/18/3034") dipakai buat tahu suffix ANGKA
#      (3034) sekaligus TAHUN (26 -> 2026) target-nya; "NO_JO" (cuma angka,
#      mis. "3034") dipakai sebagai fallback suffix kalau kolom "JO" kosong.
#
#      Sumber & cara cocokkan:
#        - Validasi (VAL_1)          : suffix JO SAJA (tanpa tahun). Baris
#          ditampilkan kalau kolom JUMLAH atau JUMLAH_MASUK_REWIND ada
#          isinya (teks/angka apa saja, bukan cuma "-"/kosong).
#        - Form Serah Terima (FORM_ST_1) : suffix JO SAJA (tanpa tahun).
#          Baris DITAMPILKAN HANYA kalau MASUK_REWIND ada isinya (bukan
#          "-") DAN HASIL_RIWEN kosong/"-" (masih di-antrian rewind,
#          belum ada hasilnya) -- baris lain disembunyikan.
#        - Gudang Barang Jadi Baru/Lama (BJB_KATEGORI/BJL_KATEGORI): suffix
#          JO **+ TAHUN** (soalnya nomor JO bisa kepakai ulang di tahun
#          beda -- JO 3034 tahun 2026 != JO 3034 tahun lain). Baris tanpa
#          JO sama sekali (mis. "TIDAK ADA NO JO") otomatis tidak relevan
#          karena tidak ada suffix buat dicocokkan. Tahun baris diambil
#          dari teks JO_DAN_STATUS (macam2 format, lihat _stok_extract_tahun),
#          fallback dari kolom JO itu sendiri kalau formatnya lengkap juga.
#          Kalau tahun baris tidak bisa ditebak sama sekali, baris TETAP
#          ditampilkan (lebih baik kelihatan lalu dicek manual daripada
#          hilang) -- ini asumsi, longgarkan/ketatkan lagi kalau ternyata
#          kebanyakan noise.
#
#      VAL_1 & FORM_ST_1 ada di spreadsheet STOK_SPREADSHEET_ID_A (gspread
#      "A"), BJB_KATEGORI & BJL_KATEGORI ada di spreadsheet
#      STOK_SPREADSHEET_ID_B (gspread "B", spreadsheet "Monitor Bahan
#      Baku" -- sama dengan target import_form_st_2.py/import_val_2.py &
#      classify_gudang_sheets.py, lihat bagian 6b/6d di atas).
# --------------------------------------------------------------------------
STOK_SPREADSHEET_ID_A = "1FRWpza_fa65jt8-n1-rN4rFFrfNBLixRxOLS_uUgYYU"
STOK_SPREADSHEET_ID_B = "1-ZyKSwXLzZaA6uNYRcpJNQZWX_ssYzvX45Z51xERipI"

BJB_KATEGORI_SHEET_NAME = "BJB_KATEGORI"
BJL_KATEGORI_SHEET_NAME = "BJL_KATEGORI"

_stok_spreadsheet_handle_a = {"sh": None}
_stok_spreadsheet_lock_a = threading.Lock()
_stok_spreadsheet_handle_b = {"sh": None}
_stok_spreadsheet_lock_b = threading.Lock()

# Header yang ditampilkan ke frontend, PERSIS urutan & nama kolom yang
# diminta (lihat wrwStokTable() di index.html -- ambil Object.keys(rows[0])
# apa adanya jadi header tabel, jadi urutan dict di sini = urutan kolom).
WRW_STOK_VALIDASI_COLUMNS = ("AREA", "JO", "NAMA_PRODUK", "JUMLAH", "JUMLAH_MASUK_REWIND")
WRW_STOK_FORM_ST_COLUMNS = (
    "TANGGAL", "JO", "NAMA_PRODUK", "JUMLAH_MASUK_GBJ", "BERAT/KG",
    "STATUS", "MASUK_REWIND", "HASIL_RIWEN", "DARI_SLITTING",
)
WRW_STOK_KATEGORI_COLUMNS = (
    "UKURAN_PRODUK", "PRODUK", "SISA_STOCK_AKHIR", "JO_DAN_STATUS",
    "KETERANGAN", "JO", "STATUS", "KATEGORI",
)

# Keyword pencarian kolom per nama field di atas -- dipisah dari nama field
# tampilan karena beberapa header asli sheet ejaannya bisa beda2 (mis.
# "BERAT/KG" vs "BERAT_KG" vs "BERAT KG").
_WRW_STOK_COL_KEYWORDS = {
    "AREA": ("AREA",),
    "JO": ("JO",),
    "NAMA_PRODUK": ("NAMA_PRODUK", "NAMA PRODUK", "NAMA"),
    "JUMLAH": ("JUMLAH",),
    "JUMLAH_MASUK_REWIND": ("JUMLAH_MASUK_REWIND", "JUMLAH MASUK REWIND"),
    "TANGGAL": ("TANGGAL",),
    "JUMLAH_MASUK_GBJ": ("JUMLAH_MASUK_GBJ", "JUMLAH MASUK GBJ", "JUMLAH_MASUK"),
    "BERAT/KG": ("BERAT/KG", "BERAT_KG", "BERAT KG", "BERAT"),
    "STATUS": ("STATUS",),
    "MASUK_REWIND": ("MASUK_REWIND", "MASUK REWIND"),
    "HASIL_RIWEN": ("HASIL_RIWEN", "HASIL RIWEN"),
    "DARI_SLITTING": ("DARI_SLITTING", "DARI SLITTING"),
    "UKURAN_PRODUK": ("UKURAN_PRODUK", "UKURAN PRODUK"),
    "PRODUK": ("PRODUK",),
    "SISA_STOCK_AKHIR": ("SISA_STOCK_AKHIR", "SISA STOCK AKHIR"),
    "JO_DAN_STATUS": ("JO_DAN_STATUS", "JO DAN STATUS"),
    "KETERANGAN": ("KETERANGAN",),
    "KATEGORI": ("KATEGORI",),
}


def _stok_spreadsheet_a():
    """Spreadsheet gspread "A" -- VAL_1 & FORM_ST_1."""
    with _stok_spreadsheet_lock_a:
        if _stok_spreadsheet_handle_a["sh"] is not None:
            return _stok_spreadsheet_handle_a["sh"]
    client = get_client()
    sh = client.open_by_key(STOK_SPREADSHEET_ID_A)
    with _stok_spreadsheet_lock_a:
        _stok_spreadsheet_handle_a["sh"] = sh
    return sh


def _stok_spreadsheet_b():
    """Spreadsheet gspread "B" -- BJB_KATEGORI & BJL_KATEGORI (spreadsheet
    Monitor Bahan Baku, sama dengan target classify_gudang_sheets.py)."""
    with _stok_spreadsheet_lock_b:
        if _stok_spreadsheet_handle_b["sh"] is not None:
            return _stok_spreadsheet_handle_b["sh"]
    client = get_client()
    sh = client.open_by_key(STOK_SPREADSHEET_ID_B)
    with _stok_spreadsheet_lock_b:
        _stok_spreadsheet_handle_b["sh"] = sh
    return sh


def _stok_col(header, field_name):
    """Cari index kolom (0-based) di `header` buat field logis `field_name`
    (key di WRW_STOK_*_COLUMNS / _WRW_STOK_COL_KEYWORDS), lewat
    _fstl_find_col() (exact match dulu, baru substring)."""
    keywords = _WRW_STOK_COL_KEYWORDS.get(field_name, (field_name,))
    return _fstl_find_col(header, *keywords)


def _stok_cell(row, col_idx):
    return row[col_idx].strip() if col_idx is not None and col_idx < len(row) else ""


def _stok_has_content(text):
    """True kalau sel ada isinya beneran (bukan kosong / "-")."""
    t = str(text or "").strip()
    return bool(t) and t != "-"


_STOK_JO_YEAR_RE = re.compile(r"JO(?:-[A-Z]+)?\s*/\s*(\d{2})\s*/", re.IGNORECASE)
_STOK_PAREN_YEAR_RE = re.compile(r"\((?:[A-Za-z]{3,9})?\s*(\d{4})\)")
_STOK_BARE_YEAR_RE = re.compile(r"\b(20[0-3]\d)\b")


def _stok_extract_tahun(text):
    """Coba tebak TAHUN (4 digit, mis. 2026) dari sebuah teks JO/JO_DAN_STATUS.
    Nyoba beberapa pola sekaligus soalnya penulisan JO_DAN_STATUS di
    BJB_KATEGORI/BJL_KATEGORI (apalagi BJL) tidak konsisten:
      - "JO/26/VIII/18/3034"        -> segmen ke-2 ("26") = tahun 2026
      - "JO-DDCT/26/IX/2/3217"      -> sama, segmen ke-2 tetap tahun
      - "JO 1161 (MAR2022)"         -> 4 digit tahun di dalam kurung
      - "JO, 2506 2022"             -> 4 digit tahun lepas di teks
    Balikin None kalau tidak ada pola yang cocok (biar baris TETAP
    ditampilkan -- lihat catatan di komentar blok di atas)."""
    text = str(text or "")
    m = _STOK_JO_YEAR_RE.search(text)
    if m:
        return 2000 + int(m.group(1))
    m = _STOK_PAREN_YEAR_RE.search(text)
    if m:
        return int(m.group(1))
    m = _STOK_BARE_YEAR_RE.search(text)
    if m:
        return int(m.group(1))
    return None


def _wrw_cek_stok_val1(target_suffix):
    """VAL_1 (gspread A) -- cocok suffix JO saja, baris ditampilkan kalau
    JUMLAH atau JUMLAH_MASUK_REWIND ada isinya."""
    ws = _stok_spreadsheet_a().worksheet("VAL_1")
    values = ws.get_all_values()
    if not values:
        return {"ditemukan": False, "rows": [], "pesan": "Sheet VAL_1 kosong.", "total_stok": ""}
    header = [str(h).strip() for h in values[0]]
    cols = {f: _stok_col(header, f) for f in WRW_STOK_VALIDASI_COLUMNS}
    col_jo = cols["JO"]
    if col_jo is None:
        return {"ditemukan": False, "rows": [], "pesan": "Kolom JO tidak ketemu di VAL_1.", "total_stok": ""}

    rows_out = []
    for row in values[1:]:
        jo_cell = _stok_cell(row, col_jo)
        if not jo_cell or _fstl_suffix_key(jo_cell) != target_suffix:
            continue
        jumlah = _stok_cell(row, cols["JUMLAH"])
        jumlah_masuk_rewind = _stok_cell(row, cols["JUMLAH_MASUK_REWIND"])
        if not (_stok_has_content(jumlah) or _stok_has_content(jumlah_masuk_rewind)):
            continue
        rows_out.append({
            "AREA": _stok_cell(row, cols["AREA"]),
            "JO": jo_cell,
            "NAMA_PRODUK": _stok_cell(row, cols["NAMA_PRODUK"]),
            "JUMLAH": jumlah,
            "JUMLAH_MASUK_REWIND": jumlah_masuk_rewind,
        })
    if not rows_out:
        return {"ditemukan": False, "rows": [],
                "pesan": f"Tidak ditemukan baris VAL_1 dengan suffix JO '{target_suffix}' yang ada isi JUMLAH/JUMLAH_MASUK_REWIND.",
                "total_stok": ""}
    return {"ditemukan": True, "rows": rows_out, "pesan": None, "total_stok": ""}


def _wrw_cek_stok_form_st(target_suffix):
    """FORM_ST_1 (gspread A) -- cocok suffix JO saja. Baris ditampilkan
    HANYA kalau MASUK_REWIND ada isinya (bukan "-") DAN HASIL_RIWEN kosong
    atau "-" (masih di-antrian rewind, belum ada hasilnya) -- baris lain
    (MASUK_REWIND kosong, atau HASIL_RIWEN sudah keisi) disembunyikan."""
    ws = _stok_spreadsheet_a().worksheet(FORM_ST1_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return {"ditemukan": False, "rows": [], "pesan": f"Sheet {FORM_ST1_SHEET_NAME} kosong."}
    header = [str(h).strip() for h in values[0]]
    cols = {f: _stok_col(header, f) for f in WRW_STOK_FORM_ST_COLUMNS}
    col_jo = cols["JO"]
    if col_jo is None:
        return {"ditemukan": False, "rows": [], "pesan": f"Kolom JO tidak ketemu di {FORM_ST1_SHEET_NAME}."}

    rows_out = []
    for row in values[1:]:
        jo_cell = _stok_cell(row, col_jo)
        if not jo_cell or jo_cell == "-" or _fstl_suffix_key(jo_cell) != target_suffix:
            continue
        masuk_rewind = _stok_cell(row, cols["MASUK_REWIND"])
        hasil_riwen = _stok_cell(row, cols["HASIL_RIWEN"])
        if not (_stok_has_content(masuk_rewind) and not _stok_has_content(hasil_riwen)):
            continue  # bukan baris "masih di-antrian rewind" -- gausah ditampilkan
        rows_out.append({
            "TANGGAL": _stok_cell(row, cols["TANGGAL"]),
            "JO": jo_cell,
            "NAMA_PRODUK": _stok_cell(row, cols["NAMA_PRODUK"]),
            "JUMLAH_MASUK_GBJ": _stok_cell(row, cols["JUMLAH_MASUK_GBJ"]),
            "BERAT/KG": _stok_cell(row, cols["BERAT/KG"]),
            "STATUS": _stok_cell(row, cols["STATUS"]),
            "MASUK_REWIND": masuk_rewind,
            "HASIL_RIWEN": hasil_riwen,
            "DARI_SLITTING": _stok_cell(row, cols["DARI_SLITTING"]),
        })
    if not rows_out:
        return {"ditemukan": False, "rows": [],
                "pesan": f"Tidak ditemukan baris {FORM_ST1_SHEET_NAME} dengan suffix JO '{target_suffix}' yang relevan ditampilkan."}
    return {"ditemukan": True, "rows": rows_out, "pesan": None}


def _wrw_cek_stok_kategori(sheet_name, target_suffix, target_tahun):
    """BJB_KATEGORI / BJL_KATEGORI (gspread B) -- cocok suffix JO **+
    TAHUN** (lihat catatan panjang di komentar blok di atas). Balikin juga
    total_stok_utuh (jumlah SISA_STOCK_AKHIR numerik dari baris yang
    ketemu) & perlu_review (daftar baris yang KATEGORI-nya PERLU_REVIEW) --
    ASUMSI definisi, sesuaikan lagi kalau beda dari yang dimaksud."""
    try:
        ws = _stok_spreadsheet_b().worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return {"ditemukan": False, "rows": [], "pesan": f"Sheet {sheet_name} tidak ditemukan.",
                "total_stok_utuh": "", "perlu_review": ""}
    values = ws.get_all_values()
    if not values:
        return {"ditemukan": False, "rows": [], "pesan": f"Sheet {sheet_name} kosong.",
                "total_stok_utuh": "", "perlu_review": ""}
    header = [str(h).strip() for h in values[0]]
    cols = {f: _stok_col(header, f) for f in WRW_STOK_KATEGORI_COLUMNS}
    col_jo = cols["JO"]
    if col_jo is None:
        return {"ditemukan": False, "rows": [], "pesan": f"Kolom JO tidak ketemu di {sheet_name}.",
                "total_stok_utuh": "", "perlu_review": ""}

    rows_out = []
    total_stok_utuh = 0.0
    ada_angka = False
    review_count = 0
    for row in values[1:]:
        jo_cell = _stok_cell(row, col_jo)
        if not jo_cell or jo_cell == "-":
            continue  # tidak ada JO -> tidak relevan (mis. "TIDAK ADA NO JO")
        suffix = _fstl_suffix_key(jo_cell)
        if not suffix or suffix != target_suffix:
            continue
        jo_dan_status = _stok_cell(row, cols["JO_DAN_STATUS"])
        tahun_baris = _stok_extract_tahun(jo_dan_status) or _stok_extract_tahun(jo_cell)
        if target_tahun and tahun_baris and tahun_baris != target_tahun:
            continue  # suffix sama tapi tahunnya beda -> bukan JO yang sama

        kategori = _stok_cell(row, cols["KATEGORI"])
        sisa = _stok_cell(row, cols["SISA_STOCK_AKHIR"])
        num = import_engine._parse_flexible_number(sisa) if sisa else None
        if isinstance(num, (int, float)):
            total_stok_utuh += num
            ada_angka = True
        if kategori.strip().upper() == "PERLU_REVIEW":
            review_count += 1

        rows_out.append({
            "UKURAN_PRODUK": _stok_cell(row, cols["UKURAN_PRODUK"]),
            "PRODUK": _stok_cell(row, cols["PRODUK"]),
            "SISA_STOCK_AKHIR": sisa,
            "JO_DAN_STATUS": jo_dan_status,
            "KETERANGAN": _stok_cell(row, cols["KETERANGAN"]),
            "JO": jo_cell,
            "STATUS": _stok_cell(row, cols["STATUS"]),
            "KATEGORI": kategori,
        })
    if not rows_out:
        return {"ditemukan": False, "rows": [],
                "pesan": f"Tidak ditemukan baris {sheet_name} dengan JO '{target_suffix}'"
                         + (f" tahun {target_tahun}" if target_tahun else "") + ".",
                "total_stok_utuh": "", "perlu_review": ""}
    return {
        "ditemukan": True,
        "rows": rows_out,
        "pesan": None,
        "total_stok_utuh": total_stok_utuh if ada_angka else "",
        "perlu_review": review_count,
    }


@app.route("/api/waste-rewind/cek-stok", methods=["POST"])
def waste_rewind_cek_stok():
    spk, no_jo, _user = _wrw_body()
    try:
        _sh, _ws, header, _row_no, row = _wrw_load_target(spk, no_jo)
        find = import_engine._find_col_index
        c_nama = find(header, "Nama_Produk")
        c_nojo = find(header, "NO_JO")
        # sengaja pakai _fstl_find_col (exact match diprioritaskan) buat
        # kolom "JO", BUKAN import_engine._find_col_index -- header sheet
        # ini juga punya kolom "NO_JO" yang mengandung teks "JO" sebagai
        # substring, jadi kalau pencariannya substring-based bisa salah
        # kepilih kolom "NO_JO".
        c_jo = _fstl_find_col(header, "JO")
        nama_produk = row[c_nama].strip() if c_nama is not None and c_nama < len(row) else ""
        no_jo_val = row[c_nojo].strip() if c_nojo is not None and c_nojo < len(row) else str(no_jo)
        jo_full = row[c_jo].strip() if c_jo is not None and c_jo < len(row) else ""
        if not nama_produk and not no_jo_val and not jo_full:
            raise ValueError("JO/Nama_Produk kosong di baris ini, tidak bisa cari stok.")

        target_suffix = _fstl_suffix_key(jo_full or no_jo_val)
        target_tahun = _stok_extract_tahun(jo_full)
        if not target_suffix:
            raise ValueError(f"Tidak bisa membaca nomor JO dari baris ini (JO='{jo_full}', NO_JO='{no_jo_val}').")

        validasi_out = _wrw_cek_stok_val1(target_suffix)
        form_st_out = _wrw_cek_stok_form_st(target_suffix)
        bjb_out = _wrw_cek_stok_kategori(BJB_KATEGORI_SHEET_NAME, target_suffix, target_tahun)
        bjl_out = _wrw_cek_stok_kategori(BJL_KATEGORI_SHEET_NAME, target_suffix, target_tahun)

        return jsonify({
            "success": True,
            "nama_produk": nama_produk,
            "no_jo": no_jo_val,
            "jo": jo_full or no_jo_val,
            "tahun": target_tahun,
            "validasi": validasi_out,
            "form_st": form_st_out,
            "bjb": bjb_out,
            "bjl": bjl_out,
        })
    except Exception as e:
        return _wrw_error_response(e)


def _apply_rewind_revisi_into_rewind_py():
    """LANGKAH TERAKHIR refresh Waste Rewind (setelah semua sync lain, karena
    sync lain menghitung ulang kolom turunan dari data mentah). Baca
    REWIND_PY_REVISI; untuk tiap baris revisi (urut dari atas -> yang lebih
    baru menimpa), tiap kolom revisi yang TERISI dan BEDA dari nilai di
    REWIND_PY dipakai (per sel, bukan seluruh baris). Kolom identitas/Status
    tidak pernah ditimpa. Balikin jumlah sel yang diubah."""
    sh = _waste_rewind_spreadsheet()
    try:
        ws_rev = sh.worksheet(WASTE_REWIND_REVISI_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        return 0
    rev = ws_rev.get_all_values()
    if len(rev) < 2:
        return 0
    rev_header = [str(h).strip() for h in rev[0]]
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if len(values) < 2:
        return 0
    header = [str(h).strip() for h in values[0]]
    find = import_engine._find_col_index
    norm = import_engine._norm
    rc_spk, rc_nojo = find(rev_header, "SPK"), find(rev_header, "NO_JO")
    c_spk, c_nojo = find(header, "SPK"), find(header, "NO_JO")
    if None in (rc_spk, rc_nojo, c_spk, c_nojo):
        return 0

    row_by_key = {}
    for i, row in enumerate(values[1:], start=2):
        r = _wrw_pad(row, len(header))
        row_by_key[(str(r[c_spk]).strip(), str(r[c_nojo]).strip())] = i

    col_map = _wrw_revisi_col_map(rev_header, header)  # (kolom di REVISI, kolom di REWIND_PY)

    desired = {}
    for row in rev[1:]:
        r = _wrw_pad(row, len(rev_header))
        rn = row_by_key.get((str(r[rc_spk]).strip(), str(r[rc_nojo]).strip()))
        if rn is None:
            continue
        for j, c in col_map:
            v = str(r[j]).strip()
            if v != "":
                desired[(rn, c)] = v

    updates = []
    for (rn, c), v in desired.items():
        cur = _wrw_pad(values[rn - 1], len(header))[c]
        if _wrw_same_value(cur, v):
            continue
        updates.append({"range": gspread.utils.rowcol_to_a1(rn, c + 1), "values": [[v]]})
    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        _wrw_invalidate_cache()
    return len(updates)


def _read_rewind_finish_restore_map():
    """Isi Status + 4 kolom waste dari REWIND_PY_FINISH, per (SPK, NO_JO).
    Dipakai refresh penuh REWIND_PY supaya baris yang sudah Finish tidak
    kehilangan Status/hasil hitungnya. Sheet tidak ada / gagal dibaca ->
    {} (refresh tetap jalan)."""
    sh = _waste_rewind_spreadsheet()
    try:
        ws = sh.worksheet(WASTE_REWIND_FINISH_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        print(f"[restore Finish] sheet {WASTE_REWIND_FINISH_SHEET} tidak ada, dilewati")
        return {}
    # Sengaja TIDAK ditelan: kalau gagal baca (mis. 429), refresh dibatalkan
    # SEBELUM REWIND_PY dihapus, supaya Status/waste yang terkunci tidak hilang.
    values = ws.get_all_values()
    if not values:
        return {}
    header = [str(h).strip() for h in values[0]]
    find = import_engine._find_col_index
    c_spk, c_nojo = find(header, "SPK"), find(header, "NO_JO")
    if c_spk is None or c_nojo is None:
        return {}
    names = ("Status",) + _WRW_CALC_COLS
    cols = {n: find(header, n) for n in names}
    out = {}
    for row in values[1:]:
        r = _wrw_pad(row, len(header))
        key = (str(r[c_spk]).strip(), str(r[c_nojo]).strip())
        if not key[0] or not key[1]:
            continue
        out[key] = {n: r[c] for n, c in cols.items() if c is not None}  # baris terbawah menang
    return out


# --------------------------------------------------------------------------
# 6e. REWIND KECIL — REFRESH PENUH (hapus baris 2 ke bawah + tulis ulang)
#     sheet REWIND_PY dari sheet mentah REWIND_PY_RAW
# --------------------------------------------------------------------------
# BUKAN tabel/list terpisah -- pasangan (SPK, NO_JO) unik dari sheet MENTAH
# "REWIND_PY_RAW" (hasil import_rewind_kecil.py, spreadsheet SAMA dengan REWIND_PY)
# ditulis LANGSUNG jadi baris di REWIND_PY itu sendiri, supaya tetap tampil
# di SATU tabel yang sama yang sudah dibaca lewat /api/waste-rewind.
# NO_JO di sini SAMA PERSIS dengan kolom "JO" di sheet REWIND_PY_RAW -- tidak ada
# kolom "NO_JO" terpisah di sheet sumber, cuma beda label kolom di REWIND_PY.
#
# !!! PERUBAHAN PERILAKU (atas permintaan user) !!!
# Versi SEBELUMNYA cuma APPEND pasangan (SPK, NO_JO) yang belum ada & tidak
# pernah menyentuh baris lama sama sekali -- supaya kolom lain yang diisi
# manual/formula (Persentase_Waste_(%), Meter_Hilang_Rewind,
# Hasil_Slitting_(Rol), Waste_Slitting_After_Rewind_Presentase, dst) aman.
#
# Versi SEKARANG SENGAJA menghapus SEMUA baris data (baris 2 ke bawah,
# SELURUH kolom -- termasuk kolom manual/formula tadi) lalu menulis ulang
# dari nol tiap kali tombol Refresh ditekan, supaya sheet-nya benar2
# "ke-refresh" total sesuai isi REWIND_PY_RAW saat itu. Konsekuensinya:
# apa pun yang pernah diisi manual/dihitung pakai formula di baris data
# REWIND_PY akan HILANG setiap refresh dan HARUS diisi ulang. Baris header
# (baris 1) TIDAK ikut terhapus.
#
# Aturan:
#   1. Hanya baris REWIND dengan TANGGAL >= REWIND_KECIL_START_DATE yang
#      dipertimbangkan, lalu di-unique-kan per pasangan (SPK, JO).
#   1b. Kalau SPK atau JO pada baris itu bukan angka murni (teks seperti
#       "EX", "RETUR", "xxx"), SELURUH baris itu diabaikan -- bukan cuma
#       sisi yang teksnya, supaya tidak ada pasangan (SPK, NO_JO) yang
#       jomplang (satu sisi keambil, sisi lain kosong/tidak match).
#   2. SEMUA baris data lama di REWIND_PY (baris 2 s.d. baris terakhir,
#      seluruh lebar sheet) DIHAPUS lebih dulu -- lihat catatan
#      "PERUBAHAN PERILAKU" di atas.
#   3. Baris baru ditulis untuk SETIAP pasangan (SPK, NO_JO) unik yang ada
#      di REWIND_PY_RAW saat ini. Selain kolom SPK & NO_JO, kolom JO
#      (lengkap), Nama_Produk, Planning_Order, Planning_Meter & Potongan
#      JUGA ikut diisi otomatis -- dicari lewat NO_JO dicocokkan ke suffix
#      (angka belakang) kode JO di sheet JO_1 punya spreadsheet FSTL
#      (FSTL_SPREADSHEET_ID, lihat blok "FSTL -- LAMPIRAN WASTE" di
#      bawah), caranya SAMA PERSIS kayak lookup JO/NAMA/ORDER/METER di
#      halaman Update Stock (import_engine.sync_update_stock_from_jo()):
#      kolom F JO_1 = kode JO lengkap, kolom G = NAMA (KEMASAN), header
#      "ORDER" = Planning Order, header "METER" = Planning Meter, header
#      "POTONGAN" = kolom Potongan -- lihat _fstl_lookup_jo1_by_suffix_map().
#      Kalau NO_JO tidak ketemu di JO_1 (belum ada / suffix tidak match),
#      kolom2 itu dikosongkan (menunggu diisi manual). Kolom lain di luar
#      daftar ini (Persentase_Waste_(%), dst) SELALU kosong sampai diisi
#      manual lagi setelah refresh.
#   4. Dipanggil dari _run_rewind_kecil_worker() (tombol Refresh di halaman
#      Waste Rewind). import_rewind_kecil.py TIDAK lagi dijalankan dari sini
#      -- REWIND_PY_RAW diisi lewat Refresh Semua / Jalankan Satu Script di
#      halaman Input Data Produksi (hindari limit 429). Setelah ini,
#      _sync_bahan_awal_printing_into_rewind_py() jalan lagi buat isi
#      ulang kolom Bahan_Awal_Printing_(Meter) di baris-baris yang baru
#      ditulis.
REWIND_KECIL_RAW_SHEET_NAME = "REWIND_PY_RAW"
REWIND_KECIL_START_DATE = date(2026, 9, 1)  # 01/09/2026
_RW_KECIL_NUMERIC_RE = re.compile(r"\d+")  # SPK & NO_JO harus SELURUHNYA angka -- kalau ada huruf (EX, RETUR, xxx, dll) dianggap teks & baris diabaikan


def _read_rewind_kecil_spk_jo():
    """Baca sheet mentah REWIND_PY_RAW, filter TANGGAL >= REWIND_KECIL_START_DATE,
    balikin list {"spk", "noJo"} unik. Selalu baca langsung dari sheet
    (dipanggil sekali per proses refresh, tidak perlu cache sendiri)."""
    sh = _waste_rewind_spreadsheet()  # spreadsheet sama dgn REWIND_PY, handle dipakai bareng
    ws = sh.worksheet(REWIND_KECIL_RAW_SHEET_NAME)
    values = ws.get_all_values()

    rows = []
    if values:
        header = [str(h).strip() for h in values[0]]
        col_tanggal = import_engine._find_col_index(header, "TANGGAL")
        col_spk = import_engine._find_col_index(header, "SPK")
        col_jo = import_engine._find_col_index(header, "JO")

        seen = set()
        for raw_row in values[1:]:
            def _cell(idx, _row=raw_row):
                return str(_row[idx]).strip() if idx is not None and idx < len(_row) else ""

            tgl_parsed = import_engine._parse_date_flexible(_cell(col_tanggal))
            if tgl_parsed is None or tgl_parsed < REWIND_KECIL_START_DATE:
                continue

            spk = _cell(col_spk)
            jo = _cell(col_jo)
            if not spk and not jo:
                continue

            # SPK & NO_JO seharusnya berupa angka. Kalau salah satu berisi
            # teks (mis. "EX", "RETUR", "xxx"), seluruh baris ini DIABAIKAN
            # -- bukan cuma sisi yang teks -- supaya tidak ada baris baru di
            # REWIND_PY yang cuma kolom SPK atau NO_JO-nya saja terisi
            # (jomplang, pasangannya hilang).
            if not (_RW_KECIL_NUMERIC_RE.fullmatch(spk) and _RW_KECIL_NUMERIC_RE.fullmatch(jo)):
                continue

            key = (spk, jo)
            if key in seen:
                continue
            seen.add(key)
            rows.append({"spk": spk, "noJo": jo})

    return rows


def _fstl_lookup_jo1_by_suffix_map():
    """Baca sheet JO_1 di spreadsheet FSTL (FSTL_SPREADSHEET_ID), balikin
    dict {suffix_key (angka belakang kode JO): {"jo","nama","order",
    "meter","potongan"}}.

    Dipakai buat auto-isi kolom JO/Nama_Produk/Planning_Order/
    Planning_Meter/Potongan waktu baris SPK/NO_JO baru di-append ke
    REWIND_PY -- lihat _sync_rewind_kecil_spk_jo_into_rewind_py().

    Kolom JO (F) & NAMA/KEMASAN (G) dipakai lewat FSTL_JO1_COL_JO /
    FSTL_JO1_COL_PRODUK (sama seperti fstl_lookup_produk()). Kolom
    ORDER/METER/POTONGAN dicari lewat NAMA HEADER-nya sendiri (bukan
    index tetap) pakai _fstl_find_col(), soalnya sengaja disamakan
    caranya dengan sync_update_stock_from_jo() di import_engine.py
    (halaman Update Stock) yang juga baca kolom "ORDER"/"METER" dari
    JO_1, ditambah kolom "POTONGAN".

    Kalau ada beberapa baris JO_1 dengan suffix sama, baris yang
    PALING BAWAH (jadi paling baru) yang dipakai -- overwrite biasa
    lewat urutan loop dari atas ke bawah."""
    try:
        sh = _fstl_spreadsheet()
        rows = _fstl_get_sheet_values(sh, FSTL_JO1_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        return {}
    if not rows:
        return {}

    header = rows[0]
    col_jo = FSTL_JO1_COL_JO
    col_nama = FSTL_JO1_COL_PRODUK
    col_order = _fstl_find_col(header, "ORDER")
    col_meter = _fstl_find_col(header, "METER")
    col_potongan = _fstl_find_col(header, "POTONGAN")

    def _cell(row, idx):
        return str(row[idx]).strip() if idx is not None and idx < len(row) else ""

    lookup = {}
    for row in rows[1:]:
        jo_cell = _cell(row, col_jo)
        if not jo_cell:
            continue
        key = _fstl_suffix_key(jo_cell)
        if key == "" or key is None:
            continue
        lookup[key] = {
            "jo": jo_cell,
            "nama": _cell(row, col_nama),
            "order": _cell(row, col_order),
            "meter": _cell(row, col_meter),
            "potongan": _cell(row, col_potongan),
        }
    return lookup


def _lp1_bahan_awal_printing_lookup():
    """Baca sheet LP_1 (spreadsheet FSTL), balikin dict {(spk_key, jo_key):
    total_meter_awal} buat auto-isi kolom Bahan_Awal_Printing_(Meter) di
    REWIND_PY -- lihat _sync_bahan_awal_printing_into_rewind_py().

    Header "SPK/JO" di LP_1 isinya kode gabungan, mis.
    "2674/26/VIII/8/3095" (segmen PALING DEPAN = SPK "2674", segmen PALING
    BELAKANG = JO "3095") atau bentuk pendek "3685/3123" (SPK "3685" /
    JO "3123") -- segmen tengah (kode bulan/urutan romawi dst, kalau ada)
    diabaikan sepenuhnya, cuma segmen pertama & terakhir yang dipakai.

    Untuk tiap baris LP_1 yang kolom "Urutan_Proses"-nya bernilai 1, kolom
    "Meter_Awal" dijumlahkan ke pasangan (SPK, JO) itu -- kalau ada
    beberapa baris LP_1 dengan pasangan SPK/JO sama & Urutan_Proses 1,
    nilai Meter_Awal-nya DIJUMLAHKAN (bukan dipakai satu baris saja),
    sesuai definisi user."""
    try:
        sh = _fstl_spreadsheet()
        rows = _fstl_get_sheet_values(sh, FSTL_LP1_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        return {}
    if not rows:
        return {}

    header = rows[0]
    col_spkjo = _fstl_find_col(header, "SPK/JO")
    col_urutan = _fstl_find_col(header, "Urutan_Proses")
    col_meter_awal = _fstl_find_col(header, "Meter_Awal")
    if col_spkjo is None or col_urutan is None or col_meter_awal is None:
        return {}

    def _cell(row, idx):
        return str(row[idx]).strip() if idx is not None and idx < len(row) else ""

    lookup = {}
    for row in rows[1:]:
        spkjo_cell = _cell(row, col_spkjo)
        if not spkjo_cell:
            continue

        urutan_val = import_engine._parse_flexible_number(_cell(row, col_urutan))
        if urutan_val != 1:
            continue  # cuma proses pertama (Urutan_Proses = 1) yang dipakai

        segments = [s.strip() for s in spkjo_cell.split("/") if s.strip()]
        if len(segments) < 2:
            continue  # bukan format "SPK/.../JO" atau "SPK/JO" yang valid

        spk_key = import_engine._numeric_key_prefix(segments[0])
        jo_key = import_engine._numeric_key_prefix(segments[-1])
        if spk_key == "" or jo_key == "":
            continue

        meter_val = import_engine._parse_flexible_number(_cell(row, col_meter_awal)) or 0.0
        key = (spk_key, jo_key)
        lookup[key] = lookup.get(key, 0.0) + meter_val

    return lookup


def _sync_bahan_awal_printing_into_rewind_py():
    """Isi ulang kolom Bahan_Awal_Printing_(Meter) di REWIND_PY, untuk tiap
    baris yang SPK & NO_JO-nya (keduanya harus angka murni, sama seperti
    aturan sinkron SPK/NO_JO) cocok dengan pasangan SPK/JO hasil
    _lp1_bahan_awal_printing_lookup() (dari LP_1, difilter Urutan_Proses = 1,
    Meter_Awal dijumlahkan kalau cocok lebih dari satu baris).

    Baris yang TIDAK ketemu pasangannya di LP_1 dibiarkan apa adanya (TIDAK
    dikosongkan) -- dianggap belum ada laporan produksi Printing-nya, bukan
    berarti harus ditulis 0. Baris yang nilainya sudah sama persis juga TIDAK
    ditulis ulang (hemat kuota API). Balikin jumlah sel yang benar-benar
    diupdate.

    Dipanggil dari _run_rewind_kecil_worker() (tombol Refresh di halaman
    Waste Rewind), SETELAH _sync_rewind_kecil_spk_jo_into_rewind_py() --
    supaya baris SPK/NO_JO yang baru saja di-append juga langsung kebagian
    nilai Bahan_Awal_Printing-nya di refresh yang sama."""
    lookup = _lp1_bahan_awal_printing_lookup()
    if not lookup:
        return 0

    sh = _waste_rewind_spreadsheet()  # spreadsheet sama dgn REWIND_PY, handle dipakai bareng
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    col_spk = import_engine._find_col_index(header, "SPK")
    col_nojo = import_engine._find_col_index(header, "NO_JO")
    col_bahan = import_engine._find_col_index(header, "Bahan_Awal_Printing_(Meter)")
    if col_spk is None or col_nojo is None or col_bahan is None:
        raise RuntimeError(
            f"Kolom SPK/NO_JO/Bahan_Awal_Printing_(Meter) tidak ketemu di header {WASTE_REWIND_SHEET_NAME}"
        )

    updates = []
    for i, row in enumerate(values[1:], start=2):  # baris 2 = data pertama di sheet
        spk = row[col_spk].strip() if col_spk < len(row) else ""
        nojo = row[col_nojo].strip() if col_nojo < len(row) else ""
        if not (spk.isdigit() and nojo.isdigit()):
            continue  # SPK/NO_JO harus angka murni, sama seperti aturan sinkron SPK/NO_JO

        total = lookup.get((
            import_engine._numeric_key_prefix(spk),
            import_engine._numeric_key_prefix(nojo),
        ))
        if total is None:
            continue  # tidak ketemu pasangannya di LP_1 -- biarkan sel apa adanya

        new_value = import_engine._format_number(total)
        current = row[col_bahan].strip() if col_bahan < len(row) else ""
        if current == new_value:
            continue  # sudah sama, tidak perlu ditulis ulang

        updates.append({
            "range": gspread.utils.rowcol_to_a1(i, col_bahan + 1),
            "values": [[new_value]],
        })

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        # invalidate cache REWIND_PY biar GET /api/waste-rewind berikutnya
        # baca nilai Bahan_Awal_Printing_(Meter) yang baru
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0

    return len(updates)


PRINTING_SHEET_NAMES = ["PRINTING_1", "PRINTING_2", "PRINTING_3", "PRINTING_4", "PRINTING_5"]
PRINTING_METER_HEADER = "METER_AKHIR_JADI"

# Nama kolom tujuan di REWIND_PY, urutan SAMA dengan PRINTING_SHEET_NAMES
# (index ke-0 = PRINTING_1 -> kolom Printing_1, dst).
REWIND_PY_PRINTING_COLS = ["Printing_1", "Printing_2", "Printing_3", "Printing_4", "Printing_5"]
REWIND_PY_TOTAL_PRINTING_COL = "Total_Hasil_Printing"


def _printing_meter_lookup(sheet_name):
    """Baca satu sheet PRINTING_X (spreadsheet FSTL), balikin dict
    {(spk_key, jo_key): total_meter_akhir_jadi} -- SPK & JO di sini kolom
    TERPISAH (bukan kode gabungan "SPK/JO" macam LP_1/SL_1), jadi TIDAK
    ada split by '/' sama sekali, langsung dicocokkan apa adanya lewat
    kolom "SPK" & "JO" masing-masing.

    Kalau sheet-nya tidak ada sama sekali di spreadsheet FSTL (mis.
    PRINTING_1, yang sumber datanya memang belum ada) atau header
    SPK/JO/METER_AKHIR_JADI tidak ketemu, balikin dict kosong -- ini yang
    bikin _sync_printing_kolom_into_rewind_py() nulis 0 buat sheet itu
    (lihat pemanggilnya)."""
    try:
        sh = _fstl_spreadsheet()
        rows = _fstl_get_sheet_values(sh, sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        return {}
    if not rows:
        return {}

    header = rows[0]
    col_spk = import_engine._find_col_index(header, "SPK")
    col_jo = import_engine._find_col_index(header, "JO")
    col_meter = import_engine._find_col_index(header, PRINTING_METER_HEADER)
    if col_spk is None or col_jo is None or col_meter is None:
        return {}
    max_col = max(col_spk, col_jo, col_meter)

    lookup = {}  # (spk_key, jo_key) -> total (float)
    for row in rows[1:]:  # lewati header
        if len(row) <= max_col:
            continue
        spk_cell = str(row[col_spk]).strip()
        jo_cell = str(row[col_jo]).strip()
        if not spk_cell or spk_cell == "-" or not jo_cell or jo_cell == "-":
            continue

        spk_key = import_engine._numeric_key_prefix(spk_cell)
        jo_key = import_engine._numeric_key_prefix(jo_cell)
        if spk_key == "" or jo_key == "":
            continue

        val = import_engine._parse_flexible_number(row[col_meter])
        if val is None:
            continue

        key = (spk_key, jo_key)
        lookup[key] = lookup.get(key, 0.0) + val

    return lookup


def _sync_printing_kolom_into_rewind_py():
    """Isi kolom Printing_1..Printing_5 & Total_Hasil_Printing di
    REWIND_PY, untuk tiap baris yang SPK & NO_JO-nya (keduanya harus
    angka murni, sama seperti aturan sinkron SPK/NO_JO lainnya):
      - Printing_1..5 = jumlah METER_AKHIR_JADI dari sheet PRINTING_1..5
        (spreadsheet FSTL) yang SPK & JO-nya (kolom TERPISAH, BUKAN kode
        gabungan macam SL_1/LP_1) sama-sama cocok. Sheet yang tidak
        ketemu/tidak ada (mis. PRINTING_1, sumber datanya belum ada)
        ATAU pasangan SPK/JO-nya tidak ketemu di situ -> ditulis 0 (BEDA
        dari Bahan_Awal_Printing_(Meter)/Hasil_Slitting yang dibiarkan
        apa adanya kalau tidak ketemu -- di sini SENGAJA ditulis 0,
        sesuai permintaan user).
      - Total_Hasil_Printing = jumlah Printing_1 s/d Printing_5 yang baru
        saja dihitung di atas (bukan dibaca ulang dari sheet).

    Tiap kolom dicek TERPISAH: kalau nilainya sudah sama persis, kolom
    itu saja yang tidak ditulis ulang (hemat kuota API). Kolom yang
    header-nya tidak ketemu di REWIND_PY dilewati begitu saja (tidak
    menggagalkan kolom lain). Balikin jumlah SEL yang benar-benar
    diupdate (gabungan Printing_1..5 + Total_Hasil_Printing).

    Dipanggil dari _run_rewind_kecil_worker() (tombol Refresh di halaman
    Waste Rewind), sama seperti sync-sync lainnya."""
    lookups = [_printing_meter_lookup(name) for name in PRINTING_SHEET_NAMES]
    # Sengaja TIDAK early-return kalau semua lookup kosong (beda dari
    # sync-sync lain di atas) -- di sini kolom tetap harus ditulis 0
    # (bukan dibiarkan apa adanya) kalau memang tidak ketemu, sesuai
    # permintaan user (lihat docstring).

    sh = _waste_rewind_spreadsheet()  # spreadsheet sama dgn REWIND_PY, handle dipakai bareng
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    col_spk = import_engine._find_col_index(header, "SPK")
    col_nojo = import_engine._find_col_index(header, "NO_JO")
    if col_spk is None or col_nojo is None:
        raise RuntimeError(f"Kolom SPK/NO_JO tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    col_printing = [import_engine._find_col_index(header, name) for name in REWIND_PY_PRINTING_COLS]
    col_total = import_engine._find_col_index(header, REWIND_PY_TOTAL_PRINTING_COL)

    updates = []
    for i, row in enumerate(values[1:], start=2):  # baris 2 = data pertama di sheet
        spk = row[col_spk].strip() if col_spk < len(row) else ""
        nojo = row[col_nojo].strip() if col_nojo < len(row) else ""
        if not (spk.isdigit() and nojo.isdigit()):
            continue  # SPK/NO_JO harus angka murni, sama seperti aturan sinkron SPK/NO_JO

        key = (import_engine._numeric_key_prefix(spk), import_engine._numeric_key_prefix(nojo))

        per_line_totals = []
        for col_idx, lookup in zip(col_printing, lookups):
            total = lookup.get(key, 0.0)
            per_line_totals.append(total)
            if col_idx is None:
                continue
            new_value = import_engine._format_number(total)
            current = row[col_idx].strip() if col_idx < len(row) else ""
            if current != new_value:
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(i, col_idx + 1),
                    "values": [[new_value]],
                })

        if col_total is not None:
            total_value = import_engine._format_number(sum(per_line_totals))
            current = row[col_total].strip() if col_total < len(row) else ""
            if current != total_value:
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(i, col_total + 1),
                    "values": [[total_value]],
                })

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        # invalidate cache REWIND_PY biar GET /api/waste-rewind berikutnya
        # baca nilai Printing_1..5/Total_Hasil_Printing yang baru
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0

    return len(updates)
SL1_COL_UP_HEADER = "UP"
SL1_COL_TOTAL_METER_HEADER = "TOTAL_METER"


def _sl1_slitting_lookup():
    """Baca sheet SL_1 (spreadsheet FSTL) SEKALI, balikin dict
    {(spk_key, jo_key): {"sl_matches": [...], "up_values": [...],
    "total_meter": float}} -- dipakai bareng buat isi TIGA kolom di
    REWIND_PY sekaligus (lihat _sync_slitting_kolom_into_rewind_py()):
      - Hasil_Slitting_(Rol)   <- import_engine._compute_hasil_slitting(sl_matches)
      - UP_Slitting            <- import_engine._modus_value(up_values)
      - Hasil_Slitting_(Meter) <- jumlah TOTAL_METER

    BEDA dari kolom HASIL SLITTING di halaman Data Validasi
    (import_engine.sync_validasi_header(), yang mencocokkan HANYA lewat
    suffix/angka belakang kode JO): di sini dicocokkan lewat PASANGAN
    LENGKAP (SPK, JO) -- segmen PALING DEPAN & PALING BELAKANG dari
    kolom "SPK/JO" SL_1, persis logic yang sama dengan
    _lp1_bahan_awal_printing_lookup() di atas (segmen tengah, kalau ada,
    diabaikan).

    UP & TOTAL_METER dikumpulkan dari SEMUA baris SL_1 yang pasangan
    SPK/JO-nya cocok, TERLEPAS baris itu punya HASIL_ROL/METER_ROL valid
    atau tidak (beda syarat dari sl_matches, yang cuma ikut baris dengan
    HASIL_ROL & METER_ROL dua-duanya valid) -- soalnya UP & TOTAL_METER
    adalah kolom independen, tidak terikat ke perhitungan modus rol."""
    try:
        sh = _fstl_spreadsheet()
        rows = _fstl_get_sheet_values(sh, import_engine.SL_SOURCE_SHEET_NAME)
    except gspread.exceptions.WorksheetNotFound:
        return {}
    if not rows:
        return {}

    header = rows[0]
    col_jo = import_engine._find_col_index(header, "SPK/JO")
    col_hasil_rol = import_engine._find_col_index(header, "HASIL_ROL")
    col_meter_rol = import_engine._find_col_index(header, "METER/ROL")
    col_up = import_engine._find_col_index(header, SL1_COL_UP_HEADER)
    col_total_meter = import_engine._find_col_index(header, SL1_COL_TOTAL_METER_HEADER)
    if col_jo is None:
        col_jo = import_engine.SL_COL_JO_FALLBACK
    if col_hasil_rol is None:
        col_hasil_rol = import_engine.SL_COL_HASIL_ROL_FALLBACK
    if col_meter_rol is None:
        col_meter_rol = import_engine.SL_COL_METER_ROL_FALLBACK
    # col_up / col_total_meter TIDAK punya fallback (posisinya tidak
    # didokumentasikan di tempat lain) -- kalau header-nya tidak ketemu,
    # ya sudah, bagian itu saja yang tidak terisi (lihat pemakaiannya di
    # bawah, None-checked satu-satu, bukan bikin seluruh lookup gagal).
    needed_cols = [c for c in (col_jo, col_hasil_rol, col_meter_rol, col_up, col_total_meter) if c is not None]
    max_col = max(needed_cols)

    lookup = {}  # (spk_key, jo_key) -> {"sl_matches": [...], "up_values": [...], "total_meter": float}
    for row in rows[1:]:  # lewati header
        if len(row) <= max_col:
            continue
        jo_cell = str(row[col_jo]).strip()
        if not jo_cell or jo_cell == "-":
            continue

        segments = [s.strip() for s in jo_cell.split("/") if s.strip()]
        if len(segments) < 2:
            continue  # bukan format "SPK/.../JO" atau "SPK/JO" yang valid

        spk_key = import_engine._numeric_key_prefix(segments[0])
        jo_key = import_engine._numeric_key_prefix(segments[-1])
        if spk_key == "" or jo_key == "":
            continue

        key = (spk_key, jo_key)
        entry = lookup.setdefault(key, {"sl_matches": [], "up_values": [], "total_meter": 0.0})

        k_val = import_engine._parse_flexible_number(row[col_hasil_rol])
        o_val = import_engine._parse_flexible_number(row[col_meter_rol])
        if k_val is not None and o_val is not None:
            entry["sl_matches"].append((k_val, import_engine._format_number(o_val)))

        if col_up is not None and col_up < len(row):
            entry["up_values"].append(row[col_up])

        if col_total_meter is not None and col_total_meter < len(row):
            tm_val = import_engine._parse_flexible_number(row[col_total_meter])
            if tm_val is not None:
                entry["total_meter"] += tm_val

    return lookup


def _sync_slitting_kolom_into_rewind_py():
    """Isi ulang TIGA kolom di REWIND_PY sekaligus (satu kali baca SL_1,
    satu kali batch_update), untuk tiap baris yang SPK & NO_JO-nya
    (keduanya harus angka murni, sama seperti aturan sinkron SPK/NO_JO &
    Bahan_Awal_Printing_(Meter)) cocok dengan pasangan SPK/JO hasil
    _sl1_slitting_lookup():
      - Hasil_Slitting_(Rol)   = import_engine._compute_hasil_slitting(sl_matches)
                                 (format modus/non-modus, mis. '42 + 3@530')
      - UP_Slitting            = import_engine._modus_value(up_values)
                                 (nilai UP paling sering muncul; kosong
                                 kalau tidak ada modus tunggal)
      - Hasil_Slitting_(Meter) = jumlah TOTAL_METER dari semua baris SL_1
                                 yang pasangan SPK/JO-nya cocok

    Baris yang TIDAK ketemu pasangannya di SL_1 dibiarkan apa adanya
    (TIDAK dikosongkan). Tiap kolom dicek TERPISAH: kalau nilainya sudah
    sama persis, kolom itu saja yang tidak ditulis ulang (hemat kuota
    API) -- satu baris bisa saja cuma 1-2 dari 3 kolomnya yang berubah.
    Kolom yang header-nya tidak ketemu di sheet dilewati begitu saja
    (tidak menggagalkan kolom lain). Balikin jumlah SEL yang benar-benar
    diupdate (gabungan ketiga kolom).

    Dipanggil dari _run_rewind_kecil_worker() (tombol Refresh di halaman
    Waste Rewind), sama seperti _sync_bahan_awal_printing_into_rewind_py()."""
    lookup = _sl1_slitting_lookup()
    if not lookup:
        return 0

    sh = _waste_rewind_spreadsheet()  # spreadsheet sama dgn REWIND_PY, handle dipakai bareng
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    col_spk = import_engine._find_col_index(header, "SPK")
    col_nojo = import_engine._find_col_index(header, "NO_JO")
    if col_spk is None or col_nojo is None:
        raise RuntimeError(f"Kolom SPK/NO_JO tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    col_hasil_rol = import_engine._find_col_index(header, "Hasil_Slitting_(Rol)")
    col_up = import_engine._find_col_index(header, "UP_Slitting")
    col_hasil_meter = import_engine._find_col_index(header, "Hasil_Slitting_(Meter)")

    updates = []
    for i, row in enumerate(values[1:], start=2):  # baris 2 = data pertama di sheet
        spk = row[col_spk].strip() if col_spk < len(row) else ""
        nojo = row[col_nojo].strip() if col_nojo < len(row) else ""
        if not (spk.isdigit() and nojo.isdigit()):
            continue  # SPK/NO_JO harus angka murni, sama seperti aturan sinkron SPK/NO_JO

        entry = lookup.get((
            import_engine._numeric_key_prefix(spk),
            import_engine._numeric_key_prefix(nojo),
        ))
        if entry is None:
            continue  # tidak ketemu pasangannya di SL_1 -- biarkan sel apa adanya

        if col_hasil_rol is not None:
            new_value = import_engine._compute_hasil_slitting(entry["sl_matches"])
            current = row[col_hasil_rol].strip() if col_hasil_rol < len(row) else ""
            if current != new_value:
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(i, col_hasil_rol + 1),
                    "values": [[new_value]],
                })

        if col_up is not None:
            new_value = import_engine._modus_value(entry["up_values"])
            current = row[col_up].strip() if col_up < len(row) else ""
            if current != new_value:
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(i, col_up + 1),
                    "values": [[new_value]],
                })

        if col_hasil_meter is not None:
            new_value = import_engine._format_number(entry["total_meter"])
            current = row[col_hasil_meter].strip() if col_hasil_meter < len(row) else ""
            if current != new_value:
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(i, col_hasil_meter + 1),
                    "values": [[new_value]],
                })

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        # invalidate cache REWIND_PY biar GET /api/waste-rewind berikutnya
        # baca nilai Hasil_Slitting_(Rol)/UP_Slitting/Hasil_Slitting_(Meter) yang baru
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0

    return len(updates)


def _sync_rewind_kecil_tanggal_qty_into_rewind_py():
    """Isi TIGA kolom di REWIND_PY (satu kali baca REWIND_PY_RAW, dua kali
    batch_update): Tanggal_Rewind, Qty_Awal_Rewind, Qty_Akhir_Rewind.

    Pencocokan: SPK REWIND_PY = SPK REWIND_PY_RAW, NO_JO REWIND_PY = JO
    REWIND_PY_RAW (keduanya angka murni). SEMUA baris RAW yang cocok dipakai
    (tanpa filter tanggal):
      - Tanggal_Rewind   = kolom TANGGAL, format DD/Mmm/YY, unik, urut, dipisah koma
      - Qty_Awal_Rewind  = kolom JUMLAH_AWAL, digabung lewat rewind_qty.combine_qty()
      - Qty_Akhir_Rewind = kolom JUMLAH_AKHIR, aturan sama

    Baris REWIND_PY yang SPK/NO_JO-nya angka tapi TIDAK ketemu di RAW
    dikosongkan (""). Baris dengan SPK/NO_JO non-angka dilewati. Sel yang
    nilainya sudah sama tidak ditulis ulang. Balikin jumlah SEL yang diupdate.

    Dipanggil dari _run_rewind_kecil_worker(), SETELAH
    _sync_rewind_kecil_spk_jo_into_rewind_py() (baris REWIND_PY sudah ada)."""
    sh = _waste_rewind_spreadsheet()
    raw_ws = sh.worksheet(REWIND_KECIL_RAW_SHEET_NAME)
    raw_values = raw_ws.get_all_values()
    if not raw_values:
        return 0

    raw_header = [str(h).strip() for h in raw_values[0]]
    r_tgl = import_engine._find_col_index(raw_header, "TANGGAL")
    r_spk = import_engine._find_col_index(raw_header, "SPK")
    r_jo = import_engine._find_col_index(raw_header, "JO")
    r_awal = import_engine._find_col_index(raw_header, "JUMLAH_AWAL")
    r_akhir = import_engine._find_col_index(raw_header, "JUMLAH_AKHIR")
    if r_spk is None or r_jo is None:
        raise RuntimeError(f"Kolom SPK/JO tidak ketemu di header {REWIND_KECIL_RAW_SHEET_NAME}")

    lookup = {}  # (spk_key, jo_key) -> {"dates": [], "awal": [], "akhir": []}
    for raw_row in raw_values[1:]:
        def _cell(idx, _row=raw_row):
            return str(_row[idx]).strip() if idx is not None and idx < len(_row) else ""

        spk, jo = _cell(r_spk), _cell(r_jo)
        if not (spk.isdigit() and jo.isdigit()):
            continue
        entry = lookup.setdefault(
            (import_engine._numeric_key_prefix(spk), import_engine._numeric_key_prefix(jo)),
            {"dates": [], "awal": [], "akhir": []},
        )
        entry["dates"].append(import_engine._parse_date_flexible(_cell(r_tgl)))
        entry["awal"].append(_cell(r_awal))
        entry["akhir"].append(_cell(r_akhir))

    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    col_spk = import_engine._find_col_index(header, "SPK")
    col_nojo = import_engine._find_col_index(header, "NO_JO")
    if col_spk is None or col_nojo is None:
        raise RuntimeError(f"Kolom SPK/NO_JO tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    col_tgl = import_engine._find_col_index(header, "Tanggal_Rewind")
    col_awal = import_engine._find_col_index(header, "Qty_Awal_Rewind")
    col_akhir = import_engine._find_col_index(header, "Qty_Akhir_Rewind")

    updates_text = []  # tanggal: RAW supaya "01/Sep/26" tidak diubah Sheets jadi tanggal
    updates_qty = []   # qty: USER_ENTERED
    for i, row in enumerate(values[1:], start=2):
        spk = row[col_spk].strip() if col_spk < len(row) else ""
        nojo = row[col_nojo].strip() if col_nojo < len(row) else ""
        if not (spk.isdigit() and nojo.isdigit()):
            continue

        entry = lookup.get((
            import_engine._numeric_key_prefix(spk),
            import_engine._numeric_key_prefix(nojo),
        ))

        def _queue(col, new_value, bucket, _row=row, _i=i):
            if col is None:
                return
            current = _row[col].strip() if col < len(_row) else ""
            if current != new_value:
                bucket.append({
                    "range": gspread.utils.rowcol_to_a1(_i, col + 1),
                    "values": [[new_value]],
                })

        _queue(col_tgl, rewind_qty.combine_dates(entry["dates"]) if entry else "", updates_text)
        _queue(col_awal, rewind_qty.combine_qty(entry["awal"]) if entry else "", updates_qty)
        _queue(col_akhir, rewind_qty.combine_qty(entry["akhir"]) if entry else "", updates_qty)

    if updates_text:
        ws.batch_update(updates_text, value_input_option="RAW")
    if updates_qty:
        ws.batch_update(updates_qty, value_input_option="USER_ENTERED")
    if updates_text or updates_qty:
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0

    return len(updates_text) + len(updates_qty)


FORM_ST1_SHEET_NAME = "FORM_ST_1"


def _sync_kg_bruto_into_rewind_py():
    """Isi kolom Kg_Bruto di REWIND_PY.

    Untuk tiap baris REWIND_PY dengan NO_JO angka murni: cari baris di sheet
    FORM_ST_1 (spreadsheet FSTL, lewat _fstl_spreadsheet()) yang SUFFIX kolom
    JO-nya (angka setelah '/' terakhir) sama dengan NO_JO -- kolom JO_DIGIT
    TIDAK dipakai. Hanya baris ber-STATUS yang ada di kg_bruto.STATUS_PRIORITY.
    Berat diambil dari BERAT/KG sesuai posisi bagian polos (tanpa '@') di
    JUMLAH_MASUK_GBJ, lalu Kg_Bruto = modus (aturan lengkap ada di kg_bruto.py).

    NO_JO yang tidak ketemu di FORM_ST_1 -> Kg_Bruto dikosongkan (""). Sel yang
    nilainya sudah sama tidak ditulis ulang. Balikin jumlah SEL yang diupdate."""
    fstl_sh = _fstl_spreadsheet()
    form_ws = fstl_sh.worksheet(FORM_ST1_SHEET_NAME)
    form_values = form_ws.get_all_values()
    if not form_values:
        return 0

    f_header = [str(h).strip() for h in form_values[0]]
    f_jo = import_engine._find_col_index(f_header, "JO")  # exact: JO_DIGIT beda header
    f_status = import_engine._find_col_index(f_header, "STATUS")
    f_berat = import_engine._find_col_index(f_header, "BERAT/KG")
    f_jumlah = _fstl_find_col(f_header, "JUMLAH_MASUK_GBJ", "JUMLAH_MASUK")
    missing = [n for n, c in (("JO", f_jo), ("STATUS", f_status),
                              ("BERAT/KG", f_berat), ("JUMLAH_MASUK_GBJ", f_jumlah)) if c is None]
    if missing:
        raise RuntimeError(f"Kolom {', '.join(missing)} tidak ketemu di header {FORM_ST1_SHEET_NAME}")

    by_suffix = {}  # suffix_key -> [(rank, row_idx, jumlah, berat)]
    for row_idx, row in enumerate(form_values[1:], start=2):
        def _cell(idx, _row=row):
            return str(_row[idx]).strip() if idx < len(_row) else ""

        rank = kg_bruto.status_rank(_cell(f_status))
        if rank is None:
            continue
        key = _fstl_suffix_key(_cell(f_jo))
        if not key:
            continue
        by_suffix.setdefault(key, []).append((rank, row_idx, _cell(f_jumlah), _cell(f_berat)))

    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    col_nojo = import_engine._find_col_index(header, "NO_JO")
    col_kg = import_engine._find_col_index(header, "Kg_Bruto")
    if col_nojo is None or col_kg is None:
        raise RuntimeError(f"Kolom NO_JO/Kg_Bruto tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    updates = []
    for i, row in enumerate(values[1:], start=2):
        nojo = row[col_nojo].strip() if col_nojo < len(row) else ""
        if not nojo.isdigit():
            continue
        cands = by_suffix.get(import_engine._numeric_key_prefix(nojo))
        new_value = kg_bruto.pick_kg_bruto(cands) if cands else ""
        current = row[col_kg].strip() if col_kg < len(row) else ""
        if current != new_value:
            updates.append({
                "range": gspread.utils.rowcol_to_a1(i, col_kg + 1),
                "values": [[new_value]],
            })

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0
    return len(updates)


def _sync_konversi_meter_jumbo_into_rewind_py():
    """Isi DUA kolom konversi di REWIND_PY (satu kali baca, satu kali batch_update):
      - Konversi_Meter_Jumbo_Qty_Awal_Rewind  <- dari Qty_Awal_Rewind
      - Konversi_Meter_Jumbo_Qty_Akhir_Rewind <- dari Qty_Akhir_Rewind

    Input lain (sama untuk keduanya): UP_Slitting, Potongan, Kg_Bruto. Semua
    dibaca ULANG dari sheet, jadi harus dipanggil SETELAH sync Qty/UP/Kg_Bruto.
    Rumus lengkap ada di konversi_meter_jumbo.py (satu file untuk Awal & Akhir).
    Baris yang tidak bisa dihitung (Qty kosong, UP/Potongan/Kg_Bruto kosong,
    teks tidak terbaca) -> sel dikosongkan. Sel yang nilainya sudah sama
    (dibanding secara numerik) tidak ditulis ulang. Balikin jumlah SEL yang diupdate."""
    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    find = import_engine._find_col_index
    col_up = find(header, "UP_Slitting")
    col_pot = find(header, "Potongan")
    col_kg = find(header, "Kg_Bruto")
    missing = [n for n, c in (("UP_Slitting", col_up), ("Potongan", col_pot), ("Kg_Bruto", col_kg)) if c is None]

    pairs = []  # (kolom_sumber, kolom_hasil, label)
    for src_name, dst_name in (
        ("Qty_Awal_Rewind", "Konversi_Meter_Jumbo_Qty_Awal_Rewind"),
        ("Qty_Akhir_Rewind", "Konversi_Meter_Jumbo_Qty_Akhir_Rewind"),
    ):
        c_src, c_dst = find(header, src_name), find(header, dst_name)
        if c_src is not None and c_dst is not None:
            pairs.append((c_src, c_dst, dst_name))
    if missing or not pairs:
        raise RuntimeError(
            f"Kolom {', '.join(missing) or 'Qty_*_Rewind/Konversi_Meter_Jumbo_*'} "
            f"tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    updates = []
    for i, row in enumerate(values[1:], start=2):
        def _cell(idx, _row=row):
            return str(_row[idx]).strip() if idx < len(_row) else ""

        for col_src, col_dst, label in pairs:
            total, why = konversi_meter_jumbo.hitung_meter_jumbo(
                _cell(col_src), _cell(col_up), _cell(col_pot), _cell(col_kg))
            if total is None and why not in (None, "kosong"):
                print(f"[{label}] baris {i}: dikosongkan ({why})")
            new_value = "" if total is None else import_engine._format_number(round(total, 2))

            current = _cell(col_dst)
            cur_num = import_engine._parse_flexible_number(current)
            if new_value == "":
                same = current == ""
            else:
                same = cur_num is not None and abs(cur_num - total) < 0.005
            if not same:
                updates.append({
                    "range": gspread.utils.rowcol_to_a1(i, col_dst + 1),
                    "values": [[new_value]],
                })

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0
    return len(updates)


def hitung_meter_hilang(meter, potongan, up):
    """Port dari fungsi HITUNG (Apps Script). Ubah total meter jadi teks
    "<rol utuh> + <up>@<sisa>":
      utuh = floor(round(meter) / potongan), sisa = round(meter) % potongan
      utuh == 0            -> "<up>@<meter>"
      sisa  == 0           -> "<utuh*up>"
      selain itu           -> "<utuh*up> + <up>@<sisa>"
    Balikin "" (sel dikosongkan) kalau input tidak bisa dihitung: meter/up
    kosong atau bukan angka, atau potongan <= 0. Kalau meter (dibulatkan) <= 0
    hasilnya "0", sama seperti HITUNG aslinya."""
    if meter is None or potongan is None or up is None:
        return ""
    if potongan <= 0:
        return ""
    meter_bulat = math.floor(meter + 0.5)  # = Math.round di JS (bukan round() Python yang bulatkan ke genap)
    if meter_bulat <= 0:
        return "0"
    utuh = int(meter_bulat // potongan)
    sisa = round(meter_bulat - utuh * potongan, 2)
    fmt = import_engine._format_number
    if utuh == 0:
        return f"{fmt(up)}@{fmt(meter_bulat)}"
    nilai_utuh = utuh * up
    if sisa == 0:
        return fmt(nilai_utuh)
    return f"{fmt(nilai_utuh)} + {fmt(up)}@{fmt(sisa)}"


def _sync_meter_hilang_into_rewind_py():
    """Isi DUA kolom di REWIND_PY (satu kali baca, satu kali batch_update):
      - Meter_Jumbo_Hilang_Rewind = Konversi_Meter_Jumbo_Qty_Awal_Rewind
                                    - Konversi_Meter_Jumbo_Qty_Akhir_Rewind
      - Meter_Hilang_Rewind       = hitung_meter_hilang(Meter_Jumbo_Hilang_Rewind,
                                    Potongan, UP_Slitting)   (rumus HITUNG)

    HARUS dipanggil SETELAH _sync_konversi_meter_jumbo_into_rewind_py() dan
    sync UP_Slitting/Potongan (semua dibaca ULANG dari sheet). Baris yang
    tidak bisa dihitung (Awal/Akhir kosong, Potongan kosong/<=0, UP kosong)
    -> sel dikosongkan. Sel yang nilainya sudah sama tidak ditulis ulang.
    Balikin jumlah SEL yang diupdate."""
    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    find = import_engine._find_col_index
    names = (
        "Konversi_Meter_Jumbo_Qty_Awal_Rewind", "Konversi_Meter_Jumbo_Qty_Akhir_Rewind",
        "Meter_Jumbo_Hilang_Rewind", "Meter_Hilang_Rewind", "Potongan", "UP_Slitting",
    )
    cols = {n: find(header, n) for n in names}
    missing = [n for n, c in cols.items() if c is None]
    if missing:
        raise RuntimeError(f"Kolom {', '.join(missing)} tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    c_awal = cols["Konversi_Meter_Jumbo_Qty_Awal_Rewind"]
    c_akhir = cols["Konversi_Meter_Jumbo_Qty_Akhir_Rewind"]
    c_jh = cols["Meter_Jumbo_Hilang_Rewind"]
    c_mh = cols["Meter_Hilang_Rewind"]
    c_pot = cols["Potongan"]
    c_up = cols["UP_Slitting"]
    parse = import_engine._parse_flexible_number

    updates = []
    for i, row in enumerate(values[1:], start=2):
        def _cell(idx, _row=row):
            return str(_row[idx]).strip() if idx < len(_row) else ""

        # 1) Meter_Jumbo_Hilang_Rewind = Awal - Akhir
        awal, akhir = parse(_cell(c_awal)), parse(_cell(c_akhir))
        if awal is None or akhir is None:
            jumbo_hilang = None
            new_jh = ""
        else:
            jumbo_hilang = round(awal - akhir, 2)
            new_jh = import_engine._format_number(jumbo_hilang)

        cur_jh = _cell(c_jh)
        cur_jh_num = parse(cur_jh)
        if jumbo_hilang is None:
            same = cur_jh == ""
        else:
            same = cur_jh_num is not None and abs(cur_jh_num - jumbo_hilang) < 0.005
        if not same:
            updates.append({
                "range": gspread.utils.rowcol_to_a1(i, c_jh + 1),
                "values": [[new_jh]],
            })

        # 2) Meter_Hilang_Rewind = HITUNG(Meter_Jumbo_Hilang, Potongan, UP)
        new_mh = hitung_meter_hilang(jumbo_hilang, parse(_cell(c_pot)), parse(_cell(c_up)))
        if _cell(c_mh) != new_mh:
            updates.append({
                "range": gspread.utils.rowcol_to_a1(i, c_mh + 1),
                "values": [[new_mh]],
            })

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0
    return len(updates)


def _sync_hitung_waste_kolom_into_rewind_py():
    """Hitung ulang 4 kolom waste (_WRW_CALC_COLS) untuk SEMUA baris di
    REWIND_PY sekaligus, pakai rumus yang SAMA dengan tombol "Hitung Waste"
    di modal Detail (lihat hitung_waste_rewind()) -- jadi tombol Refresh di
    halaman Waste Rewind ikut mengisi/memperbarui:
        Waste_Slitting_Meter, Persentase_Waste_(%),
        Waste_Slitting_After_Rewind_Meter, Waste_Slitting_After_Rewind_Presentase
    untuk SEMUA JO, tanpa harus buka Detail satu-satu. Tombol "Hitung Waste"
    per-JO di modal Detail TETAP ADA & tetap jalan (dipakai kalau cuma mau
    hitung ulang satu JO tertentu, mis. sesudah revisi data mentahnya).

    HARUS dipanggil SETELAH Bahan_Awal_Printing_(Meter), Hasil_Slitting_(Meter)
    & Meter_Jumbo_Hilang_Rewind terisi (jadi taruh di urutan akhir, sebelum
    _apply_rewind_revisi_into_rewind_py supaya revisi manual tetap menang).

    Baris yang sudah Finish DILEWATI (dikunci, sama seperti aturan tombol
    "Hitung Waste" per-JO di /api/waste-rewind/hitung) -- nilainya sudah
    tersimpan permanen di REWIND_PY_FINISH. Baris yang datanya belum cukup
    (Bahan Awal / Hasil Slitting kosong atau Bahan Awal = 0) juga dilewati
    apa adanya (tidak mengosongkan nilai lama), supaya baris yang memang
    belum siap dihitung tidak keliru ditampilkan kosong/error.
    Sel yang nilainya sudah sama tidak ditulis ulang. Balikin jumlah SEL
    yang diupdate."""
    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    values = ws.get_all_values()
    if not values:
        return 0

    header = [str(h).strip() for h in values[0]]
    find = import_engine._find_col_index
    parse = import_engine._parse_flexible_number
    names = ("Bahan_Awal_Printing_(Meter)", "Hasil_Slitting_(Meter)",
              "Meter_Jumbo_Hilang_Rewind", "Status") + _WRW_CALC_COLS
    cols = {n: find(header, n) for n in names}
    missing = [n for n, c in cols.items() if c is None and n != "Status"]
    if missing:
        raise RuntimeError(f"Kolom {', '.join(missing)} tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    c_bahan = cols["Bahan_Awal_Printing_(Meter)"]
    c_hasil = cols["Hasil_Slitting_(Meter)"]
    c_jumbo = cols["Meter_Jumbo_Hilang_Rewind"]
    c_status = cols["Status"]

    updates = []
    for i, row in enumerate(values[1:], start=2):
        def _cell(idx, _row=row):
            return str(_row[idx]).strip() if idx is not None and idx < len(_row) else ""

        if c_status is not None and _cell(c_status).lower() == "finish":
            continue  # terkunci, sudah punya nilai permanen di REWIND_PY_FINISH

        try:
            res = hitung_waste_rewind(parse(_cell(c_bahan)), parse(_cell(c_hasil)), parse(_cell(c_jumbo)))
        except ValueError:
            continue  # data belum cukup buat dihitung, biarkan apa adanya

        for n in _WRW_CALC_COLS:
            c = cols[n]
            if not _wrw_same_value(_cell(c), res[n]):
                updates.append({"range": gspread.utils.rowcol_to_a1(i, c + 1), "values": [[res[n]]]})

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        with _waste_rewind_cache_lock:
            _waste_rewind_cache["ts"] = 0.0
    return len(updates)


def _sync_rewind_kecil_spk_jo_into_rewind_py():
    """REFRESH PENUH sheet REWIND_PY: hapus SEMUA baris data (baris 2 ke
    bawah, seluruh lebar sheet), lalu tulis ulang dari nol satu baris per
    pasangan (SPK, NO_JO) unik dari sheet REWIND_PY_RAW -- sekaligus
    auto-isi kolom JO/Nama_Produk/Planning_Order/Planning_Meter/Potongan
    dari lookup JO_1 (spreadsheet FSTL). Baris header (baris 1) tidak
    disentuh.

    PERINGATAN: ini SENGAJA menghapus juga isi kolom yang sebelumnya
    diisi manual/formula (Persentase_Waste_(%), Meter_Hilang_Rewind,
    Hasil_Slitting_(Rol), Waste_Slitting_After_Rewind_Presentase, dst) --
    lihat catatan "PERUBAHAN PERILAKU" di komentar atas fungsi ini.
    PENGECUALIAN: Status + 4 kolom waste (Waste_Slitting_Meter,
    Persentase_Waste_(%), Waste_Slitting_After_Rewind_Meter/Presentase) dipulihkan
    dari sheet REWIND_PY_FINISH untuk baris (SPK, NO_JO) yang sudah di-Finish.

    Balikin jumlah baris data yang ditulis ulang (bukan cuma yang baru)."""
    unique_pairs = _read_rewind_kecil_spk_jo()

    sh = _waste_rewind_spreadsheet()
    ws = sh.worksheet(WASTE_REWIND_SHEET_NAME)
    header = [str(h).strip() for h in ws.row_values(1)]

    col_spk = import_engine._find_col_index(header, "SPK")
    col_nojo = import_engine._find_col_index(header, "NO_JO")
    if col_spk is None or col_nojo is None:
        raise RuntimeError(f"Kolom SPK/NO_JO tidak ketemu di header {WASTE_REWIND_SHEET_NAME}")

    # Kolom tambahan yang diisi otomatis dari lookup JO_1 (boleh None kalau
    # sheet REWIND_PY suatu saat tidak/belum punya salah satu kolom ini --
    # tetap jalan, cuma kolom itu yang dilewati/tidak diisi).
    col_jo_full = import_engine._find_col_index(header, "JO")
    col_nama = import_engine._find_col_index(header, "Nama_Produk")
    col_planning_order = import_engine._find_col_index(header, "Planning_Order")
    col_planning_meter = import_engine._find_col_index(header, "Planning_Meter")
    col_potongan = import_engine._find_col_index(header, "Potongan")

    # Lebar yang dipakai buat HAPUS = seluruh lebar sheet (bukan cuma
    # kolom SPK/NO_JO dkk), supaya kolom manual/formula di kanan (mis.
    # Waste_Slitting_After_Rewind_Presentase) ikut kehapus juga -- sesuai
    # permintaan "hapus baris 2 ke bawah, ganti data baru".
    width = max(
        len(header), ws.col_count,
        col_spk + 1, col_nojo + 1,
        (col_jo_full + 1) if col_jo_full is not None else 0,
        (col_nama + 1) if col_nama is not None else 0,
        (col_planning_order + 1) if col_planning_order is not None else 0,
        (col_planning_meter + 1) if col_planning_meter is not None else 0,
        (col_potongan + 1) if col_potongan is not None else 0,
    )

    # Status + 4 kolom waste yang sudah disimpan lewat "Set Status Finish"
    # (sheet REWIND_PY_FINISH) dipulihkan supaya tidak hilang oleh refresh penuh.
    finish_map = _read_rewind_finish_restore_map()
    restore_cols = {
        n: import_engine._find_col_index(header, n)
        for n in ("Status",) + _WRW_CALC_COLS
    }
    restore_cols = {n: c for n, c in restore_cols.items() if c is not None}

    # Bangun baris-baris baru dari pasangan (SPK, NO_JO) unik.
    seen = set()
    new_rows = []
    jo1_lookup = None  # lazy: baru dibangun kalau memang ada datanya
    for pair in unique_pairs:
        key = (pair["spk"], pair["noJo"])
        if key in seen:
            continue
        seen.add(key)
        if jo1_lookup is None:
            jo1_lookup = _fstl_lookup_jo1_by_suffix_map()
        blank_row = [""] * width
        blank_row[col_spk] = pair["spk"]
        blank_row[col_nojo] = pair["noJo"]
        info = jo1_lookup.get(_fstl_suffix_key(pair["noJo"]))
        if info:
            if col_jo_full is not None:
                blank_row[col_jo_full] = info["jo"]
            if col_nama is not None:
                blank_row[col_nama] = info["nama"]
            if col_planning_order is not None:
                blank_row[col_planning_order] = info["order"]
            if col_planning_meter is not None:
                blank_row[col_planning_meter] = info["meter"]
            if col_potongan is not None:
                blank_row[col_potongan] = info["potongan"]
        saved = finish_map.get((str(pair["spk"]).strip(), str(pair["noJo"]).strip()))
        if saved:
            for n, c in restore_cols.items():
                if saved.get(n):
                    blank_row[c] = saved[n]
        new_rows.append(blank_row)

    # HAPUS baris 2 ke bawah, seluruh lebar sheet, SEBELUM tulis data baru.
    end_col_a1 = gspread.utils.rowcol_to_a1(1, width).rstrip("0123456789")
    clear_last_row = max(ws.row_count, len(new_rows) + 1)
    ws.batch_clear([f"A2:{end_col_a1}{clear_last_row}"])

    if new_rows:
        end_row = 1 + len(new_rows)
        if end_row > ws.row_count:
            ws.add_rows(end_row - ws.row_count)
        start_a1 = gspread.utils.rowcol_to_a1(2, 1)
        end_a1 = gspread.utils.rowcol_to_a1(end_row, width)
        ws.update(f"{start_a1}:{end_a1}", new_rows, value_input_option="USER_ENTERED")

    # invalidate cache REWIND_PY biar GET /api/waste-rewind berikutnya
    # (dipanggil loadWasteRewind(true) sesudah refresh ini) baca data baru
    with _waste_rewind_cache_lock:
        _waste_rewind_cache["ts"] = 0.0

    return len(new_rows)


# --------------------------------------------------------------------------
# 7. FSTL — LAMPIRAN WASTE
# --------------------------------------------------------------------------
# DUA spreadsheet berbeda dipakai di sini, jangan ketuker:
#
#   FSTL_SPREADSHEET_ID       -- spreadsheet SUMBER DATA (JO_1, LP_1,
#                                 PRINTING_2..5, DRY_1..5, SL_1, RW_1, EX_1,
#                                 BAG_1, dst -- lihat config.json). Dipakai
#                                 buat /api/fstl/cek-jo & /api/fstl/keterangan
#                                 (baik saat input JO baru MAUPUN saat user
#                                 klik "Ambil Keterangan" lagi pas revisi --
#                                 dua-duanya butuh data sumber yang sama).
#
#   FSTL_KITIR_SPREADSHEET_ID -- spreadsheet TEMPAT NYIMPEN KITIR user
#                                 (tab "{USERNAME}_Kitir"). Dipakai buat
#                                 /api/fstl/save, /api/fstl/list, dan
#                                 /api/fstl/revisi. SENGAJA dipisah dari
#                                 spreadsheet sumber di atas.
#
# Keduanya bisa dioverride lewat env var kalau suatu saat pindah lagi.
FSTL_SPREADSHEET_ID = os.environ.get(
    "FSTL_SPREADSHEET_ID", "1FRWpza_fa65jt8-n1-rN4rFFrfNBLixRxOLS_uUgYYU"
)
FSTL_KITIR_SPREADSHEET_ID = os.environ.get(
    "FSTL_KITIR_SPREADSHEET_ID", "1goadL7s6y2F38Zqgx65Z9Vy9mSVHkovEZDfDb16AgD0"
)

FSTL_HEADER = ["TANGGAL", "USER", "JO", "PRODUK", "PROSES", "KETERANGAN", "ACTION PLAN", "STATUS"]

# Nama proses (dari frontend, HARUS dibandingkan case-insensitive) -> sheet
# sumber yang dicari buat ambil KETERANGAN-nya.
#   prefixes = sheet yang NAMANYA DIAWALI salah satu prefix ini ikut dicari
#   exact    = sheet dengan nama PERSIS ini ikut dicari
FSTL_PROCESS_SOURCES = {
    "PRINTING": {"prefixes": ["PRINTING_"], "exact": []},
    "DRY LAMINASI": {"prefixes": ["DRY_"], "exact": ["SF_1"]},
    "SLITTING": {"prefixes": [], "exact": ["SL_1"]},
    "REWINDING": {"prefixes": [], "exact": ["RW_1"]},
    "EXTRUSI": {"prefixes": [], "exact": ["EX_1"]},
    "BAG MAKING": {"prefixes": [], "exact": ["BAG_1"]},
}

FSTL_COLOR_TITLE_LABEL = "#9bc2e6"  # biru — baris judul kartu (SPK/JO..) & baris label kolom
FSTL_COLOR_PROCESS = "#a9d08e"      # hijau — cuma sel nama proses di tiap baris data
FSTL_COLOR_WHITE = "#ffffff"        # putih — sel Keterangan/Action Plan/Status di baris data


def _fstl_hex_to_rgb01(hex_color):
    """'#a9d08e' -> {'red':.., 'green':.., 'blue':..} skala 0-1, format yang
    dipakai Google Sheets API buat backgroundColor lewat Worksheet.format()."""
    hex_color = hex_color.lstrip("#")
    return {
        "red": int(hex_color[0:2], 16) / 255,
        "green": int(hex_color[2:4], 16) / 255,
        "blue": int(hex_color[4:6], 16) / 255,
    }


FSTL_JO1_SHEET = "JO_1"
FSTL_JO1_COL_JO = 5      # kolom F (0-based)
FSTL_JO1_COL_PRODUK = 6  # kolom G (0-based)
FSTL_LP1_SHEET = "LP_1"


_fstl_spreadsheet_handle = {"sh": None}
_fstl_spreadsheet_lock = threading.Lock()

_fstl_kitir_spreadsheet_handle = {"sh": None}
_fstl_kitir_spreadsheet_lock = threading.Lock()


def _fstl_spreadsheet():
    """Handle ke spreadsheet SUMBER DATA (FSTL_SPREADSHEET_ID) -- open_by_key()
    juga panggilan ke Google, dan ID-nya nggak pernah berubah selama app
    jalan, jadi cukup dibuka sekali lalu handle-nya dipakai ulang terus,
    bukan dibuka lagi dari nol tiap ada request."""
    if not FSTL_SPREADSHEET_ID:
        raise RuntimeError("FSTL_SPREADSHEET_ID belum diset")
    with _fstl_spreadsheet_lock:
        if _fstl_spreadsheet_handle["sh"] is not None:
            return _fstl_spreadsheet_handle["sh"]
    client = get_client()
    sh = client.open_by_key(FSTL_SPREADSHEET_ID)
    with _fstl_spreadsheet_lock:
        _fstl_spreadsheet_handle["sh"] = sh
    return sh


def _fstl_kitir_spreadsheet():
    """Sama seperti _fstl_spreadsheet(), tapi buat spreadsheet TEMPAT NYIMPEN
    KITIR (FSTL_KITIR_SPREADSHEET_ID) -- spreadsheet BEDA dari sumber data.
    Handle-nya dipisah total (dict + lock sendiri) dari _fstl_spreadsheet()
    supaya nggak ketuker antara baca sumber data vs baca/tulis kitir user.

    CATATAN: service account yang dipakai get_client() harus sudah di-share
    (Editor) ke spreadsheet ini dulu, kalau belum open_by_key() bakal error
    permission (403)."""
    if not FSTL_KITIR_SPREADSHEET_ID:
        raise RuntimeError("FSTL_KITIR_SPREADSHEET_ID belum diset")
    with _fstl_kitir_spreadsheet_lock:
        if _fstl_kitir_spreadsheet_handle["sh"] is not None:
            return _fstl_kitir_spreadsheet_handle["sh"]
    client = get_client()
    sh = client.open_by_key(FSTL_KITIR_SPREADSHEET_ID)
    with _fstl_kitir_spreadsheet_lock:
        _fstl_kitir_spreadsheet_handle["sh"] = sh
    return sh


# --------------------------------------------------------------------------
# CACHE RINGAN buat endpoint /api/fstl/keterangan.
#
# Sheet sumber waste (PRINTING_2, SLITTING, dst) & LP_1 bisa ribuan-puluhan
# ribu baris (lihat config.json), dan sebelumnya di-fetch ULANG dari nol
# (sh.worksheets() + ws.get_all_values()) untuk SETIAP proses yang dicentang
# dalam satu request yang sama -- padahal daftar worksheet & isi LP_1 itu
# sama persis buat semua proses. Dua cache di bawah ini:
#   1. _fstl_worksheet_titles_cache : daftar nama tab (buat _fstl_matching_sheet_names)
#   2. _fstl_sheet_values_cache     : isi get_all_values() per nama sheet
# Keduanya di-share dalam SATU request (lewat parameter get_rows/titles yang
# dioper ke fungsi-fungsi di bawah), DAN juga disimpan lintas-request dengan
# TTL pendek supaya klik "Ambil Keterangan" berkali-kali dalam waktu dekat
# nggak perlu narik ulang sheet gede dari Google Sheets API. Data sumbernya
# sendiri cuma di-refresh berkala oleh proses import terpisah (lihat
# last_import di config.json), jadi cache basi beberapa menit aman -- durasinya
# bisa diatur lewat env var FSTL_CACHE_TTL_SECONDS tanpa ubah kode.
# --------------------------------------------------------------------------
_FSTL_CACHE_TTL_SECONDS = int(os.environ.get("FSTL_CACHE_TTL_SECONDS", "3600"))

# Sheet JO_1 dipakai buat "Periksa" JO (cek nama produk) waktu input JO baru.
# Sama kayak sheet sumber "Ambil Keterangan" (data yang dicari cuma ~1 bulan
# terakhir / data lama), jadi dikasih TTL 1 jam juga -- biar nggak keseringan
# loading ke Google Sheets. Ditulis eksplisit di override (bukan cuma ngandelin
# default di atas) supaya kalau nanti default umum di-ubah lagi lewat env var,
# JO_1 tetap punya nilainya sendiri yang bisa diatur terpisah.
_FSTL_JO1_CACHE_TTL_SECONDS = int(os.environ.get("FSTL_JO1_CACHE_TTL_SECONDS", "3600"))
_FSTL_SHEET_TTL_OVERRIDES = {FSTL_JO1_SHEET: _FSTL_JO1_CACHE_TTL_SECONDS}

_fstl_worksheet_titles_cache = {"ts": 0.0, "titles": None}
_fstl_sheet_values_cache = {}  # sheet_name -> (timestamp, rows)
_fstl_cache_lock = threading.Lock()


def _fstl_get_worksheet_titles(sh):
    """Daftar nama semua tab di spreadsheet FSTL, di-cache TTL pendek."""
    now = time.time()
    with _fstl_cache_lock:
        cached = _fstl_worksheet_titles_cache
        if cached["titles"] is not None and now - cached["ts"] < _FSTL_CACHE_TTL_SECONDS:
            return cached["titles"]
    titles = [w.title for w in sh.worksheets()]
    with _fstl_cache_lock:
        _fstl_worksheet_titles_cache["ts"] = now
        _fstl_worksheet_titles_cache["titles"] = titles
    return titles


def _fstl_get_sheet_values(sh, sheet_name):
    """get_all_values() satu sheet, di-cache TTL pendek per nama sheet.
    Dipanggil lewat helper ini supaya dalam SATU request /api/fstl/keterangan,
    sheet yang sama (mis. LP_1) cuma benar-benar di-fetch sekali walau
    dipakai berkali-kali (sekali per proses yang dicentang).

    TTL per-sheet bisa berbeda -- lihat _FSTL_SHEET_TTL_OVERRIDES (mis. JO_1
    dikasih TTL 1 jam karena datanya jauh lebih jarang berubah)."""
    ttl = _FSTL_SHEET_TTL_OVERRIDES.get(sheet_name, _FSTL_CACHE_TTL_SECONDS)
    now = time.time()
    with _fstl_cache_lock:
        cached = _fstl_sheet_values_cache.get(sheet_name)
        if cached is not None and now - cached[0] < ttl:
            return cached[1]
    try:
        rows = sh.worksheet(sheet_name).get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        rows = []
    with _fstl_cache_lock:
        _fstl_sheet_values_cache[sheet_name] = (now, rows)
    return rows


def _invalidate_waste_rewind_source_cache():
    """Hapus cache sheet sumber yang dipakai sinkron Waste Rewind saja
    (LP_1, JO_1, SL_1, PRINTING_1..5), supaya klik Refresh di halaman
    Waste Rewind selalu baca data terbaru. Cache sheet lain (halaman FSTL
    dll) tidak disentuh."""
    names = {
        FSTL_LP1_SHEET,
        FSTL_JO1_SHEET,
        import_engine.SL_SOURCE_SHEET_NAME,
        *PRINTING_SHEET_NAMES,
    }
    with _fstl_cache_lock:
        for name in names:
            _fstl_sheet_values_cache.pop(name, None)


def _fstl_invalidate_cache():
    """Panggil ini setelah nulis data baru ke spreadsheet FSTL (mis. abis
    fstl_save nambah worksheet baru "{user}_Kitir"), biar cache worksheet
    titles nggak ketinggalan tab yang baru dibuat."""
    with _fstl_cache_lock:
        _fstl_worksheet_titles_cache["ts"] = 0.0
        _fstl_worksheet_titles_cache["titles"] = None


def _fstl_batch_prefetch_sheets(sh, sheet_names):
    """Ambil isi BANYAK sheet SEKALIGUS lewat SATU panggilan batch ke Google
    Sheets API (values.batchGet), taruh semuanya ke cache -- dipanggil di
    awal /api/fstl/keterangan SEBELUM proses satu-satu jalan.

    Ini beda dari cache TTL biasa di atas. Cache TTL cuma nyegah baca ULANG
    sheet yang SAMA berkali-kali. Tapi kalau user centang banyak proses
    sekaligus (mis. semua 6 proses), itu bisa nyentuh 10-15 sheet yang
    BEDA-BEDA, dan kalau belum ada satupun yang ke-cache (pertama kali buka
    app, atau cache-nya udah lewat batas TTL), sebelumnya tiap sheet itu
    dibaca lewat request TERPISAH ke Google Sheets API satu-satu secara
    berurutan. Selain kena latency jaringan berkali-kali, Google Sheets API
    juga punya BATAS JUMLAH REQUEST per menit per akun -- kalau kena limit
    itu, request berikutnya otomatis ditunda/di-retry, dan itu yang paling
    mungkin bikin kerasa lelet sampai hitungan menit walau data per sheet-nya
    sendiri nggak segede itu. Gabungin semua sheet yang dibutuhkan jadi SATU
    request besar (bukan banyak request kecil) menghindari masalah ini.

    valueRenderOption=UNFORMATTED_VALUE dipakai juga karena Google Sheets
    butuh waktu ekstra buat "merender" tampilan berformat (format tanggal,
    angka, dsb) tiap baca -- kolom yang kita pakai (JO/KETERANGAN/
    KLASIFIKASI) isinya teks biasa, jadi aman dilewatin proses render itu
    dan hasilnya sama, cuma lebih cepat didapat dari sisi server Google-nya."""
    now = time.time()
    with _fstl_cache_lock:
        to_fetch = [
            name for name in sheet_names
            if name not in _fstl_sheet_values_cache
            or now - _fstl_sheet_values_cache[name][0] >= _FSTL_SHEET_TTL_OVERRIDES.get(name, _FSTL_CACHE_TTL_SECONDS)
        ]
    if not to_fetch:
        return
    print(f"[FSTL] Cache MISS, ambil ulang dari Google Sheets: {to_fetch}", flush=True)
    ranges = [f"'{name}'" for name in to_fetch]
    try:
        resp = sh.values_batch_get(ranges, params={"valueRenderOption": "UNFORMATTED_VALUE"})
    except Exception:
        # Batch gagal (mis. salah satu range bermasalah) -> jangan sampai bikin
        # seluruh request error. Biarin _fstl_get_sheet_values() di bawah baca
        # satu-satu seperti biasa sebagai fallback (lebih lambat, tapi tetap jalan).
        return
    value_ranges = resp.get("valueRanges", []) if resp else []
    now = time.time()
    with _fstl_cache_lock:
        for name, vr in zip(to_fetch, value_ranges):
            values = vr.get("values", [])
            _fstl_sheet_values_cache[name] = (now, [[str(c) for c in row] for row in values])


def _fstl_suffix_key(jo_text):
    """Cocokkan cara ambil suffix JO sama persis kayak di import_engine:
    ambil segmen paling belakang setelah '/' (mis. '123/456A' -> '456A'),
    lalu ambil angka di depannya saja (huruf nyangkut di belakang
    diabaikan, jadi '456A' == '456')."""
    return import_engine._numeric_key_prefix(import_engine._last_segment(jo_text))


def _fstl_find_col(header_row, *keywords):
    """Cari index kolom (0-based) di header_row yang cocok sama salah satu
    keyword. Prioritas: EXACT MATCH keyword pertama di SEMUA kolom dulu,
    baru exact match keyword berikutnya, dst -- baru kalau tidak ada satupun
    exact match, fallback ke substring match dengan urutan prioritas yang
    sama. Ini penting karena beberapa sheet (mis. PRINTING_5) punya kolom
    "SPK" DAN "JO" terpisah -- kalau asal ambil kolom pertama yang
    mengandung salah satu keyword, bisa kepilih kolom yang salah (SPK
    kepilih duluan padahal yang dimaksud kolom JO)."""

    def _norm(text):
        return str(text or "").strip().upper().replace("_", "").replace(" ", "")

    normed_keywords = [_norm(k) for k in keywords if k]
    normed_header = [_norm(cell) for cell in header_row]

    for kw in normed_keywords:
        for idx, cell_norm in enumerate(normed_header):
            if cell_norm and cell_norm == kw:
                return idx

    for kw in normed_keywords:
        for idx, cell_norm in enumerate(normed_header):
            if cell_norm and kw in cell_norm:
                return idx

    return None


def fstl_lookup_produk(jo_raw):
    """Cocokkan JO input ke sheet JO_1 kolom F, kembalikan (produk, error).
    produk None + error None artinya JO tidak ketemu (bukan error sistem)."""
    sh = _fstl_spreadsheet()
    try:
        ws = sh.worksheet(FSTL_JO1_SHEET)
    except gspread.exceptions.WorksheetNotFound:
        return None, f"Sheet '{FSTL_JO1_SHEET}' tidak ditemukan di spreadsheet FSTL"
    target_key = _fstl_suffix_key(jo_raw)
    if target_key == "":
        return None, "Format JO tidak valid"
    rows = _fstl_get_sheet_values(sh, FSTL_JO1_SHEET)
    for row in rows[1:]:
        jo_cell = row[FSTL_JO1_COL_JO] if len(row) > FSTL_JO1_COL_JO else ""
        if not str(jo_cell).strip():
            continue
        if _fstl_suffix_key(jo_cell) == target_key:
            produk = row[FSTL_JO1_COL_PRODUK] if len(row) > FSTL_JO1_COL_PRODUK else ""
            return (str(produk).strip() or None), None
    return None, None


def _fstl_matching_sheet_names(sh, prefixes, exact):
    names = _fstl_get_worksheet_titles(sh)
    matched = []
    for name in names:
        if name in exact:
            matched.append(name)
            continue
        if any(name.upper().startswith(p.upper()) for p in prefixes):
            matched.append(name)
    return matched


def _fstl_join_terms(terms):
    """textjoin semua keterangan yang cocok, buang kosong/'-', buang duplikat."""
    cleaned = []
    for t in terms:
        t = str(t).strip()
        if t and t != "-" and t not in cleaned:
            cleaned.append(t)
    return " | ".join(cleaned)


def _fstl_keterangan_rows(sh, sheet_name, target_key):
    rows = _fstl_get_sheet_values(sh, sheet_name)
    if not rows:
        return []
    header = rows[0]
    col_jo = _fstl_find_col(header, "JO", "SPK")
    col_ket = _fstl_find_col(header, "KETERANGAN")
    if col_jo is None or col_ket is None:
        return []
    terms = []
    for row in rows[1:]:
        jo_cell = row[col_jo] if len(row) > col_jo else ""
        if not str(jo_cell).strip():
            continue
        if _fstl_suffix_key(jo_cell) == target_key:
            terms.append(row[col_ket] if len(row) > col_ket else "")
    return terms


def _fstl_lp1_rows(sh, target_key, process_name):
    """LP_1 ambil pakai cara yang sama (suffix JO), tapi ditambah filter
    kolom KLASIFIKASI harus cocok sama proses yang lagi dicek.
    Kolom keterangan di sheet LP_1 headernya "Faktor_Penyebab_Waste"
    (BUKAN "KETERANGAN" seperti di sheet-sheet sumber proses) -- makanya
    dicari duluan, dengan "KETERANGAN" jadi fallback kalau ada versi LP_1
    lama yang headernya beda."""
    rows = _fstl_get_sheet_values(sh, FSTL_LP1_SHEET)
    if not rows:
        return []
    header = rows[0]
    col_jo = _fstl_find_col(header, "JO", "SPK")
    col_klas = _fstl_find_col(header, "KLASIFIKASI")
    col_ket = _fstl_find_col(header, "FAKTOR_PENYEBAB_WASTE", "KETERANGAN")
    if col_jo is None or col_ket is None:
        return []
    proc_norm = process_name.strip().upper()
    terms = []
    for row in rows[1:]:
        jo_cell = row[col_jo] if len(row) > col_jo else ""
        if not str(jo_cell).strip():
            continue
        if _fstl_suffix_key(jo_cell) != target_key:
            continue
        if col_klas is not None:
            klas_cell = str(row[col_klas] if len(row) > col_klas else "").strip().upper()
            if proc_norm not in klas_cell and klas_cell not in proc_norm:
                continue
        terms.append(row[col_ket] if len(row) > col_ket else "")
    return terms


def fstl_keterangan_for_process(sh, jo_raw, process_name):
    """Gabung keterangan dari sheet sumber proses + LP_1 (filter klasifikasi),
    jadi satu kalimat: '<keterangan proses> LAPORAN PROD: <keterangan LP_1>'."""
    target_key = _fstl_suffix_key(jo_raw)
    proc_key = process_name.strip().upper()
    proc_text = ""
    cfg = FSTL_PROCESS_SOURCES.get(proc_key)
    if cfg is not None:
        all_terms = []
        for name in _fstl_matching_sheet_names(sh, cfg["prefixes"], cfg["exact"]):
            all_terms.extend(_fstl_keterangan_rows(sh, name, target_key))
        proc_text = _fstl_join_terms(all_terms)
    lp1_text = _fstl_join_terms(_fstl_lp1_rows(sh, target_key, proc_key))
    if proc_text and lp1_text:
        return f"{proc_text} LAPORAN PROD: {lp1_text}"
    if lp1_text:
        return f"LAPORAN PROD: {lp1_text}"
    return proc_text


@app.route("/api/fstl/cek-jo", methods=["POST"])
def fstl_cek_jo():
    """Body: {jo}. Ambil JO belakang (abaikan huruf nyangkut), cocokkan ke
    sheet JO_1 kolom F, kembalikan produk dari kolom G."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    if not jo:
        return jsonify({"error": "JO wajib diisi"}), 400
    try:
        produk, err = fstl_lookup_produk(jo)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    if err:
        return jsonify({"error": err}), 404
    if not produk:
        return jsonify({"error": f"JO '{jo}' tidak ditemukan di sheet {FSTL_JO1_SHEET}"}), 404
    return jsonify({"produk": produk})


@app.route("/api/fstl/inspect-sheet", methods=["GET"])
def fstl_inspect_sheet():
    """DIAGNOSTIK — GET /api/fstl/inspect-sheet?sheet=PRINTING_5&jo=123/2254
    Balikin header sheet, index kolom JO/KETERANGAN yang berhasil dideteksi,
    dan contoh isi kolom JO (mentah + hasil parsing suffix-nya) biar gampang
    ketauan kenapa pencarian keterangan gak ketemu (nama header meleset,
    format JO beda, dll). Buka aja URL-nya lewat browser buat lihat hasilnya."""
    sheet_name = request.args.get("sheet", "").strip()
    jo = request.args.get("jo", "").strip()
    if not sheet_name:
        return jsonify({"error": "parameter 'sheet' wajib diisi, mis. ?sheet=PRINTING_5"}), 400
    try:
        sh = _fstl_spreadsheet()
        ws = sh.worksheet(sheet_name)
        rows = ws.get_all_values()
    except gspread.exceptions.WorksheetNotFound:
        return jsonify({"error": f"Sheet '{sheet_name}' tidak ditemukan di spreadsheet FSTL"}), 404
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    if not rows:
        return jsonify({"sheet": sheet_name, "error": "Sheet kosong (tidak ada baris sama sekali)"})

    header = rows[0]
    col_jo = _fstl_find_col(header, "JO", "SPK")
    col_ket = _fstl_find_col(header, "KETERANGAN")
    target_key = _fstl_suffix_key(jo) if jo else None

    sample_jo_values = []
    matched_rows = 0
    if col_jo is not None:
        for row in rows[1:]:
            cell = row[col_jo] if len(row) > col_jo else ""
            if not str(cell).strip():
                continue
            suffix_key = _fstl_suffix_key(cell)
            if len(sample_jo_values) < 8:
                sample_jo_values.append({
                    "raw": cell,
                    "last_segment": import_engine._last_segment(cell),
                    "suffix_key": suffix_key,
                })
            if target_key is not None and suffix_key == target_key:
                matched_rows += 1

    return jsonify({
        "sheet": sheet_name,
        "header_row": header,
        "col_jo_terdeteksi": {"index": col_jo, "nama_header": header[col_jo] if col_jo is not None else None},
        "col_keterangan_terdeteksi": {"index": col_ket, "nama_header": header[col_ket] if col_ket is not None else None},
        "total_baris_data": len(rows) - 1,
        "jo_yang_dicari": jo or None,
        "suffix_key_yang_dicari": target_key,
        "jumlah_baris_cocok": matched_rows,
        "contoh_isi_kolom_jo": sample_jo_values,
    })


@app.route("/api/fstl/keterangan", methods=["POST"])
def fstl_keterangan():
    """Body: {jo, processes:[nama_proses,...]}. Buat tiap proses, textjoin
    keterangan dari sheet sumbernya + LP_1 (LAPORAN PROD)."""
    body = request.get_json(force=True) or {}
    jo = str(body.get("jo", "")).strip()
    processes = body.get("processes") or []
    if not jo:
        return jsonify({"error": "JO wajib diisi"}), 400
    if not processes:
        return jsonify({"error": "Pilih minimal satu proses"}), 400
    try:
        sh = _fstl_spreadsheet()
        # Kumpulin dulu SEMUA nama sheet yang bakal dibutuhin buat SEMUA
        # proses yang dicentang (bisa 10-15 sheet kalau user centang banyak
        # proses), baru ambil sekaligus dalam SATU batch request -- lihat
        # penjelasan lengkap di docstring _fstl_batch_prefetch_sheets().
        needed_sheets = {FSTL_LP1_SHEET}
        for name in processes:
            cfg = FSTL_PROCESS_SOURCES.get(name.strip().upper())
            if cfg is not None:
                needed_sheets.update(_fstl_matching_sheet_names(sh, cfg["prefixes"], cfg["exact"]))
        _fstl_batch_prefetch_sheets(sh, needed_sheets)
        results = {name: fstl_keterangan_for_process(sh, jo, name) for name in processes}
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify({"results": results})


def _fstl_get_or_create_user_sheet(sh, safe_username):
    """Kalau username XX -> sheet 'XX_Kitir'. Kalau sudah ada, dipakai apa
    adanya (TIDAK menghapus sheet lama). Kalau belum ada, dibuat baru KOSONG
    -- baris 1 sengaja TIDAK ditulisi header apa pun, biar tetap kosong buat
    dipakai user sendiri (mis. baris filter Google Sheets), sama seperti pola
    di sheet contoh (kartu-kartu waste mulai dari baris 2).

    `sh` di sini SELALU handle spreadsheet KITIR (_fstl_kitir_spreadsheet()),
    BUKAN spreadsheet sumber data -- tab '{USER}_Kitir' hidup di spreadsheet
    kitir. Nggak perlu _fstl_invalidate_cache() lagi di sini: cache
    worksheet-titles (_fstl_worksheet_titles_cache) itu punya spreadsheet
    sumber data, jadi bikin tab baru di spreadsheet kitir nggak bikin cache
    itu basi."""
    sheet_name = f"{safe_username}_Kitir"
    try:
        ws = sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=200, cols=6)
    return ws, sheet_name


def _fstl_next_card_start_row(all_values):
    """Tentukan baris awal buat kartu baru, dengan aturan jeda 1 baris kosong
    di antara kartu-kartu (bukan lagi nempel/mepet seperti sebelumnya):

      - Sheet masih kosong sama sekali (belum ada kartu)  -> mulai baris 2
        (baris 1 tetap dibiarkan kosong seperti biasa, ini BUKAN "jeda antar
        kartu" jadi tidak perlu ditambah baris kosong lagi).
      - Baris terakhir yang kepakai LANGSUNG berisi data (0 baris kosong
        di bawahnya)         -> selipkan 1 baris kosong, baru mulai kartu.
      - Sudah ada TEPAT 1 baris kosong di bawah baris terakhir yang kepakai
                              -> lanjut langsung, jeda itu sudah cukup.
      - Sudah ada 2 baris kosong atau lebih                -> lanjut langsung
        juga (jangan nambah baris kosong lagi), jeda yang sudah ada dipakai
        apa adanya.

    `all_values` = hasil ws.get_all_values() (list of list of str)."""
    last_filled = 0  # nomor baris (1-based) terakhir yang punya isi
    for idx, row in enumerate(all_values, start=1):
        if any(str(cell).strip() for cell in row):
            last_filled = idx
    if last_filled == 0:
        return 2  # sheet kosong total, belum pernah ada kartu
    trailing_blank = len(all_values) - last_filled
    if trailing_blank == 0:
        return last_filled + 2  # tidak ada jeda -> selipkan 1 baris kosong
    return last_filled + 1  # sudah ada >=1 baris kosong -> lanjut langsung


@app.route("/api/fstl/save", methods=["POST"])
def fstl_save():
    """Body: {username, jo, produk, processes:[{name,keterangan,actionPlan,status}]}.
    Simpan ke sheet '{USERNAME}_Kitir' (dibuat kalau belum ada, tanpa hapus
    sheet lama) sebagai SATU "kartu" yang di-APPEND ke bagian PALING BAWAH
    sheet (bukan disisipkan di atas lagi) -- kolom A & baris 1 TIDAK pernah
    ditulisi apa pun, data mulai kolom B. Antar kartu WAJIB ada jeda 1 baris
    kosong (lihat _fstl_next_card_start_row): kalau kartu sebelumnya nempel
    tanpa jeda, disisipkan 1 baris kosong; kalau sudah ada jeda 1 baris,
    dipakai apa adanya; kalau jedanya sudah 2 baris atau lebih, tidak
    ditambah lagi:

      Baris judul (hijau #a9d08e) : SPK/JO : {jo} | Produk : {produk} |
                                    Waste Besar Proses : {p1, p2, ...}
      Baris label (biru  #9bc2e6) : Waste Besar Proses | Keterangan |
                                    Action Plan | Status   (label kolom)
      Baris data..N (hijau)       : satu baris per proses yang dicentang --
                                    {nama proses} | {keterangan} |
                                    {action plan} | {status}

    (kalau user centang 6 proses, berarti ada 6 baris hijau data di bawah
    baris label, bukan 6 kartu terpisah)."""
    body = request.get_json(force=True) or {}
    username = str(body.get("username", "")).strip()
    jo = str(body.get("jo", "")).strip()
    produk = str(body.get("produk", "")).strip()
    processes = body.get("processes") or []
    if not username:
        return jsonify({"error": "username wajib diisi"}), 400
    if not jo:
        return jsonify({"error": "JO wajib diisi"}), 400
    if not processes:
        return jsonify({"error": "Pilih minimal satu proses"}), 400

    safe_username = "".join(ch for ch in username if ch.isalnum() or ch in ("-", "_")).upper() or "USER"
    try:
        sh = _fstl_kitir_spreadsheet()
        ws, sheet_name = _fstl_get_or_create_user_sheet(sh, safe_username)

        process_names = [str(p.get("name", "")).strip() for p in processes if str(p.get("name", "")).strip()]
        card_title_row = [
            "", f"SPK/JO : {jo}", f"Produk : {produk}",
            f"Waste Besar Proses : {', '.join(process_names)}", "",
        ]
        label_row = ["", "Waste Besar Proses", "Keterangan", "Action Plan", "Status"]
        data_rows = [
            ["", p.get("name", ""), p.get("keterangan", ""), p.get("actionPlan", ""), (p.get("status") or "Open")]
            for p in processes
        ]
        rows_to_write = [card_title_row, label_row] + data_rows

        # APPEND ke bawah dengan jeda 1 baris kosong antar kartu (lihat
        # _fstl_next_card_start_row): sheet kosong -> mulai baris 2 seperti
        # biasa; kalau kartu terakhir nempel tanpa jeda -> disisipkan 1
        # baris kosong; kalau jeda sudah 1 baris atau lebih -> lanjut apa
        # adanya, tidak ditambah jeda baru.
        start_row = _fstl_next_card_start_row(ws.get_all_values())
        end_row = start_row + len(rows_to_write) - 1
        ws.update(f"A{start_row}:E{end_row}", rows_to_write, value_input_option="USER_ENTERED")

        # Pewarnaan: baris judul & baris label sama-sama BIRU (judul cuma
        # sampai kolom D, E dibiarkan putih; label sampai kolom E). Baris
        # data: cuma kolom nama proses (B) yang HIJAU, kolom
        # Keterangan/Action Plan/Status (C:E) tetap PUTIH.
        n = len(data_rows)
        title_row_num, label_row_num = start_row, start_row + 1
        ws.format(f"B{title_row_num}:D{title_row_num}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_TITLE_LABEL)})
        ws.format(f"B{label_row_num}:E{label_row_num}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_TITLE_LABEL)})
        if n:
            data_start, data_end = start_row + 2, start_row + 1 + n
            ws.format(f"B{data_start}:B{data_end}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_PROCESS)})
            ws.format(f"C{data_start}:E{data_end}", {"backgroundColor": _fstl_hex_to_rgb01(FSTL_COLOR_WHITE)})

        # Cache isi sheet ini (dipakai endpoint /api/fstl/list) jadi basi
        # begitu ada kartu baru ditulis -- buang dari cache biar list
        # berikutnya baca versi terbaru, bukan versi sebelum kartu ini ada.
        with _fstl_cache_lock:
            _fstl_sheet_values_cache.pop(sheet_name, None)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify({"success": True, "sheet": sheet_name})


FSTL_KITIR_SUFFIX_RE = re.compile(r"^(.+)_kitir$", re.IGNORECASE)


def _fstl_is_done_status(status_text):
    return str(status_text or "").strip().lower() in ("selesai", "done", "closed", "close", "ok")


def _fstl_parse_cards(rows, username, sheet_name):
    """Baca isi sheet '{username}_Kitir' (list baris dari get_all_values(),
    index kolom 0-based -- kolom B=1, C=2, D=3, E=4) balik jadi list kartu
    {username, sheet, jo, produk, processes:[...]}.

    Baris judul kartu dikenali dari kolom B yang diawali 'SPK/JO'. Baris
    label (header per-kartu) dikenali dari kolom B persis 'Waste Besar
    Proses' DAN kolom C persis 'Keterangan', lalu dilewati (bukan data).
    Baris lain yang punya isi di kolom B dianggap satu baris proses untuk
    kartu yang lagi aktif.

    Tiap baris proses ikut menyimpan nomor barisnya sendiri di sheet asli
    ("row", 1-based, sama seperti nomor baris di Google Sheets) -- dipakai
    /api/fstl/revisi buat tahu persis sel Keterangan mana yang mau ditulis
    ulang, tanpa perlu nebak lagi posisi kartunya di sheet."""

    def cell(row, idx):
        return row[idx].strip() if len(row) > idx else ""

    cards = []
    current = None
    for row_idx, row in enumerate(rows):
        row_num = row_idx + 1  # baris 1 di get_all_values() == baris 1 di sheet
        b, c, d, e = cell(row, 1), cell(row, 2), cell(row, 3), cell(row, 4)
        if not (b or c or d or e):
            continue
        if b.upper().startswith("SPK/JO"):
            if current is not None:
                cards.append(current)
            jo = b.split(":", 1)[1].strip() if ":" in b else b
            produk = c.split(":", 1)[1].strip() if ":" in c else c
            current = {"username": username, "sheet": sheet_name, "jo": jo, "produk": produk, "processes": []}
            continue
        if b == "Waste Besar Proses" and c == "Keterangan":
            continue  # baris label, dilewati
        if current is not None and b:
            current["processes"].append({
                "row": row_num, "name": b, "keterangan": c, "actionPlan": d, "status": e or "Open",
            })
    if current is not None:
        cards.append(current)
    return cards


@app.route("/api/fstl/list", methods=["GET"])
def fstl_list():
    """Ambil semua kartu waste dari SEMUA sheet '*_Kitir' (semua user)
    sekaligus, buat ditampilkan gabung di satu index. Filter yang tadinya
    berdasarkan Status (Open/Selesai) diganti jadi filter berdasarkan nama
    orang yang mengerjakan (turunan nama sheet '{USERNAME}_Kitir').

    Sengaja baca LANGSUNG dari Google Sheets (bukan lewat cache TTL yang
    dipakai /api/fstl/keterangan) -- endpoint ini cuma dipanggil sesekali
    (pas buka halaman Lampiran Waste), jadi lebih penting selalu dapat data
    paling baru (termasuk kartu yang baru saja disimpan) daripada hemat
    kuota API lewat cache basi.

    `sh` di sini spreadsheet KITIR (_fstl_kitir_spreadsheet()), BUKAN
    spreadsheet sumber data -- tab '*_Kitir' hidup di spreadsheet kitir."""
    try:
        sh = _fstl_kitir_spreadsheet()
        cards = []
        usernames = set()
        for ws_obj in sh.worksheets():
            sheet_name = ws_obj.title
            m = FSTL_KITIR_SUFFIX_RE.match(sheet_name)
            if not m:
                continue
            username = m.group(1)
            usernames.add(username)
            rows = ws_obj.get_all_values()
            cards.extend(_fstl_parse_cards(rows, username, sheet_name))
        for card in cards:
            statuses = [p["status"] for p in card["processes"]]
            card["status"] = "done" if statuses and all(_fstl_is_done_status(s) for s in statuses) else "open"
        return jsonify({"cards": cards, "usernames": sorted(usernames)})
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/fstl/revisi", methods=["POST"])
def fstl_revisi():
    """Body: {sheet, processes:[{row, keterangan}, ...]}.

    Revisi kartu waste yang SUDAH TERSIMPAN -- tapi CUMA kolom Keterangan
    (kolom C) di baris-baris proses yang dikirim yang ditulis ulang. JO,
    Produk, nama proses (centang), Action Plan, dan Status sengaja TIDAK
    bisa diubah lewat endpoint ini (sudah dikunci juga di frontend) --
    endpoint ini murni buat kasus "keterangannya salah/kurang lengkap,
    tolong dibetulkan", bukan buat ganti kartu jadi JO/proses lain.

    "row" per proses adalah nomor baris asli di sheet '{username}_Kitir'
    (dikirim balik oleh /api/fstl/list, lihat _fstl_parse_cards), jadi
    revisi ini langsung nulis ke sel yang tepat tanpa perlu cari ulang
    posisi kartunya."""
    body = request.get_json(force=True) or {}
    sheet_name = str(body.get("sheet", "")).strip()
    processes = body.get("processes") or []
    if not sheet_name or not FSTL_KITIR_SUFFIX_RE.match(sheet_name):
        return jsonify({"error": "Sheet tidak valid"}), 400
    if not processes:
        return jsonify({"error": "Tidak ada keterangan yang direvisi"}), 400

    updates = []
    for p in processes:
        row_num = p.get("row")
        if not isinstance(row_num, int) or row_num < 2:
            continue  # baris 1 nggak pernah dipakai buat data kartu, abaikan kalau ada yang aneh
        keterangan = str(p.get("keterangan", ""))
        updates.append({"range": f"C{row_num}", "values": [[keterangan]]})
    if not updates:
        return jsonify({"error": "Tidak ada baris valid untuk direvisi"}), 400

    try:
        sh = _fstl_kitir_spreadsheet()
        try:
            ws = sh.worksheet(sheet_name)
        except gspread.exceptions.WorksheetNotFound:
            return jsonify({"error": f"Sheet '{sheet_name}' tidak ditemukan di spreadsheet kitir"}), 404
        ws.batch_update(updates, value_input_option="USER_ENTERED")
        # Cache isi sheet ini jadi basi begitu keterangan direvisi, biar
        # /api/fstl/list berikutnya nunjukin versi terbaru.
        with _fstl_cache_lock:
            _fstl_sheet_values_cache.pop(sheet_name, None)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500
    return jsonify({"success": True})


# --------------------------------------------------------------------------
# 7. CHATBOT "TANYA JO" — AI yang memutuskan sendiri data apa yang perlu
#    dicari & di sheet mana, lewat tool-use ke query_group. Lihat
#    chatbot_engine.py untuk detail skema sheet & prompt yang dipakai.
# --------------------------------------------------------------------------
# Histori percakapan per sesi disimpan in-memory (sederhana, per session_id
# dari frontend) supaya user bisa tanya susulan ("kalau di Dry gimana?")
# tanpa perlu sebut ulang nomor JO. Kalau nanti dipakai multi-worker/lebih
# dari 1 proses, ganti ke penyimpanan bersama (Redis dsb).
_chat_sessions = {}
_chat_sessions_lock = threading.Lock()
CHAT_HISTORY_MAX_TURNS = 6  # jumlah giliran tanya-jawab yang disimpan per sesi


@app.route("/api/chatbot/ask", methods=["POST"])
def chatbot_ask():
    body = request.get_json(force=True) or {}
    message = str(body.get("message", "")).strip()
    session_id = str(body.get("session_id", "")).strip() or "default"
    if not message:
        return jsonify({"answer": "Pertanyaannya kosong, coba ketik dulu ya."}), 400

    with _chat_sessions_lock:
        history = _chat_sessions.get(session_id, [])

    try:
        result = chatbot_engine.run_agent(get_sheet, message, history=history)
    except Exception as exc:
        return jsonify({"answer": f"Gagal memproses pertanyaan lewat AI: {exc}"}), 500

    with _chat_sessions_lock:
        _chat_sessions[session_id] = chatbot_engine.trim_history(
            result["messages"], max_user_turns=CHAT_HISTORY_MAX_TURNS
        )

    return jsonify({"answer": result["answer"], "tool_calls": result["tool_calls"]})


@app.route("/api/chatbot/reset", methods=["POST"])
def chatbot_reset():
    """Mulai percakapan baru (buang histori) -- dipanggil kalau user pindah
    topik/JO dan mau chatbot-nya tidak kebawa konteks lama."""
    body = request.get_json(force=True) or {}
    session_id = str(body.get("session_id", "")).strip() or "default"
    with _chat_sessions_lock:
        _chat_sessions.pop(session_id, None)
    return jsonify({"success": True})


# --------------------------------------------------------------------------
# HEALTH CHECK (untuk memastikan servis & koneksi sheet hidup)
# --------------------------------------------------------------------------
@app.route("/api/health", methods=["GET"])
def health():
    return jsonify({"ok": True, "spreadsheet_configured": bool(SPREADSHEET_ID)})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)
