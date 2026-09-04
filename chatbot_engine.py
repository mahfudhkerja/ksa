"""
chatbot_engine.py
==================
Chatbot "Tanya JO" — AI-driven, pakai OpenRouter API (kompatibel format
OpenAI: chat.completions + "tools"/function calling), model
deepseek/deepseek-v4-flash-0731.

ALUR LOGIKA (contoh: user tanya "hasil produksi dry JO 1234 gimana?")
----------------------------------------------------------------------
1. Pesan user dikirim ke DeepSeek beserta:
   - SYSTEM_PROMPT (instruksi peran + aturan "jangan ngarang")
   - TOOL_DEF (definisi tool `query_group`, isinya daftar semua grup
     sheet + kolom yang ada di masing-masing grup)
2. DeepSeek baca pertanyaan, "mikir": kata kunci "dry" & "hasil produksi"
   -> cocok dengan grup "dry" (kolom HASIL_PRODUKSI_METER/KG ada di situ),
   dan nomor JO "1234" ada di kalimat.
   DeepSeek TIDAK menjawab langsung -- dia balikin response yang isinya
   `finish_reason = "tool_calls"` dengan permintaan panggil
   `query_group(group="dry", jo="1234")`.
3. Kode Python (bukan AI) yang benar-benar eksekusi: buka sheet
   DRY_1..DRY_5 satu-satu (karena belum tahu JO itu diproses di mesin Dry
   nomor berapa), filter baris yang kolom JO-nya == "1234", kembalikan
   baris yang ketemu (mis. dari DRY_3, ada HASIL_PRODUKSI_METER=850,
   HASIL_PRODUKSI_KG=210, dst).
4. Hasil tool itu (JSON mentah, data asli dari sheet) dikirim BALIK ke
   DeepSeek sebagai pesan role "tool".
5. DeepSeek baca data itu, lalu menyusun jawaban akhir dalam Bahasa
   Indonesia -- HANYA memakai angka yang ada di data tsb. Kalau baris
   kosong (JO tidak ketemu di Dry manapun), dia wajib bilang "tidak
   ditemukan", bukan menebak.
6. Kalau pertanyaannya gabungan (mis. "hasil produksi dry DAN sisa
   stocknya"), DeepSeek akan minta panggil `query_group` lagi untuk grup
   "validasi_stock" sebelum menjawab -- makanya ini jalan sebagai LOOP
   (lihat run_agent), bukan cuma 1x tanya-jawab.

Kebutuhan:
  pip install openai      (dipakai cuma sebagai HTTP client, base_url-nya
                            diarahkan ke OpenRouter, bukan ke OpenAI)
  export OPENROUTER_API_KEY=...
"""

import os
import re
import json

from openai import OpenAI

import import_engine

MODEL = os.environ.get("CHATBOT_MODEL", "deepseek/deepseek-v4-flash-0731")
MAX_AGENT_STEPS = 6  # batas jaga-jaga biar nggak looping tool call terus-terusan

_client = None


def get_ai_client():
    global _client
    if _client is None:
        _client = OpenAI(
            api_key=os.environ["OPENROUTER_API_KEY"],
            base_url="https://openrouter.ai/api/v1",
        )
    return _client


# --------------------------------------------------------------------------
# HELPER: normalisasi JO & angka
# --------------------------------------------------------------------------

def normalize_jo(value):
    """Samakan format JO: '1234', 'JO1234', 'SPK123/1234' dianggap sama --
    kalau ada '/', ambil bagian setelah '/' terakhir, lalu buang semua
    karakter selain digit."""
    if value is None:
        return ""
    s = str(value).strip()
    if "/" in s:
        s = s.rsplit("/", 1)[-1]
    return re.sub(r"\D", "", s)


def to_number(value):
    if value is None:
        return None
    s = str(value).strip()
    if not s:
        return None
    s2 = s.replace(".", "").replace(",", ".")
    try:
        return float(s2)
    except ValueError:
        return s  # bukan angka, balikin apa adanya (mis. teks keterangan)


def _produk_tokens(value):
    """Normalisasi nama produk jadi string "padat" yang tahan beda spasi/
    tanda baca, mis. 'RCE 56G D3', 'RCE56G D3', 'rce-56g-d3' semuanya jadi
    'RCE56GD3'. Sengaja TIDAK memecah huruf & angka yang nempel (mis. '56G'
    atau kode varian 'D3'/'D5') supaya kode varian yang mirip tapi beda
    (D3 vs D5) tidak keanggap sama.

    Nama fungsi tetap `_produk_tokens` (dipertahankan) walau isinya sekarang
    satu string padat, bukan list token, biar pemanggilnya tidak berubah."""
    if value is None:
        return ""
    s = str(value).upper()
    return re.sub(r"[^A-Z0-9]", "", s)


def _produk_tokens_match(query_compact, row_compact):
    """True kalau nama produk yang dicari user dianggap "sama produk"
    dengan nama produk di satu baris sheet -- LONGGAR (menoleransi nama
    yang lebih pendek/lebih panjang), karena nama di laporan printing/dry/
    slitting/bag sering nggak ditulis lengkap/konsisten satu sama lain.

    Aturan: cocok kalau versi "padat" (tanpa spasi/tanda baca) salah satu
    adalah AWALAN dari yang lain -- jadi user cari "RCE 56G" bisa nemu
    baris "RCE 56G D3" (varian lebih spesifik), dan sebaliknya user cari
    "RCE 56G D3" tetap nemu baris yang cuma nulis "RCE 56G" (nama lebih
    pendek). Dengan begini kode varian di akhir seperti D3/D5/LAM dst
    tetap dibedakan dengan tegas -- bukan cuma "mirip"."""
    q, r = query_compact, row_compact
    if not q or not r:
        return False
    return q.startswith(r) or r.startswith(q)


def _compute_slitting_summary(raw_rows):
    """raw_rows: list baris mentah SL_1 (hasil ws.get_all_records) yang
    JO-nya sudah cocok. Hitung ringkasan "HASIL SLITTING" pakai rumus
    yang SAMA PERSIS dengan kolom E sheet Validasi (lihat
    import_engine._compute_hasil_slitting): kumpulkan pasangan
    (HASIL_ROLL, METER/ROLL) tiap baris, kelompokkan per METER/ROLL,
    jumlahkan HASIL_ROLL per kelompok -- kelompok dengan jumlah BARIS
    terbanyak (modus, syarat >1 & tidak seri) ditulis sebagai angka
    polos, kelompok lain ditulis '{jumlah}@{meter}'.

    Dihitung di sini (Python, deterministik) supaya jawaban chatbot soal
    "hasil slitting" selalu sama dengan angka di sheet Validasi -- tidak
    dibiarkan dihitung ulang manual oleh AI dari baris-baris mentah
    (rawan salah kelompok/salah jumlah)."""
    sl_matches = []
    for r in raw_rows:
        k_val = import_engine._parse_flexible_number(r.get("HASIL_ROLL"))
        o_val = import_engine._parse_flexible_number(r.get("METER/ROLL"))
        if k_val is not None and o_val is not None:
            sl_matches.append((k_val, import_engine._format_number(o_val)))
    if not sl_matches:
        return None
    return import_engine._compute_hasil_slitting(sl_matches)


# --------------------------------------------------------------------------
# SKEMA SHEET — daftar semua tab produksi, dikelompokkan per proses.
# Kalau nanti nama tab / kolom berubah, cukup update di sini saja.
# --------------------------------------------------------------------------

SHEET_GROUPS = {
    "printing": {
        "label": "Printing (mesin 2-5)",
        "sheets": ["PRINTING_2", "PRINTING_3", "PRINTING_4", "PRINTING_5"],
        "jo_column": "JO",
        "produk_column": "NAMA_PRODUK",
        "columns": [
            "TANGGAL", "SHIFT/OPERATOR", "JAM_AWAL", "JAM_AKHIR", "SPK", "JO",
            "NAMA_PRODUK", "UKURAN_PRODUK", "JENIS_BAHAN", "MICRON_BAHAN",
            "UK_BAHAN", "NO_LOT_BAHAN", "METER_BAHAN", "KG_BAHAN",
            "NO_ROL_JADI", "METER_AKHIR_JADI", "KG_JADI", "JAM_TURUN_JADI",
            "WIP_RAK", "WIP_BARIS", "KETERANGAN_JADI",
        ],
    },
    "rewind": {
        "label": "Rewind",
        "sheets": ["RW_1"],
        "jo_column": "JO",
        "produk_column": "NAMA_PRODUK",
        "columns": [
            "TANGGAL", "MESIN", "OPERATOR", "SHIFT", "JAM_KERJA", "FINISH",
            "SPK", "JO", "NAMA_PRODUK", "HPREW_MC", "HPREW_NO", "HPREW_METER",
            "HPREW_KG", "HPREW_METER_AKHIR", "HPREW_KG_AKHIR", "WASTE",
            "LETAK_WIP", "LETAK_BARIS", "WAKTU_NAIK", "WAKTU_TURUN",
            "KETERANGAN_WASTE",
        ],
    },
    "dry": {
        "label": "Dry Laminasi (mesin 1-5)",
        "sheets": ["DRY_1", "DRY_2", "DRY_3", "DRY_4", "DRY_5"],
        "jo_column": "JO",
        "produk_column": "NAMA_PRODUK",
        "columns": [
            "TANGGAL", "SHIFT/OPERATOR", "JAM_AWAL", "JAM_AKHIR", "SPK", "JO",
            "NAMA_PRODUK", "LAP1/LAP2", "HPREW_NO_BAHAN", "HPREW_NO",
            "HPREW_METER", "HPREW_KG", "HPREW_JAM_NAIK", "LAPISAN_JENIS",
            "LAPISAN_MIC", "LAPISAN_UKURAN", "LAPISAN_METER", "LAPISAN_KG",
            "LAPISAN_NO_BAHAN", "HASIL_PRODUKSI_NO", "HASIL_PRODUKSI_METER",
            "HASIL_PRODUKSI_KG", "HASIL_PRODUKSI_JAM_TURUN",
            "HASIL_PRODUKSI_WIP", "HASIL_PRODUKSI_BARIS", "AGING_ROOM",
            "WASTE", "KETERANGAN_WASTE",
        ],
    },
    "extrusi": {
        "label": "Extrusi",
        "sheets": ["EX_1"],
        "jo_column": "JO",
        "produk_column": "NAMA_PRODUK",
        "columns": [
            "TANGGAL", "SHIFT/OPERATOR", "JAM_AWAL", "JAM_AKHIR", "SPK", "JO",
            "NAMA_PRODUK", "BAHAN_BAKU_LOT", "BAHAN_BAKU_NO",
            "BAHAN_BAKU_METER", "BAHAN_BAKU_KG", "BAHAN_BAKU_JAM_NAIK",
            "BAHAN_LAPISAN_PP_RANDOM", "BAHAN_LAPISAN_MASTER_BATCH",
            "HASIL_PRODUKSI_NO", "HASIL_PRODUKSI_METER", "HASIL_PRODUKSI_KG",
            "HASIL_PRODUKSI_JAM_TURUN", "WASTE", "KETERANGAN_WASTE",
        ],
    },
    "solvent_free": {
        "label": "Solvent Free",
        "sheets": ["SF_1"],
        "jo_column": "JO",
        "produk_column": "NAMA_PRODUK",
        "columns": [
            "TANGGAL", "SHIFT/OPERATOR", "JAM_AWAL", "JAM_AKHIR", "SPK", "JO",
            "NAMA_PRODUK", "LAP1/LAP2", "HPREW_NO_BAHAN", "HPREW_NO",
            "HPREW_METER", "HPREW_KG", "LAPISAN_JENIS", "LAPISAN_BERAT_JENIS",
            "LAPISAN_MIC", "LAPISAN_UKURAN", "LAPISAN_METER", "LAPISAN_KG",
            "LAPISAN_NO_BAHAN", "HASIL_PRODUKSI_NO", "HASIL_PRODUKSI_METER",
            "HASIL_PRODUKSI_KG", "HASIL_PRODUKSI_JAM_TURUN", "WASTE",
            "KETERANGAN_WASTE",
        ],
    },
    "slitting": {
        "label": "Slitting",
        "sheets": ["SL_1"],
        "jo_column": "SPK/JO",
        "produk_column": "NAMA_PRODUK",
        "columns": [
            "TANGGAL", "MESIN", "SHIFT", "SPK/JO", "NAMA_PRODUK", "UK", "UP",
            "NO_ROLL", "METER", "KG", "HASIL_ROLL", "RW", "ROLL_BAIK",
            "ROLL_KURLEB", "METER/ROLL", "TOTAL_METER", "HASIL_RW",
            "KG_BRUTO", "KG_NETTO", "OPERATOR", "JAM", "PERSEN", "STATUS",
            "WASTE", "KETERANGAN_WASTE",
        ],
    },
    "bag_making": {
        "label": "Bag Making",
        "sheets": ["BAG_1"],
        "jo_column": "SPK/JO",
        "produk_column": "PRODUK",
        "columns": [
            "TANGGAL", "MESIN", "SHIFT", "SPK", "SPK/JO", "PRODUK",
            "UKURAN_BAG_1", "UKURAN_BAG_2", "AWAL_ROLL", "AWAL_METER",
            "AWAL_KG", "AKHIR_BAIK", "AKHIR_KW", "AKHIR_JELEK", "AKHIR_TOTAL",
            "AKHIR_METER", "AKHIR_AREA", "BERAT/PACK", "WASTE_KG",
            "KETERANGAN_WASTE",
        ],
    },
    "jo_master": {
        "label": "Kumpulan JO (identitas & spek order)",
        "sheets": ["JO_1"],
        "jo_column": "JO",
        "columns": [
            "CEK_JO_AKHIR", "STATUS_JO", "MANUAL", "CUSTOMER", "NO_PO", "JO",
            "KEMASAN", "ORDER", "UK", "UP", "POTONGAN", "BHN", "METER",
            "ROLL", "EST", "OPP", "LAP_1", "LAP_2", "KETERANGAN",
        ],
    },
    "laporan_produksi": {
        "label": "Laporan Produksi (rekap tiap proses per JO, per mesin)",
        "sheets": ["LP_1"],
        "jo_column": "SPK/JO",
        "produk_column": "Produk",
        "columns": [
            "SPK/JO1", "Tanggal", "SPK/JO", "Produk", "Customer",
            "Meter_Order", "Urutan_Proses", "Proses", "Mesin",
            "Tanggal_Proses", "Meter_Awal", "Meter_Hasil", "Meter_Waste",
            "Presentase_Waste", "Indikator", "Meter_waste2",
            "Faktor_Penyebab_Waste", "KLASIFIKASI",
        ],
    },
    "validasi_stock": {
        "label": "Validasi Stock (rekonsiliasi akhir & sisa stock)",
        "sheets": ["Validasi"],
        "jo_column": "JO",
        "produk_column": "NAMA",
        "columns": [
            "TANGGAL", "JO", "NAMA", "ORDER", "HASIL SLITTING", "POTONGAN",
            "HASIL SLIT\n(QTY)", "HASIL BAG", "VALIDASI", "FORM SERAH TERIMA",
            "TOTAL", "SELISIH", "STATUS",
        ],
    },
}


def _schema_description():
    lines = []
    for key, g in SHEET_GROUPS.items():
        lines.append(f"- \"{key}\" ({g['label']}): kolom tersedia = {', '.join(g['columns'])}")
    return "\n".join(lines)


# --------------------------------------------------------------------------
# EKSEKUSI TOOL: query_group -> beneran baca sheet & filter per JO
# --------------------------------------------------------------------------

def query_group(get_sheet_fn, group, jo=None, produk=None):
    """Cari semua baris (di semua tab dalam grup ini) yang cocok.

    Isi salah satu (boleh dua-duanya, nanti digabung pakai OR):
    - `jo`: dicocokkan ke kolom JO/SPK grup ini (exact, lewat normalize_jo).
    - `produk`: dicocokkan ke kolom nama produk grup ini (LONGGAR, lewat
      token matching -- lihat _produk_tokens_match) karena nama produk di
      laporan tiap proses sering ditulis nggak lengkap/beda singkatan satu
      sama lain (mis. sheet Printing nulis "RCE 56G D3" tapi sheet
      Slitting cuma nulis "RCE 56G")."""
    g = SHEET_GROUPS.get(group)
    if g is None:
        return {"error": f"Grup '{group}' tidak dikenal. Grup yang valid: {list(SHEET_GROUPS.keys())}"}
    if not jo and not produk:
        return {"error": "Isi salah satu: 'jo' atau 'produk'."}

    target_jo = normalize_jo(jo) if jo else None
    produk_col = g.get("produk_column")
    query_tokens = _produk_tokens(produk) if (produk and produk_col) else None
    if produk and not produk_col:
        return {"error": f"Grup '{group}' tidak punya kolom nama produk, cari pakai 'jo' saja."}

    jo_col = g["jo_column"]
    results = []
    errors = []
    matched_raw_rows = []  # dipakai khusus buat ringkasan HASIL SLITTING
    for sheet_name in g["sheets"]:
        try:
            ws = get_sheet_fn(sheet_name)
            rows = ws.get_all_records(numericise_ignore=["all"])
        except Exception as exc:
            errors.append(f"{sheet_name}: {exc}")
            continue
        for r in rows:
            jo_ok = target_jo is not None and normalize_jo(r.get(jo_col)) == target_jo
            produk_ok = (
                query_tokens is not None
                and _produk_tokens_match(query_tokens, _produk_tokens(r.get(produk_col)))
            )
            if not (jo_ok or produk_ok):
                continue
            clean_row = {k: to_number(v) for k, v in r.items() if str(v).strip() != ""}
            results.append({"sheet": sheet_name, "row": clean_row})
            matched_raw_rows.append(r)

    out = {
        "group": group,
        "jo": jo,
        "produk_dicari": produk,
        "jumlah_baris_ditemukan": len(results),
        "data": results,
    }
    if produk:
        out["catatan_pencarian_produk"] = (
            "Pencocokan nama produk di atas LONGGAR (tahan beda spasi/tanda "
            "baca; nama yang lebih pendek dianggap cocok kalau jadi AWALAN "
            "dari nama yang lebih panjang, atau sebaliknya) supaya tetap "
            "ketemu walau nama di sheet ini ditulis lebih pendek/lebih "
            "panjang dari yang dicari -- tapi kode varian di akhir seperti "
            "D3 vs D5 tetap dibedakan tegas, tidak ikut nyampur. Kalau "
            "'data' berisi lebih dari satu nama produk berbeda (cek field "
            "row NAMA_PRODUK/PRODUK/Produk/NAMA di tiap baris), JANGAN "
            "digabung asal -- sebutkan variasi namanya ke user kalau perlu "
            "klarifikasi mana yang dimaksud."
        )
    if errors:
        out["errors"] = errors

    if group == "slitting" and matched_raw_rows:
        summary = _compute_slitting_summary(matched_raw_rows)
        if summary:
            out["ringkasan_hasil_slitting"] = summary
            out["catatan_ringkasan"] = (
                "Field 'ringkasan_hasil_slitting' di atas sudah dihitung otomatis "
                "dengan rumus yang SAMA PERSIS seperti kolom HASIL SLITTING di sheet "
                "Validasi: HASIL_ROLL dijumlah per kelompok METER/ROLL yang sama, "
                "kelompok dengan jumlah baris terbanyak (modus) ditulis sebagai angka "
                "polos, kelompok lain ditulis '{jumlah}@{meter}'. Kalau user tanya "
                "total/hasil slitting suatu JO, PAKAI angka dari field ini apa adanya "
                "-- JANGAN dihitung ulang manual dari baris-baris di 'data'."
            )

    return out


def search_produk(get_sheet_fn, keyword, max_hasil=20):
    """Cari SEMUA nama produk unik (di SEMUA grup yang punya kolom nama
    produk) yang cocok dengan `keyword`, buat langkah KONFIRMASI sebelum
    tanya lebih lanjut -- jadi kalau user cuma sebut nama produk (bukan
    nomor JO) atau sebut nama pendek yang ternyata punya beberapa varian
    (mis. "RCE 56G" -> bisa jadi "RCE 56G D3", "RCE 56G LAM", dst), AI bisa
    tampilkan daftarnya dulu ke user buat dipilih/dikonfirmasi, sebelum
    query_group dipanggil.

    Pencocokan pakai token matching yang sama longgarnya dengan query_group
    (lihat _produk_tokens_match), jadi otomatis nangkep nama lengkap maupun
    singkat/beda spasi. Hasil dikelompokkan per nama produk PERSIS seperti
    tertulis di sheet (supaya user bisa lihat variasi penulisannya), + di
    grup mana ditemukan, contoh JO/SPK terkait, dan tanggal terakhir
    ditemukan."""
    query_tokens = _produk_tokens(keyword)
    if not query_tokens:
        return {"error": "Kata kunci nama produk kosong."}

    found = {}  # key: nama produk persis (as-is) -> info
    errors = []
    for group, g in SHEET_GROUPS.items():
        produk_col = g.get("produk_column")
        if not produk_col:
            continue
        jo_col = g["jo_column"]
        for sheet_name in g["sheets"]:
            try:
                ws = get_sheet_fn(sheet_name)
                rows = ws.get_all_records(numericise_ignore=["all"])
            except Exception as exc:
                errors.append(f"{sheet_name}: {exc}")
                continue
            for r in rows:
                nama = str(r.get(produk_col) or "").strip()
                if not nama:
                    continue
                if not _produk_tokens_match(query_tokens, _produk_tokens(nama)):
                    continue
                entry = found.setdefault(nama, {
                    "nama_produk": nama,
                    "grup_ditemukan": set(),
                    "contoh_jo": set(),
                    "tanggal_terakhir": None,
                    "jumlah_baris": 0,
                })
                entry["grup_ditemukan"].add(group)
                jo_val = normalize_jo(r.get(jo_col))
                if jo_val:
                    entry["contoh_jo"].add(jo_val)
                tgl = str(r.get("TANGGAL") or r.get("Tanggal") or "").strip()
                if tgl and (entry["tanggal_terakhir"] is None or tgl > entry["tanggal_terakhir"]):
                    entry["tanggal_terakhir"] = tgl
                entry["jumlah_baris"] += 1

    hasil = []
    for entry in found.values():
        hasil.append({
            "nama_produk": entry["nama_produk"],
            "grup_ditemukan": sorted(entry["grup_ditemukan"]),
            "contoh_jo": sorted(entry["contoh_jo"])[:10],
            "tanggal_terakhir_ditemukan": entry["tanggal_terakhir"],
            "jumlah_baris": entry["jumlah_baris"],
        })
    hasil.sort(key=lambda x: x["nama_produk"])

    out = {
        "keyword": keyword,
        "jumlah_nama_produk_unik_ditemukan": len(hasil),
        "hasil": hasil[:max_hasil],
        "catatan": (
            "Ini daftar nama produk UNIK (apa adanya, persis tulisan di "
            "sheet) yang cocok dengan keyword user -- termasuk yang "
            "penulisannya lebih pendek/lebih panjang. Kalau hasilnya lebih "
            "dari 1 nama, WAJIB tampilkan daftarnya ke user dan minta "
            "konfirmasi produk/JO mana yang dimaksud sebelum lanjut cari "
            "detail (jangan langsung nebak salah satu). Kalau cuma 1 nama "
            "ketemu, boleh langsung lanjut sambil bilang ke user nama "
            "produknya biar dia bisa koreksi kalau salah. Setelah user "
            "konfirmasi, panggil query_group dengan 'produk' (bukan harus "
            "'jo') untuk grup yang relevan -- pencocokannya sudah longgar "
            "jadi tetap ketemu walau di grup itu nama produknya ditulis "
            "lebih pendek/beda dari yang dikonfirmasi user."
        ),
    }
    if len(hasil) > max_hasil:
        out["catatan_potongan"] = f"Hasil dipotong ke {max_hasil} nama teratas dari {len(hasil)} yang ketemu -- minta user perjelas keyword kalau butuh yang lain."
    if errors:
        out["errors"] = errors
    return out


# Format tool DeepSeek (sama dengan format function-calling OpenAI):
# {"type": "function", "function": {name, description, parameters}}
TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "search_produk",
            "description": (
                "Cari nama produk (bukan nomor JO) di SEMUA grup sheet "
                "sekaligus, buat langkah KONFIRMASI dulu waktu user tanya "
                "pakai nama produk (mis. 'RCE 56G') bukan nomor JO. "
                "Balikin daftar nama produk UNIK yang cocok (apa adanya "
                "persis tulisan di sheet, termasuk varian nama yang lebih "
                "pendek/lebih panjang), plus grup mana ditemukan, contoh "
                "JO terkait, dan tanggal terakhir ditemukan (buat jawab "
                "'produk X terakhir produksi kapan'). Pencocokan LONGGAR "
                "(tahan beda spasi/tanda baca -- versi lebih pendek dianggap cocok kalau jadi AWALAN dari versi lebih panjang, atau sebaliknya), jadi "
                "cari 'RCE 56G' bisa nemu 'RCE 56G D3', 'RCE 56 G', dst. "
                "Panggil tool ini SEBELUM query_group tiap kali pertanyaan "
                "user menyebut nama produk (bukan nomor JO eksplisit)."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "keyword": {"type": "string", "description": "Nama/kata kunci produk yang disebut user, contoh: 'RCE 56G'"},
                },
                "required": ["keyword"],
            },
        },
    },
    {
        "type": "function",
        "function": {
            "name": "query_group",
            "description": (
                "Cari data produksi (detail per baris) di salah satu grup "
                "sheet berikut:\n" + _schema_description() + "\n\n"
                "Isi salah satu: 'jo' (nomor/kode JO, exact match) ATAU "
                "'produk' (nama produk, pencocokan LONGGAR berbasis "
                "kata/angka -- tetap ketemu walau nama di grup ini ditulis "
                "lebih pendek/panjang/beda singkatan dari yang disebut "
                "user, jadi TIDAK perlu nama super lengkap & TIDAK perlu "
                "sama persis dengan grup lain). Kalau user sudah sebut nama "
                "produk dan sudah dikonfirmasi lewat search_produk, pakai "
                "'produk' langsung di sini -- tidak perlu ubah ke JO dulu, "
                "kecuali usernya juga sebut nomor JO. Panggil tool ini "
                "sekali per grup yang relevan; boleh berkali-kali (grup "
                "berbeda-beda) kalau pertanyaannya butuh data dari "
                "beberapa proses sekaligus."
            ),
            "parameters": {
                "type": "object",
                "properties": {
                    "jo": {"type": "string", "description": "Nomor/kode JO yang dicari, contoh: '1234'. Kosongkan kalau cari pakai nama produk."},
                    "produk": {"type": "string", "description": "Nama produk yang dicari, contoh: 'RCE 56G D3'. Kosongkan kalau cari pakai nomor JO."},
                    "group": {
                        "type": "string",
                        "enum": list(SHEET_GROUPS.keys()),
                        "description": "Grup sheet yang mau dicari.",
                    },
                },
                "required": ["group"],
            },
        },
    },
]


SYSTEM_PROMPT = (
    "Kamu adalah asisten internal pabrik yang menjawab pertanyaan staff "
    "tentang data produksi berdasarkan nomor JO (job order) ATAU nama "
    "produk, dengan cara mencari langsung ke Google Sheets lewat tool "
    "`search_produk` dan `query_group`.\n\n"
    "Cara kerja kamu:\n"
    "1. Pahami dari pertanyaan user: apakah dia sebut NOMOR JO, atau NAMA "
    "PRODUK (mis. 'RCE 56G'), dan data/informasi apa persisnya yang "
    "diminta (misalnya: hasil produksi di Dry, no lot bahan saat printing, "
    "mesin apa saja yang dipakai, sisa stock, kapan terakhir produksi, JO "
    "apa saja untuk produk itu, dll).\n"
    "1b. Kalau user sebut NAMA PRODUK (bukan nomor JO eksplisit), panggil "
    "`search_produk` dulu SEBELUM `query_group`, walaupun namanya "
    "pendek/singkat. Tool ini balikin semua varian nama persis seperti "
    "tertulis di sheet (bisa lebih pendek/lebih panjang dari yang disebut "
    "user), grup mana ditemukan, contoh JO, dan tanggal terakhir. Kalau "
    "hasilnya lebih dari satu nama produk berbeda, WAJIB tampilkan "
    "daftarnya ke user dan minta konfirmasi mana yang dimaksud SEBELUM "
    "lanjut cari detail lain -- jangan menebak salah satu. Kalau cuma satu "
    "nama ketemu, boleh langsung lanjut, tapi tetap sebutkan nama "
    "produknya persis ke user (biar dia bisa koreksi kalau salah).\n"
    "2. Pilih grup sheet yang paling mungkin punya jawabannya (lihat "
    "deskripsi kolom tiap grup di definisi tool), lalu panggil "
    "`query_group` -- isi 'jo' kalau user sebut nomor JO, atau isi "
    "'produk' kalau berdasarkan nama produk yang sudah dikonfirmasi/"
    "ditemukan lewat search_produk. Kalau pertanyaannya butuh beberapa "
    "jenis data sekaligus, panggil tool untuk tiap grup yang relevan.\n"
    "2b. Untuk pencarian pakai 'produk': pencocokannya SUDAH LONGGAR "
    "(tahan beda spasi/tanda baca -- versi lebih pendek/panjang tetap cocok selama salah satunya AWALAN dari yang lain, tapi kode varian di akhir seperti D3 vs D5 tetap dibedakan), jadi nama produk yang "
    "ditulis beda-beda di tiap laporan (printing/dry/slitting/bag dll "
    "sering nggak konsisten -- ada yang lengkap ada yang disingkat) TETAP "
    "bisa ketemu. Jangan ubah-ubah nama produk sendiri berharap 'lebih "
    "cocok' ke satu sheet tertentu -- pakai saja nama yang sudah "
    "dikonfirmasi user apa adanya, tool yang urus perbedaan penulisannya.\n"
    "3. Kalau hasil query kosong tapi kamu belum yakin sudah cek semua "
    "kemungkinan grup yang relevan, coba grup lain dulu sebelum menyerah.\n"
    "4. SETELAH dapat data dari tool, jawab HANYA berdasarkan data itu. "
    "JANGAN PERNAH mengarang angka, nama mesin, tanggal, atau field apa "
    "pun yang tidak benar-benar ada di hasil tool.\n"
    "5. Kalau setelah dicari datanya memang tidak ada / JO atau produk "
    "tidak ketemu, katakan terus terang ke user bahwa datanya tidak "
    "ditemukan -- jangan menebak atau mengisi dengan asumsi.\n"
    "6. Khusus pertanyaan soal HASIL SLITTING / total hasil slitting suatu "
    "JO: kalau hasil query_group grup 'slitting' punya field "
    "'ringkasan_hasil_slitting', WAJIB pakai angka dari field itu apa "
    "adanya sebagai jawaban -- itu sudah dihitung dengan rumus yang sama "
    "persis dengan kolom HASIL SLITTING di sheet Validasi (roll "
    "dijumlah per kelompok panjang/meter yang sama, kelompok terbanyak "
    "ditulis polos, sisanya ditulis '{jumlah}@{meter}'). JANGAN pernah "
    "menjumlahkan sendiri kolom HASIL_ROLL dari baris-baris mentah di "
    "'data' untuk pertanyaan ini.\n"
    "7. Jawab singkat, jelas, dalam Bahasa Indonesia. Kalau ada beberapa "
    "baris/mesin, tampilkan sebagai list bernomor."
)


def trim_history(messages, max_user_turns=6):
    """Potong histori TANPA merusak pasangan assistant(tool_calls) <->
    tool(hasil). Satu "giliran" = mulai dari satu pesan role="user"
    (pertanyaan asli, selalu berisi teks) sampai sebelum pesan
    role="user" berikutnya -- di dalamnya bisa ada beberapa pasangan
    assistant/tool kalau AI manggil tool berkali-kali, dan itu HARUS ikut
    terbawa utuh, nggak boleh kepotong di tengah.

    Kalau motongnya asal jumlah pesan (mis. messages[-20:]), gampang
    kepotong pas di antara pesan assistant yang minta tool_calls dan
    pesan tool balasannya -- itu yang bikin DeepSeek nolak dengan error
    "Messages with role 'tool' must be a response to a preceding message
    with 'tool_calls'"."""
    system_msgs = [m for m in messages if m.get("role") == "system"]
    rest = [m for m in messages if m.get("role") != "system"]

    turn_start_idx = [
        i for i, m in enumerate(rest)
        if m.get("role") == "user" and isinstance(m.get("content"), str)
    ]
    if len(turn_start_idx) > max_user_turns:
        cut_at = turn_start_idx[-max_user_turns]
        rest = rest[cut_at:]

    return system_msgs + rest


def run_agent(get_sheet_fn, user_message, history=None):
    """Jalankan satu putaran percakapan chatbot memakai DeepSeek.

    `history` opsional: list pesan sebelumnya (format OpenAI messages)
    kalau mau multi-turn dengan konteks; kalau None, percakapan baru.

    Balikin dict {"answer": str, "tool_calls": [...], "messages": [...]}.
    """
    client = get_ai_client()
    messages = list(history) if history else [{"role": "system", "content": SYSTEM_PROMPT}]
    messages.append({"role": "user", "content": user_message})

    tool_call_log = []

    for _ in range(MAX_AGENT_STEPS):
        response = client.chat.completions.create(
            model=MODEL,
            messages=messages,
            tools=TOOLS,
        )
        msg = response.choices[0].message

        if not msg.tool_calls:
            messages.append({"role": "assistant", "content": msg.content})
            return {
                "answer": (msg.content or "Maaf, saya belum bisa menjawab itu.").strip(),
                "tool_calls": tool_call_log,
                "messages": messages,
            }

        # DeepSeek minta panggil satu atau beberapa tool -> eksekusi semua,
        # lalu kirim balik hasilnya sebagai pesan role "tool".
        messages.append({
            "role": "assistant",
            "content": msg.content,
            "tool_calls": [tc.model_dump() for tc in msg.tool_calls],
        })
        for tc in msg.tool_calls:
            try:
                args = json.loads(tc.function.arguments)
            except (json.JSONDecodeError, TypeError):
                args = {}

            if tc.function.name == "query_group":
                result = query_group(get_sheet_fn, args.get("group"), args.get("jo"), args.get("produk"))
            elif tc.function.name == "search_produk":
                result = search_produk(get_sheet_fn, args.get("keyword"))
            else:
                result = {"error": f"Tool '{tc.function.name}' tidak dikenal"}

            tool_call_log.append({"tool": tc.function.name, "input": args, "result": result})
            messages.append({
                "role": "tool",
                "tool_call_id": tc.id,
                "content": json.dumps(result, ensure_ascii=False),
            })

    return {
        "answer": "Maaf, pertanyaannya kompleks dan saya belum berhasil mengumpulkan datanya. Coba tanya lebih spesifik ya (sebutkan nomor JO & proses yang dimaksud).",
        "tool_calls": tool_call_log,
        "messages": messages,
    }
