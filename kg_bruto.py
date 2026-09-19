"""
kg_bruto.py
Logika murni (tanpa akses Google Sheets) untuk mengisi kolom Kg_Bruto di
REWIND_PY dari sheet FORM_ST_1.

ATURAN
  1. Baris FORM_ST_1 dicocokkan lewat SUFFIX kolom JO (angka setelah '/' terakhir)
     = NO_JO di REWIND_PY. Kolom JO_DIGIT TIDAK dipakai.
  2. Hanya baris ber-STATUS: HASIL SLITTING, HASIL SORTIR, KIRIM, STOCK,
     STOCK BELUM SORTIR, KARANTINA. (Urutan ini = URUTAN PRIORITAS.)
  3. Tiap baris: JUMLAH_MASUK_GBJ dan BERAT/KG sama-sama dipecah per '+'
     (koma = desimal, BUKAN pemisah). Bagian JUMLAH yang angka bulat polos
     (tanpa '@') dipasangkan dengan berat di POSISI yang sama. Bagian ber-'@'
     diabaikan. Baris tanpa bagian polos dilewati. Kalau ada beberapa
     bagian polos, semua beratnya ikut.
  4. Kg_Bruto = MODUS dari semua berat terkumpul. Kalau tidak ada yang
     berulang (semua muncul 1x) atau seri, ambil yang STATUS-nya paling atas
     di urutan prioritas (kalau status sama: baris paling atas di sheet).
"""

import re

STATUS_PRIORITY = [
    "HASIL SLITTING",
    "HASIL SORTIR",
    "KIRIM",
    "STOCK",
    "STOCK BELUM SORTIR",
    "KARANTINA",
]
_RANK = {name: i for i, name in enumerate(STATUS_PRIORITY)}


def normalize_status(text):
    t = " ".join(str(text or "").upper().split())
    return t.replace("SLITING", "SLITTING")  # toleransi salah ketik 1 T


def status_rank(text):
    """Index prioritas (0 = paling atas); None kalau status tidak dipakai."""
    return _RANK.get(normalize_status(text))


def _to_float(text):
    t = str(text).strip().replace(" ", "")
    if not t:
        return None
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    try:
        return float(t)
    except ValueError:
        return None


def plain_weights(jumlah_text, berat_text):
    """Return list float: berat untuk tiap bagian JUMLAH yang angka bulat polos."""
    jumlah = [p.strip() for p in str(jumlah_text or "").split("+") if p.strip()]
    berat = [p.strip() for p in str(berat_text or "").split("+") if p.strip()]
    out = []
    for i, part in enumerate(jumlah):
        if not re.fullmatch(r"\d+", part):
            continue
        if i >= len(berat):
            continue
        v = _to_float(berat[i])
        if v is not None:
            out.append(v)
    return out


def format_weight(v):
    if abs(v - round(v)) < 1e-9:
        return str(int(round(v)))
    return f"{v:.10g}".replace(".", ",")


def pick_kg_bruto(candidates):
    """candidates: list (rank:int, row_idx:int, jumlah_text, berat_text) --
    hanya baris yang statusnya sudah lolos filter. Return teks berat atau ''."""
    entries = []  # (rank, row_idx, value)
    for rank, row_idx, jumlah, berat in candidates:
        for v in plain_weights(jumlah, berat):
            entries.append((rank, row_idx, v))
    if not entries:
        return ""

    counts, best = {}, {}
    for rank, row_idx, v in entries:
        k = round(v, 6)
        counts[k] = counts.get(k, 0) + 1
        cur = best.get(k)
        if cur is None or (rank, row_idx) < cur[:2]:
            best[k] = (rank, row_idx, v)

    top = max(counts.values())
    tied = [best[k] for k, c in counts.items() if c == top]
    tied.sort(key=lambda e: (e[0], e[1]))
    return format_weight(tied[0][2])
