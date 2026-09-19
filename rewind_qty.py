"""
rewind_qty.py
Logika murni (tanpa akses Google Sheets) untuk mengisi kolom
Tanggal_Rewind, Qty_Awal_Rewind, Qty_Akhir_Rewind di sheet REWIND_PY
dari baris-baris REWIND_PY_RAW.

ATURAN QTY (satu sel bisa berisi beberapa bagian dipisah '+' atau baris baru):
  1. Angka bulat polos ('10', '18')            -> DIJUMLAHKAN jadi satu total di depan.
  2. n@panjang berstatus METER                  -> n dijumlahkan kalau panjangnya SAMA
       (1@100 + 2@100 = 3@100); panjang beda tidak digabung.
       METER ditentukan oleh satuan yang tertulis: 'm', 'mtr', 'meter'
       (dengan/tanpa spasi, mis. '2@500', '2@500m', '2@500 m', '2@400 meter').
       Kalau satuan TIDAK ditulis: panjang >= 100 -> meter, < 100 -> kg.
  3. n@panjang berstatus KG ('kg' tertulis, atau tanpa satuan & < 100)
     dan angka ber-kg tanpa '@' ('1,8kg')       -> TIDAK dijumlahkan, tampil apa adanya
       (dinormalkan jadi 'x kg'), tidak digabung, tidak di-unique.
  4. Bagian lain yang tidak dikenali (mis. '21,5' polos) -> tampil apa adanya.

Urutan output: total bulat + n@meter (panjang naik) + sisanya (urutan muncul).
"""

import re
from collections import OrderedDict

_MONTHS_ID = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun",
              "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]

_NUM = r"\d+(?:[.,]\d+)?"
_AT_RE = re.compile(rf"^({_NUM})\s*@\s*({_NUM})\s*(.*)$", re.IGNORECASE)
_KG_ONLY_RE = re.compile(rf"^({_NUM})\s*(?:kg|kilo)\.?$", re.IGNORECASE)
_METER_UNIT_RE = re.compile(r"^(?:m|mtr|meter|mtrs|meters)\.?$", re.IGNORECASE)
_KG_UNIT_RE = re.compile(r"^(?:kg|kilo)\.?$", re.IGNORECASE)


def _to_float(text):
    """'1.600' -> 1600 (titik = ribuan), '4,21' -> 4.21 (koma = desimal)."""
    t = str(text).strip().replace(" ", "")
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".")
    elif "," in t:
        t = t.replace(",", ".")
    elif "." in t:
        t = t.replace(".", "")
    return float(t)


def _fmt(value):
    """Angka -> teks; bulat tanpa desimal, selain itu koma desimal."""
    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))
    return f"{value:.10g}".replace(".", ",")


def parse_qty_text(text):
    """Return (whole_total, have_whole, meter_groups, others).
    meter_groups: dict panjang(float) -> jumlah(float); others: list str."""
    whole = 0
    have_whole = False
    meter = {}
    others = []

    raw = str(text or "").replace("\r", "\n")
    parts = []
    for line in raw.split("\n"):
        parts.extend(line.split("+"))

    for p in parts:
        p = p.strip()
        if not p or p == "-":
            continue

        if re.fullmatch(r"\d+", p):
            whole += int(p)
            have_whole = True
            continue

        m = _AT_RE.match(p)
        if m:
            n_txt, len_txt, unit = m.group(1), m.group(2), m.group(3).strip()
            try:
                n_val, len_val = _to_float(n_txt), _to_float(len_txt)
            except ValueError:
                others.append(p)
                continue
            if _METER_UNIT_RE.match(unit):
                kind = "meter"
            elif _KG_UNIT_RE.match(unit):
                kind = "kg"
            elif unit == "":
                kind = "meter" if len_val >= 100 else "kg"
            else:
                others.append(p)
                continue
            if kind == "meter":
                meter[len_val] = meter.get(len_val, 0) + n_val
            else:
                others.append(f"{n_txt}@{len_txt} kg")
            continue

        m = _KG_ONLY_RE.match(p)
        if m:
            others.append(f"{m.group(1)} kg")
            continue

        others.append(p)

    return whole, have_whole, meter, others


def combine_qty(texts):
    """texts: list teks sel (satu per baris RAW yang cocok). Return string
    hasil gabungan, '' kalau tidak ada isi sama sekali."""
    total = 0
    have_total = False
    meter = {}
    others = []
    for t in texts:
        w, hw, mg, oth = parse_qty_text(t)
        if hw:
            total += w
            have_total = True
        for k, v in mg.items():
            meter[k] = meter.get(k, 0) + v
        others.extend(oth)

    out = []
    if have_total:
        out.append(str(total))
    for length in sorted(meter):
        out.append(f"{_fmt(meter[length])}@{_fmt(length)}")
    out.extend(others)
    return " + ".join(out)


def format_date_ddmmmyy(d):
    """date -> 'DD/Mmm/YY' (nama bulan Indonesia, tidak tergantung locale)."""
    return f"{d.day:02d}/{_MONTHS_ID[d.month - 1]}/{d.year % 100:02d}"


def combine_dates(dates):
    """dates: iterable date. Unik, urut kronologis, pemisah koma."""
    uniq = sorted(set(d for d in dates if d is not None))
    return ", ".join(format_date_ddmmmyy(d) for d in uniq)
