"""Konversi teks Qty_Awal_Rewind (mis. '42 + 6@830 + 1@6,05 kg') jadi METER JUMBO.

Variabel:
  up        = UP_Slitting  (jumlah lajur hasil slitting per jumbo)
  potongan  = Potongan     (panjang 1 rol utuh, meter)
  kg_bruto  = Kg_Bruto     (berat 1 rol utuh, kg)

Aturan per bagian (dipisah '+'):
  N               angka polos (rol utuh)   -> N / up * potongan
  N@X  / N@Xm     rol sisa, meter diketahui -> N / up * X
  N@X kg          rol sisa, hanya kg        -> N / up * (X / kg_bruto) * potongan
  X kg  (tanpa @) sama dengan 1@X kg        -> 1 / up * (X / kg_bruto) * potongan
  X m   (tanpa @) sama dengan 1@X m         -> 1 / up * X
"""
import re

_NUM = r"[\d.,]+"
_AT_RE = re.compile(rf"^({_NUM})\s*@\s*({_NUM})\s*(kg|m)?$", re.IGNORECASE)
_UNIT_RE = re.compile(rf"^({_NUM})\s*(kg|m)$", re.IGNORECASE)
_PLAIN_RE = re.compile(rf"^({_NUM})$")
_SPLIT_RE = re.compile(r"\s*\+\s*|,\s+")  # '+' atau koma+spasi; '6,05' (desimal) tidak terpecah


def parse_number(text):
    """Format Indonesia: '6,05' -> 6.05, '1340.' -> 1340, '1.600' -> 1600."""
    text = str(text).strip().replace(" ", "")
    if not text:
        return None
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")
    elif "." in text:
        text = text.replace(".", "")  # titik saja = pemisah ribuan / titik nyasar
    try:
        return float(text)
    except ValueError:
        return None


def hitung_meter_jumbo(qty_text, up, potongan, kg_bruto):
    """Return (meter_jumbo | None, alasan_gagal | None). None = kosongkan sel."""
    qty_text = str(qty_text or "").strip()
    if not qty_text:
        return None, "kosong"
    up, potongan, kg_bruto = (parse_number(v) for v in (up, potongan, kg_bruto))
    if not up:
        return None, "UP_Slitting kosong/0"

    plain, meter, kg = 0.0, 0.0, 0.0
    need_potongan = need_kg = False
    for raw in _SPLIT_RE.split(qty_text):
        part = raw.strip()
        if not part:
            continue
        m = _AT_RE.match(part)
        if m:
            left, right, unit = parse_number(m.group(1)), parse_number(m.group(2)), (m.group(3) or "").lower()
        else:
            m = _UNIT_RE.match(part)
            if m:
                left, right, unit = 1.0, parse_number(m.group(1)), m.group(2).lower()
            elif _PLAIN_RE.match(part):
                n = parse_number(part)
                if n is None:
                    return None, f"bagian tidak terbaca: '{part}'"
                plain += n
                need_potongan = True
                continue
            else:
                return None, f"bagian tidak terbaca: '{part}'"
        if left is None or right is None:
            return None, f"angka tidak terbaca: '{part}'"
        if unit == "kg":
            need_kg = need_potongan = True
            kg += (left / up) * right  # dikali (1/kg_bruto)*potongan di akhir
        else:  # '@X' atau '@Xm' atau 'Xm': meter sudah diketahui
            meter += (left / up) * right

    if need_potongan and not potongan:
        return None, "Potongan kosong/0"
    if need_kg and not kg_bruto:
        return None, "Kg_Bruto kosong/0"
    total = meter
    if plain:
        total += (plain / up) * potongan
    if kg:
        total += kg / kg_bruto * potongan
    return total, None
