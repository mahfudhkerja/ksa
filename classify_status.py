"""
classify_status.py
===================
Klasifikasi kondisi stok gudang (kolom KETERANGAN + STATUS yang isinya teks
bebas, diketik manual oleh operator) menjadi kategori baku:

    BAIK          - stok siap kirim, kondisi bagus
    BISA_REWORK   - ex retur/reject tapi masih bisa diproses ulang jadi baik
    BISA_REWIND   - sisa rewind yang statusnya OK / siap dipakai
    KARANTINA     - ditahan RnD/QC, belum ada keputusan final
    RETUR         - baru masuk retur, belum ada info kondisi lanjutan
    NOT_OK        - reject final, tidak bisa dipakai/dikirim
    PERLU_REVIEW  - tidak cocok pola manapun -> jangan ditebak, cek manual

ATURAN UTAMA (kenapa bukan sekadar "cari kata kunci lalu tandai"):
-------------------------------------------------------------------
Banyak baris menceritakan PROSES, bukan kondisi akhir, misalnya:
    "NOT OK, SUDAH DICUTTER, BISA REWORK"   -> hasil akhirnya BISA_REWORK
    "BAIK SISA REWIND"                       -> hasil akhirnya BISA_REWIND
    "EX RETUR BELUM DI REWORK"               -> hasil akhirnya masih RETUR
      (belum dikerjakan, walau kata REWORK ada)

Maka pengecekan dilakukan sebagai CASCADE berurutan (bukan independen):
begitu satu aturan cocok, langsung berhenti -- jangan cek aturan di
bawahnya lagi. Urutan aturan disusun dari yang paling "final/menyimpulkan"
ke yang paling umum.
"""

import re


def _norm(text):
    """Uppercase + rapikan spasi supaya regex konsisten. None jadi string kosong."""
    if text is None:
        return ""
    s = str(text).upper()
    s = re.sub(r"\s+", " ", s).strip()
    return s


# Setiap aturan: (nama_kategori, pola_regex_yang_HARUS_ada, pola_regex_yang_TIDAK_BOLEH_ada)
# pola dicek pada gabungan KETERANGAN + " " + STATUS.
_RULES = [
    # 1. Sudah eksplisit dinyatakan bisa dirework -> ini SELALU jadi kesimpulan akhir,
    #    walau di depannya ada kata NOT OK / JELEK / REJECT.
    ("BISA_REWORK",
     r"BISA\s*(DI\s*)?REWORK",
     None),

    # 2. Sedang/masih menunggu dirework -> belum final jadi baik, tapi juga bukan
    #    reject permanen. Kita kelompokkan ke RETUR (perlu tindak lanjut produksi).
    ("RETUR",
     r"BELUM\s*(DI\s*)?REWORK|TUNGGU\s*REWORK|HARUS\s*REWORK",
     None),

    # 3. Sisa rewind yang statusnya OK/baik.
    ("BISA_REWIND",
     r"(SISA\s*RE?WIND|REWIND)\s*(KECIL\s*)?(OK|OKE|BAIK)|BAIK[, ]+SISA\s*RE?WIND|BAIK\s*SUDAH\s*RE?WIND",
     r"NOT\s*OK|BELUM\s*RE?WIND"),

    # 4. Karantina / masih ditahan RnD-QC, belum ada keputusan.
    ("KARANTINA",
     r"KARANTINA|TUNGGU\s*KONFIRMASI|BELUM\s*(DI\s*)?CEK|RND\s*BELUM\s*CEK",
     None),

    # 5. Retur yang belum ada info kondisi lanjutan (baru "PALET X RETUR" polos,
    #    atau ada catatan tapi belum ada verdict BAIK/NOT OK/REWORK).
    ("RETUR",
     r"\bRETUR\b",
     r"NOT\s*OKE?|BAIK|BISA\s*(DI\s*)?REWORK|JELEK"),

    # 6. Reject final -- NOT OK / JELEK, dan TIDAK disertai kata yang menyatakan
    #    itu masih bisa diproses (rework/rewind/baik).
    ("NOT_OK",
     r"NOT\s*OKE?|JELEK|\bBAP\b",
     r"BISA\s*(DI\s*)?REWORK|SISA\s*RE?WIND.*(OK|OKE|BAIK)|BAIK"),

    # 7. Baik / siap kirim, tanpa embel-embel reject.
    ("BAIK",
     r"\bBAIK\b|VISUAL\s*OK",
     r"NOT\s*OKE?|JELEK"),
]


def classify_status(keterangan, status):
    """
    Klasifikasi satu baris data gudang.

    Parameters
    ----------
    keterangan : str atau None -- isi kolom KETERANGAN
    status     : str atau None -- isi kolom STATUS

    Returns
    -------
    str -- salah satu dari:
        "BAIK", "BISA_REWORK", "BISA_REWIND", "KARANTINA",
        "RETUR", "NOT_OK", "PERLU_REVIEW"
    """
    gabungan = _norm(keterangan) + " " + _norm(status)
    gabungan = gabungan.strip()

    if not gabungan:
        return "PERLU_REVIEW"

    for kategori, pola_wajib, pola_larangan in _RULES:
        if re.search(pola_wajib, gabungan):
            if pola_larangan and re.search(pola_larangan, gabungan):
                continue  # ada kata yang membatalkan aturan ini, lanjut ke aturan berikut
            return kategori

    return "PERLU_REVIEW"


if __name__ == "__main__":
    # Contoh cepat pakai beberapa baris nyata dari data gudang
    contoh = [
        ("", "BAIK"),
        ("", "NOT OK"),
        ("PALET 3 REWORK", "HASIL REWORK EX RETUR"),
        ("", "SISA REWIND OK"),
        ("EX RAK 103.5 TURUN U/ DIREWORK DIPRODUKSI.", "BISA DI REWORK"),
        ("", "EX RETUR BISA DI REWORK"),
        ("PALET 132 RETUR", ""),
        ("", "KARANTINA"),
        ("RETUR BLM REWORK.", "RETUR REWORK KIRIM"),
        ("", "BELUM REWORK QC"),
        ("MASUK 06,30 BAP DARI RAK 146.8 PALET 13", "KERIPUT"),
        ("PALET 6.2 RETUR,(JELEK DI BAP)", "(JELEK DI BAP)"),
        ("", "BAIK SUDAH REWIND"),
        ("", "SISA REWIND NOT OK"),
    ]
    for ket, stat in contoh:
        print(f"{classify_status(ket, stat):15s} <- KETERANGAN={ket!r} STATUS={stat!r}")
