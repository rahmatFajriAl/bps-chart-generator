"""Narasi otomatis (bahasa Indonesia) dari data tabel. Murni teks, tidak menyentuh HTML/DB."""
import re

import numpy as np

# kolom yang nilainya TIDAK boleh dijumlahkan / dipersentasekan terhadap total
_NON_ADDITIVE = ("kepadatan", "rasio", "persen", "rata", "indeks", "laju", "angka", "%", "per ")
_SUP = {"km2": "km²", "m2": "m²", "km3": "km³", "m3": "m³"}


def fmt_id(v, max_dec=2):
    """12500.5 -> '12.500,5' (format angka Indonesia)."""
    s = f"{float(v):,.{max_dec}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    if "," in s:
        s = s.rstrip("0").rstrip(",")
    return s


def _unit(unit):
    u = (unit or "").strip()
    return _SUP.get(u.lower(), u)


def _is_total(name):
    n = str(name).strip().lower()
    return n.startswith("jumlah") or n.startswith("total")


def _join(names):
    names = list(names)
    if len(names) <= 1:
        return "".join(names)
    return ", ".join(names[:-1]) + " dan " + names[-1]


def _clean_label(label):
    return re.sub(r"\s*\(.*?\)", "", str(label)).strip()


def narrate_column(data, column, unit="", exclude_totals=True):
    """Narasi satu kolom. Return string (kosong kalau data tidak cukup)."""
    cats = list(data["categories"])
    vals = np.array(data["series_values"][column], dtype=float)
    keep = [i for i in range(len(cats))
            if not np.isnan(vals[i]) and not (exclude_totals and _is_total(cats[i]))]
    if not keep:
        return ""
    names = [str(cats[i]) for i in keep]
    v = vals[keep]
    n = len(v)

    u = _unit(unit)
    u_txt = f" {u}" if u else ""
    cat_lbl = (data.get("category_label") or "").strip()
    prefix = f"{cat_lbl} " if cat_lbl and cat_lbl.lower() != "wilayah" else ""
    attr = _clean_label(column).lower()
    title = str(data.get("title") or "").strip()
    intro = f"Berdasarkan {title}, " if title else ""

    def show(idx):   # daftar nama untuk nilai yang sama (maks 3)
        return _join([f"{prefix}{names[i]}" for i in idx[:3]])

    if n == 1:
        return f"{intro}{prefix}{names[0]} memiliki {attr} sebesar {fmt_id(v[0])}{u_txt}.".replace(
            f"{intro}{prefix}", f"{intro}{prefix}", 1)

    hi, lo = v.max(), v.min()
    hi_idx = [i for i in range(n) if v[i] == hi]
    lo_idx = [i for i in range(n) if v[i] == lo]

    if hi == lo:
        return f"{intro}seluruh {n} {cat_lbl.lower() or 'wilayah'} memiliki {attr} yang sama, yaitu {fmt_id(hi)}{u_txt}."

    s1 = (f"{intro}{show(hi_idx)} memiliki {attr} terbesar, yaitu {fmt_id(hi)}{u_txt}, "
          f"sedangkan {show(lo_idx)} terkecil, yaitu {fmt_id(lo)}{u_txt}.")
    if s1[0].islower():
        s1 = s1[0].upper() + s1[1:]

    parts = [s1]
    additive = not any(k in str(column).lower() for k in _NON_ADDITIVE)
    mean = v.mean()
    parts.append(f"Rata-rata {attr} dari {n} {cat_lbl.lower() or 'wilayah'} adalah {fmt_id(mean)}{u_txt}.")

    if additive and n >= 5 and v.sum() > 0 and (v >= 0).all():
        order = np.argsort(-v)[:3]
        share = v[order].sum() / v.sum() * 100
        parts.append(f"Tiga teratas ({_join([prefix + names[i] for i in order])}) menyumbang "
                     f"{fmt_id(share, 1)} persen dari total {fmt_id(v.sum())}{u_txt}.")
    return " ".join(parts)


def build_narrative(data, columns, unit="", exclude_totals=True):
    """Narasi untuk beberapa kolom, dipisah baris kosong."""
    out = []
    for c in columns:
        t = narrate_column(data, c, unit, exclude_totals)
        if t:
            out.append(t)
    return "\n\n".join(out)