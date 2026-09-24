import matplotlib
matplotlib.use('Agg')
import re
import json
import textwrap
import numpy as np
import openpyxl
import matplotlib.pyplot as plt

# =========================================================================
# 1. PARSER ANGKA (format ID/EN, plus jaring pengaman utk Number korup)
# =========================================================================
_MISSING_TOKENS = {"", "-", "--", "...", "..", "n/a", "na", "null", "none"}

def parse_id_number(raw):
    if raw is None:
        return float("nan")
    if isinstance(raw, bool):
        return float(raw)
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip()
    if s.lower() in _MISSING_TOKENS:
        return float("nan")
    s = s.replace(" ", "").replace("\xa0", "")
    has_dot, has_comma = "." in s, "," in s
    try:
        if has_dot and has_comma:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif has_comma:
            s = s.replace(",", ".")
        elif has_dot:
            parts = s.split(".")
            if len(parts) > 1 and all(len(p) == 3 for p in parts[1:]) and len(parts[0]) <= 3:
                s = s.replace(".", "")
        return float(s)
    except ValueError:
        return float("nan")


def fix_thousands(v):
    if not (isinstance(v, float) and not np.isnan(v) and not v.is_integer()):
        return v
    s = repr(v)
    if "." not in s:
        return v
    sign = ""
    if s.startswith("-"):
        sign, s = "-", s[1:]
    intpart, frac = s.split(".")
    if not (intpart.isdigit() and frac.isdigit()):
        return v
    frac = frac.ljust(3, "0")[:3]
    return float(sign + intpart + frac)


def _looks_like_broken_count_series(values):
    nums = [v for v in values if isinstance(v, (int, float)) and not (isinstance(v, float) and np.isnan(v))]
    if len(nums) < 2:
        return False
    frac_like = [v for v in nums if isinstance(v, float) and not v.is_integer()]
    int_like_count = len(nums) - len(frac_like)
    if not frac_like or int_like_count == 0:
        return False
    return int_like_count >= len(frac_like)


def normalize_series_values(series_values):
    parsed = {label: [parse_id_number(v) for v in values] for label, values in series_values.items()}
    fixed = {}
    for label, values in parsed.items():
        fixed[label] = [fix_thousands(v) for v in values] if _looks_like_broken_count_series(values) else values
    return fixed


# =========================================================================
# 2. PEMBACA SHEET — otomatis mengenali header 1 baris ATAU header 2 baris
#    bertingkat (pakai info MERGE CELL ASLI, bukan tebak-tebakan), dan
#    otomatis mengabaikan kolom "hantu" (header kosong & bukan bagian merge)
# =========================================================================
_CODE_ROW_RE = re.compile(r"^\(\d+\)$")


def _find_code_row(ws, start_row=2, max_scan=6):
    """Cari baris '(1) (2) (3) ...' -> penanda akhir header, awal data."""
    for r in range(start_row, start_row + max_scan):
        vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
        non_none = [v for v in vals if v is not None and str(v).strip() != ""]
        if non_none and all(_CODE_ROW_RE.match(str(v).strip()) for v in non_none):
            return r
    return None


def _merge_map_for_row(ws, row):
    """kolom -> kolom kiri-atas dari merge yang menaungi baris ini (kalau ada)."""
    m = {}
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= row <= rng.max_row:
            for c in range(rng.min_col, rng.max_col + 1):
                m[c] = rng.min_col
    return m


def _read_header(ws):
    """
    Baca header BERAPA PUN jumlah lapisnya (1 baris, 2 baris, dst) secara
    generik: cari dulu baris kode '(1)(2)(3)...' -> semua baris DI ATASNYA
    (mulai baris 2) dianggap lapisan header, digabung per kolom pakai info
    merge cell ASLI (bukan tebak-tebakan blank = lanjutan).
    """
    code_row = _find_code_row(ws, start_row=2, max_scan=6)
    header_rows = list(range(2, code_row)) if code_row else [2]
    data_start_row = (code_row + 1) if code_row else 3

    max_col = ws.max_column
    row_maps = []
    for hr in header_rows:
        vals = [ws.cell(hr, c).value for c in range(1, max_col + 1)]
        row_maps.append((vals, _merge_map_for_row(ws, hr)))

    def label_at(vals, mmap, col_idx):
        src = mmap.get(col_idx, col_idx)
        v = vals[src - 1]
        return None if v is None else str(v).strip()

    labels, valid_cols = [], []
    for c in range(2, max_col + 1):  # kolom 1 = kategori, dilewati di sini
        parts = [lab for vals, mmap in row_maps if (lab := label_at(vals, mmap, c))]
        if not parts:
            continue  # kolom kosong asli (bukan bagian merge apa pun) -> lewati
        labels.append(" ".join(parts))
        valid_cols.append(c)

    category_label = row_maps[0][0][0]
    return category_label, labels, valid_cols, data_start_row


def _is_total_like(label, all_labels):
    if not label.strip():
        return False
    low = label.strip().lower()
    if "jumlah" in low or "total" in low or low.startswith("kecamatan"):
        return True
    if label != label.lstrip():
        return False  # baris berindentasi = item biasa, bukan total
    # header kategori tanpa indentasi padahal saudara2nya berindentasi -> subtotal grup
    siblings_indented = any(o != o.lstrip() for o in all_labels if o.strip())
    if siblings_indented and label.strip() != "":
        return True
    return False


def load_sheet(path, sheet_name):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet_name]

    title = ws.cell(1, 1).value or sheet_name
    category_label, series_labels, valid_cols, data_start_row = _read_header(ws)

    categories_raw, series_values_raw = [], {lbl: [] for lbl in series_labels}
    for r in range(data_start_row, ws.max_row + 1):
        cat = ws.cell(r, 1).value
        row_vals = [ws.cell(r, c).value for c in valid_cols]
        if cat is None and all(v is None for v in row_vals):
            continue  # baris benar-benar kosong -> lewati
        cat = "" if cat is None else str(cat)
        categories_raw.append(cat)
        for lbl, v in zip(series_labels, row_vals):
            series_values_raw[lbl].append(v)

    series_values = normalize_series_values(series_values_raw)
    categories_clean = [c.strip() for c in categories_raw]
    is_total_row = [_is_total_like(c, categories_raw) for c in categories_raw]

    return {
        "title": str(title).strip(),
        "category_label": category_label,
        "categories": categories_clean,
        "series_labels": series_labels,
        "series_values": series_values,
        "is_total_row": is_total_row,
    }


import re
import textwrap
import numpy as np
import matplotlib.pyplot as plt

# =========================================================================
# Palet warna
# =========================================================================
SINGLE_COLOR = "#F0932B"
TWO_COLOR = ["#FBC384", "#C98A15"]
MULTI_PALETTE = ["#F0932B", "#EB6E24", "#B34A0E", "#6B2E0A",
                 "#FBC384", "#F6D186", "#8B5E1E", "#4A3113"]
VERTICAL_MAX_CATEGORIES = 8


def pick_palette(n):
    if n == 1:
        return [SINGLE_COLOR]
    if n == 2:
        return TWO_COLOR
    if n <= len(MULTI_PALETTE):
        return MULTI_PALETTE[:n]
    cmap = plt.get_cmap("YlOrBr", n + 2)
    return [cmap(i + 1) for i in range(n)]


def choose_orientation(n_categories, n_series, orientation="auto"):
    if orientation in ("horizontal", "vertical"):
        return orientation
    if n_series == 1 and n_categories <= VERTICAL_MAX_CATEGORIES:
        return "vertical"
    return "horizontal"


# =========================================================================
# Label seri di samping angka batang — mis. "1200 Laki-laki"
# =========================================================================
def _common_prefix(strs):
    if not strs:
        return ""
    s1, s2 = min(strs), max(strs)
    i = 0
    while i < len(s1) and i < len(s2) and s1[i] == s2[i]:
        i += 1
    return s1[:i]


def short_series_labels(series_labels):
    """
    Tebakan OTOMATIS (dipakai sbg nilai default di kotak isian panel,
    boleh ditimpa manual oleh pengguna lewat series_label_overrides).
    Membuang bagian AWAL yg sama di semua kolom, mis.
    ['Penduduk Laki','Penduduk Perempuan'] -> ['Laki','Perempuan'].
    """
    prefix = _common_prefix(series_labels)
    if prefix:
        cut = prefix.rfind(" ")
        prefix = prefix[: cut + 1] if cut != -1 else ""
    if not prefix:
        return list(series_labels)
    out = [l[len(prefix):].strip() for l in series_labels]
    return [o if o else l for o, l in zip(out, series_labels)]


def resolve_short_labels(series_labels, overrides=None):
    """
    Keterangan FINAL yg dipakai di samping angka: pakai isian manual dari
    panel (overrides) kalau ada, kalau tidak fallback ke tebakan otomatis.
    overrides: dict {nama_kolom_asli: teks_keterangan_manual}
    """
    auto = short_series_labels(series_labels)
    if not overrides:
        return auto
    return [overrides.get(lbl, "").strip() or a for lbl, a in zip(series_labels, auto)]


def _value_text(r, unit_hint, series_label, show_series_label):
    txt = f"{r:g}"
    if unit_hint:
        txt += f" {unit_hint}"
    if show_series_label and series_label:
        txt += f" {series_label}"
    return txt


# =========================================================================
# Klasifikasi kolom (utama / total / persen / rasio / kepadatan)
# =========================================================================
_COLUMN_TOTAL_KEYWORDS = {"jumlah", "total"}
_PCT_KEYWORDS = {"persen", "persentase", "distribusi", "%"}
_RATIO_KEYWORDS = {"rasio"}
_DENSITY_KEYWORDS = {"kepadatan"}

_TAG = {"total": " (total)", "pct": " (%)", "ratio": " (rasio)",
        "density": " (kepadatan)", "main": ""}


def classify_column(label):
    low = label.lower()
    if any(k in low for k in _COLUMN_TOTAL_KEYWORDS):
        return "total"
    if any(k in low for k in _PCT_KEYWORDS):
        return "pct"
    if any(k in low for k in _RATIO_KEYWORDS):
        return "ratio"
    if any(k in low for k in _DENSITY_KEYWORDS):
        return "density"
    return "main"


def guess_unit_from_title(title):
    m = re.search(r"\(([^()]+)\)(?!.*\()", title)
    if m:
        cand = m.group(1).strip()
        if cand and not cand.isdigit() and len(cand) <= 15:
            return cand
    return ""


# =========================================================================
# Penyaringan data
# =========================================================================
def _filtered_data(data, exclude_totals, columns):
    categories = list(data["categories"])
    series_labels = list(data["series_labels"])
    series_values = {k: list(v) for k, v in data["series_values"].items()}
    is_total_row = list(data.get("is_total_row", [False] * len(categories)))

    if columns:
        wanted = set(columns)
        matched = [l for l in series_labels if l in wanted]
        if not matched:
            print(f"[WARNING] Tidak ada kolom yang cocok dengan {columns}. "
                  f"Kolom tersedia: {series_labels}")
        else:
            series_labels = matched
            series_values = {l: series_values[l] for l in series_labels}

    if exclude_totals:
        keep = [i for i, is_tot in enumerate(is_total_row) if not is_tot]
        categories = [categories[i] for i in keep]
        series_values = {l: [series_values[l][i] for i in keep] for l in series_labels}

    return categories, series_labels, series_values


def _footer(fig, data, figure_number):
    footer = "Sumber: BPS, Pendataan Potensi Desa (Podes) 2024"
    if figure_number:
        footer = f"Gambar {figure_number}   |   " + footer
    fig.text(0.02, 0.005, footer, fontsize=8, style="italic")


def _title(ax, data, width=78):
    ax.set_title("\n".join(textwrap.wrap(data["title"], width)),
                  fontsize=10.5, fontweight="bold", loc="left", pad=14)


def plot_horizontal(data, categories, series_labels, series_values,
                     figure_number="", unit_hint="", show_series_label=True,
                     series_label_overrides=None):
    n_cat, n_series = len(categories), len(series_labels)
    colors = pick_palette(n_series)
    short_labels = (resolve_short_labels(series_labels, series_label_overrides)
                     if (show_series_label and n_series > 1) else [""] * n_series)

    row_h_in = 0.30 if n_series <= 2 else 0.24
    fig_h = max(4, n_cat * n_series * row_h_in + 1.8)
    fig, ax = plt.subplots(figsize=(9.5, fig_h))

    label_fontsize = 8 if n_series <= 3 else (7 if n_series <= 5 else 6.5)

    y = np.arange(n_cat)
    bar_h = 0.8 / n_series
    max_val = max((v for vals in series_values.values() for v in vals if not np.isnan(v)), default=1)

    for i, (label, color, short_lbl) in enumerate(zip(series_labels, colors, short_labels)):
        raw = series_values[label]
        vals = np.nan_to_num(raw, nan=0.0)
        offset = (i - (n_series - 1) / 2) * bar_h
        bars = ax.barh(y + offset, vals, height=bar_h * 0.85, color=color, label=label)
        for rect, r in zip(bars, raw):
            if np.isnan(r):
                continue
            txt = _value_text(r, unit_hint, short_lbl, show_series_label and n_series > 1)
            ax.text(rect.get_width() + max_val * 0.015,
                     rect.get_y() + rect.get_height() / 2,
                     txt, va="center", ha="left", fontsize=label_fontsize)

    ax.set_yticks(y)
    ax.set_yticklabels(categories)
    ax.invert_yaxis()
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.get_xaxis().set_visible(False)
    ax.set_xlim(0, max_val * 1.30 if max_val > 0 else 1)
    if n_series > 1:
        ax.legend(loc="lower right", frameon=False, fontsize=9)

    _title(ax, data)
    _footer(fig, data, figure_number)
    plt.tight_layout(rect=[0, 0.03, 1, 1])
    return fig


def plot_vertical(data, categories, series_labels, series_values,
                   figure_number="", unit_hint="", show_series_label=True,
                   series_label_overrides=None):
    n_cat, n_series = len(categories), len(series_labels)
    colors = pick_palette(n_series)
    short_labels = (resolve_short_labels(series_labels, series_label_overrides)
                     if (show_series_label and n_series > 1) else [""] * n_series)
    fig_w = max(6, n_cat * (1.3 if n_series == 1 else 1.9) + 1.5)
    fig, ax = plt.subplots(figsize=(fig_w, 5.8))

    x = np.arange(n_cat)
    bar_w = 0.75 / n_series
    max_val = max((v for vals in series_values.values() for v in vals if not np.isnan(v)), default=1)
    label_fontsize = 9 if n_series <= 2 else 7.5

    for i, (label, color, short_lbl) in enumerate(zip(series_labels, colors, short_labels)):
        raw = series_values[label]
        vals = np.nan_to_num(raw, nan=0.0)
        offset = (i - (n_series - 1) / 2) * bar_w
        bars = ax.bar(x + offset, vals, width=bar_w * 0.92, color=color, label=label)
        for rect, r in zip(bars, raw):
            if np.isnan(r):
                continue
            txt = _value_text(r, unit_hint, short_lbl, show_series_label and n_series > 1)
            ax.text(rect.get_x() + rect.get_width() / 2,
                     rect.get_height() + max_val * 0.03,
                     txt, ha="center", va="bottom", fontsize=label_fontsize,
                     rotation=90 if n_series > 1 else 0)

    wrap_width = max(10, int(70 / max(n_cat, 1)))
    wrapped_labels = ["\n".join(textwrap.wrap(c, wrap_width)) for c in categories]
    ax.set_xticks(x)
    ax.set_xticklabels(wrapped_labels, rotation=0, ha="center", fontsize=8.5)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)
    ax.get_yaxis().set_visible(False)
    ax.set_ylim(0, max_val * 1.55 if max_val > 0 else 1)
    if n_series > 1:
        ax.legend(loc="upper right", frameon=False, fontsize=9)

    _title(ax, data)
    _footer(fig, data, figure_number)
    plt.tight_layout(rect=[0, 0.05, 1, 1])
    return fig


def plot_chart(data, exclude_totals=True, figure_number="", save_path=None,
               unit_hint="", columns=None, orientation="auto", show_series_label=True,
               series_label_overrides=None):
    categories, series_labels, series_values = _filtered_data(data, exclude_totals, columns)

    n_cat, n_series = len(categories), len(series_labels)
    if n_cat == 0 or n_series == 0:
        print(f"[SKIP] '{data['title']}' tidak punya data untuk digambar.")
        return None

    final_orientation = choose_orientation(n_cat, n_series, orientation)
    if final_orientation == "vertical":
        fig = plot_vertical(data, categories, series_labels, series_values,
                             figure_number=figure_number, unit_hint=unit_hint,
                             show_series_label=show_series_label,
                             series_label_overrides=series_label_overrides)
    else:
        fig = plot_horizontal(data, categories, series_labels, series_values,
                               figure_number=figure_number, unit_hint=unit_hint,
                               show_series_label=show_series_label,
                               series_label_overrides=series_label_overrides)

    if save_path:
        fig.savefig(save_path, dpi=200, bbox_inches="tight")
        print(f"Chart disimpan: {save_path}")

    return fig