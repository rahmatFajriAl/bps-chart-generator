import matplotlib
matplotlib.use('Agg')
import re
import io as _io
import json
import textwrap
import numpy as np
import pandas as pd
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
# 2. PEMBACA SHEET EXCEL — otomatis mengenali header 1 baris ATAU header 2
#    baris bertingkat (pakai info MERGE CELL ASLI, bukan tebak-tebakan), dan
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


def _read_meta(wb):
    """Baca metadata dari sheet tersembunyi '_meta' (dibuat oleh wizard 'Buat Tabel Baru'),
    kalau tidak ada, kembalikan dict kosong -> nilai default dipakai di tempat lain."""
    if "_meta" not in wb.sheetnames:
        return {}
    meta = {}
    for row in wb["_meta"].iter_rows(values_only=True):
        if row and row[0]:
            meta[str(row[0]).strip()] = row[1] if len(row) > 1 else None
    return meta


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
    meta = _read_meta(wb)
    source = (str(meta.get("Sumber Data")).strip() if meta.get("Sumber Data") else "") or None

    return {
        "title": str(title).strip(),
        "category_label": category_label,
        "categories": categories_clean,
        "series_labels": series_labels,
        "series_values": series_values,
        "is_total_row": is_total_row,
        "source": source,
    }


# =========================================================================
# 2b. PEMBACA CSV — struktur lebih sederhana daripada Excel (tidak ada merge
#     cell / header bertingkat), tapi tetap dikembalikan dalam bentuk data
#     yang SAMA dengan load_sheet() supaya bisa dipakai plot_chart() apa
#     adanya. Delimiter & encoding dideteksi otomatis supaya file CSV dari
#     berbagai sumber (koma, titik-koma, Excel Windows) tetap terbaca.
# =========================================================================
def _decode_csv_bytes(raw):
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _read_csv_bytes(file_bytes):
    text = _decode_csv_bytes(file_bytes)
    try:
        df = pd.read_csv(_io.StringIO(text), sep=None, engine="python",
                         dtype=str, keep_default_na=False)
    except Exception:
        # fallback kalau deteksi delimiter otomatis gagal (mis. file 1 kolom)
        df = pd.read_csv(_io.StringIO(text), sep=";", dtype=str, keep_default_na=False)
    # buang kolom "Unnamed" kosong yang kadang muncul dari trailing delimiter
    df = df.loc[:, ~df.columns.str.match(r"^Unnamed", na=False) | (df.astype(str).ne("").any())]
    return df


def load_csv(file_bytes, title=None):
    """
    Baca file CSV jadi struktur data yang sama dengan load_sheet(): kolom
    pertama dianggap kategori (baris/wilayah), kolom lainnya jadi seri data
    numerik. CSV tidak punya konsep sheet, jadi seluruh file diperlakukan
    sebagai satu "sheet" tunggal.
    """
    if isinstance(file_bytes, (bytes, bytearray)):
        raw = bytes(file_bytes)
    else:
        raw = file_bytes.read()

    df = _read_csv_bytes(raw)
    if df.shape[1] < 2:
        raise ValueError("File CSV harus punya minimal 2 kolom: kategori + 1 kolom data.")

    category_label = str(df.columns[0]).strip()
    series_labels = [str(c).strip() for c in df.columns[1:]]

    categories_raw = df.iloc[:, 0].astype(str).tolist()
    series_values_raw = {lbl: df.iloc[:, i + 1].tolist() for i, lbl in enumerate(series_labels)}

    keep_idx = [
        i for i, cat in enumerate(categories_raw)
        if cat.strip() or any(str(series_values_raw[lbl][i]).strip() for lbl in series_labels)
    ]
    categories_raw = [categories_raw[i] for i in keep_idx]
    series_values_raw = {lbl: [vals[i] for i in keep_idx] for lbl, vals in series_values_raw.items()}

    series_values = normalize_series_values(series_values_raw)
    categories_clean = [c.strip() for c in categories_raw]
    is_total_row = [_is_total_like(c, categories_raw) for c in categories_raw]

    return {
        "title": (title or "Data CSV").strip(),
        "category_label": category_label,
        "categories": categories_clean,
        "series_labels": series_labels,
        "series_values": series_values,
        "is_total_row": is_total_row,
    }


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


DEFAULT_SOURCE = "BPS, Pendataan Potensi Desa (Podes) 2024"


def _footer(fig, data, figure_number):
    footer = f"Sumber: {data.get('source') or DEFAULT_SOURCE}"
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


# =========================================================================
# 3. PEMBUAT TABEL BARU DARI NOL (wizard "Buat Tabel Baru")
#    Menghasilkan file .xlsx dengan struktur PERSIS sama dengan yang dibaca
#    load_sheet() (judul di A1, header 1 baris, baris kode "(0)(1)(2)...",
#    lalu data), plus sheet tersembunyi "_meta" utk menyimpan metadata
#    (judul, no tabel, sumber data) supaya ikut terbawa kalau file ini
#    dibuka lagi nanti.
# =========================================================================
import datetime as _dt


def _safe_sheet_name(name, fallback="Data"):
    name = (name or fallback).strip()
    for ch in ("\\", "/", "*", "[", "]", ":", "?"):
        name = name.replace(ch, "-")
    return (name[:31] or fallback)


def build_table_workbook(judul, no_tabel, sumber, category_label, row_names, col_names, values):
    """
    values: list-of-list angka (boleh None/NaN utk sel kosong), ukuran
    len(row_names) x len(col_names), urutan sama dgn row_names/col_names.
    Tampilan dibuat netral (abu-abu/hitam-putih) meniru gaya tabel dokumen
    resmi BPS: header abu-abu tebal, baris kode miring, garis kotak penuh.
    """
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

    thin = Side(style="thin", color="808080")
    medium = Side(style="medium", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9D9D9")
    stripe_fill = PatternFill("solid", fgColor="F5F5F5")

    n_cols = len(col_names) + 1   # +1 utk kolom kategori (A)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = _safe_sheet_name(no_tabel)

    # --- baris 1: judul, digabung (merge) selebar tabel, garis tebal di bawahnya ---
    ws.cell(1, 1, judul or "Tabel Baru")
    if n_cols > 1:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    for j in range(1, n_cols + 1):
        ws.cell(1, j).border = Border(bottom=medium)
    title_cell = ws.cell(1, 1)
    title_cell.font = Font(bold=True, size=12)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 24

    # --- baris 2: header kolom (abu-abu, tebal, rata tengah, bisa 2 baris) ---
    header_cells = [ws.cell(2, 1, category_label or "Wilayah")]
    for j, col in enumerate(col_names, start=2):
        header_cells.append(ws.cell(2, j, col))
    for c in header_cells:
        c.font = Font(bold=True, color="000000")
        c.fill = header_fill
        c.border = border_all
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[2].height = 34

    # --- baris 3: baris kode "(1) (2) (3) ..." (miring, tipis) ---
    code_cells = [ws.cell(3, 1, "(1)")]
    for j in range(2, n_cols + 1):
        code_cells.append(ws.cell(3, j, f"({j})"))
    for c in code_cells:
        c.font = Font(italic=True, size=9, color="595959")
        c.border = border_all
        c.alignment = Alignment(horizontal="center", vertical="center")

    # --- baris data ---
    for i, row_name in enumerate(row_names):
        r = 4 + i
        name_cell = ws.cell(r, 1, row_name)
        name_cell.border = border_all
        name_cell.alignment = Alignment(horizontal="left", vertical="center")
        for j, _col in enumerate(col_names):
            val = values[i][j] if i < len(values) and j < len(values[i]) else None
            if isinstance(val, float) and np.isnan(val):
                val = None
            cell = ws.cell(r, j + 2, val)
            cell.border = border_all
            cell.alignment = Alignment(horizontal="center", vertical="center")
        if i % 2 == 1:  # baris selang-seling, biar gampang dibaca kalau datanya panjang
            for j in range(1, n_cols + 1):
                ws.cell(r, j).fill = stripe_fill

    ws.freeze_panes = "B4"
    ws.column_dimensions["A"].width = max(20, min(32, max((len(n) for n in row_names), default=14) + 4))
    for j, col in enumerate(col_names, start=2):
        letter = ws.cell(2, j).column_letter
        ws.column_dimensions[letter].width = max(14, min(22, len(col) // 2 + 10))

    meta = wb.create_sheet("_meta")
    meta.sheet_state = "hidden"
    meta.append(["Judul Tabel", judul or ""])
    meta.append(["No Tabel", no_tabel or ""])
    meta.append(["Sumber Data", sumber or ""])
    meta.append(["Dibuat", _dt.datetime.now().strftime("%Y-%m-%d %H:%M")])

    buf = _io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
