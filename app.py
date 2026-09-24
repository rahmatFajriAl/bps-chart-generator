import io
import base64
import zipfile
from pathlib import Path

import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

from chart_core import (
    load_sheet, plot_chart, classify_column, short_series_labels,
    guess_unit_from_title, _TAG,
)

BASE_DIR = Path(__file__).parent
LOGO_PATH = BASE_DIR / "logo-bps.webp"
BRAND = "BPS Lampung Utara"


def load_logo():
    """Return (PIL image, base64 string) atau (None, None) kalau file logo tidak ada."""
    try:
        from PIL import Image
        img = Image.open(LOGO_PATH)
        b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        return img, b64
    except Exception:
        return None, None


LOGO_IMG, LOGO_B64 = load_logo()

st.set_page_config(page_title=f"Chart Generator | {BRAND}",
                   page_icon=LOGO_IMG if LOGO_IMG else ":material/bar_chart:",
                   layout="wide", initial_sidebar_state="collapsed")

# ------------------------------------------------------------------ CSS ----
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"], .stApp { font-family: 'Plus Jakarta Sans', sans-serif; }
.stApp { background: linear-gradient(180deg, #FFF3E3 0%, #FFF8F0 280px); }

#MainMenu, footer, [data-testid="stDecoration"], [data-testid="stToolbar"] { display: none !important; }
header[data-testid="stHeader"] { background: transparent; height: 0; }

.block-container { max-width: 1100px; padding: 2rem 1.5rem 4rem; }

/* HERO */
.hero {
  background: linear-gradient(135deg, #F0932B 0%, #EB6E24 55%, #B34A0E 100%);
  border-radius: 24px; padding: 2rem 2rem 1.8rem; color: #fff;
  box-shadow: 0 18px 40px -18px rgba(179, 74, 14, .55);
  position: relative; overflow: hidden; margin-bottom: 1.4rem;
}
.hero::after {
  content: ""; position: absolute; right: -60px; top: -60px; width: 220px; height: 220px;
  border-radius: 50%; background: rgba(255,255,255,.12);
}
.hero .badge {
  display: inline-block; background: rgba(255,255,255,.2); backdrop-filter: blur(6px);
  padding: .25rem .8rem; border-radius: 999px; font-size: .75rem; font-weight: 600;
  letter-spacing: .04em; margin-bottom: .8rem;
}
.hero h1 { color: #fff; margin: 0; font-weight: 800; line-height: 1.15;
           font-size: clamp(1.5rem, 4.5vw, 2.4rem); padding: 0; }
.hero p { margin: .6rem 0 0; opacity: .92; font-size: clamp(.85rem, 2.4vw, 1rem); max-width: 620px; }

/* BRAND */
.brand { display: flex; align-items: center; gap: .8rem; margin-bottom: 1rem; position: relative; z-index: 1; }
.brand .logo { background: #fff; border-radius: 16px; padding: .4rem .55rem; display: flex; align-items: center;
               box-shadow: 0 8px 18px -8px rgba(0,0,0,.35); }
.brand .logo img { height: 44px; width: auto; display: block; }
.brand .name { font-weight: 700; font-size: clamp(.85rem, 2.6vw, 1.05rem); letter-spacing: .02em; line-height: 1.2; }
.brand .name small { display: block; font-weight: 500; opacity: .85; font-size: .72rem; letter-spacing: .06em; text-transform: uppercase; }

/* HOW-TO (satu kartu panduan, bukan tombol) */
.howto { background: #fff; border: 1px solid #F6E3CE; border-radius: 20px; padding: 1.1rem 1.3rem 1.2rem;
         box-shadow: 0 8px 26px -14px rgba(120, 60, 10, .3); margin-bottom: 1rem; cursor: default; }
.howto .head { font-size: .72rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase;
               color: #9A7B5C; margin-bottom: .9rem; }
.howto .row { display: flex; gap: .9rem; align-items: flex-start; position: relative; padding-bottom: 1.1rem; }
.howto .row:last-child { padding-bottom: 0; }
.howto .row:not(:last-child)::before { content: ""; position: absolute; left: 15px; top: 34px; bottom: 2px;
                                       width: 2px; background: #F6D9B8; }
.howto .num { flex: 0 0 auto; width: 32px; height: 32px; border-radius: 50%; border: 2px solid #F0932B;
              background: #FFF1DE; color: #B34A0E; font-weight: 800; font-size: .85rem;
              display: flex; align-items: center; justify-content: center; }
.howto .t { font-weight: 700; color: #2B2118; font-size: 1rem; line-height: 32px; }
.howto .d { color: #7A6650; font-size: .85rem; margin-top: -.2rem; }

/* STAT CARDS */
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: .8rem; margin: 0 0 1.2rem; }
.stat { background: #fff; border-radius: 18px; padding: .9rem 1rem;
        box-shadow: 0 6px 20px -10px rgba(120, 60, 10, .25); border: 1px solid #F6E3CE; }
.stat .k { font-size: .7rem; color: #9A7B5C; font-weight: 600; text-transform: uppercase; letter-spacing: .05em; }
.stat .v { font-size: 1.5rem; font-weight: 800; color: #B34A0E; line-height: 1.2; margin-top: .15rem; }

/* CONTAINERS / CARDS */
[data-testid="stVerticalBlockBorderWrapper"] {
  background: #fff; border-radius: 20px !important; border: 1px solid #F6E3CE !important;
  box-shadow: 0 8px 26px -14px rgba(120, 60, 10, .3);
}
h3 { font-weight: 700 !important; color: #2B2118; }

/* UPLOADER */
[data-testid="stFileUploader"] section {
  background: #FFFAF3; border: 2px dashed #F0932B; border-radius: 18px; padding: 1.2rem;
}
[data-testid="stFileUploader"] section:hover { background: #FFF1DE; }

/* BUTTONS */
.stButton > button, .stDownloadButton > button {
  width: 100%; border-radius: 14px; font-weight: 700; padding: .65rem 1rem;
  border: 1px solid #EB6E24; transition: all .15s ease;
}
.stDownloadButton > button, .stButton > button[kind="primary"] {
  background: linear-gradient(135deg, #F0932B, #EB6E24); color: #fff; border: none;
  box-shadow: 0 8px 18px -8px rgba(235, 110, 36, .8);
}
.stButton > button:hover, .stDownloadButton > button:hover { transform: translateY(-1px); filter: brightness(1.05); }

/* TABS */
.stTabs [data-baseweb="tab-list"] { gap: .4rem; background: #fff; padding: .35rem; border-radius: 16px;
                                     border: 1px solid #F6E3CE; overflow-x: auto; }
.stTabs [data-baseweb="tab"] { border-radius: 12px; padding: .5rem 1rem; font-weight: 600; white-space: nowrap; }
.stTabs [aria-selected="true"] { background: #FFF1DE; color: #B34A0E; }
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none; }

/* IMAGE */
[data-testid="stImage"] img { border-radius: 14px; background: #fff; }

.footer-note { text-align: center; color: #9A7B5C; font-size: .78rem; margin-top: 2rem; }

/* MOBILE */
@media (max-width: 768px) {
  .block-container { padding: 1rem .8rem 3rem; }
  .hero { padding: 1.4rem 1.2rem; border-radius: 20px; }
  .stats { grid-template-columns: repeat(2, 1fr); }
  .stat .v { font-size: 1.25rem; }
  .stTabs [data-baseweb="tab"] { padding: .45rem .8rem; font-size: .85rem; }
}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------- helper ----
@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    names = wb.sheetnames
    wb.close()
    return names


@st.cache_data(show_spinner=False)
def get_sheet_data(file_bytes: bytes, sheet: str):
    return load_sheet(io.BytesIO(file_bytes), sheet)


def fig_to_png(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    return buf.getvalue()


def stat_cards(items):
    html = "".join(f'<div class="stat"><div class="k">{k}</div><div class="v">{v}</div></div>'
                   for k, v in items)
    st.markdown(f'<div class="stats">{html}</div>', unsafe_allow_html=True)


# ------------------------------------------------------------------ hero ----
if LOGO_B64:
    logo_html = f'<div class="logo"><img src="data:image/webp;base64,{LOGO_B64}" alt="Logo BPS"></div>'
else:
    logo_html = ('<div class="logo"><svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#EB6E24" '
                 'stroke-width="2.2" stroke-linecap="round"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg></div>')

st.markdown(f"""
<div class="hero">
  <div class="brand">
    {logo_html}
    <div class="name">{BRAND}<small>Badan Pusat Statistik</small></div>
  </div>
  <h1>Chart Generator</h1>
  <p>Ubah tabel Excel <b>Kecamatan Dalam Angka</b> menjadi grafik batang yang rapi dalam hitungan detik.</p>
</div>
""", unsafe_allow_html=True)

# ------------------------------------------------ panduan + upload ----
tutorial_slot = st.empty()  # tempat panduan, tampil di ATAS kotak upload
uploaded = st.file_uploader("Mulai di sini: unggah file Excel (.xlsx)", type=["xlsx", "xlsm"])

if not uploaded:
    steps = [
        ("Unggah", "Pilih file Excel Kecamatan Dalam Angka lewat kotak di bawah."),
        ("Atur", "Pilih sheet, kolom, satuan, dan orientasi grafik."),
        ("Unduh", "Simpan grafik sebagai PNG, satu per satu atau semua sekaligus."),
    ]
    rows = "".join(
        f'<div class="row"><div class="num">{i}</div>'
        f'<div><div class="t">{t}</div><div class="d">{d}</div></div></div>'
        for i, (t, d) in enumerate(steps, 1)
    )
    tutorial_slot.markdown(
        f'<div class="howto"><div class="head">Cara menggunakan</div>{rows}</div>',
        unsafe_allow_html=True,
    )
    st.stop()

file_bytes = uploaded.getvalue()
sheet_names = get_sheet_names(file_bytes)

# -------------------------------------------------------------- controls ----
with st.container(border=True):
    st.markdown("### :material/tune: Pengaturan")
    c1, c2, c3 = st.columns([2, 1, 1])
    sheet = c1.selectbox("Sheet", sheet_names)
    data = get_sheet_data(file_bytes, sheet)
    labels = data["series_labels"]
    unit = c2.text_input("Satuan", value=guess_unit_from_title(data["title"]),
                         key=f"unit_{sheet}", placeholder="km, ha, jiwa")
    orientation = c3.selectbox("Orientasi", ["auto", "horizontal", "vertical"])

    default_cols = [l for l in labels if classify_column(l) == "main"] or labels
    chosen = st.multiselect("Kolom yang ditampilkan", labels, default=default_cols,
                            format_func=lambda l: l + _TAG[classify_column(l)],
                            key=f"cols_{sheet}")

    o1, o2 = st.columns(2)
    include_totals = o1.toggle("Sertakan baris Jumlah/Total", value=False)
    no_label = o2.toggle("Tanpa keterangan (angka saja)", value=False)

    overrides = {}
    if len(chosen) > 1 and not no_label:
        with st.expander(":material/edit: Keterangan per kolom (teks di samping angka)"):
            for lbl, default in zip(chosen, short_series_labels(chosen)):
                overrides[lbl] = st.text_input(lbl, value=default, key=f"lbl_{sheet}_{lbl}")

# ----------------------------------------------------------------- stats ----
stat_cards([
    ("Total Sheet", len(sheet_names)),
    ("Baris Data", len(data["categories"])),
    ("Kolom Dipilih", f"{len(chosen)}/{len(labels)}"),
    ("Sheet Aktif", sheet),
])

# ------------------------------------------------------------------ tabs ----
tab_chart, tab_data, tab_all = st.tabs([":material/bar_chart: Grafik", ":material/table_chart: Data", ":material/folder_zip: Semua Sheet"])

with tab_chart:
    if not chosen:
        st.warning("Pilih minimal satu kolom dulu ya.")
    else:
        fig = plot_chart(data, exclude_totals=not include_totals, figure_number=sheet,
                         unit_hint=unit.strip(), columns=chosen, orientation=orientation,
                         series_label_overrides=overrides, show_series_label=not no_label)
        if fig is None:
            st.warning("Sheet ini tidak punya data untuk digambar.")
        else:
            png = fig_to_png(fig)
            plt.close(fig)
            with st.container(border=True):
                st.image(png)
            st.download_button(":material/download: Download PNG", png,
                               file_name=f"chart_{sheet.replace('.', '_')}.png",
                               mime="image/png")

with tab_data:
    st.markdown(f"**{data['title']}**")
    df = pd.DataFrame(data["series_values"], index=data["categories"])
    st.dataframe(df)

with tab_all:
    st.write("Tiap sheet memakai kolom **utama**-nya saja. Hasilnya dikemas dalam satu file ZIP.")
    if st.button(":material/play_arrow: Generate semua grafik", type="primary"):
        zbuf = io.BytesIO()
        progress = st.progress(0.0, text="Memproses...")
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, sh in enumerate(sheet_names):
                try:
                    d = get_sheet_data(file_bytes, sh)
                    main_cols = [l for l in d["series_labels"] if classify_column(l) == "main"]
                    f = plot_chart(d, exclude_totals=True, figure_number=sh,
                                   unit_hint=guess_unit_from_title(d["title"]),
                                   columns=main_cols or None)
                    if f is not None:
                        zf.writestr(f"chart_{sh.replace('.', '_')}.png", fig_to_png(f))
                        plt.close(f)
                except Exception as e:
                    st.warning(f"Sheet '{sh}' dilewati: {e}")
                progress.progress((i + 1) / len(sheet_names), text=f"Sheet {sh}")
        progress.empty()
        st.success("Selesai!")
        st.download_button(":material/download: Download ZIP", zbuf.getvalue(),
                           file_name="charts.zip", mime="application/zip")

st.markdown('<div class="footer-note">Chart Generator • BPS Lampung Utara</div>',
            unsafe_allow_html=True)
