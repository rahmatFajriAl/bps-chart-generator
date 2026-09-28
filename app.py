import io
import base64
import hashlib
import mimetypes
import zipfile
from datetime import datetime, timedelta, timezone
from html import escape
from pathlib import Path

import numpy as np
import openpyxl
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

from chart_core import (
    load_sheet, load_csv, plot_chart, classify_column, short_series_labels,
    guess_unit_from_title, build_table_workbook, _TAG, DEFAULT_SOURCE,
)
import auth
import db
import security
from security import md_escape
from narrative import build_narrative
from app_theme import inject_app_theme, render_topbar

BASE_DIR = Path(__file__).parent
LOGO_PATH = BASE_DIR / "logo-bps.png"
BRAND = "BPS Lampung Utara"
MAX_HISTORY = 6
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
MIME = {"xlsx": XLSX_MIME, "xlsm": XLSX_MIME, "csv": "text/csv"}

# ------------------------------------------------------------ ikon SVG ----
# Semua ikon dipakai sebagai SVG murni (bukan emoji) supaya tajam di layar apa pun
# dan warnanya bisa diatur lewat CSS/atribut, konsisten dengan logo fallback.
ICON_CHART = ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
             'stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round">'
             '<path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg>')
ICON_DB = ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
          'stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round">'
          '<ellipse cx="12" cy="5" rx="8" ry="3"/>'
          '<path d="M4 5v6c0 1.7 3.6 3 8 3s8-1.3 8-3V5"/>'
          '<path d="M4 11v6c0 1.7 3.6 3 8 3s8-1.3 8-3v-6"/></svg>')
ICON_LOCK = ('<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            'stroke-width="2.1" stroke-linecap="round" stroke-linejoin="round">'
            '<rect x="5" y="11" width="14" height="10" rx="2.5"/>'
            '<path d="M8 11V7a4 4 0 0 1 8 0v4"/></svg>')


def load_logo():
    """Return (PIL image, base64 string, mime) atau (None, None, None) kalau file logo tidak ada."""
    try:
        from PIL import Image
        img = Image.open(LOGO_PATH)
        b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode()
        mime = mimetypes.guess_type(LOGO_PATH.name)[0] or "image/png"
        return img, b64, mime
    except Exception:
        return None, None, None


LOGO_IMG, LOGO_B64, LOGO_MIME = load_logo()
FOOTER_HTML = f'<div class="footer-note">Chart Generator • {escape(BRAND)}<br>by AlFajri MagangHub Batch 1 | 2026</div>'


@st.cache_resource
def _init_db():
    db.init_db()
    return True


_init_db()
AUTHED = auth.is_authenticated()

# Sidebar tidak dipakai lagi (isinya sudah pindah ke halaman "Tabel Tersimpan"),
# jadi default-nya tertutup supaya tidak menutupi layar HP.
st.set_page_config(page_title=f"Chart Generator | {BRAND}",
                   page_icon=LOGO_IMG if LOGO_IMG else ":material/bar_chart:",
                   layout="wide", initial_sidebar_state="collapsed")

# ------------------------------------------------------------------ CSS ----
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800;900&display=swap');

:root {
  --brand-1: #F0932B; --brand-2: #EB6E24; --brand-3: #B34A0E; --brand-4: #7A3410;
  --ink: #2B2118; --muted: #8A7259; --line: #F1DFC4;
  --card: #FFFFFF; --bg-1: #FFF3E3; --bg-2: #FFFBF5;
  --shadow-1: 0 8px 26px -14px rgba(120, 60, 10, .32);
  --shadow-2: 0 16px 40px -18px rgba(120, 60, 10, .45);
  --ease: cubic-bezier(.22, 1, .36, 1);
}

@keyframes fadeSlideUp { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
@keyframes floatBlob   { 0%, 100% { transform: translate(0, 0) scale(1); } 50% { transform: translate(-14px, 16px) scale(1.08); } }
@keyframes pulseRing   { 0% { box-shadow: 0 0 0 0 rgba(235,110,36,.35); } 70% { box-shadow: 0 0 0 10px rgba(235,110,36,0); } 100% { box-shadow: 0 0 0 0 rgba(235,110,36,0); } }

@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after { animation-duration: .001ms !important; animation-iteration-count: 1 !important; transition-duration: .001ms !important; }
}

html, body, [class*="css"], .stApp { font-family: 'Plus Jakarta Sans', sans-serif; }
.stApp {
  background:
    radial-gradient(circle at 8% 4%, rgba(240,147,43,.10), transparent 40%),
    radial-gradient(circle at 96% 18%, rgba(179,74,14,.08), transparent 45%),
    linear-gradient(180deg, var(--bg-1) 0%, var(--bg-2) 320px);
}

#MainMenu, footer, [data-testid="stDecoration"], [data-testid="stToolbar"] { display: none !important; }
header[data-testid="stHeader"] {
  background: transparent !important; backdrop-filter: none !important;
  -webkit-backdrop-filter: none !important; box-shadow: none !important; z-index: 999990;
}

.block-container { max-width: 1100px; padding: 2rem 1.5rem 4rem; }

/* SIDEBAR */
[data-testid="stSidebar"] { background: var(--bg-2); border-right: 1px solid var(--line); }
[data-testid="stSidebar"] h3 { font-size: 1.05rem; }
[data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"] { transition: transform .18s var(--ease), box-shadow .18s var(--ease); }
[data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"]:hover { transform: translateX(2px); box-shadow: var(--shadow-2); }

/* Tombol buka/tutup sidebar BAWAAN Streamlit -- dipercantik, TIDAK disembunyikan.
   Ini yang membuat sidebar bisa dibuka & ditutup berkali-kali tanpa hilang. */
[data-testid="stSidebarCollapsedControl"] button,
[data-testid="collapsedControl"] button {
  background: #fff !important; border: 1px solid var(--brand-1) !important;
  border-radius: 12px !important; box-shadow: var(--shadow-1) !important;
  color: var(--brand-2) !important; transition: transform .18s var(--ease), box-shadow .18s var(--ease);
}
[data-testid="stSidebarCollapsedControl"] button:hover,
[data-testid="collapsedControl"] button:hover {
  transform: translateY(-1px) scale(1.06); box-shadow: var(--shadow-2) !important;
}
[data-testid="stSidebar"] [data-testid="stSidebarHeader"] button,
[data-testid="stSidebar"] [data-testid="stBaseButton-headerNoPadding"] {
  border-radius: 10px !important; color: var(--brand-2) !important;
}

/* HERO */
.hero {
  background: linear-gradient(135deg, var(--brand-1) 0%, var(--brand-2) 55%, var(--brand-3) 100%);
  background-size: 180% 180%;
  border-radius: 26px; padding: 2.1rem 2.2rem 1.9rem; color: #fff;
  box-shadow: var(--shadow-2);
  position: relative; overflow: hidden; margin-bottom: 1.4rem;
  animation: fadeSlideUp .55s var(--ease) both;
}
.hero::before, .hero::after {
  content: ""; position: absolute; border-radius: 50%; background: rgba(255,255,255,.14);
  filter: blur(2px); animation: floatBlob 9s ease-in-out infinite;
}
.hero::before { width: 220px; height: 220px; right: -60px; top: -70px; }
.hero::after  { width: 150px; height: 150px; right: 90px; bottom: -60px; background: rgba(255,255,255,.10); animation-delay: 2.2s; }
.hero .badge {
  display: inline-flex; align-items: center; gap: .4rem; background: rgba(255,255,255,.22);
  backdrop-filter: blur(8px); -webkit-backdrop-filter: blur(8px);
  padding: .3rem .85rem; border-radius: 999px; font-size: .72rem; font-weight: 700;
  letter-spacing: .06em; text-transform: uppercase; margin-bottom: .9rem;
  border: 1px solid rgba(255,255,255,.3); position: relative; z-index: 1;
}
.hero .hero-title {
  display: block; color: #fff; margin: 0; font-weight: 900; line-height: 1.12; letter-spacing: -.01em;
  font-size: clamp(1.55rem, 4.6vw, 2.5rem); padding: 0; position: relative; z-index: 1;
  text-shadow: 0 2px 18px rgba(0,0,0,.12);
}
.hero p { margin: .7rem 0 0; opacity: .95; font-size: clamp(.86rem, 2.3vw, 1.02rem); max-width: 640px;
          position: relative; z-index: 1; line-height: 1.55; }

/* BRAND */
.brand { display: flex; align-items: center; gap: .85rem; margin-bottom: 1.1rem; position: relative; z-index: 1; }
.brand .logo {
  background: #fff; border-radius: 16px; padding: .4rem .55rem; display: flex; align-items: center;
  box-shadow: 0 10px 22px -10px rgba(0,0,0,.4); transition: transform .25s var(--ease);
}
.brand .logo:hover { transform: rotate(-4deg) scale(1.04); }
.brand .logo img { height: 44px; width: auto; display: block; }
.brand .name { font-weight: 800; font-size: clamp(.85rem, 2.6vw, 1.08rem); letter-spacing: .01em; line-height: 1.2; }
.brand .name small { display: block; font-weight: 600; opacity: .88; font-size: .7rem; letter-spacing: .08em; text-transform: uppercase; margin-top: 1px; }

/* LANDING FEATURES */
.features { display: grid; grid-template-columns: repeat(3, 1fr); gap: .9rem; margin: 0 0 1.2rem; }
.feature {
  background: var(--card); border-radius: 20px; padding: 1.2rem 1.3rem; border: 1px solid var(--line);
  box-shadow: var(--shadow-1); transition: transform .2s var(--ease), box-shadow .2s var(--ease);
  animation: fadeSlideUp .5s var(--ease) both;
}
.feature:hover { transform: translateY(-4px); box-shadow: var(--shadow-2); }
.feature .icon {
  width: 38px; height: 38px; border-radius: 12px; display: flex; align-items: center; justify-content: center;
  background: linear-gradient(135deg, var(--brand-1), var(--brand-2)); color: #fff;
  margin-bottom: .65rem; box-shadow: 0 6px 14px -6px rgba(235,110,36,.55);
}
.feature .t { font-weight: 800; color: var(--brand-3); margin-bottom: .3rem; font-size: .98rem; }
.feature .d { color: var(--muted); font-size: .87rem; line-height: 1.5; }

/* HOW-TO */
.howto {
  background: var(--card); border: 1px solid var(--line); border-radius: 22px; padding: 1.2rem 1.35rem 1.25rem;
  box-shadow: var(--shadow-1); margin-bottom: 1rem; cursor: default;
  animation: fadeSlideUp .55s var(--ease) .05s both;
}
.howto .head { font-size: .72rem; font-weight: 800; letter-spacing: .1em; text-transform: uppercase;
               color: var(--muted); margin-bottom: 1rem; }
.howto .row { display: flex; gap: .95rem; align-items: flex-start; position: relative; padding-bottom: 1.15rem;
              opacity: 0; animation: fadeSlideUp .45s var(--ease) forwards; }
.howto .row:nth-child(2) { animation-delay: .08s; }
.howto .row:nth-child(3) { animation-delay: .16s; }
.howto .row:nth-child(4) { animation-delay: .24s; }
.howto .row:last-child { padding-bottom: 0; }
.howto .row:not(:last-child)::before {
  content: ""; position: absolute; left: 16px; top: 36px; bottom: 2px; width: 2px;
  background: linear-gradient(180deg, var(--brand-1), transparent);
}
.howto .num {
  flex: 0 0 auto; width: 34px; height: 34px; border-radius: 50%; color: #fff; font-weight: 800; font-size: .9rem;
  display: flex; align-items: center; justify-content: center;
  background: linear-gradient(135deg, var(--brand-1), var(--brand-2));
  box-shadow: 0 6px 14px -6px rgba(235,110,36,.6);
}
.howto .t { font-weight: 800; color: var(--ink); font-size: 1.02rem; line-height: 34px; }
.howto .d { color: var(--muted); font-size: .86rem; margin-top: -.2rem; line-height: 1.5; }

/* STAT CARDS */
.stats { display: grid; grid-template-columns: repeat(4, 1fr); gap: .85rem; margin: 0 0 1.2rem; }
.stat {
  background: var(--card); border-radius: 18px; padding: .95rem 1.05rem;
  box-shadow: var(--shadow-1); border: 1px solid var(--line);
  transition: transform .18s var(--ease), box-shadow .18s var(--ease);
  opacity: 0; animation: fadeSlideUp .4s var(--ease) forwards;
}
.stat:nth-child(1) { animation-delay: 0s; } .stat:nth-child(2) { animation-delay: .05s; }
.stat:nth-child(3) { animation-delay: .1s; } .stat:nth-child(4) { animation-delay: .15s; }
.stat:hover { transform: translateY(-3px); box-shadow: var(--shadow-2); }
.stat .k { font-size: .7rem; color: var(--muted); font-weight: 700; text-transform: uppercase; letter-spacing: .06em; }
.stat .v { font-size: 1.5rem; font-weight: 800; color: var(--brand-3); line-height: 1.2; margin-top: .18rem; overflow-wrap: anywhere; }

/* CONTAINERS / CARDS */
[data-testid="stVerticalBlockBorderWrapper"] {
  background: var(--card); border-radius: 22px !important; border: 1px solid var(--line) !important;
  box-shadow: var(--shadow-1); transition: box-shadow .2s var(--ease);
}
h3 { font-weight: 800 !important; color: var(--ink); letter-spacing: -.01em; }

/* UPLOADER */
[data-testid="stFileUploader"] section {
  background: #FFFAF3; border: 2px dashed var(--brand-1); border-radius: 18px; padding: 1.3rem;
  transition: background .2s var(--ease), border-color .2s var(--ease), transform .2s var(--ease);
}
[data-testid="stFileUploader"] section:hover { background: #FFF1DE; border-color: var(--brand-2); transform: scale(1.003); }

/* BUTTONS */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
  width: 100%; border-radius: 14px; font-weight: 700; padding: .68rem 1rem;
  border: 1px solid var(--brand-2); transition: transform .16s var(--ease), box-shadow .16s var(--ease), filter .16s var(--ease);
}
.stDownloadButton > button, .stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
  background: linear-gradient(135deg, var(--brand-1), var(--brand-2) 55%, var(--brand-3));
  background-size: 200% 200%; color: #fff; border: none;
  box-shadow: 0 10px 20px -9px rgba(235, 110, 36, .8);
}
.stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
  transform: translateY(-2px); filter: brightness(1.06); box-shadow: var(--shadow-2);
}
.stButton > button:active, .stDownloadButton > button:active { transform: translateY(0) scale(.98); }

/* ZOOM LINK */
.zoom-link {
  display: block; width: 100%; box-sizing: border-box; text-align: center; border-radius: 14px;
  font-weight: 700; padding: .68rem 1rem; border: 1px solid var(--brand-2); background: #fff; color: var(--brand-3);
  text-decoration: none; transition: all .18s var(--ease);
}
.zoom-link:hover { background: #FFF1DE; transform: translateY(-2px); box-shadow: var(--shadow-1); }

/* TABS */
.stTabs [data-baseweb="tab-list"] { gap: .4rem; background: var(--card); padding: .4rem; border-radius: 18px;
                                     border: 1px solid var(--line); overflow-x: auto; box-shadow: var(--shadow-1); }
.stTabs [data-baseweb="tab"] { border-radius: 13px; padding: .55rem 1.05rem; font-weight: 650; white-space: nowrap;
                               transition: background .18s var(--ease), color .18s var(--ease); }
.stTabs [aria-selected="true"] { background: linear-gradient(135deg, #FFF1DE, #FFE3C2); color: var(--brand-3); font-weight: 800; }
.stTabs [data-baseweb="tab-highlight"], .stTabs [data-baseweb="tab-border"] { display: none; }

/* tombol layar penuh bawaan st.image: dibuat kontras supaya terlihat */
[data-testid="StyledFullScreenButton"], [data-testid="stImageFullScreenButton"] {
  background: var(--brand-2) !important; color: #fff !important; border-radius: 10px !important;
  opacity: 1 !important; box-shadow: var(--shadow-1) !important;
}

/* IMAGE PREVIEW */
.chart-frame { overflow: auto; text-align: center; -webkit-overflow-scrolling: touch; touch-action: pan-x pan-y pinch-zoom;
               border-radius: 14px; }
.chart-frame img { max-width: none; border-radius: 14px; background: #fff; transition: width .25s var(--ease); }

.footer-note { text-align: center; color: var(--muted); font-size: .78rem; margin-top: 2.2rem; letter-spacing: .02em; line-height: 1.6; }

/* MOBILE */
@media (max-width: 768px) {
  .block-container { padding: 1rem .8rem 3rem; }
  .hero { padding: 1.5rem 1.25rem; border-radius: 22px; }
  .stats { grid-template-columns: repeat(2, 1fr); }
  .features { grid-template-columns: 1fr; }
  .stat .v { font-size: 1.25rem; }
  .stTabs [data-baseweb="tab"] { padding: .5rem .85rem; font-size: .85rem; }
}

/* ===== ATURAN PAKSA, ditaruh PALING BAWAH supaya menang dari rule lain ===== */
/* Tombol buka sidebar WAJIB selalu terlihat & bisa disentuh, apa pun rule di atas */
[data-testid="stSidebarCollapsedControl"] {
  display: flex !important; visibility: visible !important; opacity: 1 !important;
  pointer-events: auto !important; z-index: 1000000 !important;
}
[data-testid="stSidebarCollapsedControl"] button,
[data-testid="collapsedControl"],
[data-testid="collapsedControl"] button {
  display: flex !important; visibility: visible !important; opacity: 1 !important;
  pointer-events: auto !important; min-width: 42px !important; min-height: 42px !important;
}
</style>
""", unsafe_allow_html=True)

if not AUTHED:  # sembunyikan sidebar total di halaman utama / login
    st.markdown("<style>[data-testid='stSidebar'], [data-testid='stSidebarCollapsedControl'], "
                "[data-testid='collapsedControl'] { display: none !important; }</style>",
                unsafe_allow_html=True)

if LOGO_B64:
    logo_html = (f'<div class="logo"><img src="data:{escape(LOGO_MIME)};base64,{LOGO_B64}" alt="Logo BPS"></div>')
else:
    logo_html = ('<div class="logo"><svg width="44" height="44" viewBox="0 0 24 24" fill="none" stroke="#EB6E24" '
                 'stroke-width="2.2" stroke-linecap="round"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg></div>')


def render_hero(desc):
    st.markdown(f"""
<div class="hero">
  <div class="brand">
    {logo_html}
    <div class="name">{escape(BRAND)}<small>Badan Pusat Statistik</small></div>
  </div>
  <span class="badge">Chart Generator</span>
  <div class="hero-title" role="heading" aria-level="1">Selamat Datang di Chart Generator</div>
  <p>{desc}</p>
</div>
""", unsafe_allow_html=True)


# ================================================== HALAMAN UTAMA + LOGIN ====
def render_landing():
    render_hero("Ubah tabel Excel/CSV <b>Kecamatan Dalam Angka</b> menjadi grafik batang yang rapi "
                "dalam hitungan detik.")

    if st.session_state.get("auth_msg"):
        st.info(st.session_state.pop("auth_msg"))

    if not auth.is_configured():
        st.error("Akun admin belum dikonfigurasi. Jalankan `python make_hash.py`, lalu isi "
                 "`.streamlit/secrets.toml` (lihat `secrets.toml.example`).")
        st.stop()

    if st.session_state.get("view", "home") == "home":
        st.markdown(f"""
<div class="features">
  <div class="feature" style="animation-delay:.05s">
    <div class="icon">{ICON_CHART}</div>
    <div class="t">Grafik otomatis</div>
    <div class="d">Unggah Excel/CSV, pilih kolom dan satuan, grafik langsung jadi.</div></div>
  <div class="feature" style="animation-delay:.12s">
    <div class="icon">{ICON_DB}</div>
    <div class="t">Tabel tersimpan</div>
    <div class="d">Tabel yang dibuat disimpan di database, tinggal buka, edit, atau unduh kapan saja.</div></div>
  <div class="feature" style="animation-delay:.19s">
    <div class="icon">{ICON_LOCK}</div>
    <div class="t">Akses terkunci</div>
    <div class="d">Hanya petugas berwenang yang bisa masuk dan mengelola data.</div></div>
</div>
""", unsafe_allow_html=True)
        _, mid, _ = st.columns([1, 1, 1])
        if mid.button(":material/login: Masuk", type="primary", use_container_width=True, key="go_login"):
            st.session_state["view"] = "login"
            st.rerun()
    else:
        _, mid, _ = st.columns([1, 2, 1])
        with mid, st.container(border=True):
            st.markdown("### :material/lock: Masuk")
            with st.form("login_form"):
                username = st.text_input("Username", max_chars=64)
                password = st.text_input("Password", type="password", max_chars=128)
                submitted = st.form_submit_button("Masuk", type="primary", use_container_width=True)
            if submitted:
                ok, msg = auth.login(username, password)
                if ok:
                    st.rerun()
                else:
                    st.error(msg)
            if st.button("Kembali ke halaman utama", use_container_width=True, key="back_home"):
                st.session_state["view"] = "home"
                st.rerun()
    st.markdown(FOOTER_HTML, unsafe_allow_html=True)
    st.stop()


if not AUTHED:
    render_landing()

inject_app_theme()  # tema baru khusus tampilan setelah login

# ------------------------------------------------------------ helper ----
WIB = timezone(timedelta(hours=7))


def fmt_size(n):
    return f"{n / 1024:.0f} KB" if n < 1024 * 1024 else f"{n / 1048576:.1f} MB"


def fmt_dt(iso):
    try:
        return datetime.fromisoformat(iso).astimezone(WIB).strftime("%d %b %Y %H:%M")
    except Exception:
        return ""


def flash(msg):
    st.session_state["flash"] = msg


@st.cache_data(show_spinner=False)
def get_sheet_names(file_bytes: bytes, ext: str):
    if ext == "csv":
        return ["Data (CSV)"]
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True)
    names = [n for n in wb.sheetnames if n != "_meta"]
    wb.close()
    return names


@st.cache_data(show_spinner=False)
def get_sheet_data(file_bytes: bytes, sheet: str, ext: str, filename: str = ""):
    if ext == "csv":
        title = Path(filename).stem.replace("_", " ").replace("-", " ").strip() or "Data CSV"
        return load_csv(file_bytes, title=title)
    return load_sheet(io.BytesIO(file_bytes), sheet)


def fig_to_png(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight")
    return buf.getvalue()


# Chrome/Safari memblokir link data:image di tab baru (halaman jadi kosong), jadi
# "gambar penuh" ditampilkan lewat dialog Streamlit. st.image punya tombol layar penuh bawaan.
if hasattr(st, "dialog"):
    @st.dialog("Grafik ukuran penuh", width="large")
    def show_full_chart(png_bytes):
        st.caption("Gambar ditampilkan selebar jendela. Untuk file resolusi penuh, pakai tombol Download PNG.")
        st.image(png_bytes)
else:
    def show_full_chart(png_bytes):
        st.image(png_bytes)


# Dialog "isi nama file lalu download". Tombol download ada DI DALAM dialog, jadi
# nama yang diketik langsung dipakai. Nama dibersihkan (security.safe_filename).
def _download_chart_body(png_bytes, default_name):
    name = st.text_input("Nama file", value=default_name, max_chars=80, key="png_name_input")
    fname = security.safe_filename(name.strip() or default_name, "png")
    st.caption(f"Akan tersimpan sebagai: **{md_escape(fname)}**")
    st.download_button(":material/download: Download", png_bytes, file_name=fname, mime="image/png",
                       type="primary", use_container_width=True, key="png_dl_final")


if hasattr(st, "dialog"):
    @st.dialog("Simpan grafik sebagai PNG")
    def download_chart_dialog(png_bytes, default_name):
        _download_chart_body(png_bytes, default_name)
else:
    def download_chart_dialog(png_bytes, default_name):
        _download_chart_body(png_bytes, default_name)


def stat_cards(items):
    # nilai (mis. nama file) berasal dari pengguna -> WAJIB di-escape (anti XSS)
    html = "".join(f'<div class="stat"><div class="k">{escape(str(k))}</div><div class="v">{escape(str(v))}</div></div>'
                   for k, v in items)
    st.markdown(f'<div class="stats">{html}</div>', unsafe_allow_html=True)


def push_history(entry):
    hist = [h for h in st.session_state.history
            if h["name"] != entry["name"] and not (entry.get("db_id") and h.get("db_id") == entry["db_id"])]
    hist.insert(0, entry)
    st.session_state.history = hist[:MAX_HISTORY]
    st.session_state.active_name = entry["name"]


def open_record(rid):
    rec = db.get_table(rid)
    if not rec:
        flash("Tabel tidak ditemukan.")
        return
    push_history({"name": security.safe_filename(rec["name"], rec["ext"]), "bytes": rec["content"],
                  "ext": rec["ext"], "db_id": rec["id"]})


def save_new(name, ext, content):
    """Return (id, error)."""
    name = security.safe_name(name, db.MAX_NAME_LEN)
    if not name:
        return None, "Nama tidak boleh kosong."
    if db.count_tables() >= db.MAX_SAVED:
        return None, f"Database penuh (maks {db.MAX_SAVED} tabel). Hapus sebagian dulu."
    err = security.validate_upload(f"x.{ext}", content)
    if err:
        return None, err
    return db.create_table(name, ext, content), None


def update_existing(rid, name, content):
    name = security.safe_name(name, db.MAX_NAME_LEN)
    if not name:
        return "Nama tidak boleh kosong."
    err = security.validate_upload("x.xlsx", content)
    if err:
        return err
    db.update_table(rid, name, "xlsx", content)
    return None


def start_edit(rid):
    rec = db.get_table(rid)
    if not rec:
        flash("Tabel tidak ditemukan.")
        return
    try:
        if rec["ext"] == "csv":
            d = load_csv(rec["content"], title=rec["name"])
            sheet0, n_sheets = "", 1
        else:
            names = get_sheet_names(rec["content"], rec["ext"])
            sheet0, n_sheets = names[0], len(names)
            d = load_sheet(io.BytesIO(rec["content"]), sheet0)
    except Exception as e:
        flash(f"Tabel tidak bisa dibuka untuk diedit: {e}")
        return
    st.session_state.builder_meta = {
        "judul": d["title"], "no_tabel": sheet0, "cat_label": d["category_label"] or "Wilayah",
        "sumber": d.get("source") or "",
    }
    st.session_state.builder_df = pd.DataFrame(d["series_values"], index=d["categories"])
    st.session_state.builder_ver += 1
    st.session_state.editing_id = rid
    st.session_state["_pending_dbname"] = rec["name"]
    st.session_state.edit_warning = (n_sheets > 1)


def reset_builder():
    st.session_state.builder_df = None
    st.session_state.editing_id = None
    st.session_state.edit_warning = False
    st.session_state.builder_ver += 1


def apply_preview_edit(active, data, sheet, edited_df, new_labels, new_title, save_db=False):
    """Bangun ulang workbook dari hasil edit, cek hasil bacanya cocok, lalu pasang ke file aktif.
    Return pesan error (str) atau None kalau sukses."""
    old_labels = data["series_labels"]
    n_rows = len(data["categories"])
    if len(edited_df) != n_rows:
        return "Jumlah baris berubah, edit dibatalkan."

    labels_clean = [security.safe_name(n, 100) or o for n, o in zip(new_labels, old_labels)]
    if len({l.lower() for l in labels_clean}) != len(labels_clean):
        return "Ada nama kolom yang kembar. Beri nama berbeda untuk tiap kolom."

    def _txt(v):
        return "" if (v is None or (isinstance(v, float) and np.isnan(v))) else v
    row_names = [security.safe_name(_txt(n), 100) or o
                 for n, o in zip(edited_df["c_row"].tolist(), data["categories"])]

    cols_vals = [pd.to_numeric(edited_df[f"c{j}"], errors="coerce").to_numpy(dtype=float)
                 for j in range(len(old_labels))]
    values = [[cols_vals[j][i] for j in range(len(old_labels))] for i in range(n_rows)]
    title = security.safe_name(new_title, 200) or data["title"]

    new_bytes = build_table_workbook(title, sheet, data.get("source") or "",
                                     data.get("category_label") or "Wilayah",
                                     row_names, labels_clean, values)
    err = security.validate_upload("x.xlsx", new_bytes)
    if err:
        return err

    # pengaman: baca ulang lewat pembaca yang sama, pastikan angkanya tidak berubah diam-diam
    try:
        chk = load_sheet(io.BytesIO(new_bytes), get_sheet_names(new_bytes, "xlsx")[0])
        if chk["series_labels"] != labels_clean or len(chk["categories"]) != n_rows:
            return "Hasil edit tidak terbaca konsisten (jumlah baris/kolom berubah). Edit dibatalkan."
        for j, lbl in enumerate(labels_clean):
            a = np.array(chk["series_values"][lbl], dtype=float)
            if not np.allclose(a, cols_vals[j], equal_nan=True):
                return f"Angka di kolom '{lbl}' berubah saat dibaca ulang. Edit dibatalkan supaya data tidak rusak."
    except Exception as e:
        return f"Gagal memeriksa hasil edit: {e}"

    if active["ext"] == "csv":
        active["name"] = security.safe_filename(Path(active["name"]).stem, "xlsx")
    active["bytes"], active["ext"] = new_bytes, "xlsx"
    st.session_state.active_name = active["name"]
    if save_db and active.get("db_id"):
        rec = db.get_table(active["db_id"])
        if rec:
            db.update_table(active["db_id"], rec["name"], "xlsx", new_bytes)
    return None


# --------------------------------------------------------- session init ----
for k, v in {"history": [], "active_name": None, "builder_df": None, "builder_meta": {},
             "builder_ver": 0, "editing_id": None, "edit_warning": False, "confirm_delete": None,
             "last_upload_sig": None, "last_reupload_sig": None, "view_page": "main",
             "preview_edit_target": None, "prev_edit_ver": 0}.items():
    st.session_state.setdefault(k, v)

if "_pending_dbname" in st.session_state:   # widget belum dibuat di run ini -> aman diisi
    st.session_state["builder_dbname"] = st.session_state.pop("_pending_dbname")

if "flash" in st.session_state:
    st.toast(st.session_state.pop("flash"))

active = next((h for h in st.session_state.history if h["name"] == st.session_state.active_name), None)

# ==================================================== HALAMAN: AKUN ====
LOG_LABEL = {
    "success": "Berhasil masuk",
    "fail": "Gagal (username/password salah)",
    "locked": "Ditolak (akun terkunci sementara)",
    "pw_changed": "Kata sandi diubah",
    "pw_fail": "Gagal ganti kata sandi",
}

if st.session_state.view_page == "account":
    render_topbar(BRAND, logo_html, "Akun", "Ganti kata sandi dan lihat riwayat login", auth.current_user())

    ac1, ac2 = st.columns([1, 3])
    if ac1.button(":material/arrow_back: Kembali", use_container_width=True, key="back_from_account"):
        st.session_state.view_page = "main"
        st.rerun()
    ac2.caption(f"Masuk sebagai **{md_escape(auth.current_user())}**")

    with st.container(border=True):
        st.markdown("### :material/key: Ganti kata sandi")
        with st.form("change_pw_form", clear_on_submit=True):
            pw_cur = st.text_input("Kata sandi saat ini", type="password", max_chars=128)
            pw_new = st.text_input(f"Kata sandi baru (minimal {auth.MIN_PW_LEN} karakter)", type="password",
                                   max_chars=128)
            pw_new2 = st.text_input("Ulangi kata sandi baru", type="password", max_chars=128)
            pw_submit = st.form_submit_button(":material/save: Simpan kata sandi", type="primary",
                                              use_container_width=True)
        if pw_submit:
            ok_pw, msg_pw = auth.change_password(pw_cur, pw_new, pw_new2)
            (st.success if ok_pw else st.error)(msg_pw)
        st.caption("Kata sandi baru disimpan di database dan menggantikan yang ada di Secrets. Di Streamlit Cloud, "
                   "jika database ter-reset saat aplikasi restart, kata sandi kembali ke yang ada di Secrets.")

    with st.container(border=True):
        st.markdown("### :material/history: Riwayat login")
        st.caption("30 catatan terakhir. Yang dicatat hanya waktu dan hasilnya, bukan teks yang diketik.")
        log_rows = db.list_login_log(30)
        if not log_rows:
            st.caption("Belum ada catatan.")
        else:
            st.dataframe(
                pd.DataFrame({"Waktu (WIB)": [fmt_dt(r["ts"]) for r in log_rows],
                              "Hasil": [LOG_LABEL.get(r["status"], r["status"]) for r in log_rows]}),
                hide_index=True, use_container_width=True)

    st.markdown(FOOTER_HTML, unsafe_allow_html=True)
    st.stop()

# ============================================== HALAMAN: TABEL TERSIMPAN ====
# Halaman terpisah (bukan expander) -- dibuka lewat tombol di halaman utama,
# dan kembali ke halaman utama lewat tombol "Kembali". Perpindahannya lewat
# session_state, bukan folder pages/ Streamlit, supaya alur login yang sudah
# ada tidak perlu diubah.
if st.session_state.view_page == "tables":
    render_topbar(BRAND, logo_html, "Tabel Tersimpan", "Buka, edit, unduh, atau hapus tabel", auth.current_user())

    top1, top2 = st.columns([1, 3])
    if top1.button(":material/arrow_back: Kembali", use_container_width=True, key="back_to_main"):
        st.session_state.view_page = "main"
        st.rerun()
    top2.caption(f"Masuk sebagai **{md_escape(auth.current_user())}**")
    if st.button(":material/logout: Keluar", key="logout_btn"):
        auth.logout()
        st.rerun()
    st.divider()

    q = st.text_input("Cari tabel", placeholder="ketik nama...", key="side_search", max_chars=60)
    records = db.list_tables()
    if q.strip():
        records = [r for r in records if q.strip().lower() in r["name"].lower()]
    if not records:
        st.caption("Belum ada tabel tersimpan." if not q.strip() else "Tidak ada yang cocok.")

    active_db_id = active.get("db_id") if active else None
    for r in records:
        rid = r["id"]
        with st.container(border=True):
            st.markdown(f"**{md_escape(r['name'])}**")
            st.caption(f"{r['ext'].upper()} · {fmt_size(r['size'])} · {fmt_dt(r['updated_at'])}")
            if st.session_state.confirm_delete == rid:
                st.warning("Hapus permanen?")
                y, n = st.columns(2)
                if y.button("Ya", key=f"delyes_{rid}", type="primary", use_container_width=True):
                    db.delete_table(rid)
                    st.session_state.history = [h for h in st.session_state.history if h.get("db_id") != rid]
                    if active_db_id == rid:
                        st.session_state.active_name = None
                    if st.session_state.editing_id == rid:
                        reset_builder()
                    st.session_state.confirm_delete = None
                    flash("Tabel dihapus.")
                    st.rerun()
                if n.button("Batal", key=f"delno_{rid}", use_container_width=True):
                    st.session_state.confirm_delete = None
                    st.rerun()
            else:
                b1, b2, b3 = st.columns(3)
                if b1.button(":material/folder_open: Buka", key=f"open_{rid}", use_container_width=True,
                             type="primary" if active_db_id == rid else "secondary"):
                    open_record(rid)
                    st.session_state.view_page = "main"
                    st.rerun()
                if b2.button(":material/edit: Edit", key=f"edit_{rid}", use_container_width=True):
                    start_edit(rid)
                    st.session_state.view_page = "main"
                    st.rerun()
                if b3.button(":material/delete: Hapus", key=f"del_{rid}", use_container_width=True):
                    st.session_state.confirm_delete = rid
                    st.rerun()
                if active_db_id == rid:  # blob hanya dimuat untuk tabel yang sedang dibuka
                    rec = db.get_table(rid)
                    if rec:
                        st.download_button(":material/download: Download", rec["content"],
                                           file_name=security.safe_filename(rec["name"], rec["ext"]),
                                           mime=MIME.get(rec["ext"], "application/octet-stream"),
                                           key=f"dl_{rid}", use_container_width=True)

    st.markdown(FOOTER_HTML, unsafe_allow_html=True)
    st.stop()

# ------------------------------------------------------------------ hero ----
render_topbar(BRAND, logo_html, "Chart Generator",
              "Ubah tabel <b>Kecamatan Dalam Angka</b> jadi grafik batang", auth.current_user())

# --------------------------------------------- tombol ke halaman tabel ----
saved_count = db.count_tables()
nav1, nav2 = st.columns([4, 1])
if nav1.button(f":material/database: Tabel Tersimpan ({saved_count})",
               use_container_width=True, key="go_tables"):
    st.session_state.view_page = "tables"
    st.rerun()
if nav2.button(":material/manage_accounts: Akun", use_container_width=True, key="go_account"):
    st.session_state.view_page = "account"
    st.rerun()

tutorial_slot = st.empty()  # panduan tampil di ATAS kotak upload, hilang setelah ada file aktif
uploaded = st.file_uploader("Mulai di sini: unggah file Excel atau CSV", type=["xlsx", "xlsm", "csv"])

if uploaded is not None:
    fb = uploaded.getvalue()
    sig = hashlib.sha256(fb + uploaded.name.encode()).hexdigest()
    if sig != st.session_state.last_upload_sig:   # proses SEKALI per file (bukan tiap rerun)
        err = security.validate_upload(uploaded.name, fb)
        if err:
            st.error(err)
        else:
            st.session_state.last_upload_sig = sig
            push_history({"name": security.safe_name(uploaded.name, 100),
                          "bytes": fb, "ext": Path(uploaded.name).suffix.lower().lstrip("."), "db_id": None})

# ------------------------------------------------- buat / edit tabel ----
editing = st.session_state.editing_id is not None
with st.expander(":material/add_box: Atau, buat tabel baru dari nol (tanpa file Excel)",
                 expanded=st.session_state.builder_df is not None):
    st.caption("Cocok kalau struktur Excel sumbernya sering berubah-ubah. Isi metadata, "
               "aplikasi bikin tabel kosongnya, tinggal isi angkanya.")

    if st.session_state.builder_df is None:
        # ---- langkah 1: metadata tabel ----
        with st.form("builder_meta_form"):
            b1, b2 = st.columns(2)
            judul = b1.text_input("Judul Tabel", placeholder="mis. Penduduk Menurut Jenis Kelamin", max_chars=200)
            no_tabel = b2.text_input("No Tabel", placeholder="mis. 3.1", max_chars=31,
                                     help="Dipakai sebagai nama sheet & nomor gambar, seperti file BPS asli.")
            cat_label = st.text_input("Label Kolom Pertama", value="Wilayah", max_chars=60,
                                      help="Nama kolom kategori/baris, mis. 'Wilayah' atau 'Desa'.")
            sumber = st.text_input("Sumber Data", placeholder="mis. BPS, Pendataan Potensi Desa (Podes) 2025",
                                   max_chars=200)
            r1, r2 = st.columns(2)
            row_text = r1.text_area("Nama Baris (satu per baris)", height=140, max_chars=5000,
                                    placeholder="Desa A\nDesa B\nDesa C")
            col_text = r2.text_area("Nama Kolom (satu per baris)", height=140, max_chars=3000,
                                    placeholder="Penduduk Laki-laki\nPenduduk Perempuan")
            submitted = st.form_submit_button(":material/table: Bikin Tabel Kosong", type="primary")

        if submitted:
            row_names = [security.safe_name(r, 100) for r in row_text.splitlines() if r.strip()]
            col_names = [security.safe_name(c, 100) for c in col_text.splitlines() if c.strip()]
            if not judul.strip():
                st.warning("Judul tabel wajib diisi.")
            elif not row_names or not col_names:
                st.warning("Isi minimal satu Nama Baris dan satu Nama Kolom.")
            elif len(row_names) > 300 or len(col_names) > 30:
                st.warning("Maksimal 300 baris dan 30 kolom.")
            else:
                st.session_state.builder_meta = {
                    "judul": security.safe_name(judul, 200), "no_tabel": security.safe_name(no_tabel, 31),
                    "cat_label": security.safe_name(cat_label, 60) or "Wilayah",
                    "sumber": security.safe_name(sumber, 200),
                }
                st.session_state.builder_df = pd.DataFrame(np.nan, index=row_names, columns=col_names)
                st.session_state.builder_ver += 1
                st.session_state["_pending_dbname"] = security.safe_name(
                    f"{no_tabel} {judul}".strip(), db.MAX_NAME_LEN)
                st.rerun()
    else:
        # ---- langkah 2: isi angka ----
        meta = st.session_state.builder_meta
        if editing:
            st.info(f"Mode edit: tabel tersimpan **{md_escape(st.session_state.get('builder_dbname', ''))}**")
            if st.session_state.edit_warning:
                st.warning("File ini punya lebih dari satu sheet. Hanya sheet pertama yang diedit; "
                           "kalau kamu klik *Perbarui*, sheet lainnya tidak ikut tersimpan. "
                           "Pakai *Simpan sebagai baru* kalau ingin file aslinya tetap utuh.")
        st.markdown(f"**{md_escape(meta['judul'])}**  ·  No. {md_escape(meta['no_tabel'] or '-')}")

        blank_xlsx = build_table_workbook(
            meta["judul"], meta["no_tabel"], meta["sumber"], meta["cat_label"],
            list(st.session_state.builder_df.index), list(st.session_state.builder_df.columns),
            st.session_state.builder_df.values.tolist(),
        )
        t1, t2 = st.columns([3, 1])
        t1.download_button(":material/download: Download Template Kosong (isi di Excel)",
                           blank_xlsx, file_name=security.safe_filename(f"template_{meta['no_tabel'] or 'baru'}", "xlsx"),
                           mime=XLSX_MIME, use_container_width=True)
        if t2.button(":material/close: Batal" if editing else ":material/restart_alt: Buat Ulang",
                     use_container_width=True):
            reset_builder()
            st.rerun()

        filled = st.file_uploader("Atau, unggah kembali template yang sudah diisi", type=["xlsx"],
                                  key=f"builder_reupload_{st.session_state.builder_ver}")
        if filled is not None:
            fb = filled.getvalue()
            sig = hashlib.sha256(fb).hexdigest()
            if sig != st.session_state.last_reupload_sig:
                err = security.validate_upload(filled.name, fb)
                if err:
                    st.error(err)
                else:
                    try:
                        wb_tmp = openpyxl.load_workbook(io.BytesIO(fb), read_only=True)
                        sheet0 = next(n for n in wb_tmp.sheetnames if n != "_meta")
                        wb_tmp.close()
                        d = load_sheet(io.BytesIO(fb), sheet0)
                        st.session_state.last_reupload_sig = sig
                        st.session_state.builder_meta = {
                            "judul": d["title"], "no_tabel": sheet0,
                            "cat_label": d["category_label"] or "Wilayah", "sumber": d.get("source") or "",
                        }
                        st.session_state.builder_df = pd.DataFrame(d["series_values"], index=d["categories"])
                        st.session_state.builder_ver += 1
                        flash("Template berhasil dimuat ulang.")
                        st.rerun()
                    except Exception as e:
                        st.error(f"Gagal membaca template: {e}")

        st.write("Atau isi langsung di sini:")
        edited = st.data_editor(st.session_state.builder_df, use_container_width=True, num_rows="fixed",
                                key=f"builder_editor_{st.session_state.builder_ver}")
        st.session_state.builder_df = edited

        final_xlsx = build_table_workbook(
            meta["judul"], meta["no_tabel"], meta["sumber"], meta["cat_label"],
            list(edited.index), list(edited.columns), edited.values.tolist(),
        )

        db_name = st.text_input("Nama di database", key="builder_dbname", max_chars=db.MAX_NAME_LEN)
        f1, f2, f3 = st.columns(3)
        f1.download_button(":material/download: Download .xlsx", final_xlsx,
                           file_name=security.safe_filename(f"{meta['no_tabel'] or 'tabel'}_{(meta['judul'] or 'baru')[:20]}", "xlsx"),
                           mime=XLSX_MIME, use_container_width=True)
        if f2.button(":material/bar_chart: Gunakan Sekarang", use_container_width=True):
            fname = security.safe_filename(meta["no_tabel"] or "tabel-baru", "xlsx")
            push_history({"name": fname, "bytes": final_xlsx, "ext": "xlsx", "db_id": None})
            reset_builder()
            st.rerun()

        if editing:
            if f3.button(":material/save: Perbarui", type="primary", use_container_width=True, key="b_update"):
                err = update_existing(st.session_state.editing_id, db_name, final_xlsx)
                if err:
                    st.error(err)
                else:
                    rid = st.session_state.editing_id
                    reset_builder()
                    open_record(rid)
                    flash("Tabel diperbarui.")
                    st.rerun()
            if st.button(":material/library_add: Simpan sebagai baru", use_container_width=True, key="b_saveas"):
                new_id, err = save_new(db_name, "xlsx", final_xlsx)
                if err:
                    st.error(err)
                else:
                    reset_builder()
                    open_record(new_id)
                    flash("Disimpan sebagai tabel baru.")
                    st.rerun()
        else:
            if f3.button(":material/save: Simpan ke Database", type="primary", use_container_width=True, key="b_save"):
                new_id, err = save_new(db_name, "xlsx", final_xlsx)
                if err:
                    st.error(err)
                else:
                    reset_builder()
                    open_record(new_id)
                    flash("Tabel tersimpan di database.")
                    st.rerun()

# ------------------------------------------------------ riwayat sesi ----
if st.session_state.history:
    with st.container(border=True):
        st.markdown("### :material/history: Riwayat File")
        st.caption("Klik untuk berpindah antar file yang sedang dibuka di sesi ini.")
        hist_items = st.session_state.history
        for i in range(0, len(hist_items), 3):   # maksimal 3 tombol per baris
            cols = st.columns(3)
            for col, h in zip(cols, hist_items[i:i + 3]):
                is_active = h["name"] == st.session_state.active_name
                icon = ":material/check_circle: " if is_active else ":material/description: "
                if col.button(icon + md_escape(h["name"]), key=f"hist_{h['name']}", use_container_width=True,
                              type="primary" if is_active else "secondary"):
                    st.session_state.active_name = h["name"]
                    st.rerun()
    active = next((h for h in st.session_state.history if h["name"] == st.session_state.active_name), None)

# ------------------------------------------------------------ tutorial ----
if active is None:
    steps = [
        ("Unggah", "Pilih file Excel (.xlsx) atau CSV lewat kotak di atas, atau buka lewat tombol Tabel Tersimpan."),
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

file_bytes, ext, filename = active["bytes"], active["ext"], active["name"]

try:
    sheet_names = get_sheet_names(file_bytes, ext)
except Exception as e:
    st.error(f"File **{md_escape(filename)}** tidak bisa dibaca ({e}). Coba unggah ulang atau pilih file lain di riwayat.")
    st.stop()

# ------------------------------------------- simpan file unggahan ke DB ----
if active.get("db_id") is None:
    with st.container(border=True):
        s1, s2 = st.columns([3, 1], vertical_alignment="bottom")
        save_name = s1.text_input("Simpan file ini ke database sebagai", value=Path(filename).stem,
                                  key=f"savename_{filename}", max_chars=db.MAX_NAME_LEN)
        if s2.button(":material/save: Simpan", type="primary", use_container_width=True, key=f"savebtn_{filename}"):
            new_id, err = save_new(save_name, ext, file_bytes)
            if err:
                st.error(err)
            else:
                active["db_id"] = new_id
                flash("Tersimpan di database.")
                st.rerun()
else:
    st.caption(":material/database: Tabel ini tersimpan di database (lihat tombol Tabel Tersimpan).")

# -------------------------------------------------------------- controls ----
with st.container(border=True):
    st.markdown("### :material/tune: Pengaturan")
    c1, c2, c3 = st.columns([2, 1, 1])
    if ext == "csv":
        sheet = sheet_names[0]
        c1.text_input("Sheet", value=sheet, disabled=True,
                      help="File CSV tidak punya banyak sheet seperti Excel.")
    else:
        sheet = c1.selectbox("Sheet", sheet_names)

    try:
        data = get_sheet_data(file_bytes, sheet, ext, filename)
    except Exception as e:
        st.error(f"Gagal membaca isi file: {e}")
        st.stop()

    labels = data["series_labels"]
    unit = c2.text_input("Satuan", value=guess_unit_from_title(data["title"]),
                         key=f"unit_{filename}_{sheet}", placeholder="km, ha, jiwa", max_chars=15)
    orientation = c3.selectbox("Orientasi", ["auto", "horizontal", "vertical"])

    default_cols = [l for l in labels if classify_column(l) == "main"] or labels
    chosen = st.multiselect("Kolom yang ditampilkan", labels, default=default_cols,
                            format_func=lambda l: l + _TAG[classify_column(l)],
                            key=f"cols_{filename}_{sheet}")

    o1, o2 = st.columns(2)
    include_totals = o1.toggle("Sertakan baris Jumlah/Total", value=False)
    no_label = o2.toggle("Tanpa keterangan (angka saja)", value=False)

    overrides = {}
    if len(chosen) > 1 and not no_label:
        with st.expander(":material/edit: Keterangan per kolom (teks di samping angka)"):
            for lbl, default in zip(chosen, short_series_labels(chosen)):
                overrides[lbl] = st.text_input(lbl, value=default, key=f"lbl_{filename}_{sheet}_{lbl}", max_chars=40)

    with st.expander(":material/title: Judul, sumber & nomor gambar"):
        # kunci ikut hash judul+sumber: kalau data diedit, isian kembali ke nilai baru, bukan nyangkut yang lama
        _k = hashlib.md5(f"{data['title']}|{data.get('source') or ''}".encode()).hexdigest()[:8]
        chart_title = st.text_input("Judul grafik", value=data["title"], max_chars=200,
                                    key=f"ct_{filename}_{sheet}_{_k}")
        chart_source = st.text_input("Sumber data", value=data.get("source") or DEFAULT_SOURCE, max_chars=200,
                                     key=f"cs_{filename}_{sheet}_{_k}",
                                     help="Tampil di kaki grafik sebagai 'Sumber: ...'. Kosongkan untuk teks bawaan.")
        fig_no = st.text_input("Nomor gambar", value=str(sheet), max_chars=20, key=f"cn_{filename}_{sheet}_{_k}",
                               help="Tampil sebagai 'Gambar ...' di kaki grafik. Kosongkan kalau tidak diperlukan.")

    def _mpl_safe(t):   # tanda $ akan dibaca matplotlib sebagai rumus, jadi di-escape
        return security.safe_name(t, 200).replace("$", r"\$")

    chart_data = dict(data)
    chart_data["title"] = _mpl_safe(chart_title) or _mpl_safe(data["title"])
    chart_data["source"] = _mpl_safe(chart_source) or None
    chart_fig_no = _mpl_safe(fig_no)

# ----------------------------------------------------------------- stats ----
stat_cards([
    ("Total Sheet", len(sheet_names)),
    ("Baris Data", len(data["categories"])),
    ("Kolom Dipilih", f"{len(chosen)}/{len(labels)}"),
    ("File Aktif", filename),
])

# ------------------------------------------------------------------ tabs ----
tab_chart, tab_data, tab_all = st.tabs([":material/bar_chart: Grafik", ":material/table_chart: Preview Data",
                                        ":material/folder_zip: Semua Sheet"])

with tab_chart:
    if not chosen:
        st.warning("Pilih minimal satu kolom dulu ya.")
    else:
        fig = plot_chart(chart_data, exclude_totals=not include_totals, figure_number=chart_fig_no,
                         unit_hint=unit.strip(), columns=chosen, orientation=orientation,
                         series_label_overrides=overrides, show_series_label=not no_label)
        if fig is None:
            st.warning("Sheet ini tidak punya data untuk digambar.")
        else:
            png = fig_to_png(fig)
            plt.close(fig)
            b64_png = base64.b64encode(png).decode()

            zoom = st.slider(":material/zoom_in: Ukuran tampilan grafik", 50, 200, 100, step=10,
                             format="%d%%", help="Perbesar untuk melihat detail, atau perkecil untuk "
                                                 "melihat keseluruhan grafik sekaligus.")
            with st.container(border=True):
                st.markdown(
                    f'<div class="chart-frame"><img src="data:image/png;base64,{b64_png}" '
                    f'style="width:{int(zoom)}%;"></div>',
                    unsafe_allow_html=True,
                )

            d1, d2 = st.columns(2)
            if d1.button(":material/open_in_full: Buka gambar penuh", use_container_width=True,
                         key="open_full_chart"):
                show_full_chart(png)
            if d2.button(":material/download: Download PNG", use_container_width=True, key="open_png_dialog"):
                download_chart_dialog(png, f"chart_{Path(filename).stem}_{str(sheet).replace('.', '_')}")

            with st.expander(":material/notes: Narasi otomatis (siap disalin)"):
                narasi = build_narrative(data, chosen, unit=unit.strip(),
                                         exclude_totals=not include_totals)
                st.text_area("Teks narasi", value=narasi, height=220,
                             key=f"narasi_{filename}_{sheet}_{_k}",
                             help="Bisa diedit dulu sebelum disalin.")

with tab_data:
    st.markdown(f"**{md_escape(data['title'])}**")
    st.caption(f"{len(data['categories'])} baris, {len(labels)} kolom data. "
               f"Cek dulu di sini sebelum bikin grafik, siapa tahu ada data yang aneh atau salah ketik.")

    edit_target = (filename, sheet)
    edit_on = st.session_state.get("preview_edit_target") == edit_target

    if not edit_on:
        if st.button(":material/edit: Edit nama kolom & angka", key="prev_edit_open"):
            st.session_state.preview_edit_target = edit_target
            st.session_state.prev_edit_ver += 1
            st.rerun()
        df = pd.DataFrame(data["series_values"], index=data["categories"])
        st.dataframe(df, use_container_width=True)
    else:
        ver = st.session_state.prev_edit_ver
        st.info("Mode edit: ubah judul, nama kolom, nama baris, atau angka. Klik dua kali sel untuk mengetik. "
                "Perubahan baru berlaku setelah kamu klik **Terapkan**.")
        if len(sheet_names) > 1:
            st.warning("File ini punya lebih dari satu sheet. Setelah diterapkan, hanya sheet ini yang dipertahankan "
                       "(header bertingkat juga diratakan jadi satu baris).")

        new_title = st.text_input("Judul tabel", value=data["title"], key=f"prev_title_{ver}", max_chars=200)

        st.markdown("**Nama kolom**")
        new_labels = []
        rcols = st.columns(2)
        for j, lbl in enumerate(labels):
            new_labels.append(rcols[j % 2].text_input(f"Kolom {j + 1}", value=lbl, key=f"prev_ren_{ver}_{j}",
                                                      max_chars=100))

        st.markdown("**Angka & nama baris**")
        df_edit = pd.DataFrame({"c_row": data["categories"],
                                **{f"c{j}": data["series_values"][lbl] for j, lbl in enumerate(labels)}})
        cfg = {"c_row": st.column_config.TextColumn(data.get("category_label") or "Wilayah")}
        for j, nl in enumerate(new_labels):
            cfg[f"c{j}"] = st.column_config.NumberColumn(label=nl or labels[j])
        edited_df = st.data_editor(df_edit, column_config=cfg, hide_index=True, num_rows="fixed",
                                   use_container_width=True, key=f"prev_editor_{ver}")

        has_db = bool(active.get("db_id"))
        bcols = st.columns(3 if has_db else 2)
        do_apply = bcols[0].button(":material/check: Terapkan", type="primary", use_container_width=True,
                                   key="prev_apply")
        do_apply_db = has_db and bcols[1].button(":material/save: Terapkan & simpan ke database",
                                                 use_container_width=True, key="prev_apply_db")
        if bcols[-1].button(":material/close: Batal", use_container_width=True, key="prev_cancel"):
            st.session_state.preview_edit_target = None
            st.rerun()

        if do_apply or do_apply_db:
            err = apply_preview_edit(active, data, sheet, edited_df, new_labels, new_title, save_db=do_apply_db)
            if err:
                st.error(err)
            else:
                st.session_state.preview_edit_target = None
                flash("Perubahan diterapkan dan tersimpan di database." if do_apply_db else "Perubahan diterapkan.")
                st.rerun()

with tab_all:
    if ext == "csv":
        st.info("File CSV cuma punya satu sheet, jadi fitur ini otomatis sama saja dengan tab Grafik.")
    else:
        st.write("Tiap sheet memakai kolom **utama**-nya saja. Hasilnya dikemas dalam satu file ZIP.")
        if st.button(":material/play_arrow: Generate semua grafik", type="primary"):
            zbuf = io.BytesIO()
            progress = st.progress(0.0, text="Memproses...")
            with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as zf:
                for i, sh in enumerate(sheet_names):
                    try:
                        d = get_sheet_data(file_bytes, sh, ext, filename)
                        main_cols = [l for l in d["series_labels"] if classify_column(l) == "main"]
                        f = plot_chart(d, exclude_totals=True, figure_number=sh,
                                       unit_hint=guess_unit_from_title(d["title"]),
                                       columns=main_cols or None)
                        if f is not None:
                            zf.writestr(security.safe_filename(f"chart_{sh.replace('.', '_')}", "png"), fig_to_png(f))
                            plt.close(f)
                    except Exception as e:
                        st.warning(f"Sheet '{md_escape(sh)}' dilewati: {e}")
                    progress.progress((i + 1) / len(sheet_names), text=f"Sheet {sh}")
            progress.empty()
            st.success("Selesai!")
            st.download_button(":material/download: Download ZIP", zbuf.getvalue(),
                               file_name=security.safe_filename(f"charts_{Path(filename).stem}", "zip"),
                               mime="application/zip")

st.markdown(FOOTER_HTML, unsafe_allow_html=True)
