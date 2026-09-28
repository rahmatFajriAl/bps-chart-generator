"""Tema tampilan SETELAH login (ruang kerja). Tidak mengubah logika aplikasi.
Warna tetap oranye khas BPS; dipakai sebagai aksen di atas dasar netral gelap-terang."""
from html import escape

import streamlit as st

THEME_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600;700&display=swap');

:root {
  --o1: #F0932B; --o2: #EB6E24; --o3: #B34A0E;
  --bar: #2A1A10;          /* cokelat gelap untuk top bar */
  --bg: #F5F4F2;           /* netral abu hangat */
  --panel: #FFFFFF;
  --ink: #221A14; --muted: #6F655B; --rule: #E3DFD9;
}

html, body, [class*="css"], .stApp, button, input, textarea, select {
  font-family: 'IBM Plex Sans', system-ui, sans-serif !important;
}
.stApp { background: var(--bg) !important; }
.block-container { max-width: 1180px !important; padding: 3.5rem 1.5rem 4rem !important; }

/* ---------- TOP BAR (pengganti hero besar) ---------- */
.topbar {
  display: flex; align-items: center; gap: 1rem; flex-wrap: wrap;
  background: var(--bar); color: #fff; border-radius: 12px;
  padding: .7rem 1.1rem; margin-bottom: 1.1rem;
  border-bottom: 3px solid var(--o2);
}
.topbar .logo { background: #fff; border-radius: 8px; padding: .25rem .4rem; display: flex; }
.topbar .logo img { height: 32px; width: auto; display: block; }
.topbar .org { font-weight: 600; font-size: .95rem; line-height: 1.15; }
.topbar .org small { display: block; font-weight: 400; font-size: .74rem; color: #CDBFB2; }
.topbar .sep { width: 1px; align-self: stretch; background: rgba(255,255,255,.18); margin: .1rem .2rem; }
.topbar .page { flex: 1 1 220px; min-width: 0; }
.topbar .page b { display: block; font-size: 1.08rem; font-weight: 700; letter-spacing: -.005em; }
.topbar .page span { display: block; font-size: .82rem; color: #CDBFB2; margin-top: 1px; }
.topbar .user {
  display: flex; align-items: center; gap: .5rem; font-size: .82rem; color: #EADFD3;
  background: rgba(255,255,255,.08); border-radius: 999px; padding: .3rem .8rem .3rem .35rem;
}
.topbar .user i {
  width: 24px; height: 24px; border-radius: 50%; background: var(--o2); color: #fff;
  font-style: normal; font-weight: 700; font-size: .78rem; display: grid; place-items: center;
}

/* ---------- PANEL: datar, garis tipis, tanpa bayangan ---------- */
[data-testid="stVerticalBlockBorderWrapper"] {
  background: var(--panel) !important; border: 1px solid var(--rule) !important;
  border-radius: 10px !important; box-shadow: none !important; transition: none !important;
}
[data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"]:hover { transform: none !important; }
h3 { font-size: 1.05rem !important; font-weight: 700 !important; color: var(--ink) !important; }
[data-testid="stExpander"] { border-radius: 10px !important; border: 1px solid var(--rule) !important;
  background: var(--panel) !important; box-shadow: none !important; }
[data-testid="stExpander"] summary { font-weight: 600; }

/* ---------- TOMBOL ---------- */
.stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
  border-radius: 8px !important; font-weight: 600 !important; padding: .55rem .9rem !important;
  min-height: 44px;
  background: var(--panel) !important; color: var(--ink) !important; border: 1px solid var(--rule) !important;
  box-shadow: none !important; filter: none !important; transform: none !important;
  transition: border-color .15s, background .15s, color .15s !important;
}
.stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
  border-color: var(--o2) !important; color: var(--o3) !important; background: #FFF6EC !important;
}
.stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"],
.stDownloadButton > button[kind="primary"] {
  background: var(--o2) !important; color: #fff !important; border: 1px solid var(--o2) !important;
}
.stButton > button[kind="primary"]:hover, .stFormSubmitButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover { background: var(--o3) !important; border-color: var(--o3) !important; color: #fff !important; }
.stDownloadButton > button:not([kind="primary"]) { background: var(--panel) !important; color: var(--ink) !important;
  border: 1px solid var(--rule) !important; }
:focus-visible { outline: 2px solid var(--o2) !important; outline-offset: 2px; }

/* ---------- INPUT ---------- */
[data-baseweb="input"], [data-baseweb="select"] > div, [data-baseweb="textarea"] {
  border-radius: 8px !important; background: #fff !important; border-color: var(--rule) !important;
}
[data-testid="stFileUploader"] section {
  background: #fff !important; border: 1.5px dashed #C9C2B9 !important; border-radius: 10px !important;
  transform: none !important;
}
[data-testid="stFileUploader"] section:hover { border-color: var(--o2) !important; background: #FFF9F2 !important; }

/* ---------- PANDUAN: 3 langkah berdampingan (memang urutan) ---------- */
.howto { background: var(--panel) !important; border: 1px solid var(--rule) !important; border-radius: 10px !important;
  box-shadow: none !important; padding: 1rem 1.2rem !important; animation: none !important;
  display: grid; grid-template-columns: repeat(3, 1fr); gap: 0 1.4rem; }
.howto .head { grid-column: 1 / -1; text-transform: none !important; letter-spacing: 0 !important;
  font-size: .95rem !important; font-weight: 700 !important; color: var(--ink) !important; margin-bottom: .8rem !important; }
.howto .row { opacity: 1 !important; animation: none !important; padding: 0 0 0 .9rem !important;
  border-left: 3px solid var(--o1); gap: .7rem !important; }
.howto .row::before { display: none !important; }
.howto .num { width: 26px !important; height: 26px !important; border-radius: 6px !important; font-size: .8rem !important;
  background: var(--o2) !important; box-shadow: none !important; }
.howto .t { line-height: 26px !important; font-size: .98rem !important; }
.howto .d { margin-top: .25rem !important; font-size: .84rem !important; }

/* ---------- STATISTIK: satu strip bersekat ---------- */
.stats { gap: 0 !important; background: var(--panel); border: 1px solid var(--rule); border-radius: 10px; overflow: hidden; }
.stat { border: none !important; border-radius: 0 !important; box-shadow: none !important; opacity: 1 !important;
  animation: none !important; transform: none !important; border-right: 1px solid var(--rule) !important;
  border-left: 3px solid transparent !important; padding: .8rem 1rem !important; }
.stat:last-child { border-right: none !important; }
.stat:first-child { border-left-color: var(--o2) !important; }
.stat .k { text-transform: none !important; letter-spacing: 0 !important; font-size: .78rem !important;
  font-weight: 500 !important; color: var(--muted) !important; }
.stat .v { font-size: 1.35rem !important; font-weight: 700 !important; color: var(--ink) !important; }

/* ---------- TAB: garis bawah ---------- */
.stTabs [data-baseweb="tab-list"] { background: transparent !important; border: none !important; box-shadow: none !important;
  border-bottom: 1px solid var(--rule) !important; border-radius: 0 !important; padding: 0 !important; gap: .2rem !important; }
.stTabs [data-baseweb="tab"] { border-radius: 0 !important; background: transparent !important; color: var(--muted);
  padding: .6rem 1rem !important; font-weight: 600 !important; border-bottom: 3px solid transparent; margin-bottom: -1px; }
.stTabs [aria-selected="true"] { background: transparent !important; color: var(--o3) !important;
  border-bottom: 3px solid var(--o2) !important; }

/* ---------- GRAFIK & TABEL ---------- */
.chart-frame { border: 1px solid var(--rule); border-radius: 8px !important; background: #fff;
  touch-action: pan-x pan-y pinch-zoom; }
.chart-frame img { border-radius: 8px !important; }
[data-testid="stDataFrame"], [data-testid="stDataEditor"] { border: 1px solid var(--rule); border-radius: 8px; overflow: hidden; }
.footer-note { color: var(--muted) !important; }

/* ---------- MOBILE ---------- */
@media (max-width: 768px) {
  .block-container { padding: 3.5rem .7rem 3rem !important; }
  .howto { grid-template-columns: 1fr; gap: 1rem 0; }
  .topbar { padding: .6rem .8rem; gap: .6rem; }
  .topbar .sep { display: none; }
  .topbar .page { flex: 1 1 100%; order: 3; }
  .topbar .user { margin-left: auto; max-width: 60%; overflow-wrap: anywhere; }
  .stats { grid-template-columns: repeat(2, 1fr) !important; }
  .stat { border-bottom: 1px solid var(--rule) !important; }
  .stat:nth-child(2n) { border-right: none !important; }
  .stat:nth-last-child(-n+2) { border-bottom: none !important; }
}
</style>
"""


def inject_app_theme():
    st.markdown(THEME_CSS, unsafe_allow_html=True)


def render_topbar(brand, logo_html, title, subtitle, user):
    """Top bar ringkas. `logo_html` = variabel logo_html yang sudah ada di app.py."""
    initial = escape((user or "?")[:1].upper())
    st.markdown(f"""
<div class="topbar">
  {logo_html}
  <div class="org">{escape(brand)}<small>Badan Pusat Statistik</small></div>
  <div class="sep"></div>
  <div class="page"><b>{escape(title)}</b><span>{subtitle}</span></div>
  <div class="user"><i>{initial}</i>{escape(user or "")}</div>
</div>
""", unsafe_allow_html=True)