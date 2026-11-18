"""
Monitoring & Early Warning Keterlambatan Proyek (Streamlit) - Versi 3

- Tampilan bergaya SaaS modern (netral, kartu putih, badge status lembut).
- Data tersimpan otomatis di alamat halaman (URL), sehingga tidak hilang saat di-refresh.

Menu:
  1. Dashboard Utama   -> status proyek dari 3 data (target, realisasi, sisa waktu)
  2. Detail Kurva S    -> progres per item pekerjaan + Kurva S
  3. Export Laporan    -> download PDF atau Excel

Jalankan lokal:  python -m streamlit run app.py
"""
import base64
import csv
import html
import io
import json
import math
import zlib
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from fpdf import FPDF
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Monitoring Proyek", page_icon=":material/monitoring:", layout="wide",
                   initial_sidebar_state="collapsed")
S = st.session_state
esc = html.escape

# ----------------------------------------------------------------------------
# KONSTANTA
# ----------------------------------------------------------------------------
COL_NAMA = "Item Pekerjaan"
COL_BOBOT = "Bobot (%)"
COL_MULAI = "Mulai (hari ke)"
COL_DURASI = "Durasi (hari)"
COL_REAL = "Realisasi (%)"
COL_HARI = "Hari ke"
COL_RKUM = "Realisasi kumulatif (%)"
ITEM_COLS = [COL_NAMA, COL_BOBOT, COL_MULAI, COL_DURASI, COL_REAL]
RIW_COLS = [COL_HARI, COL_RKUM]

MODE_MANUAL = "Manual (3 input)"
MODE_AUTO = "Otomatis dari Detail Kurva S"

NAV = {"dashboard": "Dashboard Utama", "kurva": "Detail Kurva S", "export": "Export Laporan",
    "history": "Riwayat Cek"}
NAV_ICON = {"dashboard": "dashboard", "kurva": "chart", "export": "file", "history": "activity"}
HISTORY_FILE = Path(__file__).resolve().with_name("riwayat_proyek.json")

BULAN = ["Jan", "Feb", "Mar", "Apr", "Mei", "Jun", "Jul", "Agu", "Sep", "Okt", "Nov", "Des"]

STATUS = {
    "On Track": {"cls": "b-ok", "color": "#047857", "rgb": (4, 120, 87),
                 "msg": "Progres berada dalam batas toleransi terhadap rencana."},
    "At Risk": {"cls": "b-risk", "color": "#B45309", "rgb": (180, 83, 9),
                "msg": "Deviasi mulai melebar. Perlu percepatan ringan."},
    "Delayed": {"cls": "b-late", "color": "#B91C1C", "rgb": (185, 28, 28),
                "msg": "Terlambat serius. Perlu percepatan segera dan laporan ke manajemen."},
}

C_INK, C_PLAN, C_PLAN_LIGHT = "#0F172A", "#94A3B8", "#CBD5E1"

# ----------------------------------------------------------------------------
# IKON (Lucide, outline, 16-18px)
# ----------------------------------------------------------------------------
ICONS = {
    "dashboard": '<rect width="7" height="9" x="3" y="3" rx="1"/><rect width="7" height="5" x="14" y="3" rx="1"/>'
                 '<rect width="7" height="9" x="14" y="12" rx="1"/><rect width="7" height="5" x="3" y="16" rx="1"/>',
    "chart": '<path d="M3 3v16a2 2 0 0 0 2 2h16"/><path d="m19 9-5 5-4-4-3 3"/>',
    "file": '<path d="M15 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V7Z"/><path d="M14 2v4a2 2 0 0 0 2 2h4"/>'
            '<path d="M12 18v-6"/><path d="m9 15 3 3 3-3"/>',
    "activity": '<path d="M22 12h-2.48a2 2 0 0 0-1.93 1.46l-2.35 8.36a.25.25 0 0 1-.48 0L9.24 2.18a.25.25 0 0 0-.48 0'
                'l-2.35 8.36A2 2 0 0 1 4.49 12H2"/>',
    "clipboard": '<rect width="8" height="4" x="8" y="2" rx="1" ry="1"/>'
                 '<path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2"/>'
                 '<path d="M12 11h4"/><path d="M12 16h4"/><path d="M8 11h.01"/><path d="M8 16h.01"/>',
    "download": '<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/>'
                '<line x1="12" x2="12" y1="15" y2="3"/>',
    "trash": '<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>'
             '<line x1="10" x2="10" y1="11" y2="17"/><line x1="14" x2="14" y1="11" y2="17"/>',
}


def svg(name: str, size: int = 16) -> str:
    return (f'<svg xmlns="http://www.w3.org/2000/svg" width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" '
            f'stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" '
            f'style="flex:none">{ICONS[name]}</svg>')


def mask(name: str) -> str:
    s = ("<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='black' stroke-width='2' "
         "stroke-linecap='round' stroke-linejoin='round'>" + ICONS[name].replace('"', "'") + "</svg>")
    return 'url("data:image/svg+xml,' + quote(s, safe="") + '")'


# ----------------------------------------------------------------------------
# CSS
# ----------------------------------------------------------------------------
CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
.stApp, .stApp p, .stApp label, .stApp input, .stApp textarea, .stApp button, .stApp li, .stApp td, .stApp th,
.stApp h1, .stApp h2, .stApp h3, .stApp h4, .stApp [data-testid="stMarkdownContainer"] {
  font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif !important; }
.stApp { background:#F8FAFC; color:#0F172A; font-variant-numeric:tabular-nums; font-feature-settings:"tnum" 1;
  -webkit-font-smoothing:antialiased; }
[data-testid="stHeader"] { background:#FFFFFF; }
[data-testid="stSidebarCollapsedControl"], [data-testid="collapsedControl"] {
    display:flex !important; visibility:visible !important; opacity:1 !important; pointer-events:auto !important; z-index:1001 !important; }
[data-testid="stSidebarCollapsedControl"], [data-testid="stExpandSidebarButton"] { color:#0F172A !important; }
[data-testid="stSidebarCollapsedControl"] svg, [data-testid="stExpandSidebarButton"] svg {
    color:#0F172A !important; stroke:#0F172A !important; }
[data-testid="stToolbar"], [data-testid="stDecoration"], #MainMenu, footer { display:none !important; }
.block-container { max-width:1180px; padding:1.6rem 2rem 3rem; }

/* sidebar */
section[data-testid="stSidebar"] { width:240px !important; min-width:240px !important; max-width:240px !important;
  background:#FFFFFF; border-right:1px solid #E2E8F0; }
section[data-testid="stSidebar"] > div { width:240px !important; }
[data-testid="stSidebarUserContent"] { padding:0.4rem 0.9rem 1.5rem !important; }
.brand { display:flex; align-items:center; gap:10px; padding:2px 4px 14px; }
.brand .lg { width:28px; height:28px; border-radius:7px; background:#0F172A; color:#fff;
  display:inline-flex; align-items:center; justify-content:center; }
.brand .nm { font-size:14px; font-weight:600; color:#0F172A; line-height:1.2; }
.brand .sb { font-size:11.5px; color:#64748B; }
.side-sec { font-size:11.5px; font-weight:600; color:#334155; margin:18px 4px 6px; }

[class*="st-key-nav_"] button { width:100%; justify-content:flex-start; gap:10px; padding:0 10px; min-height:0; height:34px;
  border:none; border-radius:6px; background:transparent; color:#475569; box-shadow:none; }
[class*="st-key-nav_"] button p { font-size:13.5px; font-weight:500; color:inherit; margin:0; }
[class*="st-key-nav_"] button::before { content:""; width:16px; height:16px; flex:none; background-color:currentColor;
  -webkit-mask:var(--ic) center/contain no-repeat; mask:var(--ic) center/contain no-repeat; }
[class*="st-key-nav_"] button:hover { background:#F1F5F9; color:#0F172A; }
.st-key-nav_dashboard button { --ic: __IC_DASHBOARD__; }
.st-key-nav_kurva button { --ic: __IC_CHART__; }
.st-key-nav_export button { --ic: __IC_FILE__; }
.st-key-top_nav_dashboard button, .st-key-top_nav_kurva button, .st-key-top_nav_export button, .st-key-top_nav_history button {
    width:100%; min-height:38px; border:1px solid #E2E8F0; border-radius:8px; background:#FFFFFF; color:#334155; }
.st-key-top_nav_dashboard button:hover, .st-key-top_nav_kurva button:hover, .st-key-top_nav_export button:hover,
.st-key-top_nav_history button:hover {
    background:#F1F5F9; border-color:#CBD5E1; color:#0F172A; }
div[data-testid="stHorizontalBlock"]:has(.st-key-top_nav_dashboard) {
    flex-direction:row !important; flex-wrap:nowrap !important; overflow-x:auto; gap:8px; }
div[data-testid="stHorizontalBlock"]:has(.st-key-top_nav_dashboard) > div[data-testid="stColumn"] {
    flex:1 0 132px !important; width:auto !important; min-width:132px !important; }

.st-key-reset button { background:#FFFFFF; border:1px solid #E2E8F0; color:#334155; height:34px; min-height:0;
  border-radius:8px; box-shadow:none; justify-content:center; gap:8px; }
.st-key-reset button p { font-size:13px; font-weight:500; color:inherit; margin:0; }
.st-key-reset button:hover { background:#F8FAFC; border-color:#CBD5E1; color:#0F172A; }
.st-key-reset button::before { content:""; width:15px; height:15px; background-color:currentColor;
  -webkit-mask:__IC_TRASH__ center/contain no-repeat; mask:__IC_TRASH__ center/contain no-repeat; }

/* input */
.stApp label p { font-size:12.5px; font-weight:500; color:#475569; }
div[data-baseweb="input"], div[data-baseweb="base-input"] { background:#FFFFFF !important; border-radius:8px !important; }
div[data-baseweb="input"] { border:1px solid #E2E8F0 !important; }
div[data-baseweb="input"]:focus-within { border-color:#0F172A !important; box-shadow:0 0 0 3px rgba(15,23,42,.08) !important; }
.stApp input { font-size:13.5px !important; }

/* kartu dari container ber-key */
[class*="st-key-card_"] { background:#FFFFFF; border:1px solid #E2E8F0; border-radius:10px; padding:14px 16px;
  margin:0 0 14px; }

/* tombol utama */
[data-testid="stDownloadButton"] button { background:#0F172A; border:1px solid #0F172A; color:#fff; height:38px; min-height:0;
  border-radius:8px; gap:8px; box-shadow:none; }
[data-testid="stDownloadButton"] button p { color:#fff; font-size:13.5px; font-weight:600; margin:0; }
[data-testid="stDownloadButton"] button:hover { background:#1E293B; border-color:#1E293B; color:#fff; }
[data-testid="stDownloadButton"] button::before { content:""; width:16px; height:16px; background-color:currentColor;
  -webkit-mask:__IC_DOWNLOAD__ center/contain no-repeat; mask:__IC_DOWNLOAD__ center/contain no-repeat; }

/* komponen HTML */
.ph { display:flex; align-items:flex-start; justify-content:space-between; gap:16px; margin:0 0 6px;
    padding:12px 14px; background:linear-gradient(135deg,#FFFFFF 0%,#EFF6FF 100%);
    border-left:4px solid #1E3A8A; border-radius:8px; }
.ph-t { font-size:22px; font-weight:600; letter-spacing:-0.015em; color:#0F172A; line-height:1.25; }
.ph-s { font-size:13.5px; color:#334155; margin-top:3px; }
.meta { display:flex; flex-wrap:wrap; align-items:center; gap:6px 14px; font-size:12.5px; color:#334155;
  margin:0 0 18px; padding-bottom:14px; border-bottom:1px solid #E2E8F0; }
.meta b { color:#334155; font-weight:600; }
.meta i { width:3px; height:3px; border-radius:50%; background:#CBD5E1; display:inline-block; }

.badge { display:inline-flex; align-items:center; gap:6px; font-size:14px; font-weight:600; line-height:1;
    padding:7px 14px; border-radius:999px; border:1px solid transparent; white-space:nowrap; }
.badge::before { content:""; width:8px; height:8px; border-radius:50%; background:currentColor; }
.ph .badge { font-size:15px; }
.b-ok { background:#ECFDF5; color:#047857; border-color:#A7F3D0; }
.b-risk { background:#FFFBEB; color:#B45309; border-color:#FDE68A; }
.b-late { background:#FEF2F2; color:#B91C1C; border-color:#FECACA; }
.b-idle { background:#F1F5F9; color:#475569; border-color:#E2E8F0; }

.kpis { display:grid; grid-template-columns:repeat(auto-fit,minmax(150px,1fr)); gap:1px; background:#E2E8F0;
  border:1px solid #E2E8F0; border-radius:10px; overflow:hidden; margin:0 0 16px; }
.kpi { background:#FFFFFF; padding:12px 14px; border-top:3px solid var(--kpi-accent,#CBD5E1); }
.kpi .l { display:table; padding:3px 6px; border-radius:4px; background:var(--kpi-soft,#F1F5F9);
    font-size:12px; font-weight:600; color:#334155; }
.kpi .v { font-size:22px; font-weight:600; letter-spacing:-0.01em; color:#0F172A; line-height:1.3; margin-top:2px; }
.kpi .u { font-size:12px; font-weight:500; color:#64748B; margin-left:3px; }
.kpi .d { font-size:11.5px; color:#334155; margin-top:1px; }

.card { background:#FFFFFF; border:1px solid #E2E8F0; border-radius:10px; padding:14px 16px; margin:0 0 14px; }
.rec-card.b-ok { background:#F0FDF4; border-left:3px solid #16A34A; }
.rec-card.b-risk { background:#FFFBEB; border-left:3px solid #D97706; }
.rec-card.b-late { background:#FEF2F2; border-left:3px solid #DC2626; }
.card-h { font-size:13px; font-weight:600; color:#0F172A; margin-bottom:8px; }
.rec { font-size:13.5px; line-height:1.6; color:#334155; margin:0; }
.prow { display:flex; align-items:center; gap:12px; margin:10px 0; font-size:13px; }
.pn { width:70px; color:#334155; font-weight:500; }
.pv { width:54px; text-align:right; font-weight:600; color:#0F172A; }
.trk { flex:1; background:#EEF2F7; border-radius:999px; height:8px; overflow:hidden; }
.fil { height:100%; border-radius:999px; }
.fil-plan { background:#CBD5E1; }
.fil-act { background:#0F172A; }
.pfoot { display:flex; align-items:center; gap:10px; margin-top:12px; padding-top:12px; border-top:1px solid #E2E8F0;
  font-size:13px; color:#475569; }

.tbl-wrap { overflow-x:auto; }
table.tbl { width:100%; border-collapse:collapse; font-size:13px; }
.tbl th { text-align:right; font-size:12px; font-weight:500; color:#64748B; padding:8px 10px; border-bottom:1px solid #E2E8F0; }
.tbl td { text-align:right; padding:9px 10px; border-bottom:1px solid #F1F5F9; color:#0F172A; }
.tbl th:first-child, .tbl td:first-child { text-align:left; }
.tbl tr:last-child td { border-bottom:none; }
.mini { display:flex; align-items:center; justify-content:flex-end; gap:8px; }
.mini .trk { flex:none; width:72px; height:6px; }

.empty { background:#FFFFFF; border:1px dashed #CBD5E1; border-radius:10px; padding:32px 24px; text-align:center; }
.empty .ic { width:36px; height:36px; border-radius:8px; border:1px solid #E2E8F0; background:#F8FAFC; color:#475569;
  display:inline-flex; align-items:center; justify-content:center; }
.empty .e2 { font-size:15px; font-weight:600; color:#0F172A; margin-top:10px; }
.empty .e3 { font-size:13px; color:#64748B; margin:4px auto 0; line-height:1.55; max-width:460px; }

.sec { font-size:14px; font-weight:600; color:#0F172A; margin:18px 0 2px; }
.sec-d { font-size:12.5px; color:#334155; margin:0 0 8px; }
.note { background:#EFF6FF; border:1px solid #DBEAFE; border-left:3px solid #2563EB; border-radius:8px;
  padding:9px 12px; font-size:13px; color:#334155; margin:8px 0 12px; }
.read-note { background:#EFF6FF; border:1px solid #DBEAFE; border-left:3px solid #2563EB; border-radius:8px;
    padding:10px 12px; font-size:13px; color:#1E293B; line-height:1.55; }
</style>
"""
for _n in ("dashboard", "chart", "file", "download", "trash"):
    CSS = CSS.replace(f"__IC_{_n.upper()}__", mask(_n))


# ----------------------------------------------------------------------------
# UTILITAS FORMAT
# ----------------------------------------------------------------------------
def tgl_id(d) -> str:
    return f"{d.day} {BULAN[d.month - 1]} {d.year}"


def fmt(v, nd=1, suffix=""):
    return "-" if v is None else f"{v:.{nd}f}{suffix}".replace(".", ",")


def badge(status):
    if status is None:
        return '<span class="badge b-idle">Belum ada data</span>'
    return f'<span class="badge {STATUS[status]["cls"]}">{status}</span>'


# ----------------------------------------------------------------------------
# DATA KOSONG + PENYIMPANAN DI URL
# ----------------------------------------------------------------------------
def make_items(rows=(), min_rows=6) -> pd.DataFrame:
    rows = [list(r) for r in rows][:60]
    while len(rows) < min_rows:
        rows.append([None] * 5)
    df = pd.DataFrame(rows, columns=ITEM_COLS)
    df[COL_NAMA] = df[COL_NAMA].astype("object")
    for c in ITEM_COLS[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    return df


def _rab_number(value):
    if value is None or pd.isna(value):
        return None
    text = str(value).strip().replace("%", "").replace("\u00a0", "").replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".") if text.rfind(",") > text.rfind(".") else text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")
    try:
        number = float(text)
    except ValueError:
        return None
    return number if math.isfinite(number) else None


def parse_rab_rows(rows):
    items, invalid_rows = [], []
    for row_number, row in enumerate(rows, start=1):
        cells = ["" if value is None or pd.isna(value) else str(value).strip() for value in row]
        while cells and not cells[-1]:
            cells.pop()
        if not cells or not any(cells):
            continue
        header_cells = {cell.lower() for cell in cells}
        if header_cells.intersection({"nama pekerjaan", "bobot (%)", "hari mulai", "durasi (hari)", "realisasi (%)"}):
            continue
        if len(cells) < 4 or len(cells) > 5 or not cells[0]:
            invalid_rows.append(row_number)
            continue
        values = [_rab_number(value) for value in cells[1:]]
        if len(values) == 3:
            values.append(0.0)
        bobot, mulai, durasi, real = values
        if (any(value is None for value in values) or not 0 <= bobot <= 100 or
                not 1 <= mulai <= 3650 or not 1 <= durasi <= 3650 or not 0 <= real <= 100):
            invalid_rows.append(row_number)
            continue
        items.append([cells[0], bobot, mulai, durasi, real])
    return items, invalid_rows


def parse_rab_text(text):
    lines = [line for line in text.splitlines() if line.strip()]
    if not lines:
        return [], []
    sample = lines[0]
    delimiter = "\t" if "\t" in sample else ";" if ";" in sample else ","
    return parse_rab_rows(csv.reader(lines, delimiter=delimiter))


def make_riwayat(rows=(), min_rows=5) -> pd.DataFrame:
    rows = [list(r) for r in rows][:60]
    while len(rows) < min_rows:
        rows.append([None] * 2)
    df = pd.DataFrame(rows, columns=RIW_COLS)
    for c in RIW_COLS:
        df[c] = pd.to_numeric(df[c], errors="coerce").astype(float)
    return df


def _py(v):
    if v is None:
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating, float)):
        v = float(v)
        return None if (math.isnan(v) or math.isinf(v)) else round(v, 4)
    if isinstance(v, str):
        return v
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    return v


def _rows(df: pd.DataFrame, cols):
    out = []
    for _, r in df.iterrows():
        vals = [_py(r[c]) for c in cols]
        if any(v not in (None, "") for v in vals):
            out.append(vals)
    return out


def _num(v, lo=None, hi=None, integer=False):
    if v is None or isinstance(v, bool):
        return None
    try:
        x = float(v)
    except Exception:
        return None
    if math.isnan(x) or math.isinf(x):
        return None
    if lo is not None and x < lo:
        return None
    if hi is not None and x > hi:
        return None
    return int(round(x)) if integer else x


def encode_state() -> str:
    d = {
        "nama_proyek": S["nama_proyek"],
        "tgl_mulai": S["tgl_mulai"].isoformat(),
        "tgl_acuan": S["tgl_acuan"].isoformat(),
        "durasi_total": _py(S["durasi_total"]),
        "toleransi_warning": _py(S["toleransi_warning"]),
        "batas_critical": _py(S["batas_critical"]),
        "mode_data": S["mode_data"],
        "m_target": _py(S["m_target"]),
        "m_real": _py(S["m_real"]),
        "m_sisa": _py(S["m_sisa"]),
        "items": _rows(S["data_pekerjaan"], ITEM_COLS),
        "riwayat": _rows(S["riwayat"], RIW_COLS),
    }
    raw = json.dumps(d, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    return base64.urlsafe_b64encode(zlib.compress(raw, 9)).decode("ascii").rstrip("=")


def decode_state(blob: str) -> dict:
    if not isinstance(blob, str) or len(blob) > 20000:
        raise ValueError("ukuran tidak valid")
    raw = base64.urlsafe_b64decode(blob + "=" * (-len(blob) % 4))
    obj = zlib.decompressobj()
    data = obj.decompress(raw, 200_000)
    if obj.unconsumed_tail:
        raise ValueError("terlalu besar")
    return json.loads(data.decode("utf-8"))


def apply_loaded(d: dict):
    if not isinstance(d, dict):
        return
    if isinstance(d.get("nama_proyek"), str):
        S["nama_proyek"] = d["nama_proyek"][:120]
    for k in ("tgl_mulai", "tgl_acuan"):
        try:
            S[k] = date.fromisoformat(str(d.get(k)))
        except Exception:
            pass
    S["durasi_total"] = _num(d.get("durasi_total"), 1, 3650, True)
    for k, dflt in (("toleransi_warning", 2.0), ("batas_critical", 10.0)):
        v = _num(d.get(k), 0, 100)
        S[k] = float(v) if v is not None else dflt
    if d.get("mode_data") in (MODE_MANUAL, MODE_AUTO):
        S["mode_data"] = d["mode_data"]
    S["m_target"] = _num(d.get("m_target"), 0, 100)
    S["m_real"] = _num(d.get("m_real"), 0, 100)
    S["m_sisa"] = _num(d.get("m_sisa"), 0, 3650, True)

    items = []
    for r in (d.get("items") or [])[:60]:
        if not isinstance(r, list) or len(r) != 5:
            continue
        nama = r[0][:80] if isinstance(r[0], str) else None
        items.append([nama, _num(r[1], 0, 100), _num(r[2], 0, 3650), _num(r[3], 0, 3650), _num(r[4], 0, 100)])
    S["data_pekerjaan"] = make_items(items)

    riw = []
    for r in (d.get("riwayat") or [])[:60]:
        if isinstance(r, list) and len(r) == 2:
            riw.append([_num(r[0], 0, 3650), _num(r[1], 0, 100)])
    S["riwayat"] = make_riwayat(riw)


def init_state():
    if "riwayat_cek" not in S:
        try:
            data = json.loads(HISTORY_FILE.read_text(encoding="utf-8")) if HISTORY_FILE.exists() else []
            S["riwayat_cek"] = data[:5000] if isinstance(data, list) else []
        except (OSError, json.JSONDecodeError):
            S["riwayat_cek"] = []
    if S.get("_loaded"):
        return
    S["nama_proyek"] = ""
    S["tgl_mulai"] = date.today()
    S["tgl_acuan"] = date.today()
    S["durasi_total"] = None
    S["toleransi_warning"] = 2.0
    S["batas_critical"] = 10.0
    S["mode_data"] = MODE_MANUAL
    S["m_target"] = None
    S["m_real"] = None
    S["m_sisa"] = None
    S["data_pekerjaan"] = make_items()
    S["riwayat"] = make_riwayat()
    S["nav"] = "dashboard"
    S["_has_url_data"] = False
    try:
        blob = st.query_params.get("d")
        S["_has_url_data"] = bool(blob)
        if blob:
            apply_loaded(decode_state(blob))
        p = st.query_params.get("p")
        if p in NAV:
            S["nav"] = p
    except Exception:
        pass
    S["_loaded"] = True


def save_state():
    try:
        blob = encode_state()
    except Exception:
        return
    if len(blob) > 8000:
        return
    if st.query_params.get("d") != blob or st.query_params.get("p") != S["nav"]:
        st.query_params.from_dict({"d": blob, "p": S["nav"]})


def reset_all():
    try:
        st.query_params.clear()
    except Exception:
        pass
    for k in list(S.keys()):
        del S[k]


def set_nav(k):
    S["nav"] = k


def _history_state(source):
    state = decode_state(encode_state())
    if source == "Mode Kurva S":
        state["mode_data"] = MODE_AUTO
    return state


def save_check_history(m: dict, source=None):
    source = source or ("Mode Kurva S" if m.get("auto") else "Mode manual")
    checked_at = datetime.now()
    record = {
        "Waktu Pengecekan": f"{tgl_id(checked_at.date())}, {checked_at:%H:%M} WIB",
        "Nama Proyek": m["nama"],
        "Sumber Data": source,
        "Target Progres (%)": round(float(m["target"]), 4),
        "Realisasi Progres (%)": round(float(m["real"]), 4),
        "Deviasi (%)": round(float(m["deviasi"]), 4),
        "Sisa Waktu (hari)": _py(m["sisa"]),
        "Status Kesehatan Proyek": m["status"],
        "_state": _history_state(source),
    }
    keys = [key for key in record if key not in ("Waktu Pengecekan", "_state")]
    history = S.setdefault("riwayat_cek", [])
    if history and all(history[0].get(key) == record.get(key) for key in keys):
        return False
    history.insert(0, record)
    del history[5000:]
    try:
        HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError as exc:
        st.warning(f"Riwayat belum dapat disimpan ke file lokal: {exc}")
    return True


def load_history_entry(record):
    state = record.get("_state")
    if isinstance(state, dict):
        apply_loaded(state)
    else:
        S["nama_proyek"] = record.get("Nama Proyek", "")
        S["m_target"] = record.get("Target Progres (%)")
        S["m_real"] = record.get("Realisasi Progres (%)")
        S["m_sisa"] = record.get("Sisa Waktu (hari)")
        S["mode_data"] = MODE_MANUAL
    S["nav"] = "dashboard"
    for key in ("m_target", "m_real", "m_sisa"):
        widget_key = f"w_{key}"
        if widget_key in S:
            S[widget_key] = S[key]
    S["_ver"] = S.get("_ver", 0) + 1
    S["_base_items"] = S["data_pekerjaan"].copy()
    S["_base_riw"] = S["riwayat"].copy()


# ----------------------------------------------------------------------------
# PERHITUNGAN
# ----------------------------------------------------------------------------
def clean_items(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df[COL_NAMA] = df[COL_NAMA].fillna("").astype(str)
    for c in [COL_BOBOT, COL_MULAI, COL_DURASI, COL_REAL]:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    df = df[(df[COL_NAMA].str.strip() != "") | (df[COL_BOBOT] > 0)]
    return df.reset_index(drop=True)


def clean_riwayat(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df[COL_HARI] = pd.to_numeric(df[COL_HARI], errors="coerce")
    df[COL_RKUM] = pd.to_numeric(df[COL_RKUM], errors="coerce")
    return df.dropna().sort_values(COL_HARI).reset_index(drop=True)


def planned_curve(items: pd.DataFrame, total_days: int):
    days = np.arange(0, total_days + 1)
    curve = np.zeros(len(days), dtype=float)
    for _, r in items.iterrows():
        mulai, durasi, bobot = float(r[COL_MULAI]), max(float(r[COL_DURASI]), 1.0), float(r[COL_BOBOT])
        curve += np.clip((days - (mulai - 1)) / durasi, 0, 1) * bobot
    total_bobot = items[COL_BOBOT].sum()
    if total_bobot > 0:
        curve = curve / total_bobot * 100
    return days, curve


def actual_total(items: pd.DataFrame) -> float:
    total_bobot = items[COL_BOBOT].sum()
    if total_bobot <= 0:
        return 0.0
    return float((items[COL_BOBOT] * items[COL_REAL] / 100).sum() / total_bobot * 100)


def hari_acuan() -> int:
    total = int(S["durasi_total"] or 0)
    return int(np.clip((S["tgl_acuan"] - S["tgl_mulai"]).days, 0, total))


def hitung_status(target, real, sisa, toleransi, batas):
    deviasi = real - target
    if real >= 100:
        status = "On Track"
    elif sisa <= 0:
        status = "Delayed"
    elif deviasi >= -toleransi:
        status = "On Track"
    elif deviasi >= -batas:
        status = "At Risk"
    else:
        status = "Delayed"
    laju = (100 - real) / sisa if sisa > 0 else None
    return deviasi, status, laju


def get_metrics() -> dict:
    total = int(S["durasi_total"] or 0)
    items = clean_items(S["data_pekerjaan"])
    bobot_sum = float(items[COL_BOBOT].sum()) if len(items) else 0.0
    has_curve = total > 0 and bobot_sum > 0
    if has_curve:
        days, plan = planned_curve(items, total)
    else:
        days, plan = np.array([0]), np.array([0.0])

    auto = S["mode_data"] == MODE_AUTO
    if auto:
        ready = has_curve
        elapsed = hari_acuan() if has_curve else None
        target = float(np.interp(elapsed, days, plan)) if has_curve else 0.0
        real = actual_total(items)
        sisa = total - (elapsed or 0)
    else:
        vals = [S["m_target"], S["m_real"], S["m_sisa"]]
        ready = all(v is not None for v in vals)
        if ready:
            target, real, sisa = float(S["m_target"]), float(S["m_real"]), int(S["m_sisa"])
            elapsed = total - sisa if (total > 0 and 0 <= sisa <= total) else None
        else:
            target, real, sisa, elapsed = 0.0, 0.0, 0, None

    if ready:
        deviasi, status, laju = hitung_status(target, real, sisa, S["toleransi_warning"], S["batas_critical"])
    else:
        deviasi, status, laju = 0.0, None, None
    laju_rata = real / elapsed if (ready and elapsed) else None

    return {
        "ready": ready, "auto": auto, "has_curve": has_curve,
        "nama": S["nama_proyek"].strip() or "Proyek Tanpa Nama",
        "target": target, "real": real, "sisa": sisa, "elapsed": elapsed, "total": total,
        "deviasi": deviasi, "status": status, "laju": laju, "laju_rata": laju_rata,
        "days": days, "plan": plan, "items": items,
        "riwayat": clean_riwayat(S["riwayat"]), "mode": S["mode_data"],
    }


def rekomendasi(m: dict) -> str:
    if m["real"] >= 100:
        return "Pekerjaan sudah selesai 100%."
    if m["sisa"] <= 0:
        return ("Waktu kontrak sudah habis tetapi pekerjaan belum selesai. Segera ajukan addendum "
                "atau percepatan, dan evaluasi sumber daya.")
    teks = f"Agar selesai tepat waktu, progres minimal yang harus dikejar adalah {m['laju']:.2f}% per hari."
    if m["laju_rata"]:
        rasio = m["laju"] / m["laju_rata"]
        if rasio > 1.05:
            teks += (f" Itu sekitar {rasio:.1f}x lebih cepat dari laju rata-rata sejauh ini "
                     f"({m['laju_rata']:.2f}%/hari).")
        else:
            teks += f" Laju rata-rata sejauh ini ({m['laju_rata']:.2f}%/hari) sudah cukup untuk mengejarnya."
    if m["status"] == "At Risk":
        teks += " Disarankan menambah tenaga kerja atau alat, atau lembur pada pekerjaan kritis."
    elif m["status"] == "Delayed":
        teks += (" Perlu tindakan percepatan segera (tambah sumber daya, shift kerja, urutan ulang pekerjaan) "
                 "dan laporkan ke manajemen.")
    return teks


def tabel_per_item(items: pd.DataFrame, hari) -> pd.DataFrame:
    cols = [COL_NAMA, COL_BOBOT, "Rencana saat ini (%)", COL_REAL, "Deviasi (poin)", "Status"]
    if hari is None or len(items) == 0:
        return pd.DataFrame(columns=cols)
    rows = []
    for _, r in items.iterrows():
        mulai, durasi = float(r[COL_MULAI]), max(float(r[COL_DURASI]), 1.0)
        rencana = float(np.clip((hari - (mulai - 1)) / durasi, 0, 1)) * 100
        real = float(r[COL_REAL])
        dev = real - rencana
        status = "On Track" if dev >= -5 else ("At Risk" if dev >= -15 else "Delayed")
        rows.append({
            COL_NAMA: r[COL_NAMA] or "(tanpa nama)", COL_BOBOT: float(r[COL_BOBOT]),
            "Rencana saat ini (%)": round(rencana, 1), COL_REAL: round(real, 1),
            "Deviasi (poin)": round(dev, 1), "Status": status,
        })
    return pd.DataFrame(rows, columns=cols)


def kurva_s_df(m: dict) -> pd.DataFrame:
    df = pd.DataFrame({"Hari": m["days"], "Rencana (%)": np.round(m["plan"], 2)})
    r = m["riwayat"]
    if len(r) >= 2:
        aktual = np.interp(m["days"], r[COL_HARI], r[COL_RKUM])
        aktual = np.where(m["days"] <= r[COL_HARI].max(), aktual, np.nan)
        df["Realisasi (%)"] = np.round(aktual, 2)
    else:
        df["Realisasi (%)"] = np.nan
    return df


def hari_untuk_detail(m: dict):
    return hari_acuan() if (m["auto"] and m["has_curve"]) else m["elapsed"]


# ----------------------------------------------------------------------------
# EXPORT
# ----------------------------------------------------------------------------
def safe(text) -> str:
    return str(text).encode("latin-1", "replace").decode("latin-1")


def build_excel(m: dict) -> bytes:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Ringkasan"
    items_sheet = workbook.create_sheet("Item Pekerjaan")
    curve_sheet = workbook.create_sheet("Kurva S")
    border_side = Side(style="thin", color="CBD5E1")
    cell_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
    header_fill = PatternFill("solid", fgColor="DBEAFE")
    header_font = Font(bold=True, color="1E3A8A")

    def format_sheet(sheet, header_row, max_row, max_col):
        for cell in sheet[header_row]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for row in sheet.iter_rows(min_row=1, max_row=max_row, max_col=max_col):
            for cell in row:
                cell.border = cell_border
                if cell.row != header_row:
                    cell.alignment = Alignment(vertical="top", wrap_text=True)
                    if cell.column > 1:
                        cell.number_format = "0.00"
        for column_cells in sheet.columns:
            column_letter = get_column_letter(column_cells[0].column)
            max_length = max((len(str(cell.value)) for cell in column_cells if cell.value is not None), default=0)
            sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 14), 52)
        sheet.freeze_panes = f"A{header_row + 1}"
        sheet.auto_filter.ref = f"A{header_row}:{sheet.cell(max_row, max_col).coordinate}"

    clean = clean_items(m["items"])
    item_headers = [COL_NAMA, COL_BOBOT, COL_MULAI, COL_DURASI, COL_REAL, "Kontribusi / Bobot Realisasi (%)"]
    items_sheet.append(item_headers)
    for row_index, (_, row) in enumerate(clean.iterrows(), start=2):
        items_sheet.append([
            str(row[COL_NAMA]), float(row[COL_BOBOT]), float(row[COL_MULAI]),
            float(row[COL_DURASI]), float(row[COL_REAL]),
            f"=IFERROR(B{row_index}*E{row_index}/$B${len(clean) + 2},0)",
        ])
    item_total_row = len(clean) + 2
    if len(clean):
        items_sheet.cell(item_total_row, 1, "Total")
        items_sheet.cell(item_total_row, 2, f"=SUM(B2:B{item_total_row - 1})")
        items_sheet.cell(item_total_row, 6, f"=SUM(F2:F{item_total_row - 1})")
    else:
        items_sheet.cell(item_total_row, 1, "Total")
        items_sheet.cell(item_total_row, 2, "=0")
        items_sheet.cell(item_total_row, 6, "=0")
    for cell in items_sheet[item_total_row]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="F1F5F9")
    format_sheet(items_sheet, 1, item_total_row, len(item_headers))

    total_days = max(int(m["total"] or 0), 0)
    days = np.arange(0, total_days + 1) if total_days else np.array([0])
    if m["has_curve"]:
        plan = np.asarray(m["plan"], dtype=float)
    else:
        plan = np.zeros(len(days), dtype=float)
    plan_increments = np.diff(np.r_[0.0, plan])
    riwayat = m["riwayat"]
    actual_values = None
    last_actual_day = None
    if len(riwayat) >= 2:
        actual_values = np.interp(days, riwayat[COL_HARI], riwayat[COL_RKUM])
        last_actual_day = float(riwayat[COL_HARI].max())
    curve_headers = ["Hari", "Rencana Harian (%)", "Rencana Kumulatif (%)",
                     "Realisasi Harian (%)", "Realisasi Kumulatif (%)", "Deviasi Kumulatif (%)"]
    curve_sheet.append(curve_headers)
    for index, day in enumerate(days):
        row_number = index + 2
        plan_increment = float(plan_increments[index]) if index < len(plan_increments) else 0.0
        if actual_values is not None and day <= last_actual_day:
            actual_increment = float(actual_values[index] - (actual_values[index - 1] if index else 0.0))
        else:
            actual_increment = None
        curve_sheet.append([
            int(day), plan_increment, f"=SUM($B$2:B{row_number})", actual_increment,
            (f'=IF(A{row_number}>{last_actual_day},"",SUM($D$2:D{row_number}))'
             if last_actual_day is not None else f"=SUM($D$2:D{row_number})"),
            f'=IF(E{row_number}="","",E{row_number}-C{row_number})',
        ])
    format_sheet(curve_sheet, 1, len(days) + 1, len(curve_headers))
    for row_number in range(2, len(days) + 2):
        curve_sheet.cell(row_number, 3).number_format = "0.00"
        curve_sheet.cell(row_number, 5).number_format = "0.00"
        curve_sheet.cell(row_number, 6).number_format = "0.00"

    summary.merge_cells("A1:B1")
    summary["A1"] = "Laporan Analisis Risiko Keterlambatan Proyek"
    summary["A1"].font = Font(bold=True, size=14, color="0F172A")
    summary["A1"].alignment = Alignment(vertical="center")
    summary.append([])
    summary.append(["Parameter", "Nilai"])
    source = "Mode Kurva S" if m["auto"] else "Mode manual"
    rows = [
        ("Nama Proyek", m["nama"]),
        ("Waktu Pengecekan", datetime.now()),
        ("Sumber Data", source),
        ("Tanggal Mulai Proyek", S["tgl_mulai"]),
        ("Tanggal Acuan", S["tgl_acuan"]),
        ("Durasi Kontrak (hari)", m["total"] or 0),
        ("Target Progres (%)", f"='Kurva S'!C{int(hari_untuk_detail(m) or 0) + 2}" if m["auto"] else float(m["target"])),
        ("Realisasi Progres (%)", f"='Item Pekerjaan'!F{item_total_row}" if m["auto"] else float(m["real"])),
        ("Deviasi (%)", "=B11-B10"),
        ("Sisa Waktu Proyek (hari)", '=IF(B6="Mode manual",B20,MAX(0,B9-MAX(0,MIN(B9,B8-B7))))'),
        ("Status Kesehatan Proyek", '=IF(B11>=100,"On Track",IF(B13<=0,"Delayed",IF(B12>=-B15,"On Track",IF(B12>=-B16,"At Risk","Delayed"))))'),
        ("Toleransi Warning (poin)", float(S["toleransi_warning"])),
        ("Batas Delayed (poin)", float(S["batas_critical"])),
        ("Laju Minimal (%/hari)", '=IF(B13>0,(100-B11)/B13,"")'),
        ("Laju Rata-rata (%/hari)", '=IF(B6="Mode manual",IF(B9-B13>0,B11/(B9-B13),""),IF(MAX(0,MIN(B9,B8-B7))>0,B11/MAX(0,MIN(B9,B8-B7)),""))'),
        ("Rekomendasi", rekomendasi(m)),
        ("Sisa Waktu Manual (hari)", int(m["sisa"] or 0)),
    ]
    for label, value in rows:
        summary.append([label, value])
    summary["B5"].number_format = "dd mmm yyyy hh:mm"
    summary["B7"].number_format = "dd mmm yyyy"
    summary["B8"].number_format = "dd mmm yyyy"
    summary.row_dimensions[20].hidden = True
    summary["B14"].fill = PatternFill("solid", fgColor=STATUS[m["status"]]["color"].lstrip("#"))
    summary["B14"].font = Font(bold=True, color="FFFFFF")
    summary["B14"].alignment = Alignment(horizontal="center")
    format_sheet(summary, 3, 20, 2)

    chart = LineChart()
    chart.title = "Kurva S"
    chart.x_axis.title = "Hari ke"
    chart.y_axis.title = "Progres kumulatif (%)"
    chart.y_axis.scaling.min = 0
    chart.y_axis.scaling.max = 100
    chart.add_data(Reference(curve_sheet, min_col=3, max_col=3, min_row=1, max_row=len(days) + 1), titles_from_data=True)
    chart.add_data(Reference(curve_sheet, min_col=5, max_col=5, min_row=1, max_row=len(days) + 1), titles_from_data=True)
    chart.set_categories(Reference(curve_sheet, min_col=1, min_row=2, max_row=len(days) + 1))
    chart.series[0].graphicalProperties.line.solidFill = "2563EB"
    chart.series[1].graphicalProperties.line.solidFill = "0F172A"
    chart.width = 18
    chart.height = 9
    curve_sheet.add_chart(chart, "H2")
    for sheet in workbook.worksheets:
        sheet.sheet_view.showGridLines = False
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    workbook.calculation.calcMode = "auto"
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def build_pdf(m: dict) -> bytes:
    detail = tabel_per_item(m["items"], hari_untuk_detail(m))
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Laporan Analisis Risiko Keterlambatan Proyek", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 6, safe(f"Proyek: {m['nama']}"), new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 6, f"Tanggal laporan: {datetime.now().strftime('%d-%m-%Y %H:%M')}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    pdf.set_fill_color(*STATUS[m["status"]]["rgb"])
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 12, safe(f"STATUS: {m['status'].upper()}"), fill=True, align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.set_text_color(0, 0, 0)
    pdf.ln(3)

    baris = [
        ("Target progres", f"{m['target']:.1f} %"),
        ("Realisasi progres", f"{m['real']:.1f} %"),
        ("Deviasi", f"{m['deviasi']:+.1f} poin"),
        ("Sisa waktu", f"{m['sisa']} hari"),
        ("Laju minimal yang harus dikejar", f"{m['laju']:.2f} %/hari" if m["laju"] is not None else "-"),
        ("Laju rata-rata sejauh ini", f"{m['laju_rata']:.2f} %/hari" if m["laju_rata"] is not None else "-"),
    ]
    for k, v in baris:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(80, 7, k, border=1)
        pdf.set_font("Helvetica", "", 11)
        pdf.cell(60, 7, v, border=1, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 6, "Rekomendasi", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 5, safe(rekomendasi(m)), new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)

    if m["has_curve"]:
        fig, ax = plt.subplots(figsize=(7.5, 3.3))
        ax.plot(m["days"], m["plan"], color=C_PLAN, label="Rencana", linewidth=2)
        r = m["riwayat"]
        if len(r) > 0:
            ax.plot(r[COL_HARI], r[COL_RKUM], color=C_INK, marker="o", label="Realisasi", linewidth=2)
        if m["elapsed"] is not None:
            ax.axvline(m["elapsed"], color="#CBD5E1", linestyle="--", linewidth=1)
            ax.scatter([m["elapsed"]], [m["real"]], color=C_INK, zorder=5, label="Posisi saat ini")
        ax.set_xlabel("Hari ke")
        ax.set_ylabel("Progres kumulatif (%)")
        ax.set_ylim(0, 100)
        ax.grid(alpha=0.3)
        ax.legend(loc="upper left")
        buf = io.BytesIO()
        fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
        plt.close(fig)
        buf.seek(0)
        pdf.image(buf, w=180)
        pdf.ln(2)

    if len(detail) > 0:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 7, "Detail Progres per Item Pekerjaan", new_x="LMARGIN", new_y="NEXT")
        lebar = [60, 22, 32, 28, 26, 22]
        judul_kol = ["Item", "Bobot %", "Rencana %", "Real. %", "Dev.", "Status"]
        pdf.set_font("Helvetica", "B", 9)
        pdf.set_fill_color(241, 245, 249)
        for w_, t in zip(lebar, judul_kol):
            pdf.cell(w_, 7, t, border=1, fill=True)
        pdf.ln()
        pdf.set_font("Helvetica", "", 9)
        for _, r_ in detail.iterrows():
            vals = [safe(r_[COL_NAMA])[:32], f"{r_[COL_BOBOT]:.1f}", f"{r_['Rencana saat ini (%)']:.1f}",
                    f"{r_[COL_REAL]:.1f}", f"{r_['Deviasi (poin)']:+.1f}", r_["Status"]]
            for w_, t in zip(lebar, vals):
                pdf.cell(w_, 6, t, border=1)
            pdf.ln()
    return bytes(pdf.output())


# ----------------------------------------------------------------------------
# KOMPONEN TAMPILAN
# ----------------------------------------------------------------------------
def page_header(title, sub="", right=""):
    st.markdown(f'<div class="ph"><div><div class="ph-t">{esc(title)}</div><div class="ph-s">{esc(sub)}</div></div>'
                f'<div>{right}</div></div>', unsafe_allow_html=True)


def meta_line(parts):
    html_ = '<i></i>'.join(f"<span>{p}</span>" for p in parts)
    st.markdown(f'<div class="meta">{html_}</div>', unsafe_allow_html=True)


def page_nav():
    cols = st.columns(len(NAV))
    for (key, label), col in zip(NAV.items(), cols):
        col.button(label, key=f"top_nav_{key}", on_click=set_nav, args=(key,),
                   type="tertiary", width="stretch")


def welcome_screen():
    image_types = {".jpg": "image/jpeg", ".jpeg": "image/jpeg", ".png": "image/png", ".webp": "image/webp"}
    image_path = next(
        (path for path in sorted(Path(__file__).resolve().parent.iterdir())
         if path.is_file() and path.suffix.lower() in image_types),
        None,
    )
    if image_path:
        image_data = base64.b64encode(image_path.read_bytes()).decode("ascii")
        background_css = (
            "background-image:linear-gradient(rgba(248,250,252,0.80),rgba(248,250,252,0.88)),"
            f"url('data:{image_types[image_path.suffix.lower()]};base64,{image_data}');"
            "background-size:cover;background-position:center;background-repeat:no-repeat;"
        )
    else:
        background_css = "background:#F8FAFC;"

    with st.container(key="welcome_content"):
        welcome_html = """
<style>
section[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"],
[data-testid="stExpandSidebarButton"] { display:none !important; }
.block-container { max-width:none; min-height:100vh; display:flex; align-items:center; padding:0 2rem; }
.st-key-welcome_content { width:100%; min-height:82vh; display:flex; flex-direction:column; align-items:center;
    justify-content:center; text-align:center; position:relative; isolation:isolate; overflow:hidden; padding:24px; box-sizing:border-box;
    __WELCOME_BACKGROUND__ }
.welcome-copy { position:relative; z-index:1; width:min(100%,700px); margin:0 auto; text-align:center; }
.welcome-title { color:#0F172A; font-size:40px; font-weight:800; letter-spacing:.08em; line-height:1.2; margin:0; }
.welcome-subtitle { color:#475569; font-size:16px; line-height:1.6; margin-top:14px; }
.st-key-start_monitoring { position:relative; z-index:1; display:flex; justify-content:center; margin-top:22px; }
.st-key-start_monitoring button { background:#1E3A8A !important; border:1px solid #1E3A8A !important;
    color:#FFFFFF !important; min-height:0; padding:10px 36px; border-radius:8px; font-size:16px; font-weight:600; }
.st-key-start_monitoring button p { color:#FFFFFF !important; font-size:16px; font-weight:600; margin:0; }
.st-key-start_monitoring button:hover { background:#0F172A !important; border-color:#0F172A !important; }
@media (max-width:640px) {
    .block-container { padding:0 1rem; }
    .st-key-welcome_content { padding:20px 8px; }
    .welcome-title { font-size:30px; letter-spacing:.06em; }
    .welcome-subtitle { font-size:15px; }
}
</style>
<div class="welcome-copy">
<div class="welcome-title">MONITORING PROYEK</div>
<div class="welcome-subtitle">Early warning keterlambatan proyek</div>
</div>
        """
        st.markdown(welcome_html.replace("__WELCOME_BACKGROUND__", background_css), unsafe_allow_html=True)
        if st.button("Mulai", key="start_monitoring", type="primary"):
            S["_welcome_started"] = True
            st.rerun()


def kpis(cells, status=None):
    out = ""
    for label, value, unit, desc in cells:
        if label.lower().startswith("target"):
            accent, soft = "#2563EB", "#EFF6FF"
        elif label.lower().startswith("realisasi"):
            accent, soft = "#0F172A", "#F1F5F9"
        elif label.lower() == "sisa waktu":
            accent = STATUS.get(status, {}).get("color", "#1E3A8A")
            soft = {"On Track": "#F0FDF4", "At Risk": "#FFFBEB", "Delayed": "#FEF2F2"}.get(status, "#EFF6FF")
        elif "laju" in label.lower():
            accent, soft = "#1E3A8A", "#EFF6FF"
        else:
            accent, soft = "#64748B", "#F1F5F9"
        u = f'<span class="u">{unit}</span>' if unit else ""
        out += (f'<div class="kpi" style="--kpi-accent:{accent};--kpi-soft:{soft}">'
                f'<div class="l">{label}</div><div class="v">{value}{u}</div><div class="d">{desc}</div></div>')
    st.markdown(f'<div class="kpis">{out}</div>', unsafe_allow_html=True)


def empty_state(judul, teks):
    st.markdown(f'<div class="empty"><span class="ic">{svg("clipboard", 18)}</span>'
                f'<div class="e2">{esc(judul)}</div><div class="e3">{esc(teks)}</div></div>', unsafe_allow_html=True)


def section(judul, desk=""):
    st.markdown(f'<div class="sec">{esc(judul)}</div><div class="sec-d">{esc(desk)}</div>', unsafe_allow_html=True)


def note(teks):
    st.markdown(f'<div class="note">{esc(teks)}</div>', unsafe_allow_html=True)


def style_fig(fig, height=380):
    fig.update_layout(
        template="plotly_white", height=height, margin=dict(t=16, b=40, l=48, r=16),
        font=dict(family="Inter, system-ui, sans-serif", size=12, color="#475569"),
        paper_bgcolor="#FFFFFF", plot_bgcolor="#FFFFFF",
        legend=dict(orientation="h", y=-0.2, x=0),
        hoverlabel=dict(font_family="Inter, sans-serif", bgcolor="#FFFFFF", bordercolor="#CBD5E1"),
    )
    fig.update_xaxes(showline=True, linecolor="#94A3B8", showgrid=False, gridcolor="#E2E8F0", zeroline=False)
    fig.update_yaxes(showline=True, linecolor="#94A3B8", gridcolor="#E2E8F0", zeroline=False)
    return fig


def progress_card(m: dict):
    t, r = min(max(m["target"], 0), 100), min(max(m["real"], 0), 100)
    dev = f"{m['deviasi']:+.1f}".replace(".", ",")
    st.markdown(
        f'<div class="card"><div class="card-h">Progres terhadap rencana</div>'
        f'<div class="prow"><span class="pn">Target</span><div class="trk"><div class="fil fil-plan" style="width:{t}%"></div></div>'
        f'<span class="pv">{fmt(m["target"])}%</span></div>'
        f'<div class="prow"><span class="pn">Realisasi</span><div class="trk"><div class="fil fil-act" style="width:{r}%"></div></div>'
        f'<span class="pv">{fmt(m["real"])}%</span></div>'
        f'<div class="pfoot">{badge(m["status"])}<span>{esc(STATUS[m["status"]]["msg"])} Deviasi {dev} poin.</span></div></div>',
        unsafe_allow_html=True)


def rec_card(m: dict):
    status_class = STATUS.get(m["status"], {}).get("cls", "b-idle")
    st.markdown(f'<div class="card rec-card {status_class}"><div class="card-h">Rekomendasi</div>'
                f'<p class="rec">{esc(rekomendasi(m))}</p></div>',
                unsafe_allow_html=True)


def item_table(detail: pd.DataFrame):
    rows = ""
    for _, r in detail.iterrows():
        real = min(max(float(r[COL_REAL]), 0), 100)
        rows += (f'<tr><td>{esc(str(r[COL_NAMA]))}</td><td>{fmt(r[COL_BOBOT])}</td><td>{fmt(r["Rencana saat ini (%)"])}</td>'
                 f'<td><div class="mini"><div class="trk"><div class="fil fil-act" style="width:{real}%"></div></div>'
                 f'<span>{fmt(real)}</span></div></td><td>{fmt(r["Deviasi (poin)"])}</td><td>{badge(r["Status"])}</td></tr>')
    st.markdown(
        '<div class="card tbl-wrap"><table class="tbl"><thead><tr><th>Item pekerjaan</th><th>Bobot (%)</th>'
        '<th>Rencana (%)</th><th>Realisasi (%)</th><th>Deviasi (poin)</th><th>Status</th></tr></thead>'
        f'<tbody>{rows}</tbody></table></div>', unsafe_allow_html=True)


def pnum(label, store, **kw):
    """number_input yang nilainya tetap ada saat berpindah halaman."""
    w = "w_" + store
    if w not in S:
        S[w] = S[store]

    def _sync():
        S[store] = S[w]

    st.number_input(label, key=w, on_change=_sync, value=None, **kw)


# ----------------------------------------------------------------------------
# SIDEBAR
# ----------------------------------------------------------------------------
def sidebar():
    return


# ----------------------------------------------------------------------------
# HALAMAN
# ----------------------------------------------------------------------------
def meta_parts():
    parts = [f'Mulai <b>{tgl_id(S["tgl_mulai"])}</b>', f'Acuan <b>{tgl_id(S["tgl_acuan"])}</b>']
    if S["durasi_total"]:
        parts.insert(1, f'Kontrak <b>{int(S["durasi_total"])} hari</b>')
    parts.append("Mode manual" if S["mode_data"] == MODE_MANUAL else "Mode otomatis")
    return parts


def project_settings():
    with st.expander("Identitas & Pengaturan Proyek", expanded=True):
        c1, c2 = st.columns(2)
        with c1:
            st.text_input("Nama proyek", key="nama_proyek", placeholder="Contoh: Gedung Kantor 2 Lantai", max_chars=120)
            st.date_input("Tanggal mulai proyek", key="tgl_mulai", format="DD/MM/YYYY")
            st.number_input("Durasi kontrak (hari)", min_value=1, max_value=3650, step=1, value=None,
                            key="durasi_total", placeholder="Contoh: 120")
        with c2:
            st.date_input("Tanggal acuan (hari ini)", key="tgl_acuan", format="DD/MM/YYYY")
            st.radio("Sumber data dashboard", [MODE_MANUAL, MODE_AUTO], key="mode_data")
            c_warning, c_critical = st.columns(2)
            c_warning.number_input("Toleransi (poin)", min_value=0.0, max_value=100.0, step=0.5,
                                   key="toleransi_warning", help="Deviasi sampai -X poin masih dianggap On Track.")
            c_critical.number_input("Batas Delayed", min_value=0.0, max_value=100.0, step=0.5,
                                    key="batas_critical", help="Deviasi lebih buruk dari -X poin dianggap Delayed.")
        st.button("Kosongkan semua data", key="reset", on_click=reset_all, type="secondary")
        st.caption("Data tersimpan otomatis di alamat halaman (URL). Simpan atau bookmark link ini untuk membuka data yang sama.")


def page_dashboard():
    nama = S["nama_proyek"].strip() or "Proyek tanpa nama"
    m = get_metrics()
    page_header(nama, "Ringkasan status kesehatan proyek", badge(m["status"] if m["ready"] else None))
    meta_line(meta_parts())
    page_nav()
    project_settings()

    if S["mode_data"] == MODE_MANUAL:
        with st.container(key="card_manual"):
            section("Input data", "Isi 3 data berikut untuk melihat status proyek.")
            c1, c2, c3 = st.columns(3)
            with c1:
                pnum("Target progres saat ini (%)", "m_target", min_value=0.0, max_value=100.0, step=0.5, placeholder="Contoh: 32")
            with c2:
                pnum("Realisasi progres saat ini (%)", "m_real", min_value=0.0, max_value=100.0, step=0.5, placeholder="Contoh: 29")
            with c3:
                pnum("Sisa waktu proyek (hari)", "m_sisa", min_value=0, max_value=3650, step=1, placeholder="Contoh: 75")
        m = get_metrics()
    else:
        note("Mode otomatis: target dihitung dari Kurva S rencana, realisasi dari tabel item, dan sisa waktu dari durasi "
             "kontrak. Isi datanya di menu Detail Kurva S dan durasi di Pengaturan proyek.")

    if not m["ready"]:
        if m["auto"]:
            empty_state("Data belum lengkap", "Isi durasi kontrak di Pengaturan proyek, lalu isi bobot, hari mulai, dan durasi "
                        "tiap item di menu Detail Kurva S.")
        else:
            empty_state("Belum ada data", "Isi target progres, realisasi progres, dan sisa waktu di atas. "
                        "Status proyek akan muncul di sini.")
        return

    kpis([
        ("Target", fmt(m["target"]), "%", "Seharusnya sudah tercapai"),
        ("Realisasi", fmt(m["real"]), "%", f"Deviasi {m['deviasi']:+.1f} poin".replace(".", ",")),
        ("Sisa waktu", f"{m['sisa']}", "hari", "Hingga akhir kontrak"),
        ("Laju minimal", fmt(m["laju"], 2), "%/hari" if m["laju"] is not None else "", "Agar selesai tepat waktu"),
        ("Laju rata-rata", fmt(m["laju_rata"], 2), "%/hari" if m["laju_rata"] is not None else "", "Sejauh ini"),
    ], status=m["status"])
    save_check_history(m)
    kiri, kanan = st.columns(2)
    with kiri:
        progress_card(m)
    with kanan:
        rec_card(m)
    if st.button("Simpan ke Riwayat", key="save_check_dashboard"):
        if save_check_history(m):
            st.success("Pengecekan tersimpan di riwayat.")
        else:
            st.info("Pengecekan ini sudah menjadi entri terbaru.")


def page_kurva_s():
    entering = S.get("_prev_page") != "kurva"
    if entering:
        S["_ver"] = S.get("_ver", 0) + 1
        S["_base_items"] = S["data_pekerjaan"].copy()
        S["_base_riw"] = S["riwayat"].copy()
    ver = S["_ver"]

    page_header("Detail Kurva S", "Progres per item pekerjaan, dari rencana sampai realisasi")
    meta_line(meta_parts())
    page_nav()
    project_settings()

    with st.container(key="card_items"):
        section("Item pekerjaan", "Isi bobot, hari mulai, durasi, dan realisasi tiap item. Klik baris kosong paling bawah untuk menambah.")
        with st.expander("Import / Tempel Data dari Excel RAB"):
            paste_text = st.text_area(
                "Tempel kolom Nama Pekerjaan, Bobot (%), Hari Mulai, Durasi (hari), dan opsional Realisasi (%)",
                height=130, key="rab_paste_text",
            )
            uploaded_rab = st.file_uploader("Atau unggah file RAB", type=["xlsx", "xls", "csv"], key="rab_upload")
            if st.button("Muat Data ke Tabel", key="load_rab", type="primary"):
                try:
                    if uploaded_rab is not None:
                        if uploaded_rab.name.lower().endswith(".csv"):
                            uploaded_text = uploaded_rab.getvalue().decode("utf-8-sig")
                            imported_rows, invalid_rows = parse_rab_text(uploaded_text)
                        else:
                            imported_frame = pd.read_excel(uploaded_rab, header=None, dtype=object)
                            imported_rows, invalid_rows = parse_rab_rows(imported_frame.values.tolist())
                    elif paste_text.strip():
                        imported_rows, invalid_rows = parse_rab_text(paste_text)
                    else:
                        imported_rows, invalid_rows = [], []
                        st.warning("Tempel data atau pilih file RAB terlebih dahulu.")

                    if imported_rows:
                        S["_base_items"] = make_items(imported_rows)
                        S["data_pekerjaan"] = S["_base_items"].copy()
                        S["_ver"] += 1
                        ver = S["_ver"]
                        st.success(f"{len(imported_rows)} item berhasil dimuat ke tabel.")
                    elif uploaded_rab is not None or paste_text.strip():
                        st.error("Tidak ada baris data valid. Periksa urutan kolom dan nilai bobot/hari.")
                    if invalid_rows:
                        st.warning(f"{len(invalid_rows)} baris dilewati karena format atau nilainya tidak valid.")
                except Exception as exc:
                    st.error(f"Gagal membaca data RAB: {exc}")

        edited = st.data_editor(
            S["_base_items"], num_rows="dynamic", width="stretch", hide_index=True, key=f"ed_items_{ver}",
            column_config={
                COL_NAMA: st.column_config.TextColumn(COL_NAMA),
                COL_BOBOT: st.column_config.NumberColumn(COL_BOBOT, min_value=0.0, max_value=100.0, step=0.5),
                COL_MULAI: st.column_config.NumberColumn(COL_MULAI, min_value=1, step=1, format="%d"),
                COL_DURASI: st.column_config.NumberColumn(COL_DURASI, min_value=1, step=1, format="%d"),
                COL_REAL: st.column_config.NumberColumn(COL_REAL, min_value=0.0, max_value=100.0, step=1.0),
            },
        )
    S["data_pekerjaan"] = edited
    items = clean_items(edited)
    total_bobot = float(items[COL_BOBOT].sum()) if len(items) else 0.0
    if len(items) and abs(total_bobot - 100) > 0.01:
        note(f"Total bobot saat ini {total_bobot:.1f}%. Perhitungan dinormalisasi ke 100%.".replace(".", ",", 1))

    with st.container(key="card_riwayat"):
        section("Riwayat realisasi kumulatif", "Opsional. Dipakai untuk menggambar garis realisasi pada Kurva S.")
        S["riwayat"] = st.data_editor(
            S["_base_riw"], num_rows="dynamic", width="stretch", hide_index=True, key=f"ed_riw_{ver}",
            column_config={
                COL_HARI: st.column_config.NumberColumn(COL_HARI, min_value=0, step=1, format="%d"),
                COL_RKUM: st.column_config.NumberColumn(COL_RKUM, min_value=0.0, max_value=100.0, step=0.5),
            },
        )

    total = int(S["durasi_total"] or 0)
    if total <= 0:
        empty_state("Durasi kontrak belum diisi", "Isi durasi kontrak (hari) di Pengaturan proyek pada panel kiri.")
        return
    if len(items) == 0 or total_bobot <= 0:
        empty_state("Belum ada item pekerjaan", "Isi minimal satu item dengan bobot, hari mulai, dan durasi pada tabel di atas.")
        return

    days, plan = planned_curve(items, total)
    hari = hari_acuan()
    target = float(np.interp(hari, days, plan))
    real = actual_total(items)
    status_kurva = hitung_status(target, real, total - hari, S["toleransi_warning"], S["batas_critical"])[1]
    selisih = real - target
    if abs(selisih) < 0.05:
        kesimpulan = "Realisasi berbobot saat ini sejalan dengan rencana."
    elif selisih > 0:
        kesimpulan = f"Realisasi berbobot saat ini melampaui rencana sebesar {fmt(selisih)} poin persentase."
    else:
        kesimpulan = f"Realisasi berbobot saat ini tertinggal dari rencana sebesar {fmt(abs(selisih))} poin persentase."
    riwayat = clean_riwayat(S["riwayat"])
    detail = tabel_per_item(items, hari)

    kpis([
        ("Hari acuan", f"ke-{hari}", "", f"dari {total} hari"),
        ("Target (rencana)", fmt(target), "%", "Dari Kurva S rencana"),
        ("Realisasi (dari item)", fmt(real), "%", f"Deviasi {real - target:+.1f} poin".replace(".", ",")),
    ], status=status_kurva)
    check = {
        "nama": S["nama_proyek"].strip() or "Proyek Tanpa Nama",
        "target": target, "real": real, "deviasi": selisih,
        "sisa": total - hari, "status": status_kurva,
    }
    save_check_history(check, source="Mode Kurva S")
    if st.button("Simpan ke Riwayat", key="save_check_curve"):
        if save_check_history(check, source="Mode Kurva S"):
            st.success("Pengecekan tersimpan di riwayat.")
        else:
            st.info("Pengecekan ini sudah menjadi entri terbaru.")

    with st.container(key="card_chart"):
        section("Kurva S", "Rencana dibanding realisasi kumulatif.")
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=days, y=plan, name="Rencana", mode="lines+markers",
                                 line=dict(color="#2563EB", width=2.5), marker=dict(size=5),
                                 hovertemplate="Hari %{x}: %{y:.1f}%<extra>Rencana</extra>"))
        if len(riwayat) > 0:
            fig.add_trace(go.Scatter(x=riwayat[COL_HARI], y=riwayat[COL_RKUM], name="Realisasi (riwayat)",
                                     mode="lines+markers", line=dict(color="#0F172A", width=2.5),
                                     line_shape="linear", marker=dict(size=7),
                                     hovertemplate="Hari %{x}: %{y:.1f}%<extra>Realisasi</extra>"))
        fig.add_trace(go.Scatter(x=[hari], y=[real], name="Posisi saat ini", mode="markers",
                                 marker=dict(size=11, color=C_INK, line=dict(color="#FFFFFF", width=2)),
                                 hovertemplate="Hari %{x}: %{y:.1f}%<extra>Saat ini</extra>"))
        fig.add_vline(x=hari, line_dash="dot", line_color="#CBD5E1", annotation_text="Hari acuan",
                      annotation_font=dict(size=11, color="#64748B"))
        fig.update_layout(xaxis_title="Hari ke", yaxis_title="Progres kumulatif (%)", yaxis_range=[0, 100])
        style_fig(fig, 380)
        st.plotly_chart(fig, width="stretch")
        st.markdown(
            f'<div class="read-note">Sumbu X menunjukkan hari proyek; sumbu Y menunjukkan '
            f'progres kumulatif (%). Garis biru adalah rencana, sedangkan garis gelap adalah riwayat realisasi jika tersedia. '
            f'{esc(kesimpulan)}</div>', unsafe_allow_html=True)

    section("Status per item pekerjaan", "Dibandingkan dengan rencana pada hari acuan.")
    item_table(detail)

    with st.container(key="card_bars"):
        bar = go.Figure()
        bar.add_bar(x=detail[COL_NAMA], y=detail["Rencana saat ini (%)"], name="Rencana", marker_color="#2563EB")
        bar.add_bar(x=detail[COL_NAMA], y=detail[COL_REAL], name="Realisasi", marker_color=C_INK)
        bar.update_layout(barmode="group", yaxis_title="Progres item (%)", yaxis_range=[0, 100], bargap=0.35)
        style_fig(bar, 320)
        st.plotly_chart(bar, width="stretch")
        st.markdown(
            f'<div class="read-note">Sumbu X menunjukkan item pekerjaan; sumbu Y menunjukkan '
            f'progres tiap item (%). Bandingkan batang rencana dengan realisasi untuk melihat posisi tiap item. '
            f'{esc(kesimpulan)}</div>', unsafe_allow_html=True)


def page_export():
    m = get_metrics()
    page_header("Export laporan", "Unduh hasil analisis untuk dibagikan ke manajemen atau klien",
                badge(m["status"] if m["ready"] else None))
    meta_line(meta_parts())
    page_nav()
    project_settings()
    if not m["ready"]:
        empty_state("Belum ada data untuk dilaporkan", "Isi data di Dashboard Utama (mode manual) atau Detail Kurva S "
                    "(mode otomatis), lalu kembali ke sini.")
        return

    kpis([
        ("Target", fmt(m["target"]), "%", ""),
        ("Realisasi", fmt(m["real"]), "%", f"Deviasi {m['deviasi']:+.1f} poin".replace(".", ",")),
        ("Sisa waktu", f"{m['sisa']}", "hari", ""),
        ("Laju minimal", fmt(m["laju"], 2), "%/hari" if m["laju"] is not None else "", ""),
    ], status=m["status"])
    rec_card(m)

    nama_file = "".join(ch if ch.isalnum() else "_" for ch in m["nama"]) or "proyek"
    tgl = datetime.now().strftime("%Y%m%d")
    c1, c2 = st.columns(2)
    try:
        c1.download_button("Download PDF", data=build_pdf(m), file_name=f"laporan_{nama_file}_{tgl}.pdf",
                           mime="application/pdf", width="stretch")
    except Exception as e:
        c1.error(f"Gagal membuat PDF: {e}")
    try:
        c2.download_button("Download Excel", data=build_excel(m), file_name=f"laporan_{nama_file}_{tgl}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
    except Exception as e:
        c2.error(f"Gagal membuat Excel: {e}")


def page_riwayat():
    history = S.get("riwayat_cek", [])
    page_header("Riwayat Cek", "Pengecekan proyek tersimpan di perangkat ini")
    page_nav()
    actions = st.columns(3)
    export_rows = [{key: value for key, value in item.items() if key != "_state"} for item in history]
    history_df = pd.DataFrame(export_rows)
    csv_data = history_df.to_csv(index=False).encode("utf-8-sig")
    excel_data = io.BytesIO()
    with pd.ExcelWriter(excel_data, engine="openpyxl") as writer:
        history_df.to_excel(writer, sheet_name="Riwayat Cek", index=False)
    actions[0].download_button(
        "Download Riwayat (CSV)", data=csv_data, file_name="riwayat_proyek.csv",
        mime="text/csv", width="stretch", disabled=not history,
    )
    actions[1].download_button(
        "Download Riwayat (Excel)", data=excel_data.getvalue(), file_name="riwayat_proyek.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        width="stretch", disabled=not history,
    )
    if actions[2].button("Hapus Semua Riwayat", key="delete_all_history", width="stretch",
                         disabled=not history, type="secondary"):
        S["riwayat_cek"] = []
        try:
            HISTORY_FILE.write_text("[]", encoding="utf-8")
        except OSError as exc:
            st.warning(f"Riwayat lokal tidak dapat diperbarui: {exc}")
        st.rerun()

    if not history:
        empty_state("Belum ada riwayat pengecekan", "Riwayat akan tersimpan otomatis saat data proyek siap di Dashboard atau Detail Kurva S.")
        return

    for index, record in enumerate(history):
        status = record.get("Status Kesehatan Proyek")
        status_badge = badge(status if status in STATUS else None)
        with st.container(border=True):
            detail_col, action_col = st.columns([5, 1.5])
            with detail_col:
                st.markdown(
                    f'{status_badge} <strong>{esc(str(record.get("Nama Proyek", "Proyek tanpa nama")))}</strong>'
                    f'<span style="color:#475569"> &middot; {esc(str(record.get("Waktu Pengecekan", "")))}</span>',
                    unsafe_allow_html=True,
                )
                st.caption(
                    f'{esc(str(record.get("Sumber Data", "")))} · '
                    f'Target {fmt(record.get("Target Progres (%)"))}% · '
                    f'Realisasi {fmt(record.get("Realisasi Progres (%)"))}% · '
                    f'Deviasi {fmt(record.get("Deviasi (%)"), suffix=" poin")} · '
                    f'Sisa {record.get("Sisa Waktu (hari)", "-")} hari'
                )
            with action_col:
                if st.button("Muat Ulang Data Ini", key=f"restore_history_{index}", width="stretch"):
                    load_history_entry(record)
                    st.rerun()


def main():
    init_state()
    if not S["_has_url_data"] and not S.get("_welcome_started"):
        welcome_screen()
        return
    st.markdown(CSS, unsafe_allow_html=True)
    sidebar()
    halaman = S["nav"]
    st.markdown(
        f'<style>.st-key-nav_{halaman} button, .st-key-top_nav_{halaman} button '
        f'{{ background:#1E3A8A !important; border-color:#1E3A8A !important; color:#FFFFFF !important; '
        f'box-shadow:inset 3px 0 0 #2563EB !important; }}</style>', unsafe_allow_html=True)
    if halaman == "dashboard":
        page_dashboard()
    elif halaman == "kurva":
        page_kurva_s()
    elif halaman == "export":
        page_export()
    else:
        page_riwayat()
    S["_prev_page"] = halaman
    save_state()


main()
