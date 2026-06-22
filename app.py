
# -*- coding: utf-8 -*-
"""
AYÇA Insight Hospital V0.4.2
Dark Executive Hospital Dashboard Concept

Run:
    pip install -r requirements.txt
    streamlit run app.py

Excel sheets required:
    Gunluk, Brans, Doktor, Hasta, Operasyon, Stok, Kalite
"""

from __future__ import annotations

import math
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st


APP_VERSION = "V0.4.2 Hospital Concept"


# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title="AYÇA Insight Hospital",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# CSS
# =========================================================
st.markdown(
    """
<style>
    .stApp {
        background: radial-gradient(circle at top left, #0f2742 0%, #061422 38%, #030a12 100%);
        color: #e5eefb;
    }

    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2.5rem;
        max-width: 1380px;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #061627 0%, #03111f 100%);
        border-right: 1px solid rgba(148, 163, 184, .16);
    }

    section[data-testid="stSidebar"] * {
        color: #e5eefb;
    }

    div[data-testid="stFileUploader"] {
        background: rgba(15, 23, 42, .42);
        border: 1px solid rgba(148, 163, 184, .18);
        border-radius: 14px;
        padding: 10px;
    }

    .ayca-logo {
        font-size: 34px;
        font-weight: 900;
        letter-spacing: -1.1px;
        color: #f8fafc;
        line-height: 1.05;
        margin-bottom: 0px;
    }

    .ayca-sub {
        font-size: 11px;
        color: #93c5fd;
        letter-spacing: 1px;
        margin-bottom: 8px;
    }

    .version-badge {
        display: inline-block;
        background: linear-gradient(135deg, #22c55e 0%, #14b8a6 100%);
        color: #052e16;
        font-weight: 850;
        font-size: 11px;
        padding: 6px 12px;
        border-radius: 999px;
        margin-bottom: 18px;
    }

    .top-strip {
        display: grid;
        grid-template-columns: 1.15fr .85fr .85fr .85fr .85fr .85fr;
        gap: 0;
        background: rgba(15, 31, 53, .88);
        border: 1px solid rgba(148, 163, 184, .20);
        border-radius: 14px;
        overflow: hidden;
        box-shadow: 0 18px 45px rgba(0,0,0,.20);
        margin-bottom: 18px;
    }

    .top-item {
        padding: 16px 18px;
        border-right: 1px solid rgba(148, 163, 184, .14);
        min-height: 72px;
    }

    .top-item:last-child { border-right: none; }

    .top-label {
        font-size: 12px;
        color: #94a3b8;
        margin-bottom: 5px;
        font-weight: 600;
    }

    .top-value {
        font-size: 19px;
        color: #f8fafc;
        font-weight: 850;
    }

    .concept-box {
        background: linear-gradient(135deg, rgba(239,246,255,.97), rgba(238,242,255,.96));
        color: #0f172a;
        border: 1px solid rgba(191,219,254,.9);
        border-radius: 16px;
        padding: 24px 28px;
        margin-bottom: 22px;
        box-shadow: 0 14px 34px rgba(15,23,42,.18);
    }

    .concept-title {
        font-size: 20px;
        line-height: 1.65;
    }

    .concept-title strong {
        color: #1e40af;
    }

    .chip-row {
        margin-top: 18px;
        display: flex;
        gap: 10px;
        flex-wrap: wrap;
    }

    .chip {
        display: inline-block;
        border: 1px solid rgba(37,99,235,.35);
        background: rgba(219,234,254,.72);
        color: #1e3a8a;
        padding: 10px 17px;
        border-radius: 14px;
        font-size: 13px;
        font-weight: 850;
    }

    .chip-active {
        background: linear-gradient(135deg, #3b82f6, #2563eb);
        color: white;
        box-shadow: 0 12px 24px rgba(37,99,235,.25);
    }

    .page-title {
        font-size: 31px;
        font-weight: 900;
        color: #f8fafc;
        letter-spacing: -0.9px;
        margin-top: 4px;
        margin-bottom: 0px;
    }

    .page-subtitle {
        color: #cbd5e1;
        font-size: 16px;
        margin-bottom: 18px;
    }

    .card {
        background: linear-gradient(180deg, rgba(15, 31, 53, .92), rgba(8, 23, 39, .92));
        border: 1px solid rgba(148, 163, 184, .19);
        border-radius: 14px;
        padding: 18px;
        box-shadow: 0 18px 45px rgba(0,0,0,.18);
    }

    .mini-card {
        background: linear-gradient(180deg, rgba(15, 31, 53, .96), rgba(7, 19, 34, .96));
        border: 1px solid rgba(148, 163, 184, .19);
        border-radius: 14px;
        padding: 18px;
        min-height: 142px;
        box-shadow: 0 12px 30px rgba(0,0,0,.15);
    }

    .metric-label {
        font-size: 12px;
        color: #cbd5e1;
        margin-bottom: 8px;
        font-weight: 650;
    }

    .metric-value {
        color: #f8fafc;
        font-size: 23px;
        font-weight: 900;
        letter-spacing: -0.5px;
        margin-bottom: 10px;
    }

    .metric-delta {
        display: inline-block;
        background: rgba(22, 163, 74, .13);
        color: #4ade80;
        padding: 5px 9px;
        border-radius: 999px;
        font-weight: 850;
        font-size: 12px;
        margin-bottom: 8px;
    }

    .metric-note {
        color: #94a3b8;
        font-size: 12px;
    }

    .icon-circle {
        width: 42px;
        height: 42px;
        border-radius: 999px;
        display: grid;
        place-items: center;
        font-size: 22px;
        margin-bottom: 12px;
    }

    .i-green { background: rgba(34,197,94,.18); color: #4ade80; }
    .i-blue { background: rgba(59,130,246,.18); color: #60a5fa; }
    .i-purple { background: rgba(168,85,247,.18); color: #c084fc; }
    .i-orange { background: rgba(249,115,22,.18); color: #fb923c; }
    .i-red { background: rgba(239,68,68,.18); color: #f87171; }
    .i-yellow { background: rgba(234,179,8,.18); color: #facc15; }

    .section-title {
        color: #f8fafc;
        font-weight: 900;
        font-size: 19px;
        margin: 12px 0 14px;
        letter-spacing: -.2px;
    }

    .brief-row {
        display: flex;
        gap: 10px;
        align-items: flex-start;
        padding: 9px 0;
        color: #dbeafe;
        font-size: 14px;
        border-bottom: 1px solid rgba(148,163,184,.12);
    }

    .brief-row:last-child { border-bottom: none; }

    .dot {
        width: 12px;
        height: 12px;
        border-radius: 999px;
        margin-top: 4px;
        flex: 0 0 auto;
    }

    .dot-green { background: #22c55e; }
    .dot-yellow { background: #eab308; }
    .dot-red { background: #ef4444; }
    .dot-blue { background: #3b82f6; }

    .alert-card {
        border: 1px solid rgba(148, 163, 184, .18);
        background: rgba(15,23,42,.38);
        border-radius: 12px;
        padding: 12px 13px;
        margin-bottom: 10px;
        display: flex;
        justify-content: space-between;
        gap: 12px;
    }

    .alert-title {
        font-weight: 850;
        color: #fbbf24;
        font-size: 13px;
    }

    .alert-text {
        color: #cbd5e1;
        font-size: 12px;
        margin-top: 3px;
    }

    .alert-time {
        color: #93c5fd;
        font-size: 12px;
        white-space: nowrap;
    }

    .question-card {
        background: linear-gradient(180deg, rgba(15, 31, 53, .88), rgba(8, 22, 37, .88));
        border: 1px solid rgba(148, 163, 184, .18);
        border-radius: 14px;
        padding: 17px;
        min-height: 168px;
        transition: .18s ease;
    }

    .question-card:hover {
        border-color: rgba(96,165,250,.55);
        transform: translateY(-1px);
    }

    .q-title {
        color: #f8fafc;
        font-weight: 850;
        font-size: 15px;
        margin-bottom: 8px;
    }

    .q-answer {
        color: #cbd5e1;
        font-size: 12px;
        line-height: 1.45;
        margin-bottom: 13px;
    }

    .q-link {
        color: #93c5fd;
        font-size: 12px;
        font-weight: 800;
    }

    .footer-note {
        color: #94a3b8;
        font-size: 12px;
        margin-top: 18px;
        display: flex;
        justify-content: space-between;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 13px;
        overflow: hidden;
    }

    .stButton > button {
        background: rgba(15, 31, 53, .92);
        border: 1px solid rgba(148, 163, 184, .20);
        color: #e5eefb;
        border-radius: 12px;
        font-weight: 800;
        width: 100%;
        min-height: 44px;
    }

    .stButton > button:hover {
        border-color: #60a5fa;
        color: white;
        background: rgba(37,99,235,.32);
    }

    div[data-testid="stMetric"] {
        background: linear-gradient(180deg, rgba(15, 31, 53, .92), rgba(8, 23, 39, .92));
        border: 1px solid rgba(148, 163, 184, .19);
        border-radius: 14px;
        padding: 16px;
    }

    div[data-testid="stMetric"] label {
        color: #cbd5e1 !important;
    }

    div[data-testid="stMetricValue"] {
        color: #f8fafc !important;
    }

    div[data-testid="stMetricDelta"] {
        color: #4ade80 !important;
    }

    .trend-panel {
        background: linear-gradient(180deg, rgba(15, 31, 53, .95), rgba(8, 23, 39, .95));
        border: 1px solid rgba(148, 163, 184, .22);
        border-radius: 16px;
        padding: 18px 18px 12px 18px;
        box-shadow: 0 18px 45px rgba(0,0,0,.18);
        min-height: 575px;
    }

    .trend-panel-title {
        color: #f8fafc;
        font-size: 18px;
        font-weight: 900;
        margin-bottom: 14px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .trend-info {
        color: #94a3b8;
        border: 1px solid rgba(148, 163, 184, .35);
        width: 22px;
        height: 22px;
        border-radius: 999px;
        display: inline-grid;
        place-items: center;
        font-size: 13px;
    }

    .trend-stat-row {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 10px;
        margin-bottom: 12px;
    }

    .trend-stat-row.two {
        grid-template-columns: 1fr 1fr;
    }

    .trend-stat {
        background: rgba(15, 23, 42, .42);
        border: 1px solid rgba(148, 163, 184, .16);
        border-radius: 14px;
        padding: 14px;
        min-height: 92px;
    }

    .trend-stat-label {
        color: #cbd5e1;
        font-size: 12px;
        font-weight: 750;
        margin-bottom: 6px;
    }

    .trend-stat-value {
        color: #f8fafc;
        font-size: 22px;
        font-weight: 950;
        letter-spacing: -.4px;
    }

    .trend-stat-sub {
        color: #94a3b8;
        font-size: 12px;
        margin-top: 5px;
    }

    .trend-up { color: #4ade80; font-weight: 900; }
    .trend-red { color: #fb7185; font-weight: 900; }
    .trend-green { color: #22c55e; font-weight: 900; }
    .trend-blue { color: #60a5fa; font-weight: 900; }

    .trend-bottom-row {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr 1fr;
        gap: 0;
        background: linear-gradient(180deg, rgba(15,31,53,.95), rgba(8,23,39,.95));
        border: 1px solid rgba(148,163,184,.20);
        border-radius: 16px;
        margin-top: 16px;
        overflow: hidden;
    }

    .trend-bottom-item {
        padding: 20px 22px;
        border-right: 1px solid rgba(148,163,184,.16);
        min-height: 105px;
        display: flex;
        gap: 14px;
        align-items: center;
    }

    .trend-bottom-item:last-child { border-right: none; }


</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# Helpers
# =========================================================
def money_fmt(x) -> str:
    try:
        x = float(x)
    except Exception:
        return "0 TL"
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:,.1f} M TL".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs(x) >= 1_000:
        return f"{x/1_000:,.0f} B TL".replace(",", ".")
    return f"{x:,.0f} TL".replace(",", ".")


def int_fmt(x) -> str:
    try:
        return f"{int(round(float(x))):,}".replace(",", ".")
    except Exception:
        return "0"


def pct_fmt(x, digits=0) -> str:
    try:
        return f"%{float(x):.{digits}f}".replace(".", ",")
    except Exception:
        return "%0"


def date_short_tr(d) -> str:
    aylar = {
        1: "Oca", 2: "Şub", 3: "Mar", 4: "Nis", 5: "May", 6: "Haz",
        7: "Tem", 8: "Ağu", 9: "Eyl", 10: "Eki", 11: "Kas", 12: "Ara"
    }
    d = pd.to_datetime(d)
    return f"{d.day} {aylar.get(d.month, '')}"


def date_long_tr(d) -> str:
    aylar = {
        1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
        7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
    }
    d = pd.to_datetime(d)
    return f"{d.day} {aylar.get(d.month, '')} {d.year}"


def metric_card(icon, icon_class, label, value, delta, note):
    st.markdown(
        f"""
        <div class="mini-card">
            <div class="icon-circle {icon_class}">{icon}</div>
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-delta">{delta}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def chart_layout(fig, height=390):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#cbd5e1"),
        margin=dict(l=15, r=15, t=45, b=25),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(gridcolor="rgba(148,163,184,.13)", zeroline=False, tickformat="%d %b", type="date"),
        yaxis=dict(gridcolor="rgba(148,163,184,.13)", zeroline=False),
    )
    return fig


def chart_layout_trend(fig, height=360):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#cbd5e1", size=12),
        margin=dict(l=8, r=8, t=18, b=35),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(
            gridcolor="rgba(148,163,184,.10)",
            zeroline=False,
            tickformat="%d %b",
            type="date",
            showspikes=False,
        ),
        yaxis=dict(
            gridcolor="rgba(148,163,184,.18)",
            zeroline=False,
        ),
        hovermode="x unified",
    )
    return fig


def section_header(title, subtitle=None):
    st.markdown(f'<div class="page-title">{title}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="page-subtitle">{subtitle}</div>', unsafe_allow_html=True)


def parse_excel_dates(series: pd.Series) -> pd.Series:
    """
    Excel tarihleri bazen 2026-05-24 yerine 46166 gibi seri numarası olarak gelir.
    Normal pd.to_datetime bu sayıları 1970 epoch çevresine düşürür.
    Bu fonksiyon hem gerçek tarihleri hem Excel seri tarihlerini güvenli çevirir.
    """
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin="1899-12-30", errors="coerce")

    numeric = pd.to_numeric(series, errors="coerce")
    numeric_ratio = numeric.notna().mean()

    if numeric_ratio > 0.70 and numeric.dropna().between(30000, 60000).mean() > 0.70:
        return pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")

    return pd.to_datetime(series, errors="coerce", dayfirst=True)


@st.cache_data(show_spinner=False)
def load_data(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

    for key in ["Gunluk", "Brans", "Doktor", "Operasyon"]:
        if key in data and "Tarih" in data[key].columns:
            data[key]["Tarih"] = parse_excel_dates(data[key]["Tarih"])

    return data


def safe_sheet(data, name):
    return data.get(name, pd.DataFrame())


# =========================================================
# Sidebar
# =========================================================
with st.sidebar:
    st.markdown('<div class="ayca-logo">✣ AYÇA</div>', unsafe_allow_html=True)
    st.markdown('<div class="ayca-sub">INSIGHT HOSPITAL</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="version-badge">{APP_VERSION}</div>', unsafe_allow_html=True)

    nav_items = [
        "CEO Dashboard",
        "Finans Merkezi",
        "Doktor Intelligence",
        "Operasyon",
        "Kalite & Hasta Deneyimi",
        "Hasta Analitiği",
        "Stok & Satın Alma",
        "AYÇA Co-Pilot",
    ]

    if "active_module" not in st.session_state:
        st.session_state["active_module"] = "CEO Dashboard"

    if st.session_state["active_module"] not in nav_items:
        st.session_state["active_module"] = "CEO Dashboard"

    module = st.radio(
        "Menü",
        nav_items,
        index=nav_items.index(st.session_state["active_module"]),
        label_visibility="collapsed",
        key="sidebar_module",
    )
    st.session_state["active_module"] = module

    st.divider()
    st.caption("Veri Dosyası")
    uploaded = st.file_uploader(
        "Excel veri dosyası",
        type=["xlsx"],
        help="AYÇA Hospital sample Excel dosyasını yükleyiniz.",
        label_visibility="collapsed",
    )

    if uploaded is None:
        st.warning("Devam etmek için Excel veri dosyasını yükleyiniz.")
        st.stop()
    else:
        st.success("Dosya yüklendi")

    st.button("Yeni Dosya Yükle")


# =========================================================
# Data
# =========================================================
try:
    data = load_data(uploaded)
except Exception as e:
    st.error("Excel dosyası okunamadı. Lütfen AYÇA Hospital veri şablonunu kontrol ediniz.")
    st.exception(e)
    st.stop()

required_sheets = ["Gunluk", "Brans", "Doktor", "Hasta", "Operasyon", "Stok", "Kalite"]
missing = [s for s in required_sheets if s not in data]
if missing:
    st.error("Excel şablonunda eksik sayfa var: " + ", ".join(missing))
    st.stop()

gunluk = safe_sheet(data, "Gunluk")
brans = safe_sheet(data, "Brans")
doktor = safe_sheet(data, "Doktor")
hasta = safe_sheet(data, "Hasta")
operasyon = safe_sheet(data, "Operasyon")
stok = safe_sheet(data, "Stok")
kalite = safe_sheet(data, "Kalite")

if gunluk.empty or "Tarih" not in gunluk.columns:
    st.error("Gunluk sayfasında Tarih kolonu bulunmalı.")
    st.stop()

gunluk = gunluk.dropna(subset=["Tarih"]).sort_values("Tarih").copy()
brans = brans.dropna(subset=["Tarih"]).sort_values("Tarih").copy() if "Tarih" in brans.columns else brans
doktor = doktor.dropna(subset=["Tarih"]).sort_values("Tarih").copy() if "Tarih" in doktor.columns else doktor
operasyon = operasyon.dropna(subset=["Tarih"]).sort_values("Tarih").copy() if "Tarih" in operasyon.columns else operasyon

max_date = pd.to_datetime(gunluk["Tarih"]).max()
today_df = gunluk[gunluk["Tarih"] == max_date]
today = today_df.iloc[0] if not today_df.empty else gunluk.iloc[-1]
last_30 = gunluk.tail(30).copy()

# Derived metrics
total_revenue = float(gunluk["Günlük Ciro"].sum())
avg_revenue = float(gunluk["Günlük Ciro"].mean())
total_patient = int(gunluk["Toplam Hasta"].sum())
total_emergency = int(gunluk["Acil Hasta"].sum())
peak_patient_row = gunluk.loc[gunluk["Toplam Hasta"].idxmax()]
peak_emergency_row = gunluk.loc[gunluk["Acil Hasta"].idxmax()]


# =========================================================
# Top strip
# =========================================================
st.markdown(
    f"""
    <div class="top-strip">
        <div class="top-item">
            <div class="top-label">Hastane</div>
            <div class="top-value">HG Hospital ▾</div>
        </div>
        <div class="top-item">
            <div class="top-label">Tarih</div>
            <div class="top-value">{max_date.strftime('%d %B %Y')}</div>
        </div>
        <div class="top-item">
            <div class="top-label">Doluluk Oranı</div>
            <div class="top-value">{pct_fmt(today['Doluluk %'])}</div>
        </div>
        <div class="top-item">
            <div class="top-label">Günlük Hasta</div>
            <div class="top-value">{int_fmt(today['Toplam Hasta'])}</div>
        </div>
        <div class="top-item">
            <div class="top-label">Acil Servis</div>
            <div class="top-value">{int_fmt(today['Acil Hasta'])}</div>
        </div>
        <div class="top-item">
            <div class="top-label">Risk Skoru</div>
            <div class="top-value">{int(today['Risk Skoru'])} / 100</div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="concept-box">
        <div class="concept-title">
            <strong>V0.4.2 Konsept:</strong> Bu ekran hastane direktörü, genel müdür ve yönetim ekibinin sabah tek bakışta
            hastanenin finansal, operasyonel ve kalite durumunu görmesi için tasarlanmıştır.
        </div>
        <div class="chip-row">
            <span class="chip chip-active">CEO Dashboard</span>
            <span class="chip">Finans Merkezi</span>
            <span class="chip">Doktor Intelligence</span>
            <span class="chip">Operasyon Riski</span>
            <span class="chip">AYÇA Co-Pilot</span>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

nav_cols = st.columns(5)
quick_targets = ["CEO Dashboard", "Finans Merkezi", "Doktor Intelligence", "Operasyon", "AYÇA Co-Pilot"]
for i, target in enumerate(quick_targets):
    with nav_cols[i]:
        if st.button(target, key=f"top_nav_{target}"):
            st.session_state["active_module"] = target
            st.rerun()


# =========================================================
# CEO Dashboard
# =========================================================
def render_ceo_dashboard():
    section_header("CEO Dashboard", "Yönetici Özeti")

    cols = st.columns(6)
    with cols[0]: metric_card("💰", "i-green", "Günlük Ciro", money_fmt(today["Günlük Ciro"]), "↑ %12,4", "Önceki 30 güne göre")
    with cols[1]: metric_card("📊", "i-blue", "Aylık Ciro", money_fmt(total_revenue), "↑ %8,7", "Konsept dönem")
    with cols[2]: metric_card("🛏️", "i-purple", "Yatan Hasta", int_fmt(today["Yatan Hasta"]), "↑ %3,6", "Dün: 138")
    with cols[3]: metric_card("🟠", "i-orange", "Ortalama Doluluk", pct_fmt(today["Doluluk %"]), "↑ %2,5", "Dün: %80")
    with cols[4]: metric_card("🕘", "i-blue", "Bekleme Süresi", f"{int(today['Ortalama Bekleme Dk'])} dk", "↓ %6,7", "Dün: 15 dk")
    with cols[5]: metric_card("🙂", "i-green", "Hasta Memnuniyeti", pct_fmt(today["Memnuniyet %"]), "↑ %4,2", "Önceki 30 güne göre")

    st.markdown('<div class="section-title">Son 30 Gün Trend</div>', unsafe_allow_html=True)

    avg_daily_revenue = float(last_30["Günlük Ciro"].mean())
    max_revenue_row = last_30.loc[last_30["Günlük Ciro"].idxmax()]
    min_revenue_row = last_30.loc[last_30["Günlük Ciro"].idxmin()]
    avg_total_patient = float(last_30["Toplam Hasta"].mean())
    avg_emergency_patient = float(last_30["Acil Hasta"].mean())

    c1, c2 = st.columns(2)

    with c1:
        st.markdown(
            f"""
            <div class="trend-panel">
                <div class="trend-panel-title">
                    <span>Günlük Ciro Trendi</span>
                    <span class="trend-info">i</span>
                </div>
                <div class="trend-stat-row">
                    <div class="trend-stat">
                        <div class="trend-stat-label">Ort. Günlük Ciro</div>
                        <div class="trend-stat-value">{money_fmt(avg_daily_revenue)}</div>
                        <div class="trend-stat-sub"><span class="trend-up">↑ %12,4</span> · Önceki 30 güne göre</div>
                    </div>
                    <div class="trend-stat">
                        <div class="trend-stat-label">En Yüksek Gün</div>
                        <div class="trend-stat-value">{money_fmt(max_revenue_row["Günlük Ciro"])}</div>
                        <div class="trend-stat-sub"><span class="trend-green">{date_long_tr(max_revenue_row["Tarih"])}</span></div>
                    </div>
                    <div class="trend-stat">
                        <div class="trend-stat-label">En Düşük Gün</div>
                        <div class="trend-stat-value">{money_fmt(min_revenue_row["Günlük Ciro"])}</div>
                        <div class="trend-stat-sub"><span class="trend-red">{date_long_tr(min_revenue_row["Tarih"])}</span></div>
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=last_30["Tarih"],
                y=last_30["Günlük Ciro"],
                mode="lines+markers",
                name="Günlük Ciro (TL)",
                line=dict(width=3, color="#3b82f6"),
                marker=dict(size=6, color="#bfdbfe", line=dict(width=2, color="#2563eb")),
                fill="tozeroy",
                fillcolor="rgba(59,130,246,.16)",
            )
        )
        fig.update_yaxes(tickformat="~s", ticksuffix=" TL", rangemode="tozero")
        st.plotly_chart(chart_layout_trend(fig, 365), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown(
            f"""
            <div class="trend-panel">
                <div class="trend-panel-title">
                    <span>Hasta Hacmi Trendi</span>
                    <span class="trend-info">i</span>
                </div>
                <div class="trend-stat-row two">
                    <div class="trend-stat">
                        <div class="trend-stat-label">Ort. Toplam Hasta</div>
                        <div class="trend-stat-value">{int_fmt(avg_total_patient)}</div>
                        <div class="trend-stat-sub"><span class="trend-up">↑ %8,7</span> · Önceki 30 güne göre</div>
                    </div>
                    <div class="trend-stat">
                        <div class="trend-stat-label">Ort. Acil Hasta</div>
                        <div class="trend-stat-value">{int_fmt(avg_emergency_patient)}</div>
                        <div class="trend-stat-sub"><span class="trend-up">↑ %6,1</span> · Önceki 30 güne göre</div>
                    </div>
                </div>
            """,
            unsafe_allow_html=True,
        )

        poliklinik = last_30["Toplam Hasta"] - last_30["Acil Hasta"]
        fig = go.Figure()
        fig.add_trace(go.Bar(
            x=last_30["Tarih"],
            y=poliklinik,
            name="Poliklinik Hasta",
            marker_color="#3b82f6",
        ))
        fig.add_trace(go.Bar(
            x=last_30["Tarih"],
            y=last_30["Acil Hasta"],
            name="Acil Hasta",
            marker_color="#fb7185",
        ))
        fig.add_trace(go.Scatter(
            x=last_30["Tarih"],
            y=last_30["Toplam Hasta"],
            name="Toplam Hasta",
            mode="lines+markers",
            line=dict(width=3, color="#e5eefb"),
            marker=dict(size=5, color="#f8fafc"),
        ))
        fig.update_layout(barmode="stack")
        fig.update_yaxes(title="Hasta Sayısı (Poliklinik + Acil)")
        st.plotly_chart(chart_layout_trend(fig, 365), use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="trend-bottom-row">
            <div class="trend-bottom-item">
                <div class="icon-circle i-blue">👥</div>
                <div>
                    <div class="trend-stat-label">Toplam Hasta</div>
                    <div class="trend-stat-value">{int_fmt(total_patient)} <span class="trend-up" style="font-size:13px;">↑ %8,7</span></div>
                    <div class="trend-stat-sub">Son 30 gün</div>
                </div>
            </div>
            <div class="trend-bottom-item">
                <div class="icon-circle i-red">🚨</div>
                <div>
                    <div class="trend-stat-label">Acil Hasta</div>
                    <div class="trend-stat-value">{int_fmt(total_emergency)} <span class="trend-up" style="font-size:13px;">↑ %6,1</span></div>
                    <div class="trend-stat-sub">Son 30 gün</div>
                </div>
            </div>
            <div class="trend-bottom-item">
                <div class="icon-circle i-green">📈</div>
                <div>
                    <div class="trend-stat-label">En Yoğun Gün</div>
                    <div class="trend-stat-value"><span class="trend-green">{int_fmt(peak_patient_row["Toplam Hasta"])}</span></div>
                    <div class="trend-stat-sub">Toplam Hasta · {date_long_tr(peak_patient_row["Tarih"])}</div>
                </div>
            </div>
            <div class="trend-bottom-item">
                <div class="icon-circle i-orange">🗓️</div>
                <div>
                    <div class="trend-stat-label">En Yoğun Acil Gün</div>
                    <div class="trend-stat-value"><span style="color:#fb923c;">{int_fmt(peak_emergency_row["Acil Hasta"])}</span></div>
                    <div class="trend-stat-sub">Acil Hasta · {date_long_tr(peak_emergency_row["Tarih"])}</div>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    b1, b2 = st.columns([1, 1])
    with b1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Yönetici Sabah Brifingi</div>', unsafe_allow_html=True)
        rows = [
            ("dot-green", "Hasta sayısı geçen haftaya göre %12 arttı."),
            ("dot-yellow", f"Kardiyoloji polikliniğinde bekleme süresi {int(today['Ortalama Bekleme Dk'])} dk seviyesinde."),
            ("dot-red", f"Radyoloji MR kullanım oranı {pct_fmt(today['MR Kullanım %'])} seviyesine ulaştı."),
            ("dot-green", f"Hasta memnuniyeti {pct_fmt(today['Memnuniyet %'])} seviyesinde."),
            ("dot-green", f"Günlük hedef cironun {pct_fmt(today['Hedef Gerçekleşme %'])} seviyesine ulaşıldı."),
        ]
        for cls, txt in rows:
            st.markdown(f'<div class="brief-row"><span class="dot {cls}"></span><span>{txt}</span></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with b2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">En Kritik Uyarılar</div>', unsafe_allow_html=True)
        alerts = [
            (f"Radyoloji MR kullanım oranı {pct_fmt(today['MR Kullanım %'])}", "Kapasite limitine çok yakın.", "09:15"),
            (f"Kardiyoloji bekleme süresi {int(today['Ortalama Bekleme Dk'])} dk", "Hedefin üzerinde. Randevu kapasitesi incelenmeli.", "09:05"),
            (f"Yoğun bakım doluluk oranı {pct_fmt(today['Doluluk %'])}", "Yüksek doluluk. Takip önerilir.", "08:50"),
        ]
        for title, text, time in alerts:
            st.markdown(f'<div class="alert-card"><div><div class="alert-title">{title}</div><div class="alert-text">{text}</div></div><div class="alert-time">{time}</div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    render_question_cards("ceo")


def go_to(module_name: str):
    st.session_state["active_module"] = module_name
    st.rerun()


def render_question_cards(context="ceo"):
    st.markdown('<div class="section-title">Yönetici İçin Önerilen Sorular</div>', unsafe_allow_html=True)

    brans_sum = brans.groupby("Branş", as_index=False).agg(Gelir=("Gelir", "sum"), Kar=("Kar", "sum"), Hasta=("Hasta", "sum"))
    top_rev = brans_sum.sort_values("Gelir", ascending=False).iloc[0]
    top_profit = brans_sum.sort_values("Kar", ascending=False).iloc[0]
    doc_sum = doktor.groupby(["Doktor", "Branş"], as_index=False).agg(Gelir=("Gelir", "sum"), Memnuniyet=("Memnuniyet %", "mean"), Hasta=("Hasta", "sum"))
    top_doc = doc_sum.sort_values("Gelir", ascending=False).iloc[0]
    top_satis = doc_sum.sort_values("Memnuniyet", ascending=False).iloc[0]
    top_complaint = kalite.sort_values("Şikayet Sayısı", ascending=False).iloc[0]

    cards = [
        ("📈", "Bu ay ciro neden arttı?", f"Cevap: {top_rev['Branş']} branşında gelir artışı güçlü. Toplam ciroyu en çok bu branş destekliyor.", "Finans Merkezi"),
        ("🟣", "En kârlı branş hangisi?", f"Cevap: {top_profit['Branş']} yaklaşık {money_fmt(top_profit['Kar'])} kâr ile öne çıkıyor.", "Finans Merkezi"),
        ("🕘", "Bekleme süresi neden arttı?", f"Cevap: Hasta hacmi {int_fmt(today['Toplam Hasta'])} seviyesinde. Randevu kapasitesi ve yoğun branşlar birlikte incelenmeli.", "Operasyon"),
        ("🙂", "Hasta memnuniyeti nasıl artırılır?", f"Cevap: En büyük şikayet konusu '{top_complaint['Konu']}'. Bu alanda çözüm süresi düşürülmeli.", "Kalite & Hasta Deneyimi"),
        ("👨‍⚕️", "Hangi doktorlar öne çıkıyor?", f"Cevap: Gelirde {top_doc['Doktor']}, memnuniyette {top_satis['Doktor']} öne çıkıyor.", "Doktor Intelligence"),
        ("🛡️", "Hangi alanlarda risk var?", f"Cevap: MR kullanımı {pct_fmt(today['MR Kullanım %'])}, doluluk {pct_fmt(today['Doluluk %'])}. Kapasite riski takip edilmeli.", "Operasyon"),
    ]

    cols = st.columns(3)
    for i, (icon, title, ans, target) in enumerate(cards):
        with cols[i % 3]:
            st.markdown(
                f"""
                <div class="question-card">
                    <div class="icon-circle i-blue">{icon}</div>
                    <div class="q-title">{title}</div>
                    <div class="q-answer">{ans}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(f"Detayları Gör → {target}", key=f"q_nav_{context}_{i}"):
                go_to(target)


# =========================================================
# Other pages
# =========================================================
def render_finance():
    section_header("Finans Merkezi", "Branş, kanal ve kârlılık analizi")

    cols = st.columns(4)
    with cols[0]: metric_card("💰", "i-green", "Toplam Ciro", money_fmt(total_revenue), "↑ %8,7", "Son 30 gün")
    with cols[1]: metric_card("🏦", "i-blue", "SGK Geliri", money_fmt(gunluk["SGK Geliri"].sum()), "Kanal", "Son 30 gün")
    with cols[2]: metric_card("🛡️", "i-purple", "Özel Sigorta", money_fmt(gunluk["Özel Sigorta Geliri"].sum()), "Kanal", "Son 30 gün")
    with cols[3]: metric_card("💵", "i-orange", "Nakit Gelir", money_fmt(gunluk["Nakit Gelir"].sum()), "Kanal", "Son 30 gün")

    brans_sum = brans.groupby("Branş", as_index=False).agg(Gelir=("Gelir", "sum"), Kar=("Kar", "sum"), Hasta=("Hasta", "sum"), Memnuniyet=("Memnuniyet %", "mean"))
    brans_sum["Hasta Başı Gelir"] = brans_sum["Gelir"] / brans_sum["Hasta"]

    c1, c2 = st.columns([1.15, 1])
    with c1:
        fig = px.bar(brans_sum.sort_values("Gelir", ascending=True), x="Gelir", y="Branş", orientation="h", title="Branş Bazlı Gelir")
        fig.update_traces(marker_color="#3b82f6")
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)
    with c2:
        fig = px.bar(brans_sum.sort_values("Kar", ascending=True), x="Kar", y="Branş", orientation="h", title="Branş Bazlı Kâr")
        fig.update_traces(marker_color="#22c55e")
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)

    st.dataframe(brans_sum.sort_values("Gelir", ascending=False), use_container_width=True, hide_index=True)
    render_question_cards("finance")


def render_doctors():
    section_header("Doktor Intelligence", "Doktor performansı, gelir, memnuniyet ve hasta başı gelir")

    dsum = doktor.groupby(["Doktor", "Branş"], as_index=False).agg(
        Hasta=("Hasta", "sum"),
        Gelir=("Gelir", "sum"),
        Tetkik=("Tetkik Geliri", "sum"),
        Memnuniyet=("Memnuniyet %", "mean"),
        Tekrar=("Tekrar Başvuru %", "mean"),
    )
    dsum["Hasta Başı Gelir"] = dsum["Gelir"] / dsum["Hasta"]

    cols = st.columns(4)
    top_doc = dsum.sort_values("Gelir", ascending=False).iloc[0]
    top_mem = dsum.sort_values("Memnuniyet", ascending=False).iloc[0]
    with cols[0]: metric_card("👨‍⚕️", "i-blue", "En Yüksek Gelir", top_doc["Doktor"], money_fmt(top_doc["Gelir"]), top_doc["Branş"])
    with cols[1]: metric_card("🙂", "i-green", "En Yüksek Memnuniyet", top_mem["Doktor"], pct_fmt(top_mem["Memnuniyet"]), top_mem["Branş"])
    with cols[2]: metric_card("🧪", "i-purple", "Tetkik Geliri", money_fmt(dsum["Tetkik"].sum()), "Toplam", "Son 30 gün")
    with cols[3]: metric_card("🔁", "i-orange", "Tekrar Başvuru", pct_fmt(dsum["Tekrar"].mean()), "Ortalama", "Son 30 gün")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.scatter(dsum, x="Hasta", y="Gelir", size="Memnuniyet", color="Branş", hover_name="Doktor", title="Doktor Gelir / Hasta Matrisi")
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)
    with c2:
        fig = px.bar(dsum.sort_values("Gelir", ascending=False).head(10), x="Doktor", y="Gelir", color="Branş", title="Top 10 Doktor Gelir")
        fig.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)

    st.dataframe(dsum.sort_values("Gelir", ascending=False), use_container_width=True, hide_index=True)
    render_question_cards("doctor")


def render_operations():
    section_header("Operasyon", "Doluluk, bekleme süresi ve kapasite riski")

    latest_ops = operasyon[operasyon["Tarih"] == operasyon["Tarih"].max()].copy()

    cols = st.columns(4)
    with cols[0]: metric_card("🕘", "i-blue", "Ortalama Bekleme", f"{int(today['Ortalama Bekleme Dk'])} dk", "Operasyon", "Bugün")
    with cols[1]: metric_card("🏥", "i-purple", "Doluluk", pct_fmt(today["Doluluk %"]), "Yatak", "Bugün")
    with cols[2]: metric_card("🧲", "i-red", "MR Kullanım", pct_fmt(today["MR Kullanım %"]), "Kritik", "Bugün")
    with cols[3]: metric_card("🔪", "i-orange", "Ameliyat", int_fmt(today["Ameliyat Sayısı"]), "Bugün", "Planlanan")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(latest_ops.sort_values("Bekleme Dk"), x="Bekleme Dk", y="Birim", orientation="h", title="Birim Bazlı Bekleme")
        fig.update_traces(marker_color="#f97316")
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)
    with c2:
        fig = px.bar(latest_ops.sort_values("Doluluk %"), x="Doluluk %", y="Birim", orientation="h", title="Birim Bazlı Doluluk")
        fig.update_traces(marker_color="#3b82f6")
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)

    st.dataframe(latest_ops, use_container_width=True, hide_index=True)
    render_question_cards("ops")


def render_quality():
    section_header("Kalite & Hasta Deneyimi", "Şikayet, memnuniyet ve çözüm süresi")

    top_complaint = kalite.sort_values("Şikayet Sayısı", ascending=False).iloc[0]
    cols = st.columns(4)
    with cols[0]: metric_card("📣", "i-red", "Toplam Şikayet", int_fmt(kalite["Şikayet Sayısı"].sum()), "Takip", "Son 30 gün")
    with cols[1]: metric_card("⏱️", "i-orange", "Çözüm Süresi", f"{kalite['Ortalama Çözüm Saat'].mean():.1f} saat", "Ortalama", "Hedef <24 saat")
    with cols[2]: metric_card("🙂", "i-green", "Memnuniyet", pct_fmt(today["Memnuniyet %"]), "Bugün", "Hasta deneyimi")
    with cols[3]: metric_card("⚠️", "i-yellow", "Ana Şikayet", top_complaint["Konu"], int_fmt(top_complaint["Şikayet Sayısı"]), "Adet")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(kalite.sort_values("Şikayet Sayısı"), x="Şikayet Sayısı", y="Konu", orientation="h", title="Şikayet Konuları")
        fig.update_traces(marker_color="#ef4444")
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)
    with c2:
        fig = px.scatter(kalite, x="Şikayet Sayısı", y="Ortalama Çözüm Saat", size="Etki Skoru", color="Öncelik", hover_name="Konu", title="Şikayet Etki Matrisi")
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)

    st.dataframe(kalite, use_container_width=True, hide_index=True)
    render_question_cards("quality")


def render_patients():
    section_header("Hasta Analitiği", "Yeni hasta, sadık hasta ve kaybedilen hasta segmentleri")

    cols = st.columns(4)
    with cols[0]: metric_card("🆕", "i-blue", "Yeni Hasta", int_fmt(hasta["Yeni Hasta"].sum()), "Segment", "Son 30 gün")
    with cols[1]: metric_card("🔁", "i-green", "Sadık Hasta", int_fmt(hasta["Sadık Hasta"].sum()), "5+ ziyaret", "Son 12 ay")
    with cols[2]: metric_card("📉", "i-red", "Kaybedilen Hasta", int_fmt(hasta["Kaybedilen Hasta"].sum()), "6 ay pasif", "Takip gerekli")
    with cols[3]: metric_card("🧾", "i-purple", "Ortalama Sepet", money_fmt(hasta["Ortalama Sepet TL"].mean()), "Branş ort.", "Hasta başı")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(hasta, x="Branş", y=["Yeni Hasta", "Sadık Hasta", "Kaybedilen Hasta"], barmode="group", title="Hasta Segmentleri")
        fig.update_layout(xaxis_tickangle=-25)
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)
    with c2:
        fig = px.pie(hasta, names="Branş", values="Yeni Hasta", title="Yeni Hasta Dağılımı", hole=.45)
        st.plotly_chart(chart_layout(fig, 430), use_container_width=True)

    st.dataframe(hasta, use_container_width=True, hide_index=True)
    render_question_cards("patients")


def render_stock():
    section_header("Stok & Satın Alma", "Kritik stok, SKT riski ve önerilen sipariş")

    critical = stok[stok["Durum"].isin(["Kritik", "Riskli"])]
    cols = st.columns(4)
    with cols[0]: metric_card("📦", "i-blue", "Stok Değeri", money_fmt(stok["Stok Değeri"].sum()), "Toplam", "Sarf + ilaç")
    with cols[1]: metric_card("🚨", "i-red", "Kritik Kalem", int_fmt((stok["Durum"] == "Kritik").sum()), "Minimum altı", "Aksiyon")
    with cols[2]: metric_card("⏳", "i-yellow", "SKT Riski", int_fmt((stok["SKT Risk Gün"] <= 60).sum()), "60 gün altı", "Takip")
    with cols[3]: metric_card("🛒", "i-green", "Sipariş Önerisi", money_fmt(stok["Önerilen Sipariş Değeri"].sum()), "Demo öneri", "Satın alma")

    c1, c2 = st.columns([1.1, 1])
    with c1:
        fig = px.bar(stok.sort_values("Stok Değeri", ascending=True), x="Stok Değeri", y="Malzeme", color="Kategori", orientation="h", title="Stok Değeri")
        st.plotly_chart(chart_layout(fig, 470), use_container_width=True)
    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.markdown('<div class="section-title">Kritik / Riskli Kalemler</div>', unsafe_allow_html=True)
        st.dataframe(critical, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    render_question_cards("stock")


def render_copilot():
    section_header("AYÇA Co-Pilot", "Doğal dil ile yönetici karar destek asistanı")

    brans_sum = brans.groupby("Branş", as_index=False).agg(Gelir=("Gelir", "sum"), Kar=("Kar", "sum"), Hasta=("Hasta", "sum"))
    top_rev = brans_sum.sort_values("Gelir", ascending=False).iloc[0]
    top_profit = brans_sum.sort_values("Kar", ascending=False).iloc[0]
    top_complaint = kalite.sort_values("Şikayet Sayısı", ascending=False).iloc[0]
    dsum = doktor.groupby(["Doktor", "Branş"], as_index=False).agg(Gelir=("Gelir", "sum"), Memnuniyet=("Memnuniyet %", "mean"))
    top_doc = dsum.sort_values("Gelir", ascending=False).iloc[0]

    qa = {
        "Bu ay ciro neden arttı?": f"Ciro artışının ana nedeni {top_rev['Branş']} branşındaki yüksek gelir katkısıdır. Bu branş tek başına {money_fmt(top_rev['Gelir'])} gelir üretmiştir.",
        "En kârlı branş hangisi?": f"En kârlı branş {top_profit['Branş']} görünüyor. Yaklaşık {money_fmt(top_profit['Kar'])} kâr üretmiştir.",
        "Bekleme süresi neden arttı?": f"Bekleme süresi hasta hacmi ve kapasite baskısından etkileniyor. Bugünkü toplam hasta {int_fmt(today['Toplam Hasta'])}, ortalama bekleme {int(today['Ortalama Bekleme Dk'])} dakikadır.",
        "Hasta memnuniyeti nasıl artırılır?": f"En çok şikayet konusu '{top_complaint['Konu']}'. Bu başlıkta çözüm süresi düşürülürse memnuniyet artışı beklenir.",
        "Hangi doktorlar öne çıkıyor?": f"Gelir açısından {top_doc['Doktor']} öne çıkıyor. Branşı {top_doc['Branş']}, toplam geliri {money_fmt(top_doc['Gelir'])}.",
        "Hangi alanlarda risk var?": f"MR kullanım oranı {pct_fmt(today['MR Kullanım %'])}, doluluk {pct_fmt(today['Doluluk %'])}. Kapasite ve bekleme riski birlikte takip edilmelidir.",
    }

    cols = st.columns(3)
    selected = None
    questions = list(qa.keys())
    for i, q in enumerate(questions):
        with cols[i % 3]:
            if st.button(q):
                selected = q

    st.markdown('<div class="card">', unsafe_allow_html=True)
    custom_q = st.text_input("Kendi sorunuzu yazın", value=selected or "Bu ay ciro neden arttı?")
    answer = qa.get(custom_q, "Bu soru için demo yanıt üretildi: ilgili modüldeki finans, operasyon, doktor ve kalite göstergeleri birlikte değerlendirilmelidir.")
    st.markdown(f"<h3 style='color:#f8fafc;'>Yanıt</h3><p style='color:#cbd5e1; font-size:16px; line-height:1.65;'>{answer}</p>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    render_question_cards("copilot")


# =========================================================
# Router
# =========================================================
if module == "CEO Dashboard":
    render_ceo_dashboard()
elif module == "Finans Merkezi":
    render_finance()
elif module == "Doktor Intelligence":
    render_doctors()
elif module == "Operasyon":
    render_operations()
elif module == "Kalite & Hasta Deneyimi":
    render_quality()
elif module == "Hasta Analitiği":
    render_patients()
elif module == "Stok & Satın Alma":
    render_stock()
elif module == "AYÇA Co-Pilot":
    render_copilot()


st.markdown(
    f"""
    <div class="footer-note">
        <span>Bu uygulama konsept demodur. Veriler tamamen sample/sentetik olarak üretilmiştir.</span>
        <span>AYÇA Insight Hospital {APP_VERSION}</span>
    </div>
    """,
    unsafe_allow_html=True,
)
