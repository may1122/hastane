
# -*- coding: utf-8 -*-
"""
AYÇA Insight Hospital V0.5
Çalışan Navigasyon + Kompakt Trend Alanı

Run:
    pip install -r requirements.txt
    streamlit run app.py

Excel sheets required:
    Gunluk, Brans, Doktor, Hasta, Operasyon, Stok, Kalite
"""

from __future__ import annotations

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import streamlit as st


APP_VERSION = "V0.6 SaaS Hospital Concept"

NAV_ITEMS = [
    "CEO Dashboard",
    "Finans Merkezi",
    "Doktor Intelligence",
    "Operasyon",
    "Yönetim Toplantısı",
    "Kalite & Hasta Deneyimi",
    "Hasta Analitiği",
    "Stok & Satın Alma",
    "AYÇA Co-Pilot",
]


# =========================================================
# Page config
# =========================================================
st.set_page_config(
    page_title="AYÇA Insight Hospital",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "active_page" not in st.session_state:
    st.session_state.active_page = "CEO Dashboard"


# =========================================================
# CSS
# =========================================================
st.markdown(
    """
<style>
    .stApp {
        background: #f8fafc;
        color: #0f172a;
    }

    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 2rem;
        max-width: 1480px;
    }

    section[data-testid="stSidebar"] {
        background: #ffffff;
        border-right: 1px solid #e2e8f0;
        box-shadow: 8px 0 24px rgba(15,23,42,.035);
    }

    section[data-testid="stSidebar"] * {
        color: #0f172a;
    }

    .ayca-logo {
        font-size: 29px;
        font-weight: 950;
        letter-spacing: -1.1px;
        color: #0f172a;
        line-height: 1.05;
    }

    .ayca-sub {
        font-size: 12px;
        color: #64748b;
        letter-spacing: .3px;
        margin-bottom: 10px;
        font-weight: 750;
    }

    .version-badge {
        display: inline-block;
        background: #eff6ff;
        color: #1d4ed8;
        border: 1px solid #bfdbfe;
        font-weight: 850;
        font-size: 11px;
        padding: 6px 12px;
        border-radius: 999px;
        margin-bottom: 18px;
    }

    .sidebar-org {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 12px 14px;
        margin: 8px 0 16px;
    }

    .sidebar-org-title {
        color: #0f172a;
        font-weight: 900;
        font-size: 14px;
    }

    .sidebar-org-sub {
        color: #64748b;
        font-size: 12px;
        margin-top: 3px;
    }

    .top-strip {
        display: grid;
        grid-template-columns: 1.15fr .85fr .85fr .85fr .85fr .85fr;
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        overflow: hidden;
        box-shadow: 0 10px 28px rgba(15,23,42,.06);
        margin-bottom: 14px;
    }

    .top-item {
        padding: 15px 18px;
        border-right: 1px solid #e2e8f0;
        min-height: 68px;
    }

    .top-item:last-child { border-right: none; }

    .top-label {
        font-size: 12px;
        color: #64748b;
        margin-bottom: 5px;
        font-weight: 700;
    }

    .top-value {
        font-size: 18px;
        color: #0f172a;
        font-weight: 900;
    }

    .concept-box {
        background: linear-gradient(135deg, #eef2ff, #eff6ff);
        color: #0f172a;
        border: 1px solid #bfdbfe;
        border-radius: 16px;
        padding: 18px 22px;
        margin-bottom: 14px;
        box-shadow: 0 10px 24px rgba(15,23,42,.06);
    }

    .concept-title {
        font-size: 16px;
        line-height: 1.55;
    }

    .concept-title strong {
        color: #1e40af;
    }

    .page-title {
        font-size: 29px;
        font-weight: 950;
        color: #0f172a;
        letter-spacing: -0.7px;
        margin-top: 8px;
        margin-bottom: 0px;
    }

    .page-subtitle {
        color: #64748b;
        font-size: 15px;
        margin-bottom: 14px;
    }

    .mini-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 16px;
        min-height: 130px;
        box-shadow: 0 8px 22px rgba(15,23,42,.045);
    }

    .metric-label {
        font-size: 12px;
        color: #64748b;
        margin-bottom: 7px;
        font-weight: 750;
    }

    .metric-value {
        color: #0f172a;
        font-size: 23px;
        font-weight: 950;
        letter-spacing: -0.5px;
        margin-bottom: 9px;
    }

    .metric-delta {
        display: inline-block;
        background: #ecfdf5;
        color: #16a34a;
        padding: 5px 9px;
        border-radius: 999px;
        font-weight: 850;
        font-size: 12px;
        margin-bottom: 7px;
    }

    .metric-note {
        color: #64748b;
        font-size: 12px;
    }

    .icon-circle {
        width: 42px;
        height: 42px;
        border-radius: 999px;
        display: grid;
        place-items: center;
        font-size: 22px;
        margin-bottom: 10px;
    }

    .i-green { background: #dcfce7; color: #16a34a; }
    .i-blue { background: #dbeafe; color: #2563eb; }
    .i-purple { background: #f3e8ff; color: #7c3aed; }
    .i-orange { background: #ffedd5; color: #f97316; }
    .i-red { background: #ffe4e6; color: #e11d48; }
    .i-yellow { background: #fef3c7; color: #d97706; }

    .section-title {
        color: #0f172a;
        font-weight: 950;
        font-size: 22px;
        margin: 18px 0 12px;
        letter-spacing: -.2px;
    }

    .trend-panel-title {
        color: #0f172a;
        font-size: 18px;
        font-weight: 950;
        margin-bottom: 10px;
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    .trend-info {
        color: #94a3b8;
        border: 1px solid #cbd5e1;
        width: 22px;
        height: 22px;
        border-radius: 999px;
        display: inline-grid;
        place-items: center;
        font-size: 13px;
    }

    .trend-stat {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 14px;
        padding: 12px 14px;
        min-height: 86px;
    }

    .trend-stat-label {
        color: #475569;
        font-size: 12px;
        font-weight: 800;
        margin-bottom: 5px;
    }

    .trend-stat-value {
        color: #0f172a;
        font-size: 22px;
        font-weight: 950;
        letter-spacing: -.4px;
    }

    .trend-stat-sub {
        color: #64748b;
        font-size: 12px;
        margin-top: 4px;
    }

    .trend-up { color: #16a34a; font-weight: 900; }
    .trend-red { color: #e11d48; font-weight: 900; }
    .trend-green { color: #16a34a; font-weight: 900; }

    .brief-row {
        display: flex;
        gap: 10px;
        align-items: flex-start;
        padding: 9px 0;
        color: #334155;
        font-size: 14px;
        border-bottom: 1px solid #e2e8f0;
    }

    .brief-row:last-child { border-bottom: none; }

    .dot {
        width: 11px;
        height: 11px;
        border-radius: 999px;
        margin-top: 4px;
        flex: 0 0 auto;
    }

    .dot-green { background: #22c55e; }
    .dot-yellow { background: #eab308; }
    .dot-red { background: #ef4444; }

    .alert-card {
        border: 1px solid #e2e8f0;
        background: #f8fafc;
        border-radius: 12px;
        padding: 12px 13px;
        margin-bottom: 10px;
        display: flex;
        justify-content: space-between;
        gap: 12px;
    }

    .alert-title {
        font-weight: 900;
        color: #d97706;
        font-size: 13px;
    }

    .alert-text {
        color: #475569;
        font-size: 12px;
        margin-top: 3px;
    }

    .alert-time {
        color: #2563eb;
        font-size: 12px;
        white-space: nowrap;
    }

    .question-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 17px;
        min-height: 160px;
        box-shadow: 0 8px 22px rgba(15,23,42,.045);
    }

    .q-title {
        color: #0f172a;
        font-weight: 950;
        font-size: 15px;
        margin-bottom: 8px;
    }

    .q-answer {
        color: #475569;
        font-size: 12px;
        line-height: 1.45;
        margin-bottom: 10px;
    }

    .footer-note {
        color: #64748b;
        font-size: 12px;
        margin-top: 18px;
        display: flex;
        justify-content: space-between;
    }

    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 16px;
        padding: 16px;
        box-shadow: 0 8px 22px rgba(15,23,42,.045);
    }

    div[data-testid="stDataFrame"] {
        border-radius: 13px;
        overflow: hidden;
    }

    .stButton > button {
        border-radius: 12px;
        border: 1px solid #cbd5e1;
        background: white;
        color: #0f172a;
        font-weight: 850;
        min-height: 42px;
    }

    .stButton > button:hover {
        border-color: #2563eb;
        color: #1d4ed8;
        background: #eff6ff;
    }

    section[data-testid="stSidebar"] .stButton > button {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        color: #334155;
        text-align: left;
        justify-content: flex-start;
        box-shadow: 0 4px 12px rgba(15,23,42,.025);
    }

    section[data-testid="stSidebar"] .stButton > button:hover {
        background: #eff6ff;
        border-color: #bfdbfe;
        color: #1d4ed8;
    }

    .brief-hero {
        background: linear-gradient(135deg, #0f172a, #1e3a8a);
        color: white;
        border-radius: 20px;
        padding: 22px 24px;
        box-shadow: 0 18px 36px rgba(15,23,42,.12);
        margin-bottom: 16px;
    }

    .brief-hero-title {
        font-size: 28px;
        font-weight: 950;
        letter-spacing: -.8px;
        margin-bottom: 6px;
    }

    .brief-hero-sub {
        color: #dbeafe;
        font-size: 14px;
        line-height: 1.5;
    }

    .score-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 18px;
        box-shadow: 0 8px 22px rgba(15,23,42,.045);
    }

    .score-big {
        font-size: 44px;
        font-weight: 950;
        color: #2563eb;
        line-height: 1;
    }

    .progress-row {
        margin: 10px 0;
    }

    .progress-label {
        display:flex;
        justify-content:space-between;
        color:#334155;
        font-size:12px;
        font-weight:800;
        margin-bottom:5px;
    }

    .progress-track {
        height: 8px;
        background: #e2e8f0;
        border-radius: 999px;
        overflow:hidden;
    }

    .progress-fill {
        height: 8px;
        background: linear-gradient(90deg, #2563eb, #22c55e);
        border-radius: 999px;
    }

    .meeting-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 18px;
        padding: 18px;
        box-shadow: 0 8px 22px rgba(15,23,42,.045);
        min-height: 150px;
    }

    .meeting-label {
        color:#64748b;
        font-size:12px;
        font-weight:800;
        margin-bottom:8px;
    }

    .meeting-value {
        color:#0f172a;
        font-size:26px;
        font-weight:950;
        margin-bottom:8px;
    }

    .action-card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-left: 5px solid #2563eb;
        border-radius: 16px;
        padding: 14px 16px;
        margin-bottom: 10px;
    }

    .action-title {
        color:#0f172a;
        font-weight:950;
        margin-bottom:4px;
    }

    .action-text {
        color:#475569;
        font-size:13px;
        line-height:1.45;
    }

    .doctor-rank {
        background:white;
        border:1px solid #e2e8f0;
        border-radius:16px;
        padding:15px;
        display:flex;
        align-items:center;
        justify-content:space-between;
        margin-bottom:10px;
        box-shadow: 0 8px 18px rgba(15,23,42,.035);
    }

    .rank-left {
        display:flex;
        gap:13px;
        align-items:center;
    }

    .rank-no {
        width:34px;
        height:34px;
        border-radius:999px;
        background:#eff6ff;
        color:#1d4ed8;
        display:grid;
        place-items:center;
        font-weight:950;
    }

    .heatmap-cell {
        text-align:center;
        padding:10px 6px;
        border-radius:12px;
        font-weight:900;
        font-size:13px;
    }

    .heat-low { background:#dcfce7; color:#166534; }
    .heat-mid { background:#fef3c7; color:#92400e; }
    .heat-high { background:#ffe4e6; color:#be123c; }
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


def date_long_tr(d) -> str:
    months = {
        1: "Ocak", 2: "Şubat", 3: "Mart", 4: "Nisan", 5: "Mayıs", 6: "Haziran",
        7: "Temmuz", 8: "Ağustos", 9: "Eylül", 10: "Ekim", 11: "Kasım", 12: "Aralık"
    }
    d = pd.to_datetime(d)
    return f"{d.day} {months.get(d.month, '')} {d.year}"


def parse_excel_dates(series: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_datetime(series, unit="D", origin="1899-12-30", errors="coerce")

    numeric = pd.to_numeric(series, errors="coerce")
    if numeric.notna().mean() > 0.70 and numeric.dropna().between(30000, 60000).mean() > 0.70:
        return pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")

    return pd.to_datetime(series, errors="coerce", dayfirst=True)


def goto(page: str):
    st.session_state.active_page = page
    st.rerun()


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


def chart_layout_light(fig, height=325):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#334155", size=12),
        margin=dict(l=8, r=8, t=18, b=35),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5),
        xaxis=dict(gridcolor="#e2e8f0", zeroline=False, tickformat="%d %b", type="date"),
        yaxis=dict(gridcolor="#e2e8f0", zeroline=False),
        hovermode="x unified",
    )
    return fig


def section_header(title, subtitle=None):
    st.markdown(f'<div class="page-title">{title}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="page-subtitle">{subtitle}</div>', unsafe_allow_html=True)


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
    st.markdown('<div class="ayca-logo">✣ AYÇA Insight</div>', unsafe_allow_html=True)
    st.markdown('<div class="ayca-sub">Hospital Analytics</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="version-badge">{APP_VERSION}</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sidebar-org"><div class="sidebar-org-title">HG Hospital</div><div class="sidebar-org-sub">Premium Plan · Yönetici Paneli</div></div>',
        unsafe_allow_html=True,
    )

    st.markdown("### Menü")
    menu_icons = {
        "CEO Dashboard": "🏠",
        "Finans Merkezi": "💰",
        "Doktor Intelligence": "👨‍⚕️",
        "Operasyon": "🏥",
        "Yönetim Toplantısı": "🎯",
        "Kalite & Hasta Deneyimi": "🙂",
        "Hasta Analitiği": "👥",
        "Stok & Satın Alma": "📦",
        "AYÇA Co-Pilot": "🤖",
    }
    for item in NAV_ITEMS:
        prefix = "●" if st.session_state.active_page == item else menu_icons.get(item, "•")
        label = f"{prefix} {item}"
        if st.button(label, key=f"side_{item}", use_container_width=True):
            goto(item)

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

total_revenue = float(gunluk["Günlük Ciro"].sum())
total_patient = int(gunluk["Toplam Hasta"].sum())
total_emergency = int(gunluk["Acil Hasta"].sum())
peak_patient_row = gunluk.loc[gunluk["Toplam Hasta"].idxmax()]
peak_emergency_row = gunluk.loc[gunluk["Acil Hasta"].idxmax()]


# =========================================================
# Top Area
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
            <div class="top-value">{date_long_tr(max_date)}</div>
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
            <strong>V0.6 Konsept:</strong> Bu ekran hastane direktörü, genel müdür ve yönetim ekibinin sabah tek bakışta
            hastanenin finansal, operasyonel ve kalite durumunu görmesi için tasarlanmıştır.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

quick_cols = st.columns(5)
quick_pages = ["CEO Dashboard", "Finans Merkezi", "Doktor Intelligence", "Yönetim Toplantısı", "AYÇA Co-Pilot"]
for i, page in enumerate(quick_pages):
    with quick_cols[i]:
        if st.button(page, key=f"quick_{page}", use_container_width=True):
            goto(page)


# =========================================================
# Question cards
# =========================================================
def render_question_cards(context="main"):
    st.markdown('<div class="section-title">Yönetici İçin Önerilen Sorular</div>', unsafe_allow_html=True)

    brans_sum = brans.groupby("Branş", as_index=False).agg(Gelir=("Gelir", "sum"), Kar=("Kar", "sum"), Hasta=("Hasta", "sum"))
    top_rev = brans_sum.sort_values("Gelir", ascending=False).iloc[0]
    top_profit = brans_sum.sort_values("Kar", ascending=False).iloc[0]

    doc_sum = doktor.groupby(["Doktor", "Branş"], as_index=False).agg(
        Gelir=("Gelir", "sum"),
        Memnuniyet=("Memnuniyet %", "mean"),
        Hasta=("Hasta", "sum")
    )
    top_doc = doc_sum.sort_values("Gelir", ascending=False).iloc[0]
    top_satis = doc_sum.sort_values("Memnuniyet", ascending=False).iloc[0]
    top_complaint = kalite.sort_values("Şikayet Sayısı", ascending=False).iloc[0]

    cards = [
        ("📈", "Bu ay ciro neden arttı?", f"Cevap: {top_rev['Branş']} branşındaki güçlü gelir artışı toplam ciroyu destekliyor.", "Finans Merkezi"),
        ("🟣", "En kârlı branş hangisi?", f"Cevap: {top_profit['Branş']} yaklaşık {money_fmt(top_profit['Kar'])} kâr ile öne çıkıyor.", "Finans Merkezi"),
        ("🕘", "Bekleme süresi neden arttı?", f"Cevap: Hasta hacmi {int_fmt(today['Toplam Hasta'])} seviyesinde. Randevu kapasitesi ve yoğun branşlar birlikte incelenmeli.", "Operasyon"),
        ("🙂", "Hasta memnuniyeti nasıl artırılır?", f"Cevap: En büyük şikayet konusu '{top_complaint['Konu']}'. Bu alanda çözüm süresi düşürülmeli.", "Kalite & Hasta Deneyimi"),
        ("👨‍⚕️", "Hangi doktorlar öne çıkıyor?", f"Cevap: Gelirde {top_doc['Doktor']}, memnuniyette {top_satis['Doktor']} öne çıkıyor.", "Doktor Intelligence"),
        ("🛡️", "Hangi alanlarda risk var?", f"Cevap: MR kullanımı {pct_fmt(today['MR Kullanım %'])}, doluluk {pct_fmt(today['Doluluk %'])}. Kapasite riski takip edilmeli.", "Operasyon"),
    ]

    cols = st.columns(3)
    for i, (icon, title, answer, target) in enumerate(cards):
        with cols[i % 3]:
            st.markdown(
                f"""
                <div class="question-card">
                    <div class="icon-circle i-blue">{icon}</div>
                    <div class="q-title">{title}</div>
                    <div class="q-answer">{answer}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if st.button(f"Detayları Gör → {target}", key=f"detail_{context}_{i}", use_container_width=True):
                goto(target)


# =========================================================
# Pages
# =========================================================
def render_ceo_dashboard():
    section_header("CEO Dashboard", "Yönetici Özeti")

    st.markdown(
        f"""
        <div class="brief-hero">
            <div class="brief-hero-title">Günaydın Ahmet Bey</div>
            <div class="brief-hero-sub">
                Bugün dikkat edilmesi gereken 4 konu var: MR kapasitesi {pct_fmt(today['MR Kullanım %'])}, ortalama bekleme {int(today['Ortalama Bekleme Dk'])} dk,
                günlük gelir hedefi {pct_fmt(today['Hedef Gerçekleşme %'])} ve hasta memnuniyeti {pct_fmt(today['Memnuniyet %'])}.
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    score_left, score_right = st.columns([1, 2])
    with score_left:
        st.markdown(
            f"""
            <div class="score-card">
                <div class="metric-label">HG Hospital Sağlık Skoru</div>
                <div class="score-big">{int(today['Risk Skoru'])}/100</div>
                <div class="metric-note">Finans, operasyon, kalite ve hasta deneyimi birleşik skoru</div>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with score_right:
        st.markdown(
            f"""
            <div class="score-card">
                <div class="progress-row"><div class="progress-label"><span>Finans</span><span>90</span></div><div class="progress-track"><div class="progress-fill" style="width:90%;"></div></div></div>
                <div class="progress-row"><div class="progress-label"><span>Operasyon</span><span>76</span></div><div class="progress-track"><div class="progress-fill" style="width:76%;"></div></div></div>
                <div class="progress-row"><div class="progress-label"><span>Kalite</span><span>82</span></div><div class="progress-track"><div class="progress-fill" style="width:82%;"></div></div></div>
                <div class="progress-row"><div class="progress-label"><span>Memnuniyet</span><span>{int(today['Memnuniyet %'])}</span></div><div class="progress-track"><div class="progress-fill" style="width:{int(today['Memnuniyet %'])}%;"></div></div></div>
            </div>
            """,
            unsafe_allow_html=True,
        )

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
        with st.container(border=True):
            st.markdown('<div class="trend-panel-title"><span>Günlük Ciro Trendi</span><span class="trend-info">i</span></div>', unsafe_allow_html=True)
            s1, s2, s3 = st.columns(3)
            with s1:
                st.markdown(f'<div class="trend-stat"><div class="trend-stat-label">Ort. Günlük Ciro</div><div class="trend-stat-value">{money_fmt(avg_daily_revenue)}</div><div class="trend-stat-sub"><span class="trend-up">↑ %12,4</span> · Önceki 30 güne göre</div></div>', unsafe_allow_html=True)
            with s2:
                st.markdown(f'<div class="trend-stat"><div class="trend-stat-label">En Yüksek Gün</div><div class="trend-stat-value">{money_fmt(max_revenue_row["Günlük Ciro"])}</div><div class="trend-stat-sub"><span class="trend-green">{date_long_tr(max_revenue_row["Tarih"])}</span></div></div>', unsafe_allow_html=True)
            with s3:
                st.markdown(f'<div class="trend-stat"><div class="trend-stat-label">En Düşük Gün</div><div class="trend-stat-value">{money_fmt(min_revenue_row["Günlük Ciro"])}</div><div class="trend-stat-sub"><span class="trend-red">{date_long_tr(min_revenue_row["Tarih"])}</span></div></div>', unsafe_allow_html=True)

            fig = go.Figure()
            fig.add_trace(
                go.Scatter(
                    x=last_30["Tarih"],
                    y=last_30["Günlük Ciro"],
                    mode="lines+markers",
                    name="Günlük Ciro (TL)",
                    line=dict(width=3, color="#2563eb"),
                    marker=dict(size=6, color="#bfdbfe", line=dict(width=2, color="#2563eb")),
                    fill="tozeroy",
                    fillcolor="rgba(37,99,235,.13)",
                )
            )
            fig.update_yaxes(tickformat="~s", ticksuffix=" TL", rangemode="tozero")
            st.plotly_chart(chart_layout_light(fig, 330), use_container_width=True)

    with c2:
        with st.container(border=True):
            st.markdown('<div class="trend-panel-title"><span>Hasta Hacmi Trendi</span><span class="trend-info">i</span></div>', unsafe_allow_html=True)
            s1, s2 = st.columns(2)
            with s1:
                st.markdown(f'<div class="trend-stat"><div class="trend-stat-label">Ort. Toplam Hasta</div><div class="trend-stat-value">{int_fmt(avg_total_patient)}</div><div class="trend-stat-sub"><span class="trend-up">↑ %8,7</span> · Önceki 30 güne göre</div></div>', unsafe_allow_html=True)
            with s2:
                st.markdown(f'<div class="trend-stat"><div class="trend-stat-label">Ort. Acil Hasta</div><div class="trend-stat-value">{int_fmt(avg_emergency_patient)}</div><div class="trend-stat-sub"><span class="trend-up">↑ %6,1</span> · Önceki 30 güne göre</div></div>', unsafe_allow_html=True)

            poliklinik = last_30["Toplam Hasta"] - last_30["Acil Hasta"]
            fig = go.Figure()
            fig.add_trace(go.Bar(x=last_30["Tarih"], y=poliklinik, name="Poliklinik Hasta", marker_color="#3b82f6"))
            fig.add_trace(go.Bar(x=last_30["Tarih"], y=last_30["Acil Hasta"], name="Acil Hasta", marker_color="#fb7185"))
            fig.add_trace(go.Scatter(x=last_30["Tarih"], y=last_30["Toplam Hasta"], name="Toplam Hasta", mode="lines+markers", line=dict(width=3, color="#0f172a"), marker=dict(size=5, color="#0f172a")))
            fig.update_layout(barmode="stack")
            fig.update_yaxes(title="Hasta Sayısı")
            st.plotly_chart(chart_layout_light(fig, 330), use_container_width=True)

    b1, b2, b3, b4 = st.columns(4)
    with b1: metric_card("👥", "i-blue", "Toplam Hasta", int_fmt(total_patient), "↑ %8,7", "Son 30 gün")
    with b2: metric_card("🚨", "i-red", "Acil Hasta", int_fmt(total_emergency), "↑ %6,1", "Son 30 gün")
    with b3: metric_card("📈", "i-green", "En Yoğun Gün", int_fmt(peak_patient_row["Toplam Hasta"]), "Toplam Hasta", date_long_tr(peak_patient_row["Tarih"]))
    with b4: metric_card("🗓️", "i-orange", "En Yoğun Acil Gün", int_fmt(peak_emergency_row["Acil Hasta"]), "Acil Hasta", date_long_tr(peak_emergency_row["Tarih"]))

    left, right = st.columns(2)
    with left:
        with st.container(border=True):
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

    with right:
        with st.container(border=True):
            st.markdown('<div class="section-title">En Kritik Uyarılar</div>', unsafe_allow_html=True)
            alerts = [
                (f"Radyoloji MR kullanım oranı {pct_fmt(today['MR Kullanım %'])}", "Kapasite limitine çok yakın.", "09:15"),
                (f"Kardiyoloji bekleme süresi {int(today['Ortalama Bekleme Dk'])} dk", "Hedefin üzerinde. Randevu kapasitesi incelenmeli.", "09:05"),
                (f"Yoğun bakım doluluk oranı {pct_fmt(today['Doluluk %'])}", "Yüksek doluluk. Takip önerilir.", "08:50"),
            ]
            for title, text, time in alerts:
                st.markdown(f'<div class="alert-card"><div><div class="alert-title">{title}</div><div class="alert-text">{text}</div></div><div class="alert-time">{time}</div></div>', unsafe_allow_html=True)

    render_question_cards("ceo")



def render_management_meeting():
    section_header("Yönetim Toplantısı", "Haftalık yönetici özeti ve aksiyon listesi")

    brans_sum = brans.groupby("Branş", as_index=False).agg(Gelir=("Gelir", "sum"), Kar=("Kar", "sum"), Hasta=("Hasta", "sum"))
    top_rev = brans_sum.sort_values("Gelir", ascending=False).iloc[0]
    top_profit = brans_sum.sort_values("Kar", ascending=False).iloc[0]
    top_complaint = kalite.sort_values("Şikayet Sayısı", ascending=False).iloc[0]

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="meeting-card"><div class="meeting-label">Gelir Trend</div><div class="meeting-value">+%8,7</div><div class="metric-note">Son 30 gün</div></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="meeting-card"><div class="meeting-label">Hasta Trend</div><div class="meeting-value">+%6,1</div><div class="metric-note">{int_fmt(total_patient)} toplam hasta</div></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="meeting-card"><div class="meeting-label">En İyi Branş</div><div class="meeting-value">{top_rev["Branş"]}</div><div class="metric-note">{money_fmt(top_rev["Gelir"])} gelir</div></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="meeting-card"><div class="meeting-label">En Riskli Konu</div><div class="meeting-value">{top_complaint["Konu"]}</div><div class="metric-note">{int_fmt(top_complaint["Şikayet Sayısı"])} şikayet</div></div>', unsafe_allow_html=True)

    left, right = st.columns([1, 1])
    with left:
        st.markdown('<div class="section-title">Toplantı Gündemi</div>', unsafe_allow_html=True)
        actions = [
            ("Radyoloji kapasitesi", f"MR kullanım oranı {pct_fmt(today['MR Kullanım %'])}. Ek slot veya mesai planı değerlendirilmeli."),
            ("Bekleme süresi", f"Ortalama bekleme {int(today['Ortalama Bekleme Dk'])} dk. Kardiyoloji ve acil servis öncelikli incelenmeli."),
            ("Finans fırsatı", f"{top_profit['Branş']} kârlılıkta öne çıkıyor. Bu branşta kapasite artırımı değerlendirilebilir."),
            ("Hasta deneyimi", f"En büyük şikayet konusu {top_complaint['Konu']}. Çözüm süresi ve kök neden çalışması önerilir."),
        ]
        for title, text in actions:
            st.markdown(f'<div class="action-card"><div class="action-title">{title}</div><div class="action-text">{text}</div></div>', unsafe_allow_html=True)

    with right:
        st.markdown('<div class="section-title">Yönetim Kurulu Raporu</div>', unsafe_allow_html=True)
        st.info("Bu konseptte rapor içeriği dashboard verilerinden otomatik oluşturulacak şekilde tasarlanmıştır.")
        report_text = f"""
HG Hospital Haftalık Yönetici Özeti

- Toplam gelir: {money_fmt(total_revenue)}
- Toplam hasta: {int_fmt(total_patient)}
- En yüksek gelir üreten branş: {top_rev['Branş']}
- En kârlı branş: {top_profit['Branş']}
- En kritik operasyonel risk: MR kapasitesi {pct_fmt(today['MR Kullanım %'])}
- Ana hasta deneyimi problemi: {top_complaint['Konu']}
"""
        st.download_button(
            "📄 Yönetim Kurulu Raporu İndir",
            data=report_text,
            file_name="ayca_hospital_yonetici_ozeti.txt",
            mime="text/plain",
            use_container_width=True,
        )

    render_question_cards("meeting")


def render_finance():
    section_header("Finans Merkezi", "Branş, kanal ve kârlılık analizi")

    cols = st.columns(4)
    with cols[0]: metric_card("💰", "i-green", "Toplam Ciro", money_fmt(total_revenue), "↑ %8,7", "Son 30 gün")
    with cols[1]: metric_card("🏦", "i-blue", "SGK Geliri", money_fmt(gunluk["SGK Geliri"].sum()), "Kanal", "Son 30 gün")
    with cols[2]: metric_card("🛡️", "i-purple", "Özel Sigorta", money_fmt(gunluk["Özel Sigorta Geliri"].sum()), "Kanal", "Son 30 gün")
    with cols[3]: metric_card("💵", "i-orange", "Nakit Gelir", money_fmt(gunluk["Nakit Gelir"].sum()), "Kanal", "Son 30 gün")

    brans_sum = brans.groupby("Branş", as_index=False).agg(Gelir=("Gelir", "sum"), Kar=("Kar", "sum"), Hasta=("Hasta", "sum"), Memnuniyet=("Memnuniyet %", "mean"))
    brans_sum["Hasta Başı Gelir"] = brans_sum["Gelir"] / brans_sum["Hasta"]

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(brans_sum.sort_values("Gelir", ascending=True), x="Gelir", y="Branş", orientation="h", title="Branş Bazlı Gelir")
        fig.update_traces(marker_color="#2563eb")
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)
    with c2:
        fig = px.bar(brans_sum.sort_values("Kar", ascending=True), x="Kar", y="Branş", orientation="h", title="Branş Bazlı Kâr")
        fig.update_traces(marker_color="#16a34a")
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)

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

    st.markdown('<div class="section-title">Doktor Performans Ligi</div>', unsafe_allow_html=True)
    rank_df = dsum.sort_values(["Gelir", "Memnuniyet"], ascending=False).head(5).reset_index(drop=True)
    for idx, row in rank_df.iterrows():
        st.markdown(
            f"""
            <div class="doctor-rank">
                <div class="rank-left">
                    <div class="rank-no">{idx+1}</div>
                    <div>
                        <div class="q-title">{row['Doktor']}</div>
                        <div class="metric-note">{row['Branş']} · Gelir: {money_fmt(row['Gelir'])} · Memnuniyet: {pct_fmt(row['Memnuniyet'])}</div>
                    </div>
                </div>
                <div class="trend-up">★★★★★</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    c1, c2 = st.columns(2)
    with c1:
        fig = px.scatter(dsum, x="Hasta", y="Gelir", size="Memnuniyet", color="Branş", hover_name="Doktor", title="Doktor Gelir / Hasta Matrisi")
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)
    with c2:
        fig = px.bar(dsum.sort_values("Gelir", ascending=False).head(10), x="Doktor", y="Gelir", color="Branş", title="Top 10 Doktor Gelir")
        fig.update_layout(xaxis_tickangle=-30)
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)

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
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)
    with c2:
        fig = px.bar(latest_ops.sort_values("Doluluk %"), x="Doluluk %", y="Birim", orientation="h", title="Birim Bazlı Doluluk")
        fig.update_traces(marker_color="#2563eb")
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)

    st.markdown('<div class="section-title">Operasyon Isı Haritası</div>', unsafe_allow_html=True)
    heat_cols = st.columns(6)
    days = ["Pzt", "Sal", "Çar", "Per", "Cum", "Cmt"]
    slots = ["08-10", "10-12", "12-14", "14-16"]
    heat_values = [
        ["🟢", "🟡", "🟡", "🔴", "🔴", "🟡"],
        ["🟡", "🟡", "🔴", "🔴", "🔴", "🔴"],
        ["🟢", "🟢", "🟡", "🟡", "🔴", "🟡"],
        ["🟢", "🟡", "🟡", "🔴", "🟡", "🟢"],
    ]
    st.write("Yoğunluk: 🟢 düşük · 🟡 orta · 🔴 yüksek")
    for row_i, slot in enumerate(slots):
        cols_h = st.columns(7)
        with cols_h[0]:
            st.markdown(f"**{slot}**")
        for col_i, day in enumerate(days):
            emoji = heat_values[row_i][col_i]
            cls = "heat-low" if emoji == "🟢" else ("heat-mid" if emoji == "🟡" else "heat-high")
            with cols_h[col_i + 1]:
                st.markdown(f'<div class="heatmap-cell {cls}">{day}<br>{emoji}</div>', unsafe_allow_html=True)

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
        fig.update_traces(marker_color="#e11d48")
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)
    with c2:
        fig = px.scatter(kalite, x="Şikayet Sayısı", y="Ortalama Çözüm Saat", size="Etki Skoru", color="Öncelik", hover_name="Konu", title="Şikayet Etki Matrisi")
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)

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
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)
    with c2:
        fig = px.pie(hasta, names="Branş", values="Yeni Hasta", title="Yeni Hasta Dağılımı", hole=.45)
        st.plotly_chart(chart_layout_light(fig, 420), use_container_width=True)

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
        st.plotly_chart(chart_layout_light(fig, 460), use_container_width=True)
    with c2:
        st.markdown('<div class="section-title">Kritik / Riskli Kalemler</div>', unsafe_allow_html=True)
        st.dataframe(critical, use_container_width=True, hide_index=True)

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
        "Bu ay ciro neden arttı?": f"Ciro artışının ana nedeni {top_rev['Branş']} branşındaki yüksek gelir katkısıdır. Bu branş {money_fmt(top_rev['Gelir'])} gelir üretmiştir.",
        "En kârlı branş hangisi?": f"En kârlı branş {top_profit['Branş']} görünüyor. Yaklaşık {money_fmt(top_profit['Kar'])} kâr üretmiştir.",
        "Bekleme süresi neden arttı?": f"Bekleme süresi hasta hacmi ve kapasite baskısından etkileniyor. Bugünkü toplam hasta {int_fmt(today['Toplam Hasta'])}, ortalama bekleme {int(today['Ortalama Bekleme Dk'])} dakikadır.",
        "Hasta memnuniyeti nasıl artırılır?": f"En çok şikayet konusu '{top_complaint['Konu']}'. Bu başlıkta çözüm süresi düşürülürse memnuniyet artışı beklenir.",
        "Hangi doktorlar öne çıkıyor?": f"Gelir açısından {top_doc['Doktor']} öne çıkıyor. Branşı {top_doc['Branş']}, toplam geliri {money_fmt(top_doc['Gelir'])}.",
        "Hangi alanlarda risk var?": f"MR kullanım oranı {pct_fmt(today['MR Kullanım %'])}, doluluk {pct_fmt(today['Doluluk %'])}. Kapasite ve bekleme riski birlikte takip edilmelidir.",
    }

    if "copilot_question" not in st.session_state:
        st.session_state.copilot_question = list(qa.keys())[0]

    st.markdown('<div class="section-title">Hazır Yönetici Soruları</div>', unsafe_allow_html=True)
    qcols = st.columns(3)
    for i, q in enumerate(qa.keys()):
        with qcols[i % 3]:
            if st.button(q, key=f"copilot_q_{i}", use_container_width=True):
                st.session_state.copilot_question = q

    selected = st.session_state.copilot_question
    st.markdown("### AYÇA Yanıtı")
    st.success(qa[selected])

    render_question_cards("copilot")


# =========================================================
# Router
# =========================================================
page = st.session_state.active_page

if page == "CEO Dashboard":
    render_ceo_dashboard()
elif page == "Finans Merkezi":
    render_finance()
elif page == "Doktor Intelligence":
    render_doctors()
elif page == "Operasyon":
    render_operations()
elif page == "Yönetim Toplantısı":
    render_management_meeting()
elif page == "Kalite & Hasta Deneyimi":
    render_quality()
elif page == "Hasta Analitiği":
    render_patients()
elif page == "Stok & Satın Alma":
    render_stock()
elif page == "AYÇA Co-Pilot":
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
