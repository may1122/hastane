
# -*- coding: utf-8 -*-
"""
AYÇA Insight Hospital - Demo App
Streamlit dashboard concept for hospital management decision support.

Run:
    pip install streamlit pandas openpyxl plotly
    streamlit run ayca_hospital_app.py
"""

from __future__ import annotations

from pathlib import Path
import math
import pandas as pd
import plotly.express as px
import streamlit as st


APP_VERSION = "V0.3 Hospital Concept"
# -----------------------------
# Page Config
# -----------------------------
st.set_page_config(
    page_title="AYÇA Insight Hospital",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded",
)


# -----------------------------
# Styling
# -----------------------------
st.markdown(
    """
    <style>
    :root {
        --ayca-navy: #0f172a;
        --ayca-blue: #2563eb;
        --ayca-soft-blue: #eff6ff;
        --ayca-green: #16a34a;
        --ayca-soft-green: #ecfdf5;
        --ayca-orange: #f97316;
        --ayca-soft-orange: #fff7ed;
        --ayca-purple: #7c3aed;
        --ayca-soft-purple: #f5f3ff;
        --ayca-red: #dc2626;
        --ayca-soft-red: #fef2f2;
        --ayca-gray: #64748b;
        --card-border: #e2e8f0;
    }

    .main .block-container {
        padding-top: 1.2rem;
        padding-bottom: 2rem;
    }

    .ayca-topbar {
        background: linear-gradient(135deg, #0f172a 0%, #1e3a8a 100%);
        border-radius: 22px;
        padding: 22px 26px;
        color: white;
        margin-bottom: 18px;
        box-shadow: 0 14px 35px rgba(15, 23, 42, 0.18);
    }

    .ayca-title {
        font-size: 30px;
        font-weight: 800;
        letter-spacing: -0.6px;
        margin-bottom: 4px;
    }

    .ayca-subtitle {
        color: rgba(255,255,255,.82);
        font-size: 14px;
    }

    .version-pill {
        display: inline-block;
        background: rgba(255,255,255,.12);
        border: 1px solid rgba(255,255,255,.22);
        padding: 6px 11px;
        border-radius: 999px;
        color: white;
        font-size: 12px;
        margin-top: 8px;
    }

    .kpi-card {
        background: white;
        border: 1px solid var(--card-border);
        border-radius: 18px;
        padding: 16px 18px;
        min-height: 112px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.04);
    }

    .kpi-label {
        font-size: 13px;
        color: #64748b;
        font-weight: 650;
        margin-bottom: 8px;
    }

    .kpi-value {
        font-size: 25px;
        font-weight: 800;
        color: #0f172a;
        letter-spacing: -0.5px;
    }

    .kpi-delta-pos { color: #16a34a; font-size: 12px; margin-top: 7px; font-weight: 700; }
    .kpi-delta-neg { color: #dc2626; font-size: 12px; margin-top: 7px; font-weight: 700; }
    .kpi-delta-mid { color: #f97316; font-size: 12px; margin-top: 7px; font-weight: 700; }

    .section-title {
        font-size: 21px;
        font-weight: 800;
        color: #0f172a;
        margin: 8px 0 12px 0;
        letter-spacing: -0.3px;
    }

    .briefing-card {
        background: white;
        border: 1px solid var(--card-border);
        border-radius: 18px;
        padding: 17px 19px;
        margin-bottom: 12px;
        box-shadow: 0 6px 18px rgba(15, 23, 42, 0.035);
    }

    .briefing-title {
        font-weight: 800;
        color: #0f172a;
        margin-bottom: 6px;
    }

    .briefing-text {
        color: #334155;
        font-size: 14px;
        line-height: 1.45;
    }

    .ai-box {
        background: linear-gradient(135deg, #f8fafc 0%, #eef2ff 100%);
        border: 1px solid #dbeafe;
        border-radius: 18px;
        padding: 18px 20px;
        margin: 10px 0 16px 0;
    }

    .ai-box strong { color: #1e3a8a; }

    .module-chip {
        display: inline-block;
        padding: 6px 10px;
        border-radius: 999px;
        background: #eff6ff;
        color: #1d4ed8;
        font-size: 12px;
        font-weight: 700;
        margin-right: 6px;
        margin-bottom: 6px;
    }

    [data-testid="stMetricValue"] {
        font-size: 25px;
        font-weight: 800;
    }

    div[data-testid="stDataFrame"] {
        border-radius: 14px;
        overflow: hidden;
    }

    .small-muted {
        color: #64748b;
        font-size: 12px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Helpers
# -----------------------------
def money_fmt(x: float | int | None) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "0 TL"
    abs_x = abs(float(x))
    if abs_x >= 1_000_000:
        return f"{x/1_000_000:,.1f} M TL".replace(",", "X").replace(".", ",").replace("X", ".")
    if abs_x >= 1_000:
        return f"{x/1_000:,.0f} B TL".replace(",", ".")
    return f"{x:,.0f} TL".replace(",", ".")


def int_fmt(x: float | int | None) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "0"
    return f"{int(round(float(x))):,}".replace(",", ".")


def pct_fmt(x: float | int | None, digits: int = 1) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return "%0"
    return f"%{float(x):.{digits}f}".replace(".", ",")


def make_kpi_card(label: str, value: str, delta: str = "", tone: str = "pos"):
    delta_class = {
        "pos": "kpi-delta-pos",
        "neg": "kpi-delta-neg",
        "mid": "kpi-delta-mid",
    }.get(tone, "kpi-delta-mid")

    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="{delta_class}">{delta}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def briefing_item(icon: str, title: str, text: str):
    st.markdown(
        f"""
        <div class="briefing-card">
            <div class="briefing-title">{icon} {title}</div>
            <div class="briefing-text">{text}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_data(uploaded_file):
    """Read AYÇA Hospital Excel template uploaded by the user."""
    xls = pd.ExcelFile(uploaded_file)
    data = {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

    for key in ["Gunluk", "Brans", "Doktor", "Operasyon"]:
        if key in data and "Tarih" in data[key].columns:
            data[key]["Tarih"] = pd.to_datetime(data[key]["Tarih"])

    return data


def safe_sheet(data: dict, name: str) -> pd.DataFrame:
    return data.get(name, pd.DataFrame())


def latest_date(df: pd.DataFrame):
    if df.empty or "Tarih" not in df.columns:
        return None
    return pd.to_datetime(df["Tarih"]).max()


# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.markdown("## 🏥 AYÇA Insight Hospital")
    st.caption("Hastane Yönetim ve Karar Destek Konsept Demo · V0.3")
    st.divider()

    uploaded = st.file_uploader(
        "Hastane Veri Dosyası",
        type=["xlsx"],
        help="AYÇA Hospital veri şablonunu yükleyiniz."
    )

    if uploaded is None:
        st.info("Devam etmek için AYÇA Hospital Excel veri dosyasını yükleyiniz.")
        st.caption("Demo için sample Excel dosyasını yükleyebilirsiniz. Veri dosyası koda gömülü değildir.")
        st.stop()

    st.markdown("### Modüller")
    module = st.radio(
        "Görünüm seç",
        [
            "Yönetici Özeti",
            "Finans Merkezi",
            "Doktor Intelligence",
            "Operasyon Merkezi",
            "Hasta Merkezi",
            "Stok & Satın Alma",
            "Kalite Merkezi",
            "AYÇA Co-Pilot",
        ],
        label_visibility="collapsed",
    )

    st.divider()
    st.caption(APP_VERSION)


try:
    data = load_data(uploaded)
except Exception as e:
    st.error("Excel dosyası okunamadı. Lütfen AYÇA Hospital veri şablonunu kontrol ediniz.")
    st.exception(e)
    st.stop()


required_sheets = ["Gunluk", "Brans", "Doktor", "Hasta", "Operasyon", "Stok", "Kalite"]
missing_sheets = [s for s in required_sheets if s not in data]

if missing_sheets:
    st.error("Excel şablonunda eksik sayfa var: " + ", ".join(missing_sheets))
    st.stop()

gunluk = safe_sheet(data, "Gunluk")
brans = safe_sheet(data, "Brans")
doktor = safe_sheet(data, "Doktor")
hasta = safe_sheet(data, "Hasta")
operasyon = safe_sheet(data, "Operasyon")
stok = safe_sheet(data, "Stok")
kalite = safe_sheet(data, "Kalite")

max_date = latest_date(gunluk)
if max_date is None:
    st.error("Gunluk sayfasında Tarih kolonu bulunamadı.")
    st.stop()

current_day = gunluk[gunluk["Tarih"] == max_date].copy()
today = current_day.iloc[0] if not current_day.empty else gunluk.iloc[-1]
last_7 = gunluk[gunluk["Tarih"] >= max_date - pd.Timedelta(days=6)].copy()
prev_7 = gunluk[(gunluk["Tarih"] < max_date - pd.Timedelta(days=6)) & (gunluk["Tarih"] >= max_date - pd.Timedelta(days=13))].copy()

# -----------------------------
# Top Bar
# -----------------------------
st.markdown(
    f"""
    <div class="ayca-topbar">
        <div class="ayca-title">AYÇA Insight Hospital</div>
        <div class="ayca-subtitle">
            Executive Intelligence · Hospital CEO Dashboard · HG Hospital Konsept Demo · Son güncelleme: {max_date.strftime('%d.%m.%Y')}
        </div>
        <div class="version-pill">{APP_VERSION}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="ai-box">
    <strong>V0.3 Konsept:</strong>
    Bu ekran hastane direktörü, genel müdür ve yönetim ekibinin sabah tek bakışta hastanenin finansal,
    operasyonel ve kalite durumunu görmesi için tasarlanmıştır.
    <br><br>
    <span class="module-chip">CEO Dashboard</span>
    <span class="module-chip">Finans Merkezi</span>
    <span class="module-chip">Doktor Intelligence</span>
    <span class="module-chip">Operasyon Riski</span>
    <span class="module-chip">AYÇA Co-Pilot</span>
    </div>
    """,
    unsafe_allow_html=True,
)


# -----------------------------
# Executive Summary
# -----------------------------
if module == "Yönetici Özeti":
    st.markdown('<div class="section-title">AYÇA Sabah Brifingi</div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1.1, 1.1, 1.1])
    with col1:
        briefing_item(
            "🟢",
            "Hasta hacmi güçlü",
            f"Bugünkü toplam hasta sayısı {int_fmt(today['Toplam Hasta'])}. Son 7 günlük ortalama {int_fmt(last_7['Toplam Hasta'].mean())}.",
        )
        briefing_item(
            "🟡",
            "Bekleme süresi izlenmeli",
            f"Ortalama bekleme süresi {int_fmt(today['Ortalama Bekleme Dk'])} dk. Kardiyoloji ve acil servis tarafında yoğunluk artışı var.",
        )
    with col2:
        briefing_item(
            "🔴",
            "MR kullanım oranı kritik seviyede",
            f"Radyoloji MR kullanım oranı %{today['MR Kullanım %']:.0f}. Kapasite planı gerekebilir.",
        )
        briefing_item(
            "🟢",
            "Ciro hedef üstü",
            f"Günlük ciro {money_fmt(today['Günlük Ciro'])}. Hedef gerçekleşme oranı %{today['Hedef Gerçekleşme %']:.0f}.",
        )
    with col3:
        briefing_item(
            "🟣",
            "Hasta memnuniyeti",
            f"Memnuniyet skoru %{today['Memnuniyet %']:.0f}. Son 7 gün ortalaması %{last_7['Memnuniyet %'].mean():.0f}.",
        )
        briefing_item(
            "🤖",
            "AYÇA yorumu",
            "Gelir artışı güçlü; fakat bekleme süresi ve görüntüleme kapasitesi, büyümenin operasyonel baskı oluşturabileceğini gösteriyor.",
        )

    st.markdown('<div class="section-title">Yönetici KPI Özeti</div>', unsafe_allow_html=True)
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k1: make_kpi_card("Günlük Ciro", money_fmt(today["Günlük Ciro"]), f"Hedef %{today['Hedef Gerçekleşme %']:.0f}", "pos")
    with k2: make_kpi_card("Aylık Ciro", money_fmt(gunluk["Günlük Ciro"].sum()), "Konsept dönem", "pos")
    with k3: make_kpi_card("Toplam Hasta", int_fmt(today["Toplam Hasta"]), "Bugün", "pos")
    with k4: make_kpi_card("Yatan Hasta", int_fmt(today["Yatan Hasta"]), f"Doluluk %{today['Doluluk %']:.0f}", "mid")
    with k5: make_kpi_card("Acil Servis", int_fmt(today["Acil Hasta"]), "Canlı yoğunluk", "mid")
    with k6: make_kpi_card("Risk Skoru", f"{int(today['Risk Skoru'])}/100", "Yönetim skoru", "pos" if today["Risk Skoru"] >= 75 else "mid")

    st.markdown('<div class="section-title">Son 30 Gün Trend</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        fig = px.line(gunluk, x="Tarih", y="Günlük Ciro", markers=True, title="Günlük Ciro Trendi")
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.line(gunluk, x="Tarih", y=["Toplam Hasta", "Acil Hasta"], markers=True, title="Hasta Hacmi")
        fig.update_layout(height=360, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)


elif module == "Finans Merkezi":
    st.markdown('<div class="section-title">Finans Merkezi</div>', unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)
    with f1: make_kpi_card("Toplam Ciro", money_fmt(gunluk["Günlük Ciro"].sum()), "Demo dönem", "pos")
    with f2: make_kpi_card("Günlük Ortalama", money_fmt(gunluk["Günlük Ciro"].mean()), "Son 30 gün", "pos")
    with f3: make_kpi_card("SGK Geliri", money_fmt(gunluk["SGK Geliri"].sum()), "Kanal bazlı", "mid")
    with f4: make_kpi_card("Özel Sigorta", money_fmt(gunluk["Özel Sigorta Geliri"].sum()), "Kanal bazlı", "pos")

    st.markdown("### Branş Bazlı Gelir")
    brans_sum = brans.groupby("Branş", as_index=False).agg(
        Gelir=("Gelir", "sum"),
        Hasta=("Hasta", "sum"),
        Kar=("Kar", "sum"),
        Memnuniyet=("Memnuniyet %", "mean"),
    )
    brans_sum["Hasta Başı Gelir"] = brans_sum["Gelir"] / brans_sum["Hasta"]

    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig = px.bar(brans_sum.sort_values("Gelir", ascending=False), x="Branş", y="Gelir", title="Branş Bazlı Gelir")
        fig.update_layout(height=410, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(
            brans_sum.sort_values("Gelir", ascending=False).assign(
                Gelir=lambda d: d["Gelir"].map(money_fmt),
                Kar=lambda d: d["Kar"].map(money_fmt),
                Memnuniyet=lambda d: d["Memnuniyet"].map(lambda x: pct_fmt(x, 0)),
                **{"Hasta Başı Gelir": lambda d: d["Hasta Başı Gelir"].map(money_fmt)}
            ),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown(
        """
        <div class="ai-box">
        <strong>AYÇA Finans Yorumu:</strong><br>
        Kardiyoloji ve Ortopedi gelir üretiminde lider. Dahiliye hasta hacmi yüksek olmasına rağmen hasta başı gelir daha düşük.
        Yönetim tarafında branş bazlı paket, tetkik ve doktor kırılımı birlikte izlenmeli.
        </div>
        """,
        unsafe_allow_html=True,
    )


elif module == "Doktor Intelligence":
    st.markdown('<div class="section-title">Doktor Intelligence</div>', unsafe_allow_html=True)

    dsum = doktor.groupby(["Doktor", "Branş"], as_index=False).agg(
        Hasta=("Hasta", "sum"),
        Gelir=("Gelir", "sum"),
        TetkikGeliri=("Tetkik Geliri", "sum"),
        Memnuniyet=("Memnuniyet %", "mean"),
        TekrarBasvuru=("Tekrar Başvuru %", "mean"),
    )
    dsum["Hasta Başı Gelir"] = dsum["Gelir"] / dsum["Hasta"]

    c1, c2 = st.columns([1.1, 1])
    with c1:
        fig = px.bar(dsum.sort_values("Gelir", ascending=False).head(10), x="Doktor", y="Gelir", color="Branş", title="En Yüksek Gelir Üreten Doktorlar")
        fig.update_layout(height=430, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        top_doc = dsum.sort_values("Gelir", ascending=False).iloc[0]
        make_kpi_card("Öne Çıkan Doktor", top_doc["Doktor"], f"{top_doc['Branş']} · {money_fmt(top_doc['Gelir'])}", "pos")
        st.write("")
        st.dataframe(
            dsum.sort_values("Gelir", ascending=False).head(8).assign(
                Gelir=lambda d: d["Gelir"].map(money_fmt),
                TetkikGeliri=lambda d: d["TetkikGeliri"].map(money_fmt),
                Memnuniyet=lambda d: d["Memnuniyet"].map(lambda x: pct_fmt(x, 0)),
                TekrarBasvuru=lambda d: d["TekrarBasvuru"].map(lambda x: pct_fmt(x, 0)),
                **{"Hasta Başı Gelir": lambda d: d["Hasta Başı Gelir"].map(money_fmt)}
            ),
            use_container_width=True,
            hide_index=True,
        )

    st.markdown(
        """
        <div class="ai-box">
        <strong>AYÇA Doktor Yorumu:</strong><br>
        Sadece gelir değil, hasta başı gelir, tetkik geliri, memnuniyet ve tekrar başvuru birlikte izlenmeli.
        Bazı doktorlarda hasta sayısı artarken hasta başı gelir düşebilir; bu durum paket, fiyatlama veya tetkik yönlendirme analizi gerektirir.
        </div>
        """,
        unsafe_allow_html=True,
    )


elif module == "Operasyon Merkezi":
    st.markdown('<div class="section-title">Operasyon Merkezi</div>', unsafe_allow_html=True)

    latest_ops = operasyon[operasyon["Tarih"] == operasyon["Tarih"].max()].copy()

    o1, o2, o3, o4 = st.columns(4)
    with o1: make_kpi_card("Ortalama Bekleme", f"{int(today['Ortalama Bekleme Dk'])} dk", "Bugün", "mid")
    with o2: make_kpi_card("Doluluk", pct_fmt(today["Doluluk %"], 0), "Yatak", "pos")
    with o3: make_kpi_card("MR Kullanım", pct_fmt(today["MR Kullanım %"], 0), "Kritik kapasite", "neg" if today["MR Kullanım %"] > 90 else "mid")
    with o4: make_kpi_card("Ameliyat", int_fmt(today["Ameliyat Sayısı"]), "Bugün", "pos")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(latest_ops.sort_values("Bekleme Dk", ascending=False), x="Birim", y="Bekleme Dk", title="Birim Bazlı Bekleme Süresi")
        fig.update_layout(height=390, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.bar(latest_ops.sort_values("Doluluk %", ascending=False), x="Birim", y="Doluluk %", title="Birim Bazlı Doluluk")
        fig.update_layout(height=390, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(latest_ops, use_container_width=True, hide_index=True)


elif module == "Hasta Merkezi":
    st.markdown('<div class="section-title">Hasta Merkezi</div>', unsafe_allow_html=True)

    h1, h2, h3, h4 = st.columns(4)
    with h1: make_kpi_card("Yeni Hasta", int_fmt(hasta["Yeni Hasta"].sum()), "Demo dönem", "pos")
    with h2: make_kpi_card("Sadık Hasta", int_fmt(hasta["Sadık Hasta"].sum()), "5+ ziyaret", "pos")
    with h3: make_kpi_card("Kaybedilen Hasta", int_fmt(hasta["Kaybedilen Hasta"].sum()), "6 ay pasif", "neg")
    with h4: make_kpi_card("Memnuniyet", pct_fmt(gunluk["Memnuniyet %"].mean(), 0), "Ortalama", "pos")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(hasta, x="Branş", y=["Yeni Hasta", "Sadık Hasta", "Kaybedilen Hasta"], barmode="group", title="Hasta Segmentleri")
        fig.update_layout(height=410, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.pie(hasta, names="Branş", values="Yeni Hasta", title="Yeni Hasta Dağılımı")
        fig.update_layout(height=410, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(hasta, use_container_width=True, hide_index=True)


elif module == "Stok & Satın Alma":
    st.markdown('<div class="section-title">Stok & Satın Alma</div>', unsafe_allow_html=True)

    critical = stok[stok["Durum"].isin(["Kritik", "Riskli"])].copy()
    s1, s2, s3, s4 = st.columns(4)
    with s1: make_kpi_card("Toplam Stok Değeri", money_fmt(stok["Stok Değeri"].sum()), "Sarf + ilaç", "pos")
    with s2: make_kpi_card("Kritik Kalem", int_fmt((stok["Durum"] == "Kritik").sum()), "Minimum stok altı", "neg")
    with s3: make_kpi_card("SKT Riski", int_fmt((stok["SKT Risk Gün"] <= 60).sum()), "60 gün altı", "mid")
    with s4: make_kpi_card("Önerilen Sipariş", money_fmt(stok["Önerilen Sipariş Değeri"].sum()), "Demo öneri", "pos")

    c1, c2 = st.columns([1.2, 1])
    with c1:
        fig = px.bar(stok.sort_values("Stok Değeri", ascending=False).head(12), x="Malzeme", y="Stok Değeri", color="Kategori", title="Stok Değeri En Yüksek Kalemler")
        fig.update_layout(height=430, margin=dict(l=10, r=10, t=50, b=10), xaxis_tickangle=-35)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.markdown("### Kritik / Riskli Kalemler")
        st.dataframe(critical, use_container_width=True, hide_index=True)

    st.markdown(
        """
        <div class="ai-box">
        <strong>AYÇA Stok Yorumu:</strong><br>
        Kritik stoklar sadece miktara göre değil, ameliyat/randevu planı ve tedarik süresiyle birlikte değerlendirilmelidir.
        SKT riski olan kalemler için kullanım hızı ve bölüm tüketimi birlikte analiz edilmelidir.
        </div>
        """,
        unsafe_allow_html=True,
    )


elif module == "Kalite Merkezi":
    st.markdown('<div class="section-title">Kalite Merkezi</div>', unsafe_allow_html=True)

    q1, q2, q3, q4 = st.columns(4)
    with q1: make_kpi_card("Toplam Şikayet", int_fmt(kalite["Şikayet Sayısı"].sum()), "Demo dönem", "mid")
    with q2: make_kpi_card("Ortalama Çözüm", f"{kalite['Ortalama Çözüm Saat'].mean():.1f} saat", "Hedef <24 saat", "pos")
    with q3: make_kpi_card("Memnuniyet", pct_fmt(gunluk["Memnuniyet %"].mean(), 0), "Genel", "pos")
    with q4: make_kpi_card("En Büyük Konu", kalite.sort_values("Şikayet Sayısı", ascending=False).iloc[0]["Konu"], "Şikayet analizi", "mid")

    c1, c2 = st.columns(2)
    with c1:
        fig = px.bar(kalite.sort_values("Şikayet Sayısı", ascending=False), x="Konu", y="Şikayet Sayısı", title="Şikayet Dağılımı")
        fig.update_layout(height=410, margin=dict(l=10, r=10, t=50, b=10), xaxis_tickangle=-25)
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        fig = px.scatter(kalite, x="Şikayet Sayısı", y="Ortalama Çözüm Saat", size="Etki Skoru", color="Konu", title="Şikayet Etki Matrisi")
        fig.update_layout(height=410, margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    st.dataframe(kalite, use_container_width=True, hide_index=True)


elif module == "AYÇA Co-Pilot":
    st.markdown('<div class="section-title">AYÇA Co-Pilot</div>', unsafe_allow_html=True)

    st.markdown(
        """
        <span class="module-chip">Yönetici asistanı</span>
        <span class="module-chip">Sebep-sonuç analizi</span>
        <span class="module-chip">Doğal dil soru-cevap</span>
        <span class="module-chip">Aksiyon önerisi</span>
        """,
        unsafe_allow_html=True,
    )

    question = st.text_input(
        "Yönetici sorusu",
        value="Bu ay ciro iyi görünürken operasyonel risk nerede oluşuyor?",
    )

    if st.button("Analiz Et", type="primary"):
        st.markdown(
            f"""
            <div class="ai-box">
            <strong>Soru:</strong> {question}<br><br>
            <strong>AYÇA Co-Pilot Yanıtı:</strong><br>
            Bu ay gelir tarafı güçlü görünmektedir. Günlük hedef gerçekleşme oranı %{today['Hedef Gerçekleşme %']:.0f} seviyesinde.
            Ancak operasyonel risk özellikle <strong>MR kapasitesi</strong> ve <strong>bekleme süresi</strong> tarafında oluşuyor.
            MR kullanım oranı %{today['MR Kullanım %']:.0f}; bu seviye sürdürülebilir kapasite sınırına yakındır.
            Kardiyoloji ve acil servis yoğunluğu artarsa hasta memnuniyetinde düşüş görülebilir.
            <br><br>
            <strong>Önerilen aksiyon:</strong>
            Radyoloji slot planlaması, yoğun branşlarda doktor çalışma saatleri ve çağrı merkezi randevu dağılımı birlikte değerlendirilmelidir.
            </div>
            """,
            unsafe_allow_html=True,
        )
    else:
        st.info("Örnek soruyu değiştirip 'Analiz Et' butonuna basabilirsiniz.")

    st.markdown("### Örnek Sorular")
    st.write("- Bu ay ciro neden yükseldi?")
    st.write("- Hangi branşta hasta başı gelir düşüyor?")
    st.write("- Hangi doktorlarda memnuniyet yüksek ama gelir düşük?")
    st.write("- Hangi stok kalemleri ameliyat planını riske atabilir?")
    st.write("- Hangi şikayet konusu hasta deneyimini en çok etkiliyor?")


st.caption("Bu uygulama konsept demodur. Veriler tamamen sample/sentetik olarak üretilmiştir.")
