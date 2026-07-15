"""
Dashboard Visualisasi Data Keuangan IDX
Streamlit App — membaca output_extractor.xlsx
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import os

# ─── KONFIGURASI PAGE ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard Keuangan IDX",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─── CSS CUSTOM ───────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
}

/* Light background */
.stApp {
    background: #f4f7fb;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #1F4E79 0%, #2E75B6 100%);
    border-right: 1px solid #2E75B6;
}
[data-testid="stSidebar"] * {
    color: #ffffff !important;
}
[data-testid="stSidebar"] .stRadio label {
    color: #ffffff !important;
}

/* Cards / Metric boxes */
.metric-card {
    background: #ffffff;
    border: 1px solid #d0e4f7;
    border-radius: 16px;
    padding: 20px 24px;
    text-align: center;
    box-shadow: 0 2px 12px rgba(46,117,182,0.08);
    transition: transform 0.2s, box-shadow 0.2s;
}
.metric-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 24px rgba(46,117,182,0.15);
}
.metric-title {
    color: #5a7a9a;
    font-size: 12px;
    font-weight: 500;
    text-transform: uppercase;
    letter-spacing: 1.2px;
    margin-bottom: 6px;
}
.metric-value {
    color: #1F4E79;
    font-size: 28px;
    font-weight: 700;
    letter-spacing: -1px;
}
.metric-sub {
    color: #8aaac4;
    font-size: 11px;
    margin-top: 4px;
}

/* Section headers */
.section-header {
    background: linear-gradient(90deg, #dbeafe 0%, #f4f7fb 100%);
    border-left: 4px solid #2E75B6;
    padding: 10px 16px;
    border-radius: 0 8px 8px 0;
    margin: 20px 0 16px 0;
    color: #1F4E79;
    font-size: 16px;
    font-weight: 600;
    letter-spacing: 0.5px;
}

/* Dataframe */
[data-testid="stDataFrame"] {
    border-radius: 10px;
    overflow: hidden;
    border: 1px solid #d0e4f7;
}

/* Title */
.main-title {
    font-size: 36px;
    font-weight: 700;
    background: linear-gradient(135deg, #1F4E79 0%, #2E75B6 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    letter-spacing: -1px;
}
.main-subtitle {
    color: #5a7a9a;
    font-size: 14px;
    margin-top: -10px;
}

/* Divider */
hr {
    border: none;
    border-top: 1px solid #d0e4f7;
    margin: 24px 0;
}

/* Sidebar label */
.sidebar-label {
    color: rgba(255,255,255,0.75);
    font-size: 11px;
    text-transform: uppercase;
    letter-spacing: 1px;
    font-weight: 600;
    margin-bottom: 4px;
}
</style>
""", unsafe_allow_html=True)

# ─── KONSTANTA ────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
DATA_FILE  = os.path.join(BASE_DIR, "output_extractor.xlsx")

PERIODS    = ["2024 TW 1","2024 TW 2","2024 TW 3","2024 Tahunan",
              "2025 TW 1","2025 TW 2","2025 TW 3","2025 Tahunan","2026 TW 1"]
VARS       = ["Pendapatan","Beban Penjualan","Laba Kotor","Laba Usaha","Penambahan Aset Tetap"]

COLOR_MAP  = {
    "Pendapatan"         : "#2E75B6",
    "Beban Penjualan"    : "#E74C3C",
    "Laba Kotor"         : "#27AE60",
    "Laba Usaha"         : "#F39C12",
    "Penambahan Aset Tetap": "#9B59B6",
}
CHART_BG   = "#ffffff"
PAPER_BG   = "#f4f7fb"
FONT_COLOR = "#1F4E79"
GRID_COLOR = "rgba(46,117,182,0.12)"

# ─── LOAD DATA ────────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="Memuat data...")
def load_data():
    wb_raw = pd.read_excel(DATA_FILE, header=None)

    # Baris 0 = periode header, baris 1 = variabel header
    row1 = wb_raw.iloc[0].tolist()
    row2 = wb_raw.iloc[1].tolist()

    # Buat nama kolom
    cols = []
    cur_period = None
    for i, (r1, r2) in enumerate(zip(row1, row2)):
        r1_str = str(r1) if r1 is not None and not (isinstance(r1, float) and pd.isna(r1)) else ""
        r2_str = str(r2) if r2 is not None and not (isinstance(r2, float) and pd.isna(r2)) else ""
        if r1_str:
            cur_period = r1_str
        if i == 0:
            cols.append("No")
        elif i == 1:
            cols.append("Kode")
        elif i == 2:
            cols.append("Nama Perusahaan")
        elif i == 3:
            cols.append("Subsektor")
        elif i == 4:
            cols.append("Satuan")
        else:
            cols.append(f"{cur_period}|{r2_str}" if cur_period and r2_str else f"col_{i}")

    df = wb_raw.iloc[2:].copy()
    df.columns = cols
    df = df.reset_index(drop=True)

    # Hapus baris tanpa ticker
    df = df[df["Kode"].notna() & (df["Kode"].astype(str).str.strip() != "")]
    df["Kode"]           = df["Kode"].astype(str).str.strip()
    df["Subsektor"]      = df["Subsektor"].astype(str).str.strip().fillna("")
    df["Satuan"]         = df["Satuan"].astype(str).str.strip().fillna("Rupiah")
    df["Nama Perusahaan"] = df["Nama Perusahaan"].astype(str).str.strip()

    # Konversi kolom numerik
    for c in df.columns:
        if "|" in str(c):
            df[c] = pd.to_numeric(df[c], errors="coerce")

    return df

def get_val(df: pd.DataFrame, period: str, var: str):
    col = f"{period}|{var}"
    return df[col] if col in df.columns else pd.Series([None]*len(df))

def fmt_num(val, satuan="Rupiah"):
    """Format angka jadi Triliun/Miliar/Juta dengan satuan"""
    if pd.isna(val):
        return "-"
    if abs(val) >= 1_000_000:
        return f"{val/1_000_000:,.2f} Juta"
    elif abs(val) >= 1_000:
        return f"{val/1_000:,.2f} Ribu"
    else:
        return f"{val:,.0f}"

def growth_pct(new_val, old_val):
    if pd.isna(new_val) or pd.isna(old_val) or old_val == 0:
        return None
    return (new_val - old_val) / abs(old_val) * 100

# ─── LOAD ─────────────────────────────────────────────────────────────────────
df = load_data()

# Subsektor list (bersih) — filter None, NaN float, dan string kosong/invalid
BAD_VALUES = {"nan", "#n/a", "#na", "", "none"}
all_subsektors = sorted([
    s for s in df["Subsektor"].unique()
    if s is not None
    and not (isinstance(s, float) and pd.isna(s))
    and str(s).strip().lower() not in BAD_VALUES
])

# ═══════════════════════════════════════════════════════════════════════════════
# SIDEBAR
# ═══════════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown('<div class="main-title">📊 IDX</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Dashboard Data Keuangan BEI</div>', unsafe_allow_html=True)
    st.markdown("---")

    st.markdown('<div class="sidebar-label">🗂️ Mode Tampilan</div>', unsafe_allow_html=True)
    page = st.radio(
        "Mode Tampilan",
        ["🏠 Overview", "🏢 Analisis Perusahaan", "📈 Tren Periode", "🏭 Analisis Sektor"],
        label_visibility="collapsed",
    )

    st.markdown("---")
    st.markdown('<div class="sidebar-label">⚙️ Filter Global</div>', unsafe_allow_html=True)

    sel_subsektors = st.multiselect(
        "Subsektor",
        all_subsektors,
        default=[],
        placeholder="Semua subsektor...",
    )

    sel_period = st.selectbox("Periode Utama", PERIODS, index=8)  # default 2026 TW1
    sel_var    = st.selectbox("Variabel", VARS, index=0)

    st.markdown("---")
    st.markdown(f'<div style="color:rgba(255,255,255,0.6);font-size:10px;text-align:center">Total: {len(df)} perusahaan · {len(PERIODS)} periode<br>Sumber: IDX XBRL</div>', unsafe_allow_html=True)

# Filter df berdasarkan subsektor
df_filtered = df.copy()
if sel_subsektors:
    df_filtered = df_filtered[df_filtered["Subsektor"].isin(sel_subsektors)]

# ─── TEMPLATE PLOTLY ──────────────────────────────────────────────────────────
def apply_dark_theme(fig):
    fig.update_layout(
        plot_bgcolor=CHART_BG,
        paper_bgcolor=PAPER_BG,
        font=dict(color=FONT_COLOR, family="Inter"),
        xaxis=dict(gridcolor=GRID_COLOR, showgrid=True, zeroline=False,
                   tickfont=dict(size=10, color=FONT_COLOR),
                   title_font=dict(size=11, color=FONT_COLOR)),
        yaxis=dict(gridcolor=GRID_COLOR, showgrid=True, zeroline=False,
                   tickfont=dict(size=10, color=FONT_COLOR),
                   title_font=dict(size=11, color=FONT_COLOR)),
        legend=dict(
            bgcolor="#ffffff",
            bordercolor="#d0e4f7",
            borderwidth=1,
            font=dict(size=10, color=FONT_COLOR),
        ),
        margin=dict(t=40, b=40, l=50, r=20),
    )
    return fig

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: OVERVIEW
# ═══════════════════════════════════════════════════════════════════════════════
if page == "🏠 Overview":
    st.markdown('<div class="main-title">📊 Dashboard Keuangan IDX</div>', unsafe_allow_html=True)
    st.markdown('<div class="main-subtitle">Visualisasi Data Laporan Keuangan Bursa Efek Indonesia (BEI) · Hasil Ekstraksi XBRL</div>', unsafe_allow_html=True)
    st.markdown("---")

    # ─── KPI Cards ───
    col_p = f"{sel_period}|{sel_var}"
    prev_idx = PERIODS.index(sel_period) - 1
    prev_period = PERIODS[prev_idx] if prev_idx >= 0 else None

    total_companies = len(df_filtered)
    data_count      = df_filtered[col_p].notna().sum() if col_p in df_filtered.columns else 0
    total_val       = df_filtered[col_p].sum() if col_p in df_filtered.columns else 0
    avg_val         = df_filtered[col_p].mean() if col_p in df_filtered.columns else 0

    if prev_period:
        prev_col = f"{prev_period}|{sel_var}"
        prev_total = df_filtered[prev_col].sum() if prev_col in df_filtered.columns else 0
        growth = growth_pct(total_val, prev_total)
        growth_str = (f"{'▲' if growth > 0 else '▼'} {abs(growth):.1f}% vs {prev_period}"
                      if growth is not None else "—")
        growth_color = "#16a34a" if (growth or 0) > 0 else "#dc2626"
    else:
        growth_str = "—"
        growth_color = FONT_COLOR

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-title">Total Perusahaan</div>
            <div class="metric-value">{total_companies:,}</div>
            <div class="metric-sub">terdaftar di IDX</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-title">Data Tersedia</div>
            <div class="metric-value">{data_count:,}</div>
            <div class="metric-sub">{sel_period} · {sel_var}</div>
        </div>""", unsafe_allow_html=True)
    with c3:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-title">Total {sel_var}</div>
            <div class="metric-value">{total_val/1e9:,.1f}T</div>
            <div class="metric-sub">(dalam satuan asli × 10⁹)</div>
        </div>""", unsafe_allow_html=True)
    with c4:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-title">Rata-rata {sel_var}</div>
            <div class="metric-value">{avg_val/1e6:,.1f}M</div>
            <div class="metric-sub">per perusahaan</div>
        </div>""", unsafe_allow_html=True)
    with c5:
        st.markdown(f"""<div class="metric-card">
            <div class="metric-title">Pertumbuhan</div>
            <div class="metric-value" style="font-size:20px;color:{growth_color}">{growth_str}</div>
            <div class="metric-sub">QoQ</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── Row 1: Top 10 Bar + Pie Subsektor ───
    col_l, col_r = st.columns([2, 1])

    with col_l:
        st.markdown(f'<div class="section-header">🏆 Top 15 Perusahaan — {sel_var} ({sel_period})</div>', unsafe_allow_html=True)
        if col_p in df_filtered.columns:
            top15 = (df_filtered[["Kode","Nama Perusahaan","Subsektor", col_p]]
                     .dropna(subset=[col_p])
                     .nlargest(15, col_p))
            fig_bar = px.bar(
                top15, x="Kode", y=col_p,
                color=col_p,
                color_continuous_scale=[[0, "#1F4E79"],[0.5, "#2E75B6"],[1, "#7ec8e3"]],
                hover_data={"Nama Perusahaan": True, "Subsektor": True, col_p: ":,.0f"},
                text=top15[col_p].apply(lambda v: f"{v/1e6:,.1f}M"),
            )
            fig_bar.update_traces(textposition="outside", textfont_size=9)
            fig_bar.update_layout(showlegend=False, coloraxis_showscale=False,
                                  xaxis_title="Kode Saham", yaxis_title=sel_var)
            apply_dark_theme(fig_bar)
            st.plotly_chart(fig_bar, use_container_width=True)

    with col_r:
        st.markdown(f'<div class="section-header">🏭 Distribusi per Subsektor</div>', unsafe_allow_html=True)
        if col_p in df_filtered.columns:
            sec_sum = (df_filtered.groupby("Subsektor")[col_p]
                       .sum().dropna()
                       .sort_values(ascending=False)
                       .head(12))
            fig_pie = px.pie(
                values=sec_sum.values,
                names=sec_sum.index,
                color_discrete_sequence=px.colors.sequential.Blues_r,
                hole=0.45,
            )
            fig_pie.update_traces(textinfo="percent+label", textfont_size=9)
            fig_pie.update_layout(showlegend=False, margin=dict(t=10,b=10,l=10,r=10))
            apply_dark_theme(fig_pie)
            st.plotly_chart(fig_pie, use_container_width=True)

    # ─── Row 2: Heatmap data availability ───
    st.markdown('<div class="section-header">🗓️ Ketersediaan Data per Periode</div>', unsafe_allow_html=True)
    avail_data = {}
    for p in PERIODS:
        avail_data[p] = {v: df_filtered[f"{p}|{v}"].notna().sum()
                         for v in VARS if f"{p}|{v}" in df_filtered.columns}

    avail_df = pd.DataFrame(avail_data).T
    fig_heat = px.imshow(
        avail_df.T,
        color_continuous_scale=[[0,"#0a1628"],[0.3,"#1F4E79"],[0.7,"#2E75B6"],[1,"#7ec8e3"]],
        aspect="auto",
        text_auto=True,
    )
    fig_heat.update_coloraxes(showscale=False)
    fig_heat.update_layout(
        xaxis_title="Periode",
        yaxis_title="Variabel",
        margin=dict(t=20,b=20,l=150,r=20),
    )
    apply_dark_theme(fig_heat)
    st.plotly_chart(fig_heat, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: ANALISIS PERUSAHAAN
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "🏢 Analisis Perusahaan":
    st.markdown('<div class="main-title">🏢 Analisis Perusahaan</div>', unsafe_allow_html=True)
    st.markdown("---")

    # Pilih perusahaan
    company_list = sorted(df["Kode"].tolist())
    col_s1, col_s2 = st.columns([1, 3])
    with col_s1:
        sel_company = st.selectbox("Pilih Perusahaan", company_list, index=0)

    comp_row = df[df["Kode"] == sel_company].iloc[0]
    nama     = comp_row["Nama Perusahaan"]
    subsektor = comp_row["Subsektor"]
    satuan   = comp_row["Satuan"]

    with col_s2:
        st.markdown(f"""
        <div style="background:#ffffff;border:1px solid #d0e4f7;
                    border-radius:12px;padding:14px 20px;margin-top:4px;
                    box-shadow: 0 2px 12px rgba(46,117,182,0.08);">
            <span style="color:#1F4E79;font-weight:700;font-size:20px">{sel_company}</span>
            <span style="color:#5a7a9a;font-size:14px;margin-left:12px">{nama}</span>
            <br>
            <span style="color:#8aaac4;font-size:12px">📌 {subsektor} &nbsp;|&nbsp; 💰 Satuan: {satuan}</span>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ─── Kumpulkan data perusahaan per periode ───
    comp_data = []
    for p in PERIODS:
        row_dict = {"Periode": p}
        for v in VARS:
            col = f"{p}|{v}"
            row_dict[v] = comp_row[col] if col in comp_row.index else None
        comp_data.append(row_dict)
    comp_df = pd.DataFrame(comp_data)

    # ─── KPI Cards untuk periode terbaru ───
    latest = comp_df.dropna(subset=["Pendapatan"]).iloc[-1] if not comp_df.dropna(subset=["Pendapatan"]).empty else None
    if latest is not None:
        st.markdown(f'<div class="section-header">📊 Data Terbaru — {latest["Periode"]}</div>', unsafe_allow_html=True)
        c1,c2,c3,c4,c5 = st.columns(5)
        cards = [("Pendapatan","c1","#2E75B6"),
                 ("Beban Penjualan","c2","#E74C3C"),
                 ("Laba Kotor","c3","#27AE60"),
                 ("Laba Usaha","c4","#F39C12"),
                 ("Penambahan Aset Tetap","c5","#9B59B6")]
        for var, col_var, color in cards:
            val = latest.get(var)
            display = f"{val:,.0f}" if pd.notna(val) else "—"
            eval(col_var).markdown(f"""<div class="metric-card">
                <div class="metric-title" style="color:{color}">{var}</div>
                <div class="metric-value" style="font-size:18px">{display}</div>
                <div class="metric-sub">{satuan}</div>
            </div>""", unsafe_allow_html=True)

    # ─── Line chart semua variabel per periode ───
    st.markdown('<div class="section-header">📈 Tren Semua Variabel</div>', unsafe_allow_html=True)

    fig_line = go.Figure()
    for v in VARS:
        valid = comp_df[comp_df[v].notna()]
        if not valid.empty:
            fig_line.add_trace(go.Scatter(
                x=valid["Periode"], y=valid[v],
                name=v, mode="lines+markers",
                line=dict(color=COLOR_MAP[v], width=2.5),
                marker=dict(size=7, symbol="circle"),
                hovertemplate=f"<b>{v}</b><br>%{{x}}<br>%{{y:,.0f}} {satuan}<extra></extra>",
            ))
    fig_line.update_layout(
        title=dict(text=f"Tren Keuangan {sel_company}", font=dict(size=14, color=FONT_COLOR)),
        xaxis_title="Periode",
        yaxis_title=f"Nilai ({satuan})",
        hovermode="x unified",
        height=420,
    )
    apply_dark_theme(fig_line)
    st.plotly_chart(fig_line, use_container_width=True)

    # ─── Waterfall: Breakdown laba terakhir ───
    col_w, col_t = st.columns([1.5, 1])
    with col_w:
        st.markdown('<div class="section-header">🌊 Waterfall Breakdown Laba (Periode Terbaru)</div>', unsafe_allow_html=True)
        if latest is not None:
            pend = latest.get("Pendapatan") or 0
            beban = latest.get("Beban Penjualan") or 0
            lk    = latest.get("Laba Kotor") or 0
            lu    = latest.get("Laba Usaha") or 0
            diff_ops = lu - lk if pd.notna(lu) and pd.notna(lk) else 0

            fig_wf = go.Figure(go.Waterfall(
                orientation="v",
                measure=["absolute","relative","total","relative","total"],
                x=["Pendapatan","Beban Penjualan","Laba Kotor","Beban Operasi","Laba Usaha"],
                y=[pend, -beban, lk, diff_ops, lu],
                connector={"line":{"color":"rgba(46,117,182,0.4)"}},
                decreasing={"marker":{"color":"#E74C3C"}},
                increasing={"marker":{"color":"#27AE60"}},
                totals={"marker":{"color":"#2E75B6"}},
                text=[f"{v:,.0f}" for v in [pend,-beban,lk,diff_ops,lu]],
                textposition="outside",
                textfont=dict(size=9, color=FONT_COLOR),
            ))
            fig_wf.update_layout(
                title=dict(text=latest["Periode"], font=dict(size=12, color=FONT_COLOR)),
                showlegend=False,
                height=350,
                xaxis_title="",
                yaxis_title=satuan,
            )
            apply_dark_theme(fig_wf)
            st.plotly_chart(fig_wf, use_container_width=True)

    with col_t:
        st.markdown('<div class="section-header">📋 Tabel Data Lengkap</div>', unsafe_allow_html=True)
        display_df = comp_df.set_index("Periode")
        for v in VARS:
            display_df[v] = display_df[v].apply(
                lambda x: f"{x:,.0f}" if pd.notna(x) else "—"
            )
        st.dataframe(display_df, use_container_width=True, height=350)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: TREN PERIODE
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "📈 Tren Periode":
    st.markdown('<div class="main-title">📈 Tren Antar Periode</div>', unsafe_allow_html=True)
    st.markdown("---")

    # ─── Pilih multiple perusahaan ───
    sel_companies_tren = st.multiselect(
        "Bandingkan Perusahaan (maks 8)",
        sorted(df_filtered["Kode"].tolist()),
        default=sorted(df_filtered["Kode"].tolist())[:5],
        max_selections=8,
    )

    if not sel_companies_tren:
        st.info("Pilih minimal 1 perusahaan untuk melihat tren.")
    else:
        # ─── Line chart per variabel ───
        for var in VARS:
            st.markdown(f'<div class="section-header">📊 {var}</div>', unsafe_allow_html=True)
            fig = go.Figure()
            for ticker in sel_companies_tren:
                row = df_filtered[df_filtered["Kode"] == ticker]
                if row.empty:
                    continue
                row = row.iloc[0]
                y_vals = [row.get(f"{p}|{var}", None) for p in PERIODS]
                # Hanya plot jika ada data
                if any(v is not None and not (isinstance(v, float) and pd.isna(v)) for v in y_vals):
                    fig.add_trace(go.Scatter(
                        x=PERIODS, y=y_vals,
                        name=ticker,
                        mode="lines+markers",
                        line=dict(width=2),
                        marker=dict(size=6),
                        connectgaps=True,
                        hovertemplate=f"<b>{ticker}</b><br>%{{x}}<br>%{{y:,.0f}}<extra></extra>",
                    ))
            fig.update_layout(
                hovermode="x unified",
                height=320,
                xaxis_title="",
                yaxis_title=var,
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
            )
            apply_dark_theme(fig)
            st.plotly_chart(fig, use_container_width=True)

    # ─── Growth Rate Heatmap ───
    st.markdown('<div class="section-header">🌡️ Heatmap Growth Rate QoQ (Pendapatan)</div>', unsafe_allow_html=True)
    growth_rows = []
    for _, row in df_filtered.iterrows():
        ticker = row["Kode"]
        g_row = {"Kode": ticker}
        for i in range(1, len(PERIODS)):
            cur  = f"{PERIODS[i]}|Pendapatan"
            prev = f"{PERIODS[i-1]}|Pendapatan"
            vc, vp = row.get(cur), row.get(prev)
            if pd.notna(vc) and pd.notna(vp) and vp != 0:
                g_row[f"{PERIODS[i-1]}→{PERIODS[i]}"] = round((vc - vp) / abs(vp) * 100, 1)
            else:
                g_row[f"{PERIODS[i-1]}→{PERIODS[i]}"] = None
        growth_rows.append(g_row)

    growth_df = pd.DataFrame(growth_rows).set_index("Kode")
    growth_df = growth_df.dropna(how="all")
    # Ambil top 40 company by data availability
    growth_df = growth_df.loc[growth_df.notna().sum(axis=1).nlargest(40).index]

    if not growth_df.empty:
        fig_gheat = px.imshow(
            growth_df.T,
            color_continuous_scale="RdYlGn",
            color_continuous_midpoint=0,
            aspect="auto",
            zmin=-50, zmax=50,
            text_auto=".1f",
        )
        fig_gheat.update_traces(textfont=dict(size=7))
        fig_gheat.update_layout(
            xaxis_title="Perusahaan",
            yaxis_title="Transisi Periode",
            height=360,
            margin=dict(t=20,b=40,l=160,r=20),
        )
        apply_dark_theme(fig_gheat)
        fig_gheat.update_layout(coloraxis_colorbar=dict(
            title=dict(text="Growth %", font=dict(color=FONT_COLOR)),
            tickfont=dict(color=FONT_COLOR),
        ))
        st.plotly_chart(fig_gheat, use_container_width=True)

# ═══════════════════════════════════════════════════════════════════════════════
# PAGE: ANALISIS SEKTOR
# ═══════════════════════════════════════════════════════════════════════════════
elif page == "🏭 Analisis Sektor":
    st.markdown('<div class="main-title">🏭 Analisis per Subsektor</div>', unsafe_allow_html=True)
    st.markdown("---")

    col1, col2 = st.columns([1, 1])

    # ─── Pilih variabel & periode ───
    with col1:
        var_sek = st.selectbox("Variabel", VARS, key="var_sek")
    with col2:
        per_sek = st.selectbox("Periode", PERIODS, index=8, key="per_sek")

    col_key = f"{per_sek}|{var_sek}"

    # ─── Bar: Total per subsektor ───
    st.markdown(f'<div class="section-header">📊 Total {var_sek} per Subsektor ({per_sek})</div>', unsafe_allow_html=True)
    if col_key in df_filtered.columns:
        sec_data = (df_filtered.groupby("Subsektor")[col_key]
                    .agg(["sum","mean","count"])
                    .dropna()
                    .sort_values("sum", ascending=True)
                    .tail(20))
        sec_data.columns = ["Total","Rata-rata","Jumlah Perusahaan"]
        sec_data = sec_data.reset_index()

        fig_sec = px.bar(
            sec_data, x="Total", y="Subsektor",
            orientation="h",
            color="Total",
            color_continuous_scale=[[0,"#1F4E79"],[0.5,"#2E75B6"],[1,"#7ec8e3"]],
            hover_data={"Rata-rata":":,.0f","Jumlah Perusahaan":True},
            text=sec_data["Total"].apply(lambda v: f"{v/1e6:,.1f}M"),
        )
        fig_sec.update_traces(textposition="outside", textfont_size=9)
        fig_sec.update_layout(
            coloraxis_showscale=False,
            showlegend=False,
            height=500,
            xaxis_title=var_sek,
            yaxis_title="",
        )
        apply_dark_theme(fig_sec)
        st.plotly_chart(fig_sec, use_container_width=True)

    # ─── Scatter: Pendapatan vs Laba Usaha colored by sektor ───
    st.markdown(f'<div class="section-header">🔵 Scatter — Pendapatan vs Laba Usaha ({per_sek})</div>', unsafe_allow_html=True)
    col_pend = f"{per_sek}|Pendapatan"
    col_laba = f"{per_sek}|Laba Usaha"
    if col_pend in df_filtered.columns and col_laba in df_filtered.columns:
        scatter_df = df_filtered[["Kode","Nama Perusahaan","Subsektor",col_pend,col_laba,"Satuan"]].dropna()
        fig_sc = px.scatter(
            scatter_df,
            x=col_pend,
            y=col_laba,
            color="Subsektor",
            hover_data={"Kode":True,"Nama Perusahaan":True,
                        col_pend:":,.0f", col_laba:":,.0f"},
            size_max=18,
            opacity=0.8,
            labels={col_pend:"Pendapatan", col_laba:"Laba Usaha"},
        )
        fig_sc.update_traces(marker=dict(size=10, line=dict(width=0)))
        fig_sc.update_layout(height=460, hovermode="closest")
        apply_dark_theme(fig_sc)
        st.plotly_chart(fig_sc, use_container_width=True)

    # ─── Box plot distribusi per sektor ───
    st.markdown(f'<div class="section-header">📦 Distribusi {var_sek} per Subsektor (Box Plot)</div>', unsafe_allow_html=True)
    if col_key in df_filtered.columns:
        box_df = df_filtered[["Subsektor", col_key]].dropna()
        top_seks = box_df.groupby("Subsektor")[col_key].median().nlargest(15).index
        box_df = box_df[box_df["Subsektor"].isin(top_seks)]

        fig_box = px.box(
            box_df, x="Subsektor", y=col_key,
            color="Subsektor",
            color_discrete_sequence=px.colors.qualitative.Bold,
        )
        fig_box.update_layout(
            showlegend=False,
            height=420,
            xaxis_tickangle=-30,
            yaxis_title=var_sek,
        )
        apply_dark_theme(fig_box)
        st.plotly_chart(fig_box, use_container_width=True)

    # ─── Tabel Ringkasan ───
    st.markdown(f'<div class="section-header">📋 Ringkasan per Subsektor ({per_sek})</div>', unsafe_allow_html=True)
    if col_key in df_filtered.columns:
        sum_tbl = df_filtered.groupby("Subsektor").agg(
            Perusahaan=("Kode","count"),
            **{f"Total {var_sek}": (col_key, "sum")},
            **{f"Rata-rata {var_sek}": (col_key, "mean")},
            **{f"Maks {var_sek}": (col_key, "max")},
        ).dropna().sort_values(f"Total {var_sek}", ascending=False)
        for c in [f"Total {var_sek}", f"Rata-rata {var_sek}", f"Maks {var_sek}"]:
            sum_tbl[c] = sum_tbl[c].apply(lambda v: f"{v:,.0f}" if pd.notna(v) else "—")
        st.dataframe(sum_tbl, use_container_width=True, height=400)

# ─── Footer ───────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    '<div style="text-align:center;color:#8aaac4;font-size:11px">'
    'Dashboard Keuangan IDX · Data Sumber: Bursa Efek Indonesia (XBRL)'
    '</div>',
    unsafe_allow_html=True
)
