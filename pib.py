# ============================================================
# TransportesNova — Dashboard de Costos por Metodología
# ============================================================
# Instalación:
#   pip install streamlit pandas plotly openpyxl
# Ejecución:
#   streamlit run app.py
# ============================================================

import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import numpy as np
import os

# ─────────────────────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TransportesNova Dashboard",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded",
)

COLORS = {
    "SCRUM":          "#2E75B6",
    "SCRUMBAN":       "#70AD47",
    "KANBAN":         "#ED7D31",
    "XP":             "#9E480E",
    "Water Scrum Fall": "#954F72",
    "WSF":            "#954F72",
    "PMI":            "#4472C4",
    "bg":             "#F7F9FC",
    "text":           "#262626",
}

METHOD_COLORS = [
    COLORS["SCRUM"],
    COLORS["KANBAN"],
    COLORS["SCRUMBAN"],
    COLORS["XP"],
    COLORS["WSF"],
]

EXCEL_PATH = os.path.join(os.path.dirname(__file__), "TransportesNova_CostosSCRUMBAN.xlsx")

# ─────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
    body, .stApp { background-color: #F7F9FC; color: #262626; }
    .main-title {
        font-size: 2rem; font-weight: 700; color: #262626;
        border-left: 6px solid #2E75B6; padding-left: 14px;
        margin-bottom: 6px;
    }
    .sub-title {
        font-size: 1rem; color: #555; margin-bottom: 20px;
        padding-left: 20px;
    }
    .kpi-card {
        background: white;
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.07);
        border-top: 4px solid;
        text-align: center;
        margin-bottom: 12px;
    }
    .kpi-label { font-size: 0.78rem; color: #777; text-transform: uppercase;
                 letter-spacing: 0.04em; margin-bottom: 6px; }
    .kpi-value { font-size: 1.55rem; font-weight: 700; color: #262626; }
    .kpi-sub   { font-size: 0.78rem; color: #999; margin-top: 4px; }
    .section-header {
        font-size: 1.15rem; font-weight: 600; color: #262626;
        border-bottom: 2px solid #e0e0e0; padding-bottom: 6px;
        margin: 20px 0 12px 0;
    }
    .note-box {
        background: #EFF6FF; border-left: 4px solid #2E75B6;
        border-radius: 6px; padding: 12px 16px; margin: 12px 0;
        font-size: 0.88rem; color: #333;
    }
    [data-testid="stSidebar"] { background: #1A2332; }
    [data-testid="stSidebar"] * { color: #E8EDF4 !important; }
    [data-testid="stSidebar"] .stRadio label { padding: 6px 0; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────
def safe_number(val):
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return 0.0
        return float(val)
    except Exception:
        return 0.0


def format_currency(val):
    try:
        v = safe_number(val)
        return f"Q {v:,.2f}"
    except Exception:
        return "Q —"


def format_percent(val):
    try:
        v = safe_number(val)
        if v <= 1.5:
            v = v * 100
        return f"{v:.2f}%"
    except Exception:
        return "—"


def kpi_card(title, value, color="#2E75B6", sub=""):
    return f"""
    <div class="kpi-card" style="border-top-color:{color};">
        <div class="kpi-label">{title}</div>
        <div class="kpi-value">{value}</div>
        {"<div class='kpi-sub'>" + sub + "</div>" if sub else ""}
    </div>"""


def section(title):
    st.markdown(f'<div class="section-header">{title}</div>', unsafe_allow_html=True)


def note(text):
    st.markdown(f'<div class="note-box">📌 {text}</div>', unsafe_allow_html=True)


def make_bar_chart(categories, values, title, color=None, colors=None, yprefix="Q "):
    fig = go.Figure()
    bar_colors = colors if colors else ([color] * len(categories) if color else METHOD_COLORS)
    fig.add_trace(go.Bar(
        x=categories, y=values,
        marker_color=bar_colors[:len(categories)],
        text=[f"{yprefix}{v:,.0f}" if yprefix else f"{v:.1f}" for v in values],
        textposition="outside",
    ))
    fig.update_layout(
        title=title, template="plotly_white",
        plot_bgcolor="#F7F9FC", paper_bgcolor="#F7F9FC",
        font=dict(color="#262626"), margin=dict(t=50, b=30, l=40, r=20),
        height=370,
    )
    return fig


def make_grouped_bar(categories, series_dict, title):
    fig = go.Figure()
    palette = list(COLORS.values())
    for i, (name, vals) in enumerate(series_dict.items()):
        fig.add_trace(go.Bar(name=name, x=categories, y=vals,
                             marker_color=palette[i % len(palette)]))
    fig.update_layout(
        barmode="group", title=title, template="plotly_white",
        plot_bgcolor="#F7F9FC", paper_bgcolor="#F7F9FC",
        font=dict(color="#262626"), margin=dict(t=50, b=30),
        height=380,
    )
    return fig


def make_pie_chart(labels, values, title, colors=None):
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        marker_colors=colors or METHOD_COLORS,
        textinfo="label+percent", hole=0.38,
    ))
    fig.update_layout(
        title=title, template="plotly_white",
        paper_bgcolor="#F7F9FC", font=dict(color="#262626"),
        margin=dict(t=50, b=20), height=360,
    )
    return fig


# ─────────────────────────────────────────────────────────────
# LOAD EXCEL
# ─────────────────────────────────────────────────────────────
@st.cache_data
def load_excel(path):
    sheets = {}
    try:
        xf = pd.ExcelFile(path)
        for s in xf.sheet_names:
            try:
                sheets[s] = pd.read_excel(path, sheet_name=s, header=None)
            except Exception:
                sheets[s] = pd.DataFrame()
    except Exception as e:
        st.error(f"No se pudo abrir el archivo Excel: {e}")
    return sheets


def clean_sheet(df):
    df = df.dropna(how="all").reset_index(drop=True)
    df = df.dropna(axis=1, how="all")
    return df


def clean_backlog_like_sheet(df):
    """Filter rows that are real backlog entries (have SCRUM-xx key)."""
    df.columns = range(len(df.columns))
    real = df[df[0].astype(str).str.match(r'^SCRUM-\d+')].copy()
    real.columns = list(range(len(real.columns)))
    # Assign headers
    col_map = {0: "Clave", 1: "Título Historia", 2: "Épica",
               3: "Descripción Épica", 4: "Tipo", 5: "SP",
               6: "Prioridad", 7: "Estado"}
    real = real.rename(columns={k: v for k, v in col_map.items() if k < len(real.columns)})
    real["SP"] = real["SP"].apply(safe_number)
    return real.reset_index(drop=True)

def hex_to_rgba(hex_color, alpha=0.5):
    hex_color = hex_color.lstrip("#")
    r = int(hex_color[0:2], 16)
    g = int(hex_color[2:4], 16)
    b = int(hex_color[4:6], 16)
    return f"rgba({r},{g},{b},{alpha})"

def extract_value(df, label_substr):
    """Search all cells for a label and return the value in adjacent cell."""
    label_substr = str(label_substr).lower().strip()
    for i in range(len(df)):
        for j in range(len(df.columns)):
            cell = str(df.iloc[i, j]).lower().strip()
            if label_substr in cell:
                # Try right cell
                if j + 1 < len(df.columns):
                    v = df.iloc[i, j + 1]
                    if pd.notna(v):
                        return v
                # Try far-right cell (col 3 or last)
                for k in range(len(df.columns) - 1, j, -1):
                    v = df.iloc[i, k]
                    if pd.notna(v):
                        return v
    return None


SHEETS = load_excel(EXCEL_PATH)

# ─────────────────────────────────────────────────────────────
# RESUMEN DATA
# ─────────────────────────────────────────────────────────────
def get_resumen_data():
    df = SHEETS.get("Resumen", pd.DataFrame())
    methods = ["SCRUM", "KANBAN", "SCRUMBAN", "XP", "WSF"]
    metrics_labels = {
        "Duración del Proyecto":              "duracion",
        "Costo Total del Proyecto (Interno)": "costo_total",
        "Ingreso Total del Proyecto (Venta)": "ingreso_total",
        "Margen Bruto":                       "margen_bruto",
        "Margen Bruto % (sobre venta)":       "margen_pct",
        "Costo por Unidad de Trabajo":        "costo_unidad",
        "Equipo Mensual (costo interno)":     "equipo_mensual",
    }
    data = {k: {} for k in metrics_labels.values()}
    if df.empty:
        return data, methods

    # Row 3 = headers (SCRUM, KANBAN, SCRUMBAN, XP, WSF)
    # Row 4+ = data rows
    for row_i in range(len(df)):
        row_label = str(df.iloc[row_i, 0]).strip()
        for ml, mk in metrics_labels.items():
            if ml.lower() in row_label.lower():
                for col_i, m in enumerate(methods):
                    col_offset = col_i + 1
                    if col_offset < len(df.columns):
                        data[mk][m] = df.iloc[row_i, col_offset]
    return data, methods


# ─────────────────────────────────────────────────────────────
# SIDEBAR NAVIGATION
# ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:20px 0 10px 0;">
        <div style="font-size:2rem;">🚛</div>
        <div style="font-weight:700; font-size:1.1rem; color:#E8EDF4;">TransportesNova</div>
        <div style="font-size:0.78rem; color:#8A9BB0;">Dashboard de Costos</div>
    </div>
    <hr style="border-color:#2E3E54; margin:8px 0 16px 0;">
    """, unsafe_allow_html=True)

    page = st.radio(
        "Navegación",
        [
            "🏠 Inicio / Resumen",
            "📊 Comparativo",
            "🔵 SCRUM",
            "🟢 SCRUMBAN",
            "🟠 KANBAN",
            "🟤 XP",
            "🟣 Water Scrum Fall",
            "💰 PMI / Finanzas y Riesgos",
            "👥 Equipo",
            "📋 Backlog",
            "🎯 MVP",
        ],
        label_visibility="collapsed",
    )
    st.markdown("<hr style='border-color:#2E3E54; margin:16px 0;'>", unsafe_allow_html=True)
    st.markdown(
        "<div style='font-size:0.72rem; color:#5A6A80; text-align:center;'>"
        "Datos: TransportesNova_CostosSCRUMBAN.xlsx<br>92 Historias · 623 SP</div>",
        unsafe_allow_html=True
    )

# ─────────────────────────────────────────────────────────────
# PAGE: INICIO
# ─────────────────────────────────────────────────────────────
if page == "🏠 Inicio / Resumen":
    st.markdown('<div class="main-title">TransportesNova — Análisis Comparativo de Costos</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">92 Historias de Usuario · 623 Story Points · 5 Metodologías Comparadas</div>', unsafe_allow_html=True)

    data, methods = get_resumen_data()
    display_names = {"WSF": "Water Scrum Fall"}
    m_colors = [COLORS["SCRUM"], COLORS["KANBAN"], COLORS["SCRUMBAN"], COLORS["XP"], COLORS["WSF"]]

    # ── KPIs principales ──
    section("Métricas por Metodología")
    for m, col_color in zip(methods, m_colors):
        label = display_names.get(m, m)
        costo = safe_number(data["costo_total"].get(m, 0))
        ingreso = safe_number(data["ingreso_total"].get(m, 0))
        margen = safe_number(data["margen_bruto"].get(m, 0))
        mpct_raw = safe_number(data["margen_pct"].get(m, 0))
        mpct = mpct_raw * 100 if mpct_raw <= 1 else mpct_raw
        dur = str(data["duracion"].get(m, "—")).split(" ")[0][:6]

    cols = st.columns(5)
    for ci, (m, col_color) in enumerate(zip(methods, m_colors)):
        label = display_names.get(m, m)
        costo = safe_number(data["costo_total"].get(m, 0))
        ingreso = safe_number(data["ingreso_total"].get(m, 0))
        margen = safe_number(data["margen_bruto"].get(m, 0))
        mpct_raw = safe_number(data["margen_pct"].get(m, 0))
        mpct = mpct_raw * 100 if mpct_raw <= 1 else mpct_raw
        dur_raw = str(data["duracion"].get(m, "—"))
        dur_num = dur_raw.split(" ")[0][:7]
        with cols[ci]:
            st.markdown(
                kpi_card(label, format_currency(costo), col_color, f"Ingreso: {format_currency(ingreso)}"),
                unsafe_allow_html=True
            )
            st.markdown(
                kpi_card("Margen Bruto", format_currency(margen), col_color, f"{mpct:.1f}%"),
                unsafe_allow_html=True
            )
            st.markdown(
                kpi_card("Duración", f"{dur_num} m", col_color),
                unsafe_allow_html=True
            )

    # ── Gráficos comparativos ──
    section("Comparativa Visual")
    labels_display = ["SCRUM", "KANBAN", "SCRUMBAN", "XP", "WSF"]
    costos  = [safe_number(data["costo_total"].get(m, 0)) for m in methods]
    ingresos = [safe_number(data["ingreso_total"].get(m, 0)) for m in methods]
    margenes = [safe_number(data["margen_bruto"].get(m, 0)) for m in methods]
    mpcts   = [safe_number(data["margen_pct"].get(m, 0)) * 100
               if safe_number(data["margen_pct"].get(m, 0)) <= 1
               else safe_number(data["margen_pct"].get(m, 0)) for m in methods]
    durs_raw = [str(data["duracion"].get(m, "0")).split(" ")[0] for m in methods]
    durs = []
    for d in durs_raw:
        try:
            durs.append(float(d.replace(",", ".")))
        except Exception:
            durs.append(0.0)

    c1, c2 = st.columns(2)
    with c1:
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Costo Interno", x=labels_display, y=costos,
                             marker_color=m_colors, text=[f"Q {v/1000:.0f}K" for v in costos],
                             textposition="outside"))
        fig.add_trace(go.Bar(name="Ingreso Total", x=labels_display, y=ingresos,
                             marker_color=[
                                "rgba(46,117,182,0.5)",
                                "rgba(237,125,49,0.5)",
                                "rgba(112,173,71,0.5)",
                                "rgba(158,72,14,0.5)",
                                "rgba(149,79,114,0.5)"
                            ],
                             text=[f"Q {v/1000:.0f}K" for v in ingresos],
                             textposition="outside"))
        fig.update_layout(barmode="group", title="Costo vs Ingreso por Metodología",
                          template="plotly_white", paper_bgcolor="#F7F9FC",
                          plot_bgcolor="#F7F9FC", height=390,
                          font=dict(color="#262626"), margin=dict(t=50, b=30))
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        fig2 = go.Figure(go.Bar(
            x=labels_display, y=mpcts, marker_color=m_colors,
            text=[f"{v:.1f}%" for v in mpcts], textposition="outside",
        ))
        fig2.add_hline(y=50, line_dash="dot", line_color="gray", annotation_text="50% ref")
        fig2.update_layout(title="Margen Bruto % por Metodología",
                           template="plotly_white", paper_bgcolor="#F7F9FC",
                           plot_bgcolor="#F7F9FC", height=390,
                           font=dict(color="#262626"), margin=dict(t=50, b=30),
                           yaxis_title="Margen %")
        st.plotly_chart(fig2, use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        fig3 = go.Figure(go.Bar(
            x=labels_display, y=durs, marker_color=m_colors,
            text=[f"{v:.1f} m" for v in durs], textposition="outside",
        ))
        fig3.update_layout(title="Duración del Proyecto (meses)",
                           template="plotly_white", paper_bgcolor="#F7F9FC",
                           plot_bgcolor="#F7F9FC", height=360,
                           font=dict(color="#262626"), margin=dict(t=50, b=30))
        st.plotly_chart(fig3, use_container_width=True)

    with c4:
        fig4 = go.Figure(go.Bar(
            x=labels_display, y=margenes, marker_color=m_colors,
            text=[f"Q {v/1000:.0f}K" for v in margenes], textposition="outside",
        ))
        fig4.update_layout(title="Margen Bruto (Q) por Metodología",
                           template="plotly_white", paper_bgcolor="#F7F9FC",
                           plot_bgcolor="#F7F9FC", height=360,
                           font=dict(color="#262626"), margin=dict(t=50, b=30))
        st.plotly_chart(fig4, use_container_width=True)

    # ── Ranking automático ──
    section("🏆 Ranking Automático")
    rc1, rc2, rc3 = st.columns(3)
    menor_costo_idx = costos.index(min(costos))
    mayor_margen_idx = mpcts.index(max(mpcts))
    menor_dur_idx = durs.index(min(d for d in durs if d > 0))

    with rc1:
        st.markdown(kpi_card("💵 Menor Costo Interno", labels_display[menor_costo_idx],
                             m_colors[menor_costo_idx],
                             format_currency(costos[menor_costo_idx])), unsafe_allow_html=True)
    with rc2:
        st.markdown(kpi_card("📈 Mayor Margen %", labels_display[mayor_margen_idx],
                             m_colors[mayor_margen_idx],
                             f"{mpcts[mayor_margen_idx]:.1f}%"), unsafe_allow_html=True)
    with rc3:
        st.markdown(kpi_card("⏱️ Menor Duración", labels_display[menor_dur_idx],
                             m_colors[menor_dur_idx],
                             f"{durs[menor_dur_idx]:.1f} meses"), unsafe_allow_html=True)

    note("PMI no es una metodología de ejecución. Se muestra en la sección 'PMI / Finanzas y Riesgos' como análisis financiero independiente (TMAR, ROI, ROA, Payback, Riesgos EMV).")

    # ── Notas metodológicas ──
    df_res = SHEETS.get("Resumen", pd.DataFrame())
    notas = {}
    if not df_res.empty:
        for i in range(len(df_res)):
            row0 = str(df_res.iloc[i, 0]).strip().rstrip(":")
            row1 = str(df_res.iloc[i, 1]).strip() if len(df_res.columns) > 1 else ""
            if row0 in ["SCRUM", "KANBAN", "SCRUMBAN", "XP", "WATER-SCRUM-FALL"]:
                notas[row0] = row1

    if notas:
        with st.expander("📝 Notas Metodológicas"):
            for k, v in notas.items():
                st.markdown(f"**{k}:** {v}")

# ─────────────────────────────────────────────────────────────
# PAGE: COMPARATIVO
# ─────────────────────────────────────────────────────────────
elif page == "📊 Comparativo":
    st.markdown('<div class="main-title">Comparativo de Metodologías</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">SCRUM · KANBAN · SCRUMBAN · XP · Water Scrum Fall</div>', unsafe_allow_html=True)

    data, methods = get_resumen_data()
    labels_display = ["SCRUM", "KANBAN", "SCRUMBAN", "XP", "WSF"]
    names = ["SCRUM", "KANBAN", "SCRUMBAN", "XP", "Water Scrum Fall"]
    m_colors = [COLORS["SCRUM"], COLORS["KANBAN"], COLORS["SCRUMBAN"], COLORS["XP"], COLORS["WSF"]]

    costos   = [safe_number(data["costo_total"].get(m, 0)) for m in methods]
    ingresos = [safe_number(data["ingreso_total"].get(m, 0)) for m in methods]
    margenes = [safe_number(data["margen_bruto"].get(m, 0)) for m in methods]
    mpcts = []
    for m in methods:
        v = safe_number(data["margen_pct"].get(m, 0))
        mpcts.append(v * 100 if v <= 1 else v)

    durs_raw = [str(data["duracion"].get(m, "0")).split(" ")[0] for m in methods]
    durs = []
    for d in durs_raw:
        try:
            durs.append(float(d.replace(",", ".")))
        except Exception:
            durs.append(0.0)

    # Tabla comparativa
    section("Tabla Comparativa")
    tabla = pd.DataFrame({
        "Metodología": names,
        "Costo Interno": [format_currency(c) for c in costos],
        "Ingreso Total": [format_currency(i) for i in ingresos],
        "Margen Bruto": [format_currency(m) for m in margenes],
        "Margen %": [f"{p:.2f}%" for p in mpcts],
        "Duración (meses)": [f"{d:.2f}" for d in durs],
    })
    st.dataframe(tabla, use_container_width=True, hide_index=True)

    tab1, tab2, tab3, tab4 = st.tabs(["Costo vs Ingreso", "Margen", "Duración", "Radar"])

    with tab1:
        fig = go.Figure()
        fig.add_trace(go.Bar(name="Costo Interno", x=labels_display, y=costos,
                             marker_color=m_colors,
                             text=[f"Q {v:,.0f}" for v in costos], textposition="outside"))
        fig.add_trace(go.Bar(name="Ingreso Total", x=labels_display, y=ingresos,
                             marker_color=[hex_to_rgba(c, 0.6) for c in m_colors],
                             text=[f"Q {v:,.0f}" for v in ingresos], textposition="outside"))
        fig.update_layout(barmode="group", title="Costo Interno vs Ingreso Total",
                          template="plotly_white", paper_bgcolor="#F7F9FC",
                          plot_bgcolor="#F7F9FC", height=420, font=dict(color="#262626"))
        st.plotly_chart(fig, use_container_width=True)

    with tab2:
        c1, c2 = st.columns(2)
        with c1:
            fig2 = go.Figure(go.Bar(x=labels_display, y=margenes, marker_color=m_colors,
                                    text=[f"Q {v:,.0f}" for v in margenes], textposition="outside"))
            fig2.update_layout(title="Margen Bruto (Q)", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400)
            st.plotly_chart(fig2, use_container_width=True)
        with c2:
            fig3 = go.Figure(go.Bar(x=labels_display, y=mpcts, marker_color=m_colors,
                                    text=[f"{v:.1f}%" for v in mpcts], textposition="outside"))
            fig3.add_hline(y=50, line_dash="dot", line_color="gray")
            fig3.update_layout(title="Margen Bruto %", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400)
            st.plotly_chart(fig3, use_container_width=True)

    with tab3:
        fig4 = go.Figure(go.Bar(x=labels_display, y=durs, marker_color=m_colors,
                                text=[f"{v:.1f} m" for v in durs], textposition="outside"))
        fig4.update_layout(title="Duración del Proyecto (meses)", template="plotly_white",
                           paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400,
                           yaxis_title="Meses")
        st.plotly_chart(fig4, use_container_width=True)

    with tab4:
        # Radar normalizado
        max_c = max(costos) if max(costos) > 0 else 1
        max_i = max(ingresos) if max(ingresos) > 0 else 1
        max_d = max(durs) if max(durs) > 0 else 1

        categories = ["Costo (inv)", "Ingreso", "Margen %", "Duración (inv)"]
        fig5 = go.Figure()
        for m, col, label in zip(methods, m_colors, labels_display):
            idx = methods.index(m)
            vals = [
                1 - costos[idx] / max_c,
                ingresos[idx] / max_i,
                mpcts[idx] / 100,
                1 - durs[idx] / max_d,
            ]
            vals += vals[:1]
            fig5.add_trace(go.Scatterpolar(
                r=vals, theta=categories + [categories[0]],
                fill="toself", name=label,
                line_color=col, opacity=0.7,
            ))
        fig5.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 1])),
            title="Radar Normalizado (mayor = mejor)",
            template="plotly_white", paper_bgcolor="#F7F9FC",
            font=dict(color="#262626"), height=460,
        )
        st.plotly_chart(fig5, use_container_width=True)

    # Conclusión automática
    section("🔍 Conclusión Automática")
    menor_costo_idx = costos.index(min(costos))
    mayor_margen_idx = mpcts.index(max(mpcts))
    menor_dur_idx = durs.index(min(d for d in durs if d > 0))
    # Mejor balance (suma de rangos inversos)
    rank_c = sorted(range(5), key=lambda i: costos[i])
    rank_m = sorted(range(5), key=lambda i: mpcts[i], reverse=True)
    rank_d = sorted(range(5), key=lambda i: durs[i])
    scores = [0] * 5
    for pos, idx in enumerate(rank_c): scores[idx] += pos
    for pos, idx in enumerate(rank_m): scores[idx] += pos
    for pos, idx in enumerate(rank_d): scores[idx] += pos
    best_balance_idx = scores.index(min(scores))

    st.markdown(f"""
    <div style="background:white; border-radius:12px; padding:20px; box-shadow:0 2px 8px rgba(0,0,0,0.07);">
        <ul style="line-height:2; margin:0; padding-left:18px;">
            <li><b>Menor costo interno:</b> <span style="color:{m_colors[menor_costo_idx]}; font-weight:700;">{names[menor_costo_idx]}</span> — {format_currency(costos[menor_costo_idx])}</li>
            <li><b>Mayor margen %:</b> <span style="color:{m_colors[mayor_margen_idx]}; font-weight:700;">{names[mayor_margen_idx]}</span> — {mpcts[mayor_margen_idx]:.2f}%</li>
            <li><b>Menor duración:</b> <span style="color:{m_colors[menor_dur_idx]}; font-weight:700;">{names[menor_dur_idx]}</span> — {durs[menor_dur_idx]:.1f} meses</li>
            <li><b>Mejor balance costo/margen/duración:</b> <span style="color:{m_colors[best_balance_idx]}; font-weight:700;">{names[best_balance_idx]}</span></li>
        </ul>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# PAGE: SCRUM
# ─────────────────────────────────────────────────────────────
elif page == "🔵 SCRUM":
    COLOR = COLORS["SCRUM"]
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">SCRUM — Estimación de Costos</div>', unsafe_allow_html=True)

    df = SHEETS.get("SCRUM", pd.DataFrame())
    if df.empty:
        st.warning("Hoja SCRUM no encontrada.")
    else:
        c_total  = safe_number(extract_value(df, "Costo Total del Proyecto (Interno)"))
        i_total  = safe_number(extract_value(df, "Ingreso Total del Proyecto (Venta)"))
        margen   = safe_number(extract_value(df, "Margen Bruto"))
        mpct_raw = safe_number(extract_value(df, "Margen Bruto %"))
        mpct     = mpct_raw * 100 if mpct_raw <= 1 else mpct_raw
        duracion = safe_number(extract_value(df, "Duración Total del Proyecto"))
        sprints  = safe_number(extract_value(df, "Total Sprints (con holgura)"))
        costo_sp = safe_number(extract_value(df, "Costo por Story Point"))

        section("KPIs Principales")
        kc = st.columns(4)
        kpi_data = [
            ("Costo Total", format_currency(c_total)),
            ("Ingreso Total", format_currency(i_total)),
            ("Margen Bruto", format_currency(margen)),
            ("Margen %", f"{mpct:.2f}%"),
        ]
        for col, (t, v) in zip(kc, kpi_data):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        kc2 = st.columns(3)
        kpi_data2 = [
            ("Duración", f"{duracion:.2f} meses"),
            ("Total Sprints", f"{sprints:.1f}"),
            ("Costo / SP", format_currency(costo_sp)),
        ]
        for col, (t, v) in zip(kc2, kpi_data2):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        # Supuestos
        section("Supuestos del Proyecto")
        t1, t2 = st.tabs(["Supuestos de Equipo", "Supuestos SCRUM"])
        with t1:
            equipo_rows = []
            in_equipo = False
            for i in range(len(df)):
                r0 = str(df.iloc[i, 0]).strip()
                if "SUPUESTOS DE EQUIPO" in r0.upper():
                    in_equipo = True
                    continue
                if in_equipo and any(k in r0.upper() for k in ["SUPUESTOS DE METODOLOGÍA", "BACKLOG", "NAN"]):
                    break
                if in_equipo and r0 and r0.lower() != "nan":
                    val = df.iloc[i, 1] if len(df.columns) > 1 else ""
                    unit = df.iloc[i, 2] if len(df.columns) > 2 else ""
                    equipo_rows.append({"Parámetro": r0, "Valor": val, "Unidad": unit})
            if equipo_rows:
                st.dataframe(pd.DataFrame(equipo_rows), use_container_width=True, hide_index=True)

        with t2:
            scrum_rows = []
            in_scrum = False
            for i in range(len(df)):
                r0 = str(df.iloc[i, 0]).strip()
                if "SUPUESTOS DE METODOLOGÍA" in r0.upper():
                    in_scrum = True
                    continue
                if in_scrum and any(k in r0.upper() for k in ["BACKLOG", "CÁLCULO", "RESUMEN", "NAN"]):
                    break
                if in_scrum and r0 and r0.lower() != "nan":
                    val = df.iloc[i, 1] if len(df.columns) > 1 else ""
                    unit = df.iloc[i, 2] if len(df.columns) > 2 else ""
                    scrum_rows.append({"Parámetro": r0, "Valor": val, "Unidad": unit})
            if scrum_rows:
                st.dataframe(pd.DataFrame(scrum_rows), use_container_width=True, hide_index=True)

        # Gráfico financiero
        section("Resumen Financiero")
        fig = go.Figure(go.Bar(
            x=["Costo Interno", "Ingreso Total", "Margen Bruto"],
            y=[c_total, i_total, margen],
            marker_color=[
                hex_to_rgba(COLOR, 0.5),
                hex_to_rgba(COLOR, 0.7),
                "#70AD47"
            ],
            text=[format_currency(v) for v in [c_total, i_total, margen]],
            textposition="outside",
        ))
        fig.update_layout(title="SCRUM — Costo, Ingreso y Margen",
                          template="plotly_white", paper_bgcolor="#F7F9FC",
                          plot_bgcolor="#F7F9FC", height=380, font=dict(color="#262626"))
        st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# PAGE: SCRUMBAN
# ─────────────────────────────────────────────────────────────
elif page == "🟢 SCRUMBAN":
    COLOR = COLORS["SCRUMBAN"]
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">SCRUMBAN — Estimación de Costos</div>', unsafe_allow_html=True)

    df = SHEETS.get("SCRUMBAN", pd.DataFrame())
    if df.empty:
        st.warning("Hoja SCRUMBAN no encontrada.")
    else:
        c_total   = safe_number(extract_value(df, "Costo Total"))
        i_total   = safe_number(extract_value(df, "Ingreso Total"))
        margen    = safe_number(extract_value(df, "Margen bruto"))
        mpct_raw  = safe_number(extract_value(df, "Margen %"))
        mpct      = mpct_raw * 100 if mpct_raw <= 1 else mpct_raw
        costo_build = safe_number(extract_value(df, "Costo Build"))
        costo_run   = safe_number(extract_value(df, "Costo Run"))
        precio_build = safe_number(extract_value(df, "Precio Build"))
        precio_run  = safe_number(extract_value(df, "Precio Run"))
        sprints_h   = safe_number(extract_value(df, "Sprints con holgura"))

        section("KPIs Principales")
        kc = st.columns(5)
        kd = [
            ("Costo Total", format_currency(c_total)),
            ("Ingreso Total", format_currency(i_total)),
            ("Margen Bruto", format_currency(margen)),
            ("Margen %", f"{mpct:.2f}%"),
            ("Sprints c/holgura", f"{sprints_h:.0f}"),
        ]
        for col, (t, v) in zip(kc, kd):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        section("Fase BUILD vs Fase RUN")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(kpi_card("Costo Build", format_currency(costo_build), "#2E75B6"), unsafe_allow_html=True)
            st.markdown(kpi_card("Precio Build", format_currency(precio_build), "#2E75B6"), unsafe_allow_html=True)
        with c2:
            st.markdown(kpi_card("Costo Run", format_currency(costo_run), COLOR), unsafe_allow_html=True)
            st.markdown(kpi_card("Precio Run", format_currency(precio_run), COLOR), unsafe_allow_html=True)

        section("Gráficos")
        col1, col2 = st.columns(2)
        with col1:
            fig1 = go.Figure(go.Bar(
                x=["Costo Build", "Costo Run"],
                y=[costo_build, costo_run],
                marker_color=["#2E75B6", COLOR],
                text=[format_currency(v) for v in [costo_build, costo_run]],
                textposition="outside",
            ))
            fig1.update_layout(title="Costo Build vs Run", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=360)
            st.plotly_chart(fig1, use_container_width=True)
        with col2:
            fig2 = go.Figure(go.Bar(
                x=["Costo Total", "Ingreso Total", "Margen Bruto"],
                y=[c_total, i_total, margen],
                marker_color=[
                hex_to_rgba(COLOR, 0.5),
                hex_to_rgba(COLOR, 0.7),
                "#70AD47"
                ],
                text=[format_currency(v) for v in [c_total, i_total, margen]],
                textposition="outside",
            ))
            fig2.update_layout(title="Resumen Financiero SCRUMBAN", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=360)
            st.plotly_chart(fig2, use_container_width=True)

        # Tabla de supuestos
        section("Parámetros del Proyecto")
        rows = []
        for i in range(len(df)):
            r0 = str(df.iloc[i, 0]).strip()
            r1 = df.iloc[i, 1] if len(df.columns) > 1 else ""
            r2 = df.iloc[i, 2] if len(df.columns) > 2 else ""
            if r0 and r0.lower() not in ["nan", ""] and pd.notna(r1) and str(r1).lower() not in ["nan", ""]:
                if not any(kw in r0.upper() for kw in ["SUPUESTO", "BACKLOG", "COSTOS", "MARGEN OBJ", "SCRUMBAN —"]):
                    rows.append({"Parámetro": r0, "Valor": r1, "Unidad": r2})
        if rows:
            with st.expander("Ver todos los parámetros"):
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────
# PAGE: KANBAN
# ─────────────────────────────────────────────────────────────
elif page == "🟠 KANBAN":
    COLOR = COLORS["KANBAN"]
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">KANBAN — Estimación de Costos</div>', unsafe_allow_html=True)

    df = SHEETS.get("KANBAN", pd.DataFrame())
    if df.empty:
        st.warning("Hoja KANBAN no encontrada.")
    else:
        c_total  = safe_number(extract_value(df, "Costo Total del Proyecto (Interno)"))
        i_total  = safe_number(extract_value(df, "Ingreso Total del Proyecto (Venta)"))
        margen   = safe_number(extract_value(df, "Margen Bruto"))
        mpct_raw = safe_number(extract_value(df, "Margen Bruto %"))
        mpct     = mpct_raw * 100 if mpct_raw <= 1 else mpct_raw
        cycle_t  = safe_number(extract_value(df, "Cycle Time promedio"))
        lead_t   = safe_number(extract_value(df, "Lead Time promedio"))
        wip      = safe_number(extract_value(df, "WIP Limit"))
        throughput = safe_number(extract_value(df, "Throughput"))
        duracion = safe_number(extract_value(df, "Duración Total del Proyecto"))
        costo_h  = safe_number(extract_value(df, "Costo por Historia"))

        section("KPIs Principales")
        k1 = st.columns(5)
        kd = [
            ("Costo Total", format_currency(c_total)),
            ("Ingreso Total", format_currency(i_total)),
            ("Margen Bruto", format_currency(margen)),
            ("Margen %", f"{mpct:.2f}%"),
            ("Duración", f"{duracion:.0f} meses"),
        ]
        for col, (t, v) in zip(k1, kd):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        k2 = st.columns(4)
        kd2 = [
            ("Cycle Time", f"{cycle_t:.0f} días"),
            ("Lead Time", f"{lead_t:.0f} días"),
            ("WIP Limit", f"{wip:.0f} tareas"),
            ("Throughput", f"{throughput:.0f} hist/mes"),
        ]
        for col, (t, v) in zip(k2, kd2):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        section("Métricas de Flujo Kanban")
        c1, c2 = st.columns(2)
        with c1:
            fig1 = go.Figure(go.Bar(
                x=["Cycle Time", "Lead Time"],
                y=[cycle_t, lead_t],
                marker_color=[
                hex_to_rgba(COLOR, 0.5),
                hex_to_rgba(COLOR, 0.7),
                "#70AD47"
                ],
                text=[f"{v:.0f} días" for v in [cycle_t, lead_t]],
                textposition="outside",
            ))
            fig1.update_layout(title="Cycle Time vs Lead Time (días)",
                               template="plotly_white", paper_bgcolor="#F7F9FC",
                               plot_bgcolor="#F7F9FC", height=360)
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            fig2 = go.Figure(go.Bar(
                x=["Costo Interno", "Ingreso Total", "Margen Bruto"],
                y=[c_total, i_total, margen],
                marker_color=[
                hex_to_rgba(COLOR, 0.5),
                hex_to_rgba(COLOR, 0.7),
                "#70AD47"
                ],
                text=[format_currency(v) for v in [c_total, i_total, margen]],
                textposition="outside",
            ))
            fig2.update_layout(title="Resumen Financiero KANBAN",
                               template="plotly_white", paper_bgcolor="#F7F9FC",
                               plot_bgcolor="#F7F9FC", height=360)
            st.plotly_chart(fig2, use_container_width=True)

        section("Parámetros Kanban")
        param_rows = []
        for i in range(len(df)):
            r0 = str(df.iloc[i, 0]).strip()
            r1 = df.iloc[i, 1] if len(df.columns) > 1 else ""
            r2 = df.iloc[i, 2] if len(df.columns) > 2 else ""
            if r0 and r0.lower() not in ["nan", ""] and pd.notna(r1) and str(r1).lower() not in ["nan", ""]:
                param_rows.append({"Parámetro": r0, "Valor": r1, "Unidad": r2})
        if param_rows:
            st.dataframe(pd.DataFrame(param_rows), use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────
# PAGE: XP
# ─────────────────────────────────────────────────────────────
elif page == "🟤 XP":
    COLOR = COLORS["XP"]
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">Extreme Programming (XP) — Estimación de Costos</div>', unsafe_allow_html=True)

    df = SHEETS.get("XP", pd.DataFrame())
    if df.empty:
        st.warning("Hoja XP no encontrada.")
    else:
        burn_rate   = safe_number(extract_value(df, "BURN RATE (Costo real mensual)"))
        total_sp    = safe_number(extract_value(df, "Total Story Points del Proyecto"))
        velocidad   = safe_number(extract_value(df, "Velocidad del Equipo"))
        sprints_h   = safe_number(extract_value(df, "Total Sprints (con holgura)"))
        duracion    = safe_number(extract_value(df, "Duración Total del Proyecto"))
        precio_mens = safe_number(extract_value(df, "Precio al cliente mensual"))
        c_total     = safe_number(extract_value(df, "Costo Total del Proyecto"))
        venta       = safe_number(extract_value(df, "Venta Proyecto"))
        ganancia    = safe_number(extract_value(df, "Ganancia Total"))
        margen_mens = safe_number(extract_value(df, "Margen Bruto mensual"))

        section("KPIs Principales")
        k1 = st.columns(5)
        kd = [
            ("Burn Rate Mensual", format_currency(burn_rate)),
            ("Total Story Points", f"{total_sp:.0f} SP"),
            ("Velocidad", f"{velocidad:.0f} SP/sprint"),
            ("Total Sprints", f"{sprints_h:.1f}"),
            ("Duración", f"{duracion:.2f} meses"),
        ]
        for col, (t, v) in zip(k1, kd):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        k2 = st.columns(4)
        kd2 = [
            ("Precio Mensual", format_currency(precio_mens)),
            ("Costo Total", format_currency(c_total)),
            ("Venta Total", format_currency(venta)),
            ("Ganancia Total", format_currency(ganancia)),
        ]
        for col, (t, v) in zip(k2, kd2):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        section("Gráficos Financieros")
        c1, c2 = st.columns(2)
        with c1:
            fig1 = go.Figure(go.Bar(
                x=["Costo Mensual", "Precio Mensual", "Margen Mensual"],
                y=[burn_rate, precio_mens, margen_mens],
                marker_color=[
                hex_to_rgba(COLOR, 0.5),
                hex_to_rgba(COLOR, 0.7),
                "#70AD47"
                ],
                text=[format_currency(v) for v in [burn_rate, precio_mens, margen_mens]],
                textposition="outside",
            ))
            fig1.update_layout(title="Resumen Financiero Mensual", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=380)
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            mpct_total = ((venta - c_total) / venta * 100) if venta > 0 else 0
            fig2 = go.Figure(go.Bar(
                x=["Costo Total", "Venta Total", "Ganancia Total"],
                y=[c_total, venta, ganancia],
                marker_color=[
                hex_to_rgba(COLOR, 0.5),
                hex_to_rgba(COLOR, 0.7),
                "#70AD47"
                ],
                text=[format_currency(v) for v in [c_total, venta, ganancia]],
                textposition="outside",
            ))
            fig2.update_layout(title="Resumen Financiero Total del Proyecto",
                               template="plotly_white", paper_bgcolor="#F7F9FC",
                               plot_bgcolor="#F7F9FC", height=380)
            st.plotly_chart(fig2, use_container_width=True)

        section("🏅 Prácticas de Calidad XP")
        xp_practices = [
            ("👥 Pair Programming", "Dos desarrolladores trabajan juntos en el mismo código. Reduce defectos y transfiere conocimiento."),
            ("✅ TDD", "Test-Driven Development: las pruebas se escriben antes del código. Garantiza calidad desde el inicio."),
            ("🔄 Integración Continua", "El código se integra y prueba múltiples veces al día. Detecta errores rápidamente."),
            ("♻️ Refactoring", "Mejora continua del código sin cambiar su funcionalidad. Mantiene la deuda técnica bajo control."),
        ]
        xp_cols = st.columns(4)
        for col, (title_p, desc) in zip(xp_cols, xp_practices):
            with col:
                st.markdown(f"""
                <div class="kpi-card" style="border-top-color:{COLOR}; text-align:left;">
                    <div style="font-size:1.1rem; font-weight:700; margin-bottom:8px;">{title_p}</div>
                    <div style="font-size:0.82rem; color:#555;">{desc}</div>
                </div>""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# PAGE: WATER SCRUM FALL
# ─────────────────────────────────────────────────────────────
elif page == "🟣 Water Scrum Fall":
    COLOR = COLORS["Water Scrum Fall"]
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">Water-Scrum-Fall — Estimación de Costos</div>', unsafe_allow_html=True)

    df = SHEETS.get("Water Scrum Fall", pd.DataFrame())
    if df.empty:
        st.warning("Hoja 'Water Scrum Fall' no encontrada.")
    else:
        costo_water = safe_number(extract_value(df, "COSTO TOTAL FASE WATER"))
        costo_scrum = safe_number(extract_value(df, "COSTO TOTAL FASE SCRUM"))
        costo_fall  = safe_number(extract_value(df, "COSTO TOTAL FASE FALL"))
        costo_extra = safe_number(extract_value(df, "COSTO TOTAL EXTRAS"))
        c_total     = safe_number(extract_value(df, "COSTO TOTAL INTERNO DEL PROYECTO"))
        precio_fijo = safe_number(extract_value(df, "PRECIO FIJO AL CLIENTE"))
        margen      = safe_number(extract_value(df, "Margen Bruto (Ganancia)"))
        dur_water   = 2.0
        dur_scrum   = safe_number(extract_value(df, "Duración fase Scrum"))
        dur_fall    = 1.5
        dur_total   = dur_water + dur_scrum + dur_fall
        mpct_raw    = safe_number(extract_value(df, "Margen objetivo (sobre venta)"))
        mpct        = mpct_raw * 100 if mpct_raw <= 1 else mpct_raw

        section("KPIs Principales")
        k1 = st.columns(5)
        kd = [
            ("Costo Interno Total", format_currency(c_total)),
            ("Precio Fijo Cliente", format_currency(precio_fijo)),
            ("Margen Bruto", format_currency(margen)),
            ("Margen %", f"{mpct:.1f}%"),
            ("Duración Total", f"{dur_total:.2f} meses"),
        ]
        for col, (t, v) in zip(k1, kd):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        section("Costos por Fase")
        k2 = st.columns(4)
        kd2 = [
            ("Fase Water", format_currency(costo_water)),
            ("Fase Scrum", format_currency(costo_scrum)),
            ("Fase Fall", format_currency(costo_fall)),
            ("Extras WSF", format_currency(costo_extra)),
        ]
        fase_colors = ["#4472C4", "#2E75B6", COLOR, "#A9A9A9"]
        for col, (t, v), fc in zip(k2, kd2, fase_colors):
            with col:
                st.markdown(kpi_card(t, v, fc), unsafe_allow_html=True)

        section("Gráficos")
        tab1, tab2, tab3 = st.tabs(["Costos por Fase", "Duración", "Costo vs Precio"])

        with tab1:
            fases = ["Fase Water", "Fase Scrum", "Fase Fall", "Extras WSF"]
            costos_fases = [costo_water, costo_scrum, costo_fall, costo_extra]
            fig1 = go.Figure(go.Bar(
                x=fases, y=costos_fases, marker_color=fase_colors,
                text=[format_currency(v) for v in costos_fases], textposition="outside",
            ))
            fig1.update_layout(title="Costo por Fase", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400)
            st.plotly_chart(fig1, use_container_width=True)

        with tab2:
            fases_d = ["Fase Water", "Fase Scrum", "Fase Fall"]
            durs_fases = [dur_water, dur_scrum, dur_fall]
            fig2 = go.Figure(go.Bar(
                x=fases_d, y=durs_fases, marker_color=fase_colors[:3],
                text=[f"{v:.2f} m" for v in durs_fases], textposition="outside",
            ))
            fig2.update_layout(title="Duración por Fase (meses)", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400)
            st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            fig3 = go.Figure(go.Bar(
                x=["Costo Total Interno", "Precio Fijo Cliente", "Margen Bruto"],
                y=[c_total, precio_fijo, margen],
                marker_color=[
                hex_to_rgba(COLOR, 0.5),
                hex_to_rgba(COLOR, 0.7),
                "#70AD47"
                ],
                text=[format_currency(v) for v in [c_total, precio_fijo, margen]],
                textposition="outside",
            ))
            fig3.update_layout(title="Costo Total vs Precio Fijo", template="plotly_white",
                               paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400)
            st.plotly_chart(fig3, use_container_width=True)

        # Timeline visual
        section("🗺️ Timeline de Fases: Water → Scrum → Fall")
        timeline_fig = go.Figure()
        phase_data = [
            ("Water", 0, dur_water, "#4472C4"),
            ("Scrum", dur_water, dur_water + dur_scrum, "#2E75B6"),
            ("Fall", dur_water + dur_scrum, dur_water + dur_scrum + dur_fall, COLOR),
        ]
        for pname, start, end, pcolor in phase_data:
            timeline_fig.add_trace(go.Bar(
                x=[end - start], y=["Proyecto WSF"],
                orientation="h", base=start,
                marker_color=pcolor, name=pname,
                text=f"  {pname}<br>{end-start:.1f} m",
                textposition="inside",
                insidetextanchor="middle",
            ))
        timeline_fig.update_layout(
            barmode="stack", title="Timeline Water → Scrum → Fall",
            template="plotly_white", paper_bgcolor="#F7F9FC",
            plot_bgcolor="#F7F9FC", height=200,
            xaxis_title="Meses", font=dict(color="#262626"),
            legend=dict(orientation="h", y=-0.5),
        )
        st.plotly_chart(timeline_fig, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# PAGE: PMI / FINANZAS
# ─────────────────────────────────────────────────────────────
elif page == "💰 PMI / Finanzas y Riesgos":
    COLOR = COLORS["PMI"]
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">Análisis Financiero PMI — TransportesNova</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Reserva de Riesgos (EMV) · TMAR · ROI · ROA · Payback</div>', unsafe_allow_html=True)
    note("PMI no es una metodología de ejecución. Esta sección presenta el análisis financiero y gestión de riesgos según el estándar PMBOK.")

    df = SHEETS.get("PMI", pd.DataFrame())
    if df.empty:
        st.warning("Hoja PMI no encontrada.")
    else:
        methods = ["SCRUM", "KANBAN", "SCRUMBAN", "XP", "WSF"]
        m_colors = [COLORS["SCRUM"], COLORS["KANBAN"], COLORS["SCRUMBAN"], COLORS["XP"], COLORS["WSF"]]

        # Extract key rows by searching for row headers
        def get_row_by_label(df, label, col_start=1, col_end=6):
            for i in range(len(df)):
                if label.lower() in str(df.iloc[i, 0]).lower():
                    return [safe_number(df.iloc[i, j]) if j < len(df.columns) else 0
                            for j in range(col_start, col_end + 1)]
            return [0] * (col_end - col_start + 1)

        presupuesto_base = get_row_by_label(df, "PRESUPUESTO BASE", 1, 5)
        reserva_riesgos  = get_row_by_label(df, "RESERVA DE RIESGOS TOTAL")[:5]
        ppto_riesgos     = get_row_by_label(df, "PRESUPUESTO INTERNO CON RIESGOS", 1, 5)
        roi_vals         = get_row_by_label(df, "ROI", 1, 5)
        roa_vals         = get_row_by_label(df, "ROA", 1, 5)
        payback_vals     = get_row_by_label(df, "Payback (meses)", 1, 5)

        # Single values
        tmar = safe_number(extract_value(df, "TMAR TOTAL"))
        reserva_total = safe_number(extract_value(df, "RESERVA DE RIESGOS TOTAL (Σ EMV)"))
        beneficio_anual = safe_number(extract_value(df, "BENEFICIO ANUAL TOTAL"))
        activos = safe_number(extract_value(df, "ACTIVOS TOTALES"))

        section("KPIs PMI")
        kc = st.columns(4)
        kd = [
            ("TMAR", f"{tmar*100:.0f}%"),
            ("Reserva de Riesgos (EMV)", format_currency(reserva_total)),
            ("Beneficio Anual Sistema", format_currency(beneficio_anual)),
            ("Activos Totales", format_currency(activos)),
        ]
        for col, (t, v) in zip(kc, kd):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        tab1, tab2, tab3, tab4 = st.tabs(["Presupuesto", "ROI / TMAR", "Payback / ROA", "Riesgos EMV"])

        with tab1:
            c1, c2 = st.columns(2)
            with c1:
                fig1 = go.Figure(go.Bar(
                    x=methods, y=presupuesto_base, marker_color=m_colors,
                    text=[format_currency(v) for v in presupuesto_base], textposition="outside",
                ))
                fig1.update_layout(title="Presupuesto Base por Metodología",
                                   template="plotly_white", paper_bgcolor="#F7F9FC",
                                   plot_bgcolor="#F7F9FC", height=400)
                st.plotly_chart(fig1, use_container_width=True)
            with c2:
                fig2 = go.Figure(go.Bar(
                    x=methods, y=ppto_riesgos, marker_color=m_colors,
                    text=[format_currency(v) for v in ppto_riesgos], textposition="outside",
                ))
                fig2.update_layout(title="Presupuesto con Riesgos por Metodología",
                                   template="plotly_white", paper_bgcolor="#F7F9FC",
                                   plot_bgcolor="#F7F9FC", height=400)
                st.plotly_chart(fig2, use_container_width=True)

        with tab2:
            fig3 = go.Figure()
            fig3.add_trace(go.Bar(name="ROI", x=methods, y=roi_vals,
                                  marker_color=m_colors,
                                  text=[f"{v:.2f}" for v in roi_vals], textposition="outside"))
            fig3.add_hline(y=tmar, line_dash="dash", line_color="red",
                          annotation_text=f"TMAR = {tmar*100:.0f}%")
            fig3.update_layout(title="ROI vs TMAR por Metodología — Viabilidad Financiera",
                               template="plotly_white", paper_bgcolor="#F7F9FC",
                               plot_bgcolor="#F7F9FC", height=420, yaxis_title="ROI")
            st.plotly_chart(fig3, use_container_width=True)

            st.markdown("**Veredicto:** Todas las metodologías tienen ROI > TMAR → Proyecto VIABLE ✓")
            viab_df = pd.DataFrame({
                "Metodología": methods,
                "ROI": [f"{v:.3f}" for v in roi_vals],
                "TMAR": [f"{tmar:.2f}"] * 5,
                "Veredicto": ["✅ VIABLE" if v > tmar else "❌ NO VIABLE" for v in roi_vals],
            })
            st.dataframe(viab_df, use_container_width=True, hide_index=True)

        with tab3:
            c1, c2 = st.columns(2)
            with c1:
                fig4 = go.Figure(go.Bar(
                    x=methods, y=payback_vals, marker_color=m_colors,
                    text=[f"{v:.2f} m" for v in payback_vals], textposition="outside",
                ))
                fig4.update_layout(title="Payback Period (meses)",
                                   template="plotly_white", paper_bgcolor="#F7F9FC",
                                   plot_bgcolor="#F7F9FC", height=380, yaxis_title="Meses")
                st.plotly_chart(fig4, use_container_width=True)
            with c2:
                fig5 = go.Figure(go.Bar(
                    x=methods, y=roa_vals, marker_color=m_colors,
                    text=[f"{v:.3f}" for v in roa_vals], textposition="outside",
                ))
                fig5.update_layout(title="ROA por Metodología",
                                   template="plotly_white", paper_bgcolor="#F7F9FC",
                                   plot_bgcolor="#F7F9FC", height=380, yaxis_title="ROA")
                st.plotly_chart(fig5, use_container_width=True)

        with tab4:
            # Extract risk table
            risk_rows = []
            in_risks = False
            for i in range(len(df)):
                r = [str(df.iloc[i, j]).strip() if j < len(df.columns) else "" for j in range(7)]
                if "Riesgo identificado" in r[0]:
                    in_risks = True
                    continue
                if in_risks and "RESERVA DE RIESGOS TOTAL" in r[0].upper():
                    break
                if in_risks and r[0] and r[0] not in ["nan", ""] and r[1] not in ["nan", ""]:
                    risk_rows.append({
                        "Riesgo": r[0],
                        "Probabilidad": r[1],
                        "Impacto (Q)": r[2],
                        "EMV (Q)": r[3],
                        "Comentario": r[6] if len(r) > 6 else "",
                    })
            if risk_rows:
                risk_df = pd.DataFrame(risk_rows)
                st.dataframe(risk_df, use_container_width=True, hide_index=True)

                # Bubble chart (Prob × Impact)
                probs = [safe_number(r["Probabilidad"]) for r in risk_rows]
                impacts = [safe_number(r["Impacto (Q)"]) for r in risk_rows]
                emvs = [safe_number(r["EMV (Q)"]) for r in risk_rows]
                labels_r = [r["Riesgo"][:30] for r in risk_rows]

                fig6 = go.Figure(go.Scatter(
                    x=probs, y=impacts,
                    mode="markers+text",
                    marker=dict(size=[e / 500 for e in emvs],
                               color=emvs, colorscale="Reds",
                               showscale=True, colorbar=dict(title="EMV (Q)")),
                    text=labels_r,
                    textposition="top center",
                ))
                fig6.update_layout(
                    title="Matriz de Riesgos (Probabilidad × Impacto)",
                    xaxis_title="Probabilidad", yaxis_title="Impacto (Q)",
                    template="plotly_white", paper_bgcolor="#F7F9FC",
                    plot_bgcolor="#F7F9FC", height=460, font=dict(color="#262626"),
                )
                st.plotly_chart(fig6, use_container_width=True)

# ─────────────────────────────────────────────────────────────
# PAGE: EQUIPO
# ─────────────────────────────────────────────────────────────
elif page == "👥 Equipo":
    COLOR = "#4472C4"
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">Equipo de Trabajo — Supuestos Compartidos</div>', unsafe_allow_html=True)

    df = SHEETS.get("Equipo", pd.DataFrame())
    if df.empty:
        st.warning("Hoja Equipo no encontrada.")
    else:
        # Find header row
        header_row = None
        for i in range(len(df)):
            if "Rol" in str(df.iloc[i, 0]):
                header_row = i
                break
        if header_row is not None:
            team_df = df.iloc[header_row + 1:].copy()
            team_df.columns = ["Rol", "Descripción", "Salario Mensual (Q)", "FTE Asignado", "Costo Asignado (Q)"] + \
                              [f"extra_{i}" for i in range(max(0, len(team_df.columns) - 5))]
            cols_keep = ["Rol", "Descripción", "Salario Mensual (Q)", "FTE Asignado", "Costo Asignado (Q)"]
            team_df = team_df[cols_keep].dropna(subset=["Rol"])
            # Remove totals row
            team_df = team_df[~team_df["Rol"].astype(str).str.contains("COSTO", na=False)]
            team_df = team_df[team_df["Rol"].astype(str).str.lower() != "nan"].reset_index(drop=True)
            team_df["Salario Mensual (Q)"] = team_df["Salario Mensual (Q)"].apply(safe_number)
            team_df["FTE Asignado"] = team_df["FTE Asignado"].apply(safe_number)
            team_df["Costo Asignado (Q)"] = team_df["Costo Asignado (Q)"].apply(safe_number)

            costo_total_mens = team_df["Costo Asignado (Q)"].sum()
            fte_total = team_df["FTE Asignado"].sum()
            n_roles = len(team_df)

            section("KPIs del Equipo")
            kc = st.columns(4)
            kd = [
                ("Costo Total Mensual", format_currency(costo_total_mens)),
                ("Costo Interno 50% FTE", format_currency(costo_total_mens)),
                ("Número de Roles", str(n_roles)),
                ("FTE Total", f"{fte_total:.1f}"),
            ]
            for col, (t, v) in zip(kc, kd):
                with col:
                    st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

            section("Distribución del Equipo")
            c1, c2 = st.columns(2)
            with c1:
                fig1 = go.Figure(go.Bar(
                    y=team_df["Rol"], x=team_df["Costo Asignado (Q)"],
                    orientation="h",
                    marker_color=COLOR,
                    text=[format_currency(v) for v in team_df["Costo Asignado (Q)"]],
                    textposition="outside",
                ))
                fig1.update_layout(title="Costo Asignado por Rol", template="plotly_white",
                                   paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=380)
                st.plotly_chart(fig1, use_container_width=True)
            with c2:
                fig2 = go.Figure(go.Bar(
                    y=team_df["Rol"], x=team_df["Salario Mensual (Q)"],
                    orientation="h",
                    marker_color=[
                    hex_to_rgba(COLOR, 0.5),
                    hex_to_rgba(COLOR, 0.7),
                    "#70AD47"
                    ],
                    text=[format_currency(v) for v in team_df["Salario Mensual (Q)"]],
                    textposition="outside",
                ))
                fig2.update_layout(title="Salario Mensual por Rol", template="plotly_white",
                                   paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=380)
                st.plotly_chart(fig2, use_container_width=True)

            section("Tabla del Equipo")
            display_df = team_df.copy()
            display_df["Salario Mensual (Q)"] = display_df["Salario Mensual (Q)"].apply(format_currency)
            display_df["Costo Asignado (Q)"]  = display_df["Costo Asignado (Q)"].apply(format_currency)
            st.dataframe(display_df, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────
# PAGE: BACKLOG
# ─────────────────────────────────────────────────────────────
elif page == "📋 Backlog":
    COLOR = "#2E75B6"
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">Backlog del Proyecto</div>', unsafe_allow_html=True)

    df_raw = SHEETS.get("Backlog", pd.DataFrame())
    if df_raw.empty:
        st.warning("Hoja Backlog no encontrada.")
    else:
        df = clean_backlog_like_sheet(df_raw)

        if df.empty:
            st.warning("No se encontraron historias válidas en el Backlog.")
        else:
            total_h = len(df)
            total_sp = int(df["SP"].sum())
            n_epicas = df["Épica"].nunique() if "Épica" in df.columns else 0
            h_high   = len(df[df["Prioridad"] == "High"]) if "Prioridad" in df.columns else 0
            h_med    = len(df[df["Prioridad"] == "Medium"]) if "Prioridad" in df.columns else 0
            h_low    = len(df[df["Prioridad"] == "Low"]) if "Prioridad" in df.columns else 0

            section("KPIs del Backlog")
            kc = st.columns(6)
            kd = [
                ("Total Historias", str(total_h)),
                ("Total Story Points", str(total_sp)),
                ("Épicas", str(n_epicas)),
                ("High", str(h_high)),
                ("Medium", str(h_med)),
                ("Low", str(h_low)),
            ]
            ep_colors = [COLOR, "#70AD47", "#ED7D31", "#C00000", "#ED7D31", "#A9A9A9"]
            for col, (t, v), c in zip(kc, kd, ep_colors):
                with col:
                    st.markdown(kpi_card(t, v, c), unsafe_allow_html=True)

            # Filtros
            section("Filtros")
            fc1, fc2, fc3 = st.columns(3)
            all_epicas = ["Todas"] + sorted(df["Épica"].dropna().unique().tolist()) if "Épica" in df.columns else ["Todas"]
            all_prios  = ["Todas"] + sorted(df["Prioridad"].dropna().unique().tolist()) if "Prioridad" in df.columns else ["Todas"]
            all_estados= ["Todos"] + sorted(df["Estado"].dropna().unique().tolist()) if "Estado" in df.columns else ["Todos"]

            with fc1:
                sel_epica  = st.selectbox("Épica", all_epicas)
            with fc2:
                sel_prio   = st.selectbox("Prioridad", all_prios)
            with fc3:
                sel_estado = st.selectbox("Estado", all_estados)

            df_f = df.copy()
            if sel_epica != "Todas" and "Épica" in df_f.columns:
                df_f = df_f[df_f["Épica"] == sel_epica]
            if sel_prio != "Todas" and "Prioridad" in df_f.columns:
                df_f = df_f[df_f["Prioridad"] == sel_prio]
            if sel_estado != "Todos" and "Estado" in df_f.columns:
                df_f = df_f[df_f["Estado"] == sel_estado]

            section("Gráficos")
            tab1, tab2, tab3 = st.tabs(["SP por Épica", "Prioridad", "Estado"])

            with tab1:
                if "Épica" in df.columns:
                    ep_sp  = df.groupby("Épica")["SP"].sum().reset_index().sort_values("SP", ascending=False)
                    ep_cnt = df.groupby("Épica").size().reset_index(name="Historias")
                    c1, c2 = st.columns(2)
                    with c1:
                        fig = go.Figure(go.Bar(
                            x=ep_sp["Épica"], y=ep_sp["SP"],
                            marker_color=COLOR,
                            text=ep_sp["SP"].astype(str), textposition="outside",
                        ))
                        fig.update_layout(title="Story Points por Épica", template="plotly_white",
                                          paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400,
                                          xaxis_tickangle=-30)
                        st.plotly_chart(fig, use_container_width=True)
                    with c2:
                        fig2 = go.Figure(go.Bar(
                            x=ep_cnt["Épica"], y=ep_cnt["Historias"],
                            marker_color=[
                            hex_to_rgba(COLOR, 0.5),
                            hex_to_rgba(COLOR, 0.7),
                            "#70AD47"
                            ],
                            text=ep_cnt["Historias"].astype(str), textposition="outside",
                        ))
                        fig2.update_layout(title="Historias por Épica", template="plotly_white",
                                           paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400,
                                           xaxis_tickangle=-30)
                        st.plotly_chart(fig2, use_container_width=True)

            with tab2:
                if "Prioridad" in df.columns:
                    prio_cnt = df["Prioridad"].value_counts()
                    prio_colors = {"High": "#C00000", "Medium": "#ED7D31", "Low": "#A9A9A9"}
                    fig = go.Figure(go.Pie(
                        labels=prio_cnt.index, values=prio_cnt.values,
                        marker_colors=[prio_colors.get(p, "#888") for p in prio_cnt.index],
                        hole=0.4,
                    ))
                    fig.update_layout(title="Distribución por Prioridad", template="plotly_white",
                                      paper_bgcolor="#F7F9FC", height=380)
                    st.plotly_chart(fig, use_container_width=True)

            with tab3:
                if "Estado" in df.columns:
                    est_cnt = df["Estado"].value_counts()
                    fig = go.Figure(go.Bar(
                        x=est_cnt.index, y=est_cnt.values,
                        marker_color=COLOR,
                        text=est_cnt.values, textposition="outside",
                    ))
                    fig.update_layout(title="Historias por Estado", template="plotly_white",
                                      paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=380)
                    st.plotly_chart(fig, use_container_width=True)

            section(f"Tabla del Backlog ({len(df_f)} historias filtradas)")
            with st.expander("Ver tabla"):
                st.dataframe(df_f, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────
# PAGE: MVP
# ─────────────────────────────────────────────────────────────
elif page == "🎯 MVP":
    COLOR = "#70AD47"
    st.markdown(f'<div class="main-title" style="border-color:{COLOR};">MVP — Minimum Viable Product</div>', unsafe_allow_html=True)

    df_raw_mvp  = SHEETS.get("MVP", pd.DataFrame())
    df_raw_back = SHEETS.get("Backlog", pd.DataFrame())

    if df_raw_mvp.empty:
        st.warning("Hoja MVP no encontrada.")
    else:
        df_mvp  = clean_backlog_like_sheet(df_raw_mvp)
        df_back = clean_backlog_like_sheet(df_raw_back) if not df_raw_back.empty else pd.DataFrame()

        total_mvp_h  = len(df_mvp)
        total_mvp_sp = int(df_mvp["SP"].sum()) if not df_mvp.empty else 0
        n_epicas_mvp = df_mvp["Épica"].nunique() if "Épica" in df_mvp.columns else 0
        total_back_h  = len(df_back)
        total_back_sp = int(df_back["SP"].sum()) if not df_back.empty else 1
        pct_h  = total_mvp_h  / max(total_back_h, 1) * 100
        pct_sp = total_mvp_sp / max(total_back_sp, 1) * 100

        section("KPIs del MVP")
        kc = st.columns(4)
        kd = [
            ("Historias MVP", str(total_mvp_h)),
            ("Story Points MVP", str(total_mvp_sp)),
            ("Épicas Incluidas", str(n_epicas_mvp)),
            ("% del Backlog (SP)", f"{pct_sp:.1f}%"),
        ]
        for col, (t, v) in zip(kc, kd):
            with col:
                st.markdown(kpi_card(t, v, COLOR), unsafe_allow_html=True)

        section("MVP vs Backlog Total")
        c1, c2 = st.columns(2)
        with c1:
            fig1 = go.Figure(go.Bar(
                x=["MVP", "Backlog Total"],
                y=[total_mvp_h, total_back_h],
                marker_color=[COLOR, "#A9A9A9"],
                text=[str(total_mvp_h), str(total_back_h)],
                textposition="outside",
            ))
            fig1.update_layout(title="Historias: MVP vs Backlog Total",
                               template="plotly_white", paper_bgcolor="#F7F9FC",
                               plot_bgcolor="#F7F9FC", height=360)
            st.plotly_chart(fig1, use_container_width=True)
        with c2:
            fig2 = go.Figure(go.Bar(
                x=["MVP SP", "Backlog SP Total"],
                y=[total_mvp_sp, total_back_sp],
                marker_color=[COLOR, "#A9A9A9"],
                text=[str(total_mvp_sp), str(total_back_sp)],
                textposition="outside",
            ))
            fig2.update_layout(title="Story Points: MVP vs Backlog Total",
                               template="plotly_white", paper_bgcolor="#F7F9FC",
                               plot_bgcolor="#F7F9FC", height=360)
            st.plotly_chart(fig2, use_container_width=True)

        if not df_mvp.empty:
            section("Gráficos del MVP")
            tab1, tab2, tab3 = st.tabs(["SP por Épica", "Historias por Épica", "Prioridad"])

            with tab1:
                if "Épica" in df_mvp.columns:
                    ep_sp = df_mvp.groupby("Épica")["SP"].sum().reset_index().sort_values("SP", ascending=False)
                    fig = go.Figure(go.Bar(
                        x=ep_sp["Épica"], y=ep_sp["SP"],
                        marker_color=COLOR,
                        text=ep_sp["SP"].astype(str), textposition="outside",
                    ))
                    fig.update_layout(title="SP del MVP por Épica", template="plotly_white",
                                      paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400,
                                      xaxis_tickangle=-30)
                    st.plotly_chart(fig, use_container_width=True)

            with tab2:
                if "Épica" in df_mvp.columns:
                    ep_cnt = df_mvp.groupby("Épica").size().reset_index(name="Historias")
                    fig = go.Figure(go.Bar(
                        x=ep_cnt["Épica"], y=ep_cnt["Historias"],
                        marker_color=[
                        hex_to_rgba(COLOR, 0.5),
                        hex_to_rgba(COLOR, 0.7),
                        "#70AD47"
                        ],
                        text=ep_cnt["Historias"].astype(str), textposition="outside",
                    ))
                    fig.update_layout(title="Historias MVP por Épica", template="plotly_white",
                                      paper_bgcolor="#F7F9FC", plot_bgcolor="#F7F9FC", height=400,
                                      xaxis_tickangle=-30)
                    st.plotly_chart(fig, use_container_width=True)

            with tab3:
                if "Prioridad" in df_mvp.columns:
                    prio_cnt = df_mvp["Prioridad"].value_counts()
                    prio_colors = {"High": "#C00000", "Medium": "#ED7D31", "Low": "#A9A9A9"}
                    fig = go.Figure(go.Pie(
                        labels=prio_cnt.index, values=prio_cnt.values,
                        marker_colors=[prio_colors.get(p, "#888") for p in prio_cnt.index],
                        hole=0.4,
                    ))
                    fig.update_layout(title="Distribución de Prioridad — MVP",
                                      template="plotly_white", paper_bgcolor="#F7F9FC", height=380)
                    st.plotly_chart(fig, use_container_width=True)

            # Filtros y tabla
            section("Tabla del MVP")
            fc1, fc2 = st.columns(2)
            all_ep  = ["Todas"] + sorted(df_mvp["Épica"].dropna().unique().tolist()) if "Épica" in df_mvp.columns else ["Todas"]
            all_pr  = ["Todas"] + sorted(df_mvp["Prioridad"].dropna().unique().tolist()) if "Prioridad" in df_mvp.columns else ["Todas"]
            with fc1:
                sel_ep = st.selectbox("Épica", all_ep, key="mvp_epica")
            with fc2:
                sel_pr = st.selectbox("Prioridad", all_pr, key="mvp_prio")

            df_f = df_mvp.copy()
            if sel_ep != "Todas" and "Épica" in df_f.columns:
                df_f = df_f[df_f["Épica"] == sel_ep]
            if sel_pr != "Todas" and "Prioridad" in df_f.columns:
                df_f = df_f[df_f["Prioridad"] == sel_pr]

            with st.expander(f"Ver tabla ({len(df_f)} historias)"):
                st.dataframe(df_f, use_container_width=True, hide_index=True)