"""Sistema de Alerta Temprana de Deserción en Educación Superior.

Aplicación Streamlit que carga el modelo XGBoost serializado y expone:
  * Predicción individual mediante formulario.
  * Carga masiva por CSV.
  * Insights 360 a nivel de cohorte.

Ejecutar: streamlit run app.py
"""
from __future__ import annotations

import io
from datetime import datetime

import numpy as np  # noqa: F401  (reservado para futuras métricas)
import pandas as pd
import plotly.express as px
import streamlit as st

from src.inference import (
    SchemaError,
    build_dataframe,
    expected_columns,
    predict,
    validate_dataframe,
)
from src.insights import cohort_summary, recommendation_text, top_drivers
from src.schema import (
    BIN_COLS,
    CATEGORIES,
    FORM_GROUPS,
    LABELS,
    NUM_COLS,
    ORD_COLS,
    RAW_COLS,
    RISK_COLORS,
)

# ---------------------------------------------------------------------------
# Configuración de página
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="SAT-DE · Alerta Temprana de Deserción",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Paleta corporativa
BRAND = {
    "primary": "#1d4ed8",      # azul corporativo
    "primary_dark": "#1e3a8a",
    "ink": "#0f172a",          # slate-900
    "ink_soft": "#334155",     # slate-700
    "muted": "#64748b",        # slate-500
    "line": "#e2e8f0",         # slate-200
    "bg": "#ffffff",
    "bg_soft": "#f8fafc",
    "accent": "#0ea5e9",
}

# Paleta cualitativa para gráficos
QUAL_PALETTE = [
    "#1d4ed8", "#0ea5e9", "#14b8a6", "#f59e0b",
    "#a855f7", "#ef4444", "#22c55e", "#64748b",
]

CUSTOM_CSS = f"""
<style>
  @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

  html, body, [class*="st-"], .stMarkdown, .stTextInput, .stSelectbox, .stNumberInput {{
    font-family: 'Inter', 'Segoe UI', system-ui, -apple-system, sans-serif !important;
  }}

  /* Layout principal */
  .main .block-container {{
    padding-top: 1.4rem; padding-bottom: 3rem; max-width: 1320px;
  }}
  h1 {{ font-weight: 700; letter-spacing: -0.02em; color: {BRAND['ink']}; }}
  h2 {{ font-weight: 700; letter-spacing: -0.015em; color: {BRAND['ink']}; }}
  h3, h4 {{ font-weight: 600; letter-spacing: -0.01em; color: {BRAND['ink']}; }}
  p, label, span {{ color: {BRAND['ink_soft']}; }}

  /* Cabecera tipo "app bar" */
  .app-header {{
    display: flex; align-items: center; justify-content: space-between;
    padding: 14px 22px; margin-bottom: 22px;
    background: linear-gradient(95deg, {BRAND['primary_dark']} 0%, {BRAND['primary']} 60%, {BRAND['accent']} 100%);
    color: #ffffff; border-radius: 16px;
    box-shadow: 0 8px 24px -12px rgba(29,78,216,0.45);
  }}
  .app-header .brand {{
    display: flex; align-items: center; gap: 14px;
  }}
  .app-header .logo {{
    width: 40px; height: 40px; border-radius: 10px;
    background: rgba(255,255,255,0.18);
    display: flex; align-items: center; justify-content: center;
    font-size: 1.3rem; backdrop-filter: blur(4px);
  }}
  .app-header .title {{
    font-weight: 700; font-size: 1.05rem; color: #fff; letter-spacing: 0.01em;
  }}
  .app-header .subtitle {{
    font-size: 0.8rem; color: rgba(255,255,255,0.78); letter-spacing: 0.04em;
    text-transform: uppercase;
  }}
  .app-header .meta {{
    font-size: 0.78rem; color: rgba(255,255,255,0.85);
    background: rgba(255,255,255,0.12);
    padding: 6px 12px; border-radius: 999px;
    border: 1px solid rgba(255,255,255,0.18);
  }}

  /* Hero por página */
  .page-hero {{
    border: 1px solid {BRAND['line']}; border-radius: 16px;
    background: linear-gradient(180deg,#ffffff 0%, {BRAND['bg_soft']} 100%);
    padding: 22px 26px; margin-bottom: 22px;
  }}
  .page-hero .eyebrow {{
    color: {BRAND['primary']}; font-weight: 600; font-size: 0.78rem;
    letter-spacing: 0.12em; text-transform: uppercase; margin-bottom: 6px;
  }}
  .page-hero h1 {{ margin: 0 0 6px 0; font-size: 1.7rem; }}
  .page-hero p {{ margin: 0; color: {BRAND['muted']}; font-size: 0.95rem; }}

  /* Risk badge */
  .risk-badge {{
    display: inline-flex; align-items: center; gap: 8px;
    padding: 8px 16px; border-radius: 999px;
    color: white; font-weight: 600; font-size: 0.92rem; letter-spacing: 0.02em;
    box-shadow: 0 4px 14px -6px rgba(15,23,42,0.25);
  }}
  .risk-badge::before {{
    content: ""; width: 8px; height: 8px; border-radius: 999px;
    background: rgba(255,255,255,0.9);
  }}

  /* KPI cards */
  .kpi-card {{
    position: relative; overflow: hidden;
    background: #ffffff;
    border: 1px solid {BRAND['line']}; border-radius: 14px;
    padding: 18px 20px;
    box-shadow: 0 1px 2px rgba(15,23,42,0.04);
    transition: transform .15s ease, box-shadow .15s ease;
  }}
  .kpi-card:hover {{
    transform: translateY(-1px);
    box-shadow: 0 8px 24px -16px rgba(15,23,42,0.18);
  }}
  .kpi-card::before {{
    content: ""; position: absolute; left: 0; top: 0; bottom: 0; width: 4px;
    background: {BRAND['primary']};
  }}
  .kpi-card.alto::before    {{ background: {RISK_COLORS['Alto']}; }}
  .kpi-card.medio::before   {{ background: {RISK_COLORS['Medio']}; }}
  .kpi-card.bajo::before    {{ background: {RISK_COLORS['Bajo']}; }}
  .kpi-card .label {{
    font-size: 0.72rem; color: {BRAND['muted']};
    text-transform: uppercase; letter-spacing: 0.1em; font-weight: 600;
  }}
  .kpi-card .value {{
    font-size: 2rem; font-weight: 700; color: {BRAND['ink']}; line-height: 1.1;
    margin-top: 4px;
  }}
  .kpi-card .sub {{
    font-size: 0.78rem; color: {BRAND['muted']}; margin-top: 2px;
  }}

  /* Form groups */
  .form-group-title {{
    display: flex; align-items: center; gap: 10px;
    font-size: 0.85rem; font-weight: 600; color: {BRAND['primary']};
    text-transform: uppercase; letter-spacing: 0.1em;
    margin: 18px 0 4px 0; padding-bottom: 6px;
    border-bottom: 1px solid {BRAND['line']};
  }}
  .form-group-title::before {{
    content: ""; width: 6px; height: 18px; border-radius: 3px;
    background: {BRAND['primary']};
  }}

  /* Tabs corporativos */
  .stTabs [data-baseweb="tab-list"] {{
    gap: 4px; border-bottom: 1px solid {BRAND['line']}; padding-bottom: 0;
  }}
  .stTabs [data-baseweb="tab"] {{
    padding: 10px 18px; border-radius: 10px 10px 0 0;
    background: transparent; color: {BRAND['muted']}; font-weight: 500;
  }}
  .stTabs [aria-selected="true"] {{
    background: {BRAND['bg_soft']}; color: {BRAND['primary']}; font-weight: 600;
  }}

  /* Botones primarios */
  .stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {{
    background: {BRAND['primary']}; border: none; border-radius: 10px;
    font-weight: 600; padding: 10px 20px;
    box-shadow: 0 6px 18px -10px rgba(29,78,216,0.55);
  }}
  .stButton > button[kind="primary"]:hover, .stDownloadButton > button[kind="primary"]:hover {{
    background: {BRAND['primary_dark']};
  }}
  .stButton > button, .stDownloadButton > button {{ border-radius: 10px; }}

  /* Inputs */
  .stTextInput input, .stNumberInput input, .stSelectbox > div > div {{
    border-radius: 10px !important;
  }}

  /* Sidebar */
  section[data-testid="stSidebar"] {{
    background: linear-gradient(180deg,#0b1220 0%, #0f172a 100%);
    border-right: 1px solid #1e293b;
  }}
  section[data-testid="stSidebar"] * {{ color: #e2e8f0 !important; }}
  section[data-testid="stSidebar"] h1,
  section[data-testid="stSidebar"] h2,
  section[data-testid="stSidebar"] h3 {{ color: #f8fafc !important; }}
  .sidebar-brand {{
    display: flex; align-items: center; gap: 10px; margin: 4px 0 14px 0;
  }}
  .sidebar-brand .logo {{
    width: 36px; height: 36px; border-radius: 9px;
    background: linear-gradient(135deg, {BRAND['primary']} 0%, {BRAND['accent']} 100%);
    display: flex; align-items: center; justify-content: center; font-size: 1.1rem;
  }}
  .sidebar-brand .name {{ font-weight: 700; font-size: 1rem; }}
  .sidebar-brand .tag {{ font-size: 0.72rem; color: #94a3b8 !important; }}
  section[data-testid="stSidebar"] hr {{ border-color: #1e293b !important; }}

  /* DataFrames */
  .stDataFrame {{ border-radius: 10px; overflow: hidden; }}

  /* Caption fina */
  .footer-note {{
    color: {BRAND['muted']}; font-size: 0.78rem; text-align: center;
    margin-top: 28px; padding-top: 14px; border-top: 1px solid {BRAND['line']};
  }}

  /* Esconder marca por defecto de Streamlit */
  #MainMenu, footer {{visibility: hidden;}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ---------------------------------------------------------------------------
# Plotly template corporativo
# ---------------------------------------------------------------------------
def _apply_chart_style(fig, height: int = 360, showlegend: bool = True):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=40, b=10),
        font=dict(family="Inter, Segoe UI, sans-serif", size=12, color=BRAND["ink_soft"]),
        title=dict(font=dict(size=15, color=BRAND["ink"]), x=0.0, xanchor="left"),
        paper_bgcolor="white",
        plot_bgcolor="white",
        showlegend=showlegend,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, font=dict(size=11)),
        xaxis=dict(showgrid=False, linecolor=BRAND["line"], tickcolor=BRAND["line"]),
        yaxis=dict(showgrid=True, gridcolor=BRAND["line"], zeroline=False),
    )
    return fig


# ---------------------------------------------------------------------------
# Helpers UI
# ---------------------------------------------------------------------------
def risk_badge(label: str) -> str:
    color = RISK_COLORS.get(label, BRAND["muted"])
    return f'<span class="risk-badge" style="background:{color}">Riesgo {label}</span>'


def kpi_card(label: str, value: str, sub: str = "", variant: str = "") -> str:
    cls = f"kpi-card {variant}".strip()
    sub_html = f'<div class="sub">{sub}</div>' if sub else ""
    return (
        f'<div class="{cls}">'
        f'<div class="label">{label}</div>'
        f'<div class="value">{value}</div>{sub_html}'
        f'</div>'
    )


def page_hero(eyebrow: str, title: str, subtitle: str) -> None:
    st.markdown(
        f'<div class="page-hero">'
        f'<div class="eyebrow">{eyebrow}</div>'
        f'<h1>{title}</h1><p>{subtitle}</p>'
        f'</div>',
        unsafe_allow_html=True,
    )


def app_header() -> None:
    today = datetime.now().strftime("%d %b %Y")
    st.markdown(
        f'<div class="app-header">'
        f'<div class="brand">'
        f'<div class="logo">🎓</div>'
        f'<div>'
        f'<div class="subtitle">MIAD · Universidad de los Andes</div>'
        f'<div class="title">Sistema de Alerta Temprana de Deserción</div>'
        f'</div>'
        f'</div>'
        f'<div class="meta">Modelo XGBoost v1.1.1 · {today}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )


def section_title(group: str) -> None:
    st.markdown(f'<div class="form-group-title">{group}</div>', unsafe_allow_html=True)


def render_individual_form() -> dict | None:
    """Renderiza el formulario agrupado y devuelve un dict con los valores."""
    record: dict = {}
    current_year = datetime.now().year
    with st.form("formulario_estudiante", clear_on_submit=False):
        for group, fields in FORM_GROUPS.items():
            section_title(group)
            cols = st.columns(3)
            for i, field in enumerate(fields):
                col = cols[i % 3]
                label = LABELS.get(field, field)
                if field in CATEGORIES:
                    record[field] = col.selectbox(label, CATEGORIES[field], key=field)
                elif field in BIN_COLS:
                    record[field] = 1 if col.radio(label, ["No", "Sí"], horizontal=True, key=field) == "Sí" else 0
                elif field == "estrato_socioeconomico":
                    record[field] = col.selectbox(label, [1, 2, 3, 4, 5, 6], index=2, key=field)
                elif field == "semestre_cursado":
                    record[field] = col.number_input(label, min_value=1, max_value=20, value=4, step=1, key=field)
                elif field == "edad_ingreso":
                    record[field] = col.number_input(label, min_value=14, max_value=60, value=18, step=1, key=field)
                elif field == "anio_registro":
                    record[field] = col.number_input(label, min_value=2000, max_value=current_year + 1, value=current_year, step=1, key=field)
                elif field == "promedio_academico":
                    record[field] = col.number_input(label, min_value=0.0, max_value=5.0, value=3.5, step=0.05, format="%.2f", key=field)
                elif field == "materias_reprobadas":
                    record[field] = col.number_input(label, min_value=0, max_value=40, value=0, step=1, key=field)
                elif field == "distancia_hogar_ies_km":
                    record[field] = col.number_input(label, min_value=0.0, max_value=500.0, value=5.0, step=0.5, key=field)
                elif field == "puntaje_icfes_percentil":
                    record[field] = col.slider(label, 0, 100, 60, key=field)
                else:
                    record[field] = col.text_input(label, key=field)
            st.markdown("")
        submitted = st.form_submit_button("Calcular riesgo", type="primary", use_container_width=True)
    return record if submitted else None


def page_inicio() -> None:
    page_hero(
        "Inicio",
        "Plataforma de Alerta Temprana de Deserción",
        "Solución analítica para consejería académica que cuantifica el riesgo de "
        "deserción y prioriza intervenciones a partir de variables académicas, "
        "demográficas y socioeconómicas.",
    )
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown(kpi_card("Predicción individual", "1", "Evaluación caso a caso"), unsafe_allow_html=True)
        st.markdown("Ingrese el perfil de un estudiante y obtenga su probabilidad de deserción junto con los factores explicativos principales.")
    with c2:
        st.markdown(kpi_card("Carga masiva", "CSV", "Procesamiento por lote"), unsafe_allow_html=True)
        st.markdown("Cargue un archivo con varios estudiantes para clasificarlos por nivel de riesgo y descargar los resultados.")
    with c3:
        st.markdown(kpi_card("Insights 360", "📊", "Visión de cohorte"), unsafe_allow_html=True)
        st.markdown("Explore el riesgo agregado por estrato, sector, área de conocimiento y desempeño académico.")

    st.markdown("&nbsp;")
    st.markdown(
        "<div class='page-hero' style='margin-top:8px'>"
        "<div class='eyebrow'>Cómo se interpreta</div>"
        "<p style='margin-top:8px'>"
        f"<b style='color:{RISK_COLORS['Bajo']}'>● Bajo</b> &lt; 30 % &nbsp;&nbsp; "
        f"<b style='color:{RISK_COLORS['Medio']}'>● Medio</b> 30 – 60 % &nbsp;&nbsp; "
        f"<b style='color:{RISK_COLORS['Alto']}'>● Alto</b> ≥ 60 %"
        "</p></div>",
        unsafe_allow_html=True,
    )


def page_individual() -> None:
    page_hero(
        "Predicción individual",
        "Evaluación de riesgo por estudiante",
        "Complete el perfil del estudiante. El modelo devolverá la probabilidad "
        "estimada de deserción y los factores más influyentes.",
    )
    record = render_individual_form()
    if record is None:
        return
    df = build_dataframe([record])
    try:
        preds = predict(df)
    except Exception as e:
        st.error(f"No fue posible calcular la predicción: {e}")
        return
    prob = float(preds.loc[0, "prob_desercion"])
    riesgo = preds.loc[0, "riesgo"]
    variant = riesgo.lower()

    st.markdown("&nbsp;")
    c1, c2 = st.columns([1, 2])
    with c1:
        st.markdown(risk_badge(riesgo), unsafe_allow_html=True)
        st.markdown("&nbsp;")
        st.markdown(
            kpi_card("Probabilidad de deserción", f"{prob*100:,.1f}%", "Probabilidad estimada", variant),
            unsafe_allow_html=True,
        )
        st.markdown(
            kpi_card("Decisión sugerida", "Intervención" if prob >= 0.6 else ("Seguimiento" if prob >= 0.3 else "Permanencia"),
                     "Según umbrales institucionales", variant),
            unsafe_allow_html=True,
        )
    with c2:
        fig = px.bar(
            x=[prob, 1 - prob], y=["Deserción", "Permanencia"], orientation="h",
            color=["Deserción", "Permanencia"],
            color_discrete_map={"Deserción": RISK_COLORS[riesgo], "Permanencia": "#e2e8f0"},
            range_x=[0, 1],
        )
        fig.update_traces(textposition="inside", texttemplate="%{x:.0%}")
        _apply_chart_style(fig, height=210, showlegend=False)
        fig.update_layout(xaxis_tickformat=".0%", yaxis_title="", xaxis_title="", title="Composición de la predicción")
        st.plotly_chart(fig, use_container_width=True)

        drivers = top_drivers(record, k=5)
        if drivers.empty:
            st.info("No fue posible computar los factores.")
        else:
            fig2 = px.bar(
                drivers.iloc[::-1], x="contribucion_rel", y="factor", orientation="h",
                color="direccion",
                color_discrete_map={"Riesgo": RISK_COLORS["Alto"], "Protección": RISK_COLORS["Bajo"]},
                text=drivers.iloc[::-1]["contribucion_rel"].map(lambda v: f"{v*100:.0f}%"),
            )
            fig2.update_traces(textposition="outside")
            _apply_chart_style(fig2, height=320)
            fig2.update_layout(title="Factores más influyentes",
                               xaxis_title="Contribución relativa", yaxis_title="",
                               legend_title_text="Dirección")
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("##### Recomendación al consejero académico")
    st.info(recommendation_text(prob, drivers))


def page_carga_masiva() -> None:
    page_hero(
        "Carga masiva",
        "Procesamiento por lote desde CSV",
        "Cargue un archivo con una fila por estudiante. La aplicación valida el "
        "esquema, clasifica el riesgo y permite descargar el resultado.",
    )
    with st.expander("📑 Ver columnas esperadas / descargar plantilla", expanded=False):
        st.code(", ".join(expected_columns()), language="text")
        plantilla = pd.DataFrame(columns=expected_columns())
        st.download_button(
            "Descargar plantilla CSV vacía",
            plantilla.to_csv(index=False).encode("utf-8"),
            file_name="plantilla_estudiantes.csv",
            mime="text/csv",
        )

    file = st.file_uploader("Archivo CSV", type=["csv"])
    if not file:
        return
    try:
        df = pd.read_csv(file)
    except Exception as e:
        st.error(f"No fue posible leer el archivo: {e}")
        return
    try:
        warns = validate_dataframe(df)
    except SchemaError as e:
        st.error(str(e))
        return
    for w in warns:
        st.warning(w)

    df_in = build_dataframe(df.to_dict(orient="records"))
    try:
        preds = predict(df_in)
    except Exception as e:
        st.error(f"Error al inferir: {e}")
        return

    st.session_state["cohort_raw"] = df_in
    st.session_state["cohort_preds"] = preds

    st.success(f"Se procesaron **{len(df_in):,}** estudiantes.")
    summary = cohort_summary(df_in, preds)
    k = summary["kpis"]
    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(kpi_card("Total estudiantes", f"{k['total']:,}"), unsafe_allow_html=True)
    c2.markdown(kpi_card("Riesgo alto", f"{k['alto']:,}", "Casos prioritarios", "alto"), unsafe_allow_html=True)
    c3.markdown(kpi_card("% riesgo alto", f"{k['pct_alto']*100:.1f}%", "Proporción de la cohorte", "alto"), unsafe_allow_html=True)
    c4.markdown(kpi_card("Probabilidad media", f"{k['prom_riesgo']*100:.1f}%", "Promedio cohorte"), unsafe_allow_html=True)

    st.markdown("##### Resultados")
    filtro = st.multiselect("Filtrar por nivel de riesgo", ["Alto", "Medio", "Bajo"], default=["Alto", "Medio", "Bajo"])
    out = pd.concat([df_in.reset_index(drop=True), preds.reset_index(drop=True)], axis=1)
    out_view = out[out["riesgo"].isin(filtro)].copy()
    out_view["prob_desercion"] = (out_view["prob_desercion"] * 100).round(2)
    st.dataframe(out_view, use_container_width=True, hide_index=True)

    buf = io.BytesIO()
    out.to_csv(buf, index=False)
    st.download_button(
        "⬇️ Descargar resultados (CSV)",
        buf.getvalue(),
        file_name="predicciones_desercion.csv",
        mime="text/csv",
        type="primary",
    )
    st.info("Continúe a **Insights 360** para explorar el detalle agregado de la cohorte.")


def page_insights() -> None:
    page_hero(
        "Insights 360",
        "Visión integral de la cohorte cargada",
        "Distribución del riesgo, cortes por dimensiones demográficas, "
        "socioeconómicas y académicas, y vista detallada por estudiante.",
    )
    raw = st.session_state.get("cohort_raw")
    preds = st.session_state.get("cohort_preds")
    if raw is None or preds is None:
        st.warning("Primero cargue un archivo en la sección **Carga masiva**.")
        return
    summary = cohort_summary(raw, preds)
    k = summary["kpis"]
    df = summary["df"]

    c1, c2, c3, c4 = st.columns(4)
    c1.markdown(kpi_card("Estudiantes", f"{k['total']:,}"), unsafe_allow_html=True)
    c2.markdown(kpi_card("Riesgo alto", f"{k['alto']:,}", f"{k['pct_alto']*100:.1f}% del total", "alto"), unsafe_allow_html=True)
    c3.markdown(kpi_card("Riesgo medio", f"{k['medio']:,}", variant="medio"), unsafe_allow_html=True)
    c4.markdown(kpi_card("Riesgo bajo", f"{k['bajo']:,}", variant="bajo"), unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs(["📈 Distribución", "👥 Demográfico / Socioeconómico", "🎓 Académico", "🔎 Detalle"])

    with tab1:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.histogram(df, x="prob_desercion", nbins=30, color="riesgo",
                               color_discrete_map=RISK_COLORS,
                               labels={"prob_desercion": "Probabilidad de deserción"},
                               title="Distribución de la probabilidad")
            _apply_chart_style(fig, height=380)
            fig.update_layout(bargap=0.05, legend_title_text="Riesgo")
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            risk_counts = df["riesgo"].value_counts().reindex(["Alto", "Medio", "Bajo"]).fillna(0).reset_index()
            risk_counts.columns = ["riesgo", "n"]
            fig = px.pie(risk_counts, names="riesgo", values="n", color="riesgo",
                         color_discrete_map=RISK_COLORS, hole=0.6,
                         title="Composición por nivel de riesgo")
            fig.update_traces(textinfo="percent+label")
            _apply_chart_style(fig, height=380)
            st.plotly_chart(fig, use_container_width=True)

    with tab2:
        c1, c2 = st.columns(2)
        with c1:
            fig = px.bar(summary["by_estrato"], x="estrato_socioeconomico", y="prob_desercion",
                         labels={"prob_desercion": "Prob. media", "estrato_socioeconomico": "Estrato"},
                         title="Riesgo medio por estrato", color="prob_desercion",
                         color_continuous_scale=[(0, "#dbeafe"), (1, BRAND["primary_dark"])])
            _apply_chart_style(fig, height=380, showlegend=False)
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        with c2:
            be = summary["by_beca"].copy()
            be["beneficiario_beca"] = be["beneficiario_beca"].map({0: "Sin beca", 1: "Con beca"}).fillna("Desconocido")
            fig = px.bar(be, x="beneficiario_beca", y="prob_desercion",
                         labels={"prob_desercion": "Prob. media", "beneficiario_beca": ""},
                         title="Riesgo medio según beca", color="prob_desercion",
                         color_continuous_scale=[(0, "#dbeafe"), (1, BRAND["primary_dark"])])
            _apply_chart_style(fig, height=380, showlegend=False)
            fig.update_layout(coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)
        fig = px.box(df, x="zona_residencia", y="prob_desercion", color="zona_residencia",
                     color_discrete_sequence=QUAL_PALETTE,
                     title="Distribución de riesgo por zona de residencia", points=False)
        _apply_chart_style(fig, height=360, showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

    with tab3:
        fig = px.bar(summary["by_area"], x="mean", y="area_conocimiento", orientation="h",
                     labels={"mean": "Prob. media", "area_conocimiento": ""},
                     title="Riesgo medio por área de conocimiento", color="mean",
                     color_continuous_scale=[(0, "#dbeafe"), (1, BRAND["primary_dark"])])
        _apply_chart_style(fig, height=420, showlegend=False)
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)
        fig = px.scatter(df, x="promedio_academico", y="materias_reprobadas", color="riesgo",
                         color_discrete_map=RISK_COLORS, opacity=0.75,
                         hover_data=["prob_desercion", "semestre_cursado"],
                         labels={"promedio_academico": "Promedio académico",
                                 "materias_reprobadas": "Materias reprobadas"},
                         title="Promedio académico vs. materias reprobadas")
        _apply_chart_style(fig, height=440)
        st.plotly_chart(fig, use_container_width=True)

    with tab4:
        st.dataframe(
            df.sort_values("prob_desercion", ascending=False).assign(
                prob_desercion=lambda d: (d["prob_desercion"] * 100).round(2)
            ),
            use_container_width=True, hide_index=True,
        )


def page_ayuda() -> None:
    page_hero(
        "Ayuda",
        "Diccionario de datos y guía rápida",
        "Definición de las variables del modelo, valores admitidos y umbrales de "
        "decisión utilizados por la plataforma.",
    )
    st.markdown("##### Diccionario de variables")
    rows = []
    for c in RAW_COLS:
        tipo = "Categórica" if c in CATEGORIES else ("Binaria" if c in BIN_COLS else ("Ordinal" if c in ORD_COLS else "Numérica"))
        valores = ", ".join(map(str, CATEGORIES[c])) if c in CATEGORIES else ""
        rows.append({"Variable": c, "Etiqueta": LABELS.get(c, c), "Tipo": tipo, "Valores": valores})
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.markdown("##### Umbrales de riesgo")
    c1, c2, c3 = st.columns(3)
    c1.markdown(kpi_card("Bajo", "< 30 %", "Seguimiento estándar", "bajo"), unsafe_allow_html=True)
    c2.markdown(kpi_card("Medio", "30 – 60 %", "Seguimiento focalizado", "medio"), unsafe_allow_html=True)
    c3.markdown(kpi_card("Alto", "≥ 60 %", "Intervención prioritaria", "alto"), unsafe_allow_html=True)

    st.markdown(
        "<div class='footer-note'>Equipo Grupo 10 · MIAD · Universidad de los Andes · "
        "Modelo XGBoost v1.1.1</div>",
        unsafe_allow_html=True,
    )


# ---------------------------------------------------------------------------
# Navegación
# ---------------------------------------------------------------------------
PAGES = {
    "🏠  Inicio": page_inicio,
    "👤  Predicción individual": page_individual,
    "📥  Carga masiva": page_carga_masiva,
    "📊  Insights 360": page_insights,
    "ℹ️  Ayuda": page_ayuda,
}

with st.sidebar:
    st.markdown(
        '<div class="sidebar-brand">'
        '<div class="logo">🎓</div>'
        '<div><div class="name">SAT-DE</div>'
        '<div class="tag">Alerta Temprana</div></div>'
        '</div>',
        unsafe_allow_html=True,
    )
    st.markdown("---")
    choice = st.radio("Navegación", list(PAGES.keys()), label_visibility="collapsed")
    st.markdown("---")
    st.caption("Modelo XGBoost v1.1.1")
    st.caption("MIAD · Grupo 10 · 2026")

app_header()
PAGES[choice]()
st.markdown(
    "<div class='footer-note'>© 2026 MIAD · Universidad de los Andes — Proyecto Aplicado en Analítica de Datos</div>",
    unsafe_allow_html=True,
)
