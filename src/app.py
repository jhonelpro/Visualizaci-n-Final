"""
Residencias Definitivas Resueltas — Chile 2000–2025
Dashboard interactivo de visualización de datos migratorios del SERMIG.

Estudiantes : Raul Hidalgo - Jhon Vallecilla
Docente     : Brian Keith Norambuena
Asignatura  : Visualización de Datos
Entrega 2   : Visualización Final
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np

# ══════════════════════════════════════════════════════════
#  CONFIGURACIÓN DE PÁGINA
# ══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Residencias Definitivas — Chile 2000–2025",
    page_icon="🇨🇱",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════════════════════
#  PALETA GLOBAL (consistente en las 3 vistas)
# ══════════════════════════════════════════════════════════
COLOR_MAP = {
    "Otorga":               "#4dac26",
    "Rechaza con Rt":       "#d01c8b",
    "Archiva":              "#f1a340",
    "Rechaza con abandono": "#998ec3",
}
TIPO_ORDER = ["Otorga", "Archiva", "Rechaza con Rt", "Rechaza con abandono"]

# ══════════════════════════════════════════════════════════
#  CSS PERSONALIZADO
# ══════════════════════════════════════════════════════════
st.markdown("""
<style>
    /* contenedor principal */
    .block-container { padding-top: 1rem; }

    /* leyenda superior */
    .legend-bar {
        display: flex; gap: 22px; align-items: center;
        background: #1e2130; border-radius: 8px;
        padding: 10px 20px; margin-bottom: 14px;
        font-size: .84rem; flex-wrap: wrap;
    }
    .legend-dot {
        width: 13px; height: 13px; border-radius: 50%;
        display: inline-block; margin-right: 5px; vertical-align: middle;
    }

    /* tarjetas de métricas */
    .metric-card {
        background: #1e2130; border-radius: 10px;
        padding: 14px 10px; text-align: center;
    }
    .metric-val { font-size: 1.55rem; font-weight: 700; color: #e8eaf6; }
    .metric-lbl { font-size: .73rem; color: #9e9e9e; margin-top: 3px; }

    /* notas pie de vista */
    .nota {
        background: #1e2130; border-left: 4px solid #4dac26;
        border-radius: 6px; padding: 9px 14px;
        font-size: .77rem; color: #bdbdbd; margin-top: 8px;
    }

    /* títulos de sección */
    .vista-title { font-size: 1.06rem; font-weight: 600;
                   color: #e8eaf6; margin-bottom: 3px; }
    .vista-sub   { font-size: .78rem; color: #9e9e9e; margin-bottom: 10px; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
#  CARGA Y PREPROCESAMIENTO
# ══════════════════════════════════════════════════════════
# Ruta al dataset relativa al propio archivo .py (funciona desde cualquier directorio)
DATA_PATH = "src/data/data.xlsx"

@st.cache_data(show_spinner="Cargando dataset…")
def load_data() -> pd.DataFrame:
    """Lee el Excel, descarta filas con Total='*' (privacidad < 6 casos)
    y convierte los tipos numéricos necesarios."""
    df = pd.read_excel(DATA_PATH, dtype=str)
    df = df[df["Total"] != "*"].copy()
    df["Total"] = pd.to_numeric(df["Total"])
    df["AÑO"]   = pd.to_numeric(df["AÑO"])
    df["TIPO_RESUELTO"] = df["TIPO_RESUELTO"].str.strip()
    return df

df_raw = load_data()

# Categorías a excluir por defecto (privacidad / sin datos)
EXCLUIR_REGION  = {"Sin Información", "Anonimizada"}
EXCLUIR_ETARIO  = {"Sin info", "Anonimizado"}
EXCLUIR_ESTUDIOS = {"No informa", "Otros estudios"}
EXCLUIR_ACT     = {"No informa", "Otras actividades", "Tripulante"}

# ══════════════════════════════════════════════════════════
#  COORDENADAS CENTROIDES POR REGIÓN (para mapa de burbujas)
# ══════════════════════════════════════════════════════════
REGION_COORDS: dict[str, tuple[float, float]] = {
    "Arica y Parinacota":                            (-18.48, -70.32),
    "Tarapacá":                                      (-20.21, -69.30),
    "Antofagasta":                                   (-23.65, -70.40),
    "Atacama":                                       (-27.37, -70.33),
    "Coquimbo":                                      (-29.95, -71.34),
    "Valparaíso":                                    (-33.04, -71.62),
    "Metropolitana de Santiago":                     (-33.46, -70.65),
    "Libertador General Bernardo O'Higgins":         (-34.58, -71.00),
    "Maule":                                         (-35.43, -71.65),
    "Ñuble":                                         (-36.72, -71.76),
    "Biobío":                                        (-37.47, -72.35),
    "La Araucanía":                                  (-38.95, -72.66),
    "Los Ríos":                                      (-39.82, -73.24),
    "Los Lagos":                                     (-41.47, -72.94),
    "Aysén del General Carlos Ibáñez del Campo":     (-45.57, -71.69),
    "Magallanes y de la Antártica Chilena":          (-53.16, -70.91),
}

# ══════════════════════════════════════════════════════════
#  BARRA LATERAL — FILTROS GLOBALES ENCADENADOS
# ══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## 🔎 Filtros Globales")

    # Toggle para incluir categorías sin información
    incluir_nulos = st.toggle(
        'Incluir "Sin Información" / "Anonimizado/a"',
        value=False,
        help="Activa para incluir filas con datos faltantes o anonimizados",
    )

    # Listas disponibles (limpias por defecto)
    paises_disp = sorted(df_raw["PAÍS"].unique())
    regiones_disp = sorted(
        r for r in df_raw["REGIÓN"].unique() if r not in EXCLUIR_REGION
    )
    sexos_disp = sorted(df_raw["SEXO"].unique())
    etarios_disp = [e for e in
                    ["17 ó menos","18-29","30-44","45-59","60-74","75 ó más"]
                    if e in df_raw["RANGO_ETARIO"].unique()]
    estudios_disp = [e for e in
                     ["Básico","Medio","Técnico","Universitario"]
                     if e in df_raw["ESTUDIOS"].unique()]

    # Selectores encadenados
    sel_paises   = st.multiselect("Nacionalidad (País)", paises_disp,
                                  placeholder="Todos los países")
    sel_regiones = st.multiselect("Región", regiones_disp,
                                  placeholder="Todas las regiones")
    año_rango    = st.slider("Rango de Año", 2000, 2025, (2000, 2025))
    sel_sexo     = st.multiselect("Sexo", sexos_disp,
                                  placeholder="Ambos sexos")
    sel_etario   = st.multiselect("Rango Etario", etarios_disp,
                                  placeholder="Todos los rangos")
    sel_estudios = st.multiselect("Nivel de Estudios", estudios_disp,
                                  placeholder="Todos los niveles")

    st.markdown("---")
    st.markdown(
        "<div style='font-size:.72rem;color:#9e9e9e;line-height:1.5;'>"
        "<b>Fuente:</b> Servicio Nacional de Migraciones (SERMIG), 2026.<br>"
        "<b>Dataset:</b> <a href='https://serviciomigraciones.cl/wp-content/uploads/estudios/Datos-abiertos/RD/RD-Resueltas 2o-semestre-2025.xlsx' "
        "target='_blank' style='color:#4dac26;'>RD-Resueltas-2o-semestre-2025.xlsx</a><br>"
        "<b>Registros válidos:</b> 1.406.334 actos administrativos"
        "</div>",
        unsafe_allow_html=True,
    )

# ══════════════════════════════════════════════════════════
#  APLICAR FILTROS AL DATAFRAME
# ══════════════════════════════════════════════════════════
df = df_raw.copy()

# Excluir categorías sin información (a menos que el toggle esté activo)
if not incluir_nulos:
    df = df[~df["REGIÓN"].isin(EXCLUIR_REGION)]
    df = df[~df["RANGO_ETARIO"].isin(EXCLUIR_ETARIO)]
    df = df[~df["ESTUDIOS"].isin(EXCLUIR_ESTUDIOS)]
    df = df[~df["ACTIVIDAD"].isin(EXCLUIR_ACT)]

# Filtros de usuario
if sel_paises:
    df = df[df["PAÍS"].isin(sel_paises)]
if sel_regiones:
    df = df[df["REGIÓN"].isin(sel_regiones)]
df = df[(df["AÑO"] >= año_rango[0]) & (df["AÑO"] <= año_rango[1])]
if sel_sexo:
    df = df[df["SEXO"].isin(sel_sexo)]
if sel_etario:
    df = df[df["RANGO_ETARIO"].isin(sel_etario)]
if sel_estudios:
    df = df[df["ESTUDIOS"].isin(sel_estudios)]

# ══════════════════════════════════════════════════════════
#  ENCABEZADO PRINCIPAL
# ══════════════════════════════════════════════════════════
st.markdown("# 🇨🇱 Residencias Definitivas Resueltas — Chile 2000–2025")

# Subtítulo dinámico según filtros activos
partes = []
if sel_paises:
    partes.append(", ".join(sel_paises[:3]) + ("…" if len(sel_paises) > 3 else ""))
if sel_regiones:
    partes.append(", ".join(sel_regiones[:2]) + ("…" if len(sel_regiones) > 2 else ""))
if año_rango != (2000, 2025):
    partes.append(f"{año_rango[0]}–{año_rango[1]}")
if sel_sexo:
    partes.append(", ".join(sel_sexo))
if sel_etario:
    partes.append(", ".join(sel_etario))
if sel_estudios:
    partes.append(", ".join(sel_estudios))
subtitulo = ("Mostrando: " + " · ".join(partes)) if partes else "Mostrando todos los registros"
st.markdown(
    f"<div style='color:#9e9e9e;font-size:.85rem;margin-bottom:12px;'>{subtitulo}</div>",
    unsafe_allow_html=True,
)

# Leyenda de colores unificada (persistente en toda la página)
st.markdown(
    '<div class="legend-bar">'
    + "".join(
        f'<span><span class="legend-dot" style="background:{COLOR_MAP[t]};"></span>{t}</span>'
        for t in TIPO_ORDER
    )
    + "</div>",
    unsafe_allow_html=True,
)

# ══════════════════════════════════════════════════════════
#  MÉTRICAS RESUMEN
# ══════════════════════════════════════════════════════════
total_regs   = int(df["Total"].sum())
n_otorga     = int(df[df["TIPO_RESUELTO"] == "Otorga"]["Total"].sum())
n_rechaza    = int(df[df["TIPO_RESUELTO"].isin(
                      ["Rechaza con Rt","Rechaza con abandono"])]["Total"].sum())
n_archiva    = int(df[df["TIPO_RESUELTO"] == "Archiva"]["Total"].sum())
tasa_otorga  = n_otorga  / total_regs * 100 if total_regs else 0
tasa_rechaza = n_rechaza / total_regs * 100 if total_regs else 0
tasa_archiva = n_archiva / total_regs * 100 if total_regs else 0

mc1, mc2, mc3, mc4 = st.columns(4)
for col, val, lbl in [
    (mc1, f"{total_regs:,.0f}", "Registros administrativos"),
    (mc2, f"{tasa_otorga:.1f}%",  "Tasa de otorgamiento"),
    (mc3, f"{tasa_rechaza:.1f}%", "Tasa de rechazo"),
    (mc4, f"{tasa_archiva:.1f}%", "Tasa de archivo"),
]:
    col.markdown(
        f'<div class="metric-card">'
        f'<div class="metric-val">{val}</div>'
        f'<div class="metric-lbl">{lbl}</div>'
        f'</div>',
        unsafe_allow_html=True,
    )

st.markdown("<br>", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════
#  VISTA 1 — MAPA DE BURBUJAS COROPLÉTICO
#  Tarea: Descubrir concentración geográfica por región
# ══════════════════════════════════════════════════════════
st.markdown(
    '<div class="vista-title">🗺️ Vista 1: Exploración Espacial — Concentración por Región</div>'
    '<div class="vista-sub">¿Qué regiones de Chile concentran la mayor demanda de '
    'Residencia Definitiva según nacionalidad? — <em>Tarea: Descubrir</em></div>',
    unsafe_allow_html=True,
)

# ── Agregación por región y tipo ──────────────────────────
df_v1_base = df[df["REGIÓN"].isin(REGION_COORDS)].copy()

df_v1_wide = (
    df_v1_base
    .groupby(["REGIÓN","TIPO_RESUELTO"])["Total"]
    .sum()
    .unstack(fill_value=0)
    .reset_index()
)
# Asegurar que existan todas las columnas aunque algún tipo no tenga datos
for t in TIPO_ORDER:
    if t not in df_v1_wide.columns:
        df_v1_wide[t] = 0

df_v1_wide["total_region"]  = df_v1_wide[TIPO_ORDER].sum(axis=1)
df_v1_wide["tasa_otorga"]   = (
    df_v1_wide["Otorga"] / df_v1_wide["total_region"].replace(0, np.nan) * 100
).round(1)
df_v1_wide["tasa_rechazo"]  = (
    (df_v1_wide["Rechaza con Rt"] + df_v1_wide["Rechaza con abandono"])
    / df_v1_wide["total_region"].replace(0, np.nan) * 100
).round(1)

# Top-3 nacionalidades por región (para tooltip)
df_top3 = (
    df_v1_base
    .groupby(["REGIÓN","PAÍS"])["Total"].sum().reset_index()
    .sort_values(["REGIÓN","Total"], ascending=[True, False])
    .groupby("REGIÓN").head(3)
    .groupby("REGIÓN")["PAÍS"]
    .apply(lambda x: " / ".join(x))
    .reset_index()
    .rename(columns={"PAÍS": "top3_paises"})
)
df_v1_wide = df_v1_wide.merge(df_top3, on="REGIÓN", how="left")

# Coordenadas
df_v1_wide["lat"] = df_v1_wide["REGIÓN"].map(lambda r: REGION_COORDS[r][0])
df_v1_wide["lon"] = df_v1_wide["REGIÓN"].map(lambda r: REGION_COORDS[r][1])

# ── Mapa ancho completo arriba, tabla y donut abajo ───────
st.markdown(
    "<div style='font-size:.78rem;color:#9e9e9e;margin-bottom:2px;'>"
    "🌍 Rueda del mouse para hacer zoom · Arrastra para mover · "
    "Hover sobre cada burbuja para ver detalle</div>",
    unsafe_allow_html=True,
)
fig_map = px.scatter_geo(
    df_v1_wide,
    lat="lat", lon="lon",
    size="total_region",
    color="tasa_otorga",
    color_continuous_scale="Blues",
    size_max=60,
    scope="world",
    hover_name="REGIÓN",
    hover_data={
        "total_region":  ":,.0f",
        "tasa_otorga":   ":.1f",
        "tasa_rechazo":  ":.1f",
        "top3_paises":   True,
        "lat": False, "lon": False,
    },
    labels={
        "total_region":  "Registros totales",
        "tasa_otorga":   "Tasa otorgamiento (%)",
        "tasa_rechazo":  "Tasa rechazo (%)",
        "top3_paises":   "Top 3 nacionalidades",
    },
)
fig_map.update_geos(
    showcountries=True,  countrycolor="#444",
    showland=True,       landcolor="#1e2235",
    showocean=True,      oceancolor="#0d1117",
    showcoastlines=True, coastlinecolor="#555",
    showframe=False,
)
fig_map.update_layout(
    paper_bgcolor="#0f1117",
    plot_bgcolor="#0f1117",
    font_color="#e8eaf6",
    height=460,
    margin=dict(l=0, r=0, t=4, b=0),
    coloraxis_colorbar=dict(
        title_text="% Otorgamiento",
        thickness=14,
        len=0.55,
        y=0.5,
        tickfont=dict(size=10),
    ),
)
st.plotly_chart(fig_map, use_container_width=True)
st.markdown(
    '<div class="nota" style="margin-top:2px;">'
    '🔵 Tamaño = volumen total · Azul oscuro = mayor tasa de otorgamiento · '
    'La RM concentra ~63,5% del total.</div>',
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tabla y donut abajo en dos columnas iguales ────────────
col_tabla, col_donut = st.columns([1, 1])

with col_tabla:
    st.markdown(
        "<div style='font-size:.85rem;font-weight:600;color:#e8eaf6;"
        "margin-bottom:4px;'>📋 Top 10 regiones</div>",
        unsafe_allow_html=True,
    )
    top10 = (
        df_v1_wide
        .nlargest(10, "total_region")
        [["REGIÓN","total_region","tasa_otorga","tasa_rechazo"]]
        .copy()
    )
    top10.columns = ["Región","Registros","% Otorga","% Rechazo"]
    top10["Registros"] = top10["Registros"].apply(lambda x: f"{x:,.0f}")
    top10["% Otorga"]  = top10["% Otorga"].apply(lambda x: f"{x:.1f}%")
    top10["% Rechazo"] = top10["% Rechazo"].apply(lambda x: f"{x:.1f}%")
    top10 = top10.reset_index(drop=True)
    top10.index += 1
    st.dataframe(top10, use_container_width=True, height=360)

with col_donut:
    st.markdown(
        "<div style='font-size:.85rem;font-weight:600;color:#e8eaf6;"
        "margin-bottom:4px;'>🍩 Distribución de resoluciones</div>",
        unsafe_allow_html=True,
    )
    df_donut = df.groupby("TIPO_RESUELTO")["Total"].sum().reset_index()
    total_fmt = f"{int(df_donut['Total'].sum()):,}".replace(",", ".")
    fig_donut = go.Figure(go.Pie(
        labels=df_donut["TIPO_RESUELTO"],
        values=df_donut["Total"],
        hole=0.58,
        marker_colors=[COLOR_MAP.get(t,"#ccc") for t in df_donut["TIPO_RESUELTO"]],
        marker_line=dict(color="#0f1117", width=2),
        showlegend=True,
        textinfo="percent",
        textfont_size=10,
        textposition="inside",
    ))
    fig_donut.add_annotation(
        text=f"<b>{total_fmt}</b><br>registros",
        x=0.5, y=0.5, showarrow=False,
        font=dict(size=11, color="#e8eaf6"),
        align="center",
    )
    fig_donut.update_layout(
        paper_bgcolor="#0f1117", font_color="#e8eaf6",
        height=360,
        margin=dict(l=0, r=0, t=0, b=0),
        legend=dict(
            orientation="v",
            x=1.01, y=0.5,
            xanchor="left",
            font=dict(size=9, color="#bdbdbd"),
            bgcolor="rgba(0,0,0,0)",
        ),
    )
    st.plotly_chart(fig_donut, use_container_width=True)

st.markdown("---")

# ══════════════════════════════════════════════════════════
#  VISTA 2 — BARRAS APILADAS AL 100 %
#  Tarea: Comparar resultados por perfil sociodemográfico
# ══════════════════════════════════════════════════════════
st.markdown(
    '<div class="vista-title">📊 Vista 2: Comparación de Perfiles Sociodemográficos</div>'
    '<div class="vista-sub">¿Existen diferencias sistemáticas en los resultados de las '
    'solicitudes según el perfil del solicitante? — <em>Tarea: Comparar</em></div>',
    unsafe_allow_html=True,
)

col_v2ctrl, col_v2main = st.columns([1, 3])

with col_v2ctrl:
    eje_y = st.selectbox(
        "Agrupar por:",
        ["PAÍS","SEXO","RANGO_ETARIO","ESTUDIOS"],
        format_func=lambda x: {
            "PAÍS":        "Nacionalidad",
            "SEXO":        "Sexo",
            "RANGO_ETARIO":"Rango Etario",
            "ESTUDIOS":    "Nivel de Estudios",
        }[x],
    )
    min_regs_v2 = st.number_input(
        "Mínimo registros por grupo",
        min_value=100, max_value=100_000, value=1000, step=500,
    )

# Preparar datos Vista 2
df_v2 = df.copy()
if eje_y == "RANGO_ETARIO":
    df_v2 = df_v2[~df_v2["RANGO_ETARIO"].isin(EXCLUIR_ETARIO)]
elif eje_y == "ESTUDIOS":
    df_v2 = df_v2[~df_v2["ESTUDIOS"].isin(EXCLUIR_ESTUDIOS)]

df_v2_agg = (
    df_v2
    .groupby([eje_y, "TIPO_RESUELTO"])["Total"]
    .sum()
    .reset_index()
)
df_v2_tot = (
    df_v2
    .groupby(eje_y)["Total"]
    .sum()
    .reset_index()
    .rename(columns={"Total": "gran_total"})
)
df_v2_agg = df_v2_agg.merge(df_v2_tot, on=eje_y)
df_v2_agg = df_v2_agg[df_v2_agg["gran_total"] >= min_regs_v2]
df_v2_agg["pct"] = (df_v2_agg["Total"] / df_v2_agg["gran_total"] * 100).round(2)

# Ordenar por tasa de Otorga descendente
orden_otorga = (
    df_v2_agg[df_v2_agg["TIPO_RESUELTO"] == "Otorga"]
    .sort_values("pct", ascending=True)[eje_y]
    .tolist()
)

# Promedio general de otorgamiento para línea de referencia
promedio_otorga = (
    df_v2_agg[df_v2_agg["TIPO_RESUELTO"] == "Otorga"]["pct"].mean()
    if not df_v2_agg.empty else 0
)

with col_v2main:
    if df_v2_agg.empty:
        st.warning("Sin datos para los filtros seleccionados.")
    else:
        fig_bar = px.bar(
            df_v2_agg,
            x="pct", y=eje_y,
            color="TIPO_RESUELTO",
            color_discrete_map=COLOR_MAP,
            category_orders={eje_y: orden_otorga, "TIPO_RESUELTO": TIPO_ORDER},
            orientation="h",
            barmode="stack",
            hover_data={
                "Total":       ":,.0f",
                "pct":         ":.2f",
                "gran_total":  ":,.0f",
            },
            labels={
                "pct":        "Proporción (%)",
                eje_y:        eje_y,
                "Total":      "Registros",
                "gran_total": "Total subgrupo",
                "TIPO_RESUELTO": "Tipo",
            },
            title=f"Distribución de resoluciones por {eje_y} (% dentro del subgrupo)",
        )
        # Línea de referencia: promedio general de otorgamiento
        fig_bar.add_vline(
            x=promedio_otorga,
            line_dash="dot", line_color="#e8eaf6", line_width=1.4,
            annotation_text=f"Promedio Otorga: {promedio_otorga:.1f}%",
            annotation_position="top right",
            annotation_font_color="#e8eaf6",
            annotation_font_size=11,
        )
        fig_bar.update_layout(
            paper_bgcolor="#0f1117", plot_bgcolor="#0f1117",
            font_color="#e8eaf6",
            height=max(380, len(orden_otorga) * 34 + 80),
            xaxis=dict(
                range=[0, 100], ticksuffix="%",
                gridcolor="#2a2a3e", zerolinecolor="#555",
            ),
            yaxis=dict(gridcolor="#2a2a3e"),
            margin=dict(l=10, r=10, t=50, b=30),
            showlegend=False,   # leyenda ya está en la barra global
            bargap=0.14,
        )
        st.plotly_chart(fig_bar, use_container_width=True)

# Nota fuera de las columnas para evitar solapamiento con la tabla
st.markdown(
    '<div class="nota">'
    'Las barras se ordenan por tasa de Otorga ascendente (mayor tasa al tope). '
    'La línea punteada blanca indica el promedio general de otorgamiento '
    'para los filtros activos. Grupos con menos registros que el umbral '
    'definido son excluidos automáticamente.'
    '</div>',
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# Tabla resumen bajo el gráfico (expansible)
with st.expander("📋 Ver tabla de proporciones por grupo"):
    if not df_v2_agg.empty:
        resumen_v2 = (
            df_v2_agg
            .pivot_table(index=eje_y, columns="TIPO_RESUELTO",
                         values="pct", aggfunc="sum")
            .round(1)
        )
        # Añadir columna de total absoluto
        resumen_v2["Total registros"] = (
            df_v2_tot.set_index(eje_y)["gran_total"]
        )
        resumen_v2 = resumen_v2.reset_index()
        # Re-ordenar igual que el gráfico
        orden_idx = [x for x in reversed(orden_otorga)
                     if x in resumen_v2[eje_y].values]
        resumen_v2 = resumen_v2.set_index(eje_y).reindex(orden_idx).reset_index()
        st.dataframe(resumen_v2, use_container_width=True)
    else:
        st.info("Sin datos para la tabla.")

st.markdown("---")

# ══════════════════════════════════════════════════════════
#  VISTA 3 — LÍNEAS TEMPORALES DE RECHAZOS
#  Tarea: Explorar evolución temporal del tipo de rechazo
# ══════════════════════════════════════════════════════════
st.markdown(
    '<div class="vista-title">📈 Vista 3: Análisis Temporal de Rechazos (2000–2025)</div>'
    '<div class="vista-sub">¿Cómo varía el tipo de rechazo (abandono vs. RT) a lo largo '
    'del tiempo según nacionalidad, región y rango etario? — <em>Tarea: Explorar</em></div>',
    unsafe_allow_html=True,
)

col_v3ctrl, col_v3main = st.columns([1, 3])

with col_v3ctrl:
    desglose_v3 = st.selectbox(
        "Desglosar líneas por:",
        ["Sin desglose (total)", "PAÍS", "REGIÓN", "RANGO_ETARIO"],
        format_func=lambda x: {
            "Sin desglose (total)": "Sin desglose (total)",
            "PAÍS":        "Nacionalidad",
            "REGIÓN":      "Región",
            "RANGO_ETARIO":"Rango Etario",
        }[x],
    )
    top_n = st.slider("Top N grupos", 1, 10, 5,
                      help="Muestra los N grupos con mayor volumen total de rechazos")
    mostrar_abs = st.toggle("Mostrar valores absolutos", value=False,
                            help="Alterna entre porcentaje y conteo absoluto")

# Filtrar sólo rechazos
df_v3 = df[df["TIPO_RESUELTO"].isin(["Rechaza con Rt","Rechaza con abandono"])].copy()

with col_v3main:
    if df_v3.empty:
        st.warning("Sin registros de rechazo para los filtros seleccionados.")
    else:
        eje_y_v3 = "Total" if mostrar_abs else "pct"
        label_y   = "Registros" if mostrar_abs else "% sobre total rechazos"

        if desglose_v3 == "Sin desglose (total)":
            # Agregar por año y tipo de rechazo
            agg = (
                df_v3
                .groupby(["AÑO","TIPO_RESUELTO"])["Total"]
                .sum()
                .reset_index()
            )
            tot_año = (
                df_v3.groupby("AÑO")["Total"].sum()
                .reset_index().rename(columns={"Total":"total_año"})
            )
            agg = agg.merge(tot_año, on="AÑO")
            agg["pct"] = (agg["Total"] / agg["total_año"] * 100).round(2)

            fig_line = px.line(
                agg, x="AÑO", y=eje_y_v3,
                color="TIPO_RESUELTO",
                color_discrete_map=COLOR_MAP,
                markers=True,
                hover_data={"Total":":,.0f","pct":":.2f","total_año":":,.0f"},
                labels={
                    "pct":        "% sobre total rechazos",
                    "Total":      "Registros",
                    "AÑO":        "Año",
                    "total_año":  "Total rechazos del año",
                    "TIPO_RESUELTO": "Tipo",
                },
                title="Evolución del tipo de rechazo — total histórico",
            )

        else:
            # Top N grupos por volumen de rechazos
            top_grupos = (
                df_v3.groupby(desglose_v3)["Total"].sum()
                .nlargest(top_n).index.tolist()
            )
            df_v3_f = df_v3[df_v3[desglose_v3].isin(top_grupos)]

            agg = (
                df_v3_f
                .groupby(["AÑO", desglose_v3, "TIPO_RESUELTO"])["Total"]
                .sum()
                .reset_index()
            )
            tot_sub = (
                df_v3_f
                .groupby(["AÑO", desglose_v3])["Total"].sum()
                .reset_index().rename(columns={"Total":"total_año"})
            )
            agg = agg.merge(tot_sub, on=["AÑO", desglose_v3])
            agg["pct"] = (agg["Total"] / agg["total_año"] * 100).round(2)

            fig_line = px.line(
                agg, x="AÑO", y=eje_y_v3,
                color=desglose_v3,
                line_dash="TIPO_RESUELTO",
                line_dash_map={
                    "Rechaza con Rt":       "solid",
                    "Rechaza con abandono": "dot",
                },
                markers=True,
                hover_data={"Total":":,.0f","pct":":.2f"},
                labels={
                    "pct":   "% sobre total rechazos",
                    "Total": "Registros",
                    "AÑO":   "Año",
                    "TIPO_RESUELTO": "Tipo",
                },
                title=f"Tipo de rechazo por {desglose_v3} — top {top_n}",
            )

        # Anotaciones contextuales fijas
        for año_an, label_an, color_an in [
            (2021, "↓ Mínimo\nCOVID",     "#f1a340"),
            (2024, "↑ Máximo\nhistórico",  "#4dac26"),
        ]:
            fig_line.add_vline(
                x=año_an,
                line_dash="dash", line_color=color_an, line_width=1.2,
                annotation_text=label_an,
                annotation_position="top",
                annotation_font_color=color_an,
                annotation_font_size=11,
            )

        # Línea de referencia horizontal: promedio histórico de % Rechaza con RT
        if not mostrar_abs and desglose_v3 == "Sin desglose (total)":
            prom_rt = agg[agg["TIPO_RESUELTO"]=="Rechaza con Rt"]["pct"].mean()
            fig_line.add_hline(
                y=prom_rt,
                line_dash="dot", line_color="#d01c8b", line_width=1,
                annotation_text=f"Promedio RT: {prom_rt:.1f}%",
                annotation_position="bottom right",
                annotation_font_color="#d01c8b",
                annotation_font_size=10,
            )

        y_range = [0, 105] if not mostrar_abs else None
        fig_line.update_layout(
            paper_bgcolor="#0f1117", plot_bgcolor="#0f1117",
            font_color="#e8eaf6", height=450,
            xaxis=dict(
                dtick=2, tickangle=-45,
                gridcolor="#2a2a3e", zerolinecolor="#555",
            ),
            yaxis=dict(
                range=y_range,
                ticksuffix=("%" if not mostrar_abs else ""),
                gridcolor="#2a2a3e", zerolinecolor="#555",
            ),
            legend=dict(
                bgcolor="#1e2130", bordercolor="#444",
                borderwidth=1, font_color="#e8eaf6",
                font_size=10,
                orientation="h",
                yanchor="top", y=-0.20,
                xanchor="center", x=0.5,
                itemwidth=40,
            ),
            margin=dict(l=10, r=10, t=50, b=130),
            hovermode="x unified",
        )
        st.plotly_chart(fig_line, use_container_width=True)
        
st.markdown(
    '<div class="nota">'
    'Línea continua = Rechaza con RT (residencia temporaria otorgada como alternativa). '
    'Línea punteada = Rechaza con abandono (salida del país). '
    'Del total histórico de rechazos, el 89,6% corresponde a "Rechaza con RT" '
    'y el 10,4% a "Rechaza con abandono".'
    '</div>',
    unsafe_allow_html=True,
)

st.markdown("<br>", unsafe_allow_html=True)

# ── Tabla de rechazos por año (expandible) ────────────────
with st.expander("📋 Ver tabla de rechazos por año"):
    if not df_v3.empty:
        tabla_v3 = (
            df_v3
            .groupby(["AÑO","TIPO_RESUELTO"])["Total"].sum()
            .unstack(fill_value=0)
            .reset_index()
        )
        tabla_v3["Total rechazos"] = tabla_v3.get("Rechaza con Rt",0) + \
                                     tabla_v3.get("Rechaza con abandono",0)
        tabla_v3["% RT"]   = (tabla_v3.get("Rechaza con Rt",0) /
                               tabla_v3["Total rechazos"].replace(0,np.nan) * 100).round(1)
        tabla_v3["% Abnd"] = (tabla_v3.get("Rechaza con abandono",0) /
                               tabla_v3["Total rechazos"].replace(0,np.nan) * 100).round(1)
        st.dataframe(tabla_v3, use_container_width=True)

st.markdown("---")

# ══════════════════════════════════════════════════════════
#  DISCUSIÓN Y CONCLUSIONES
# ══════════════════════════════════════════════════════════
st.markdown("### 💬 Discusión y Conclusiones")

with st.expander("Ver discusión completa", expanded=False):
    st.markdown("""
**Hallazgos principales**

1. **Concentración geográfica extrema:** La Región Metropolitana acumula ~63,5% de todos
   los registros administrativos, seguida a gran distancia por Valparaíso (~7%) y Antofagasta (~5%).
   Sin embargo, regiones del norte como Tarapacá y Arica presentan mayor proporción relativa de
   solicitudes bolivianas y peruanas, evidenciando patrones migratorios ligados a la geografía fronteriza.

2. **Brechas sistemáticas por nacionalidad:** La tasa de otorgamiento varía significativamente:
   Venezuela (~85–90%) contrasta con Haití (~42%), donde el 27,5% de registros corresponde a
   "Rechaza con RT" frente al 5,7% venezolano. Esta diferencia puede estar correlacionada con
   el nivel de estudios autodeclarado y la actividad laboral, aunque la alta tasa de "No informa"
   en estas variables limita conclusiones definitivas.

3. **Caída atípica en 2021 (COVID-19):** El año 2021 registró solo 28.629 actos administrativos,
   frente a 96.649 en 2022 y 185.140 en 2024 (máximo histórico), lo que revela el impacto directo
   de la pandemia en la tramitación administrativa migratoria.

4. **Dominio de "Rechaza con RT":** Del total de rechazos, el 89,6% corresponde a la modalidad
   RT (residencia temporaria como alternativa) y solo el 10,4% a "abandono". La proporción de
   abandono ha tendido a disminuir en los últimos años, sugiriendo cambios en la política
   administrativa de resolución de solicitudes.

**Limitaciones**

- Los registros son actos administrativos, no personas únicas; no es posible estimar
  el número de personas migrantes a partir de este dataset.
- El 0,3% de filas (Total=`*`) fue excluido por criterios de privacidad.
- Las variables ACTIVIDAD y ESTUDIOS tienen alta tasa de "No informa", lo que limita
  el análisis sociodemográfico detallado.

**Trabajo futuro**

- Cruzar con el Censo 2024 para contrastar registros administrativos con stock real de población.
- Incorporar datos de solicitudes de Residencia Temporaria para analizar el flujo completo.
- Modelar probabilidad de otorgamiento mediante regresión logística controlando por año,
  región, sexo, edad y nivel de estudios.
    """)

# ══════════════════════════════════════════════════════════
#  NOTA FINAL DE INTERPRETACIÓN (pie de página fijo)
# ══════════════════════════════════════════════════════════
st.markdown(
    '<div class="nota" style="margin-top:4px;">'
    '📌 <b>Nota de interpretación:</b> Los registros contabilizan <b>actos administrativos</b>, '
    'no personas únicas. Una misma persona puede aparecer en múltiples años o con múltiples '
    'solicitudes. Las cifras <b>no pueden usarse</b> para estimar el número total de extranjeros '
    'en Chile. Se excluyen 172 filas (0,3%) con Total=<code>*</code> por criterio de privacidad '
    '(menos de 6 casos por combinación de categorías). — '
    '<b>Fuente:</b> SERMIG, 2026.'
    '</div>',
    unsafe_allow_html=True,
)
