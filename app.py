"""
Dashboard interactivo — Análisis de rendimiento Planta de Litio Novandino.

Cómo usar:
1. Subir el Excel con el formato estándar (hoja "Hoja2" con columnas de planta).
2. Ajustar la meta de rendimiento si fuese necesario.
3. Explorar los slides del análisis y descargar resúmenes.

Autora: Natalia Silva — Ing. de procesos.
"""
import io
import numpy as np
import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px

# ===================== CONFIGURACIÓN GENERAL =====================
st.set_page_config(
    page_title="Análisis Rendimiento Planta — Novandino",
    page_icon="🧪",
    layout="wide",
)

# Paleta Novandino
COLOR_SOBRE = "#1FB8A6"
COLOR_BAJO = "#5C3A9E"
COLOR_DARK = "#3B1E6E"
COLOR_TEAL = "#1FB8A6"
COLOR_META = "#FF6B6B"
COLOR_ROW = "#F3EEFA"

# CSS de estilo Novandino
st.markdown("""
<style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; max-width: 1280px; }
    h1, h2, h3 { color: #3B1E6E; font-family: 'Calibri', sans-serif; }
    h1 { border-bottom: 3px solid #1FB8A6; padding-bottom: 0.5rem; }
    .stMetric { background-color: #F3EEFA; padding: 1rem; border-radius: 0.5rem; }
    div[data-testid="stMetricValue"] { color: #3B1E6E; font-weight: bold; }
    .nova-callout {
        background-color: #F3EEFA;
        border-left: 4px solid #1FB8A6;
        padding: 1rem 1.25rem;
        border-radius: 0.4rem;
        margin: 1rem 0;
    }
    .nova-dark {
        background-color: #3B1E6E;
        color: white;
        padding: 1rem 1.25rem;
        border-radius: 0.4rem;
        margin: 1rem 0;
    }
    .nova-dark h4 { color: white !important; margin-top: 0; }
</style>
""", unsafe_allow_html=True)


# ===================== HEADER =====================
col1, col2 = st.columns([3, 1])
with col1:
    st.title("Análisis de rendimiento — Planta Carbonato de Litio")
    st.markdown("**Comparación periodo sobre meta vs bajo meta · separación control / desempeño**")
with col2:
    st.markdown("<div style='text-align:right; padding-top:1.5rem;'>"
                "<span style='color:#7A6AB8; font-size:0.9em;'>"
                "Novandino Litio © 2026<br>#SomosLitioSomosFuturo"
                "</span></div>", unsafe_allow_html=True)


# ===================== CARGA DE DATOS =====================
st.sidebar.header("📁 Datos")
uploaded = st.sidebar.file_uploader(
    "Subir Excel de planta",
    type=["xlsx", "xls"],
    help="Excel con hoja llamada 'Hoja2' que contenga las columnas estándar de la planta."
)

meta = st.sidebar.number_input(
    "Meta de rendimiento [%]",
    min_value=50.0, max_value=95.0, value=78.0, step=0.5,
    help="Umbral para clasificar días como 'sobre meta' o 'bajo meta'."
) / 100.0

st.sidebar.markdown("---")
st.sidebar.markdown(
    "**Columnas requeridas** (mínimo):\n"
    "- Fecha, Rendimiento AQ\n"
    "- Flujo Bref, SCS, LM\n"
    "- Flujo PRS, SX1-2, SX3, NF\n"
    "- Flujo 605 rec.\n"
    "- T AFT, T BP2\n"
    "- Li AFT, Li Bfil, Mg Bfil\n"
    "- Li/Mg Bref, Li FT 6/7, Mg FT 6/7\n"
    "- Li BP, Mg BP, pH BP\n"
    "- Li QQ FP, Mg QQ FP\n"
    "- Carga E3, Li LM, Na LM, CO3 LM\n"
    "- Pérdidas E1, E2, E3\n"
    "- Producción\n"
    "- V inventario reactores 1E\n"
    "- Washing C1..C10"
)

if uploaded is None:
    st.info("👈 Sube el Excel de planta en la barra lateral para comenzar el análisis.")
    st.markdown("""
### Sobre este dashboard

Esta herramienta replica el análisis presentado al equipo R&D, aplicado a cualquier
periodo de operación que se desee analizar. Subiendo el Excel con el formato estándar:

1. **Identifica automáticamente** los días sobre y bajo meta.
2. **Compara parámetros de control** (entrada): flujos, T AFT, tiempo de residencia.
3. **Compara indicadores de desempeño** (salida): Mg/Li FT, Li AFT, pérdidas, Na LM.
4. **Calcula correlaciones** entre cada variable y el rendimiento.
5. **Recomienda setpoints** basados en el promedio de los días sobre meta.

Se mantiene el enfoque del feedback de R&D: separación clara entre **causa (control)** y
**efecto (desempeño)**.
""")
    st.stop()


# ===================== PROCESAMIENTO =====================
@st.cache_data(show_spinner="Procesando datos...")
def procesar(file_bytes, meta_val):
    df = pd.read_excel(file_bytes, sheet_name="Hoja2")

    # Fecha legible
    df["Fecha_dt"] = pd.to_datetime("1899-12-30") + pd.to_timedelta(df["Fecha"], unit="D")
    df["Dia"] = df["Fecha_dt"].dt.strftime("%d/%m")

    # Clasificación
    df["Grupo"] = df["Rendimiento AQ"].apply(
        lambda x: "Sobre meta" if x >= meta_val else "Bajo meta"
    )

    # Variables derivadas
    df["Li/Mg Bref"] = df["Li Bref [%]"] / df["Mg Bref [%]"]
    df["Mg FT avg"] = (df["Mg FT 6[%]"] + df["Mg FT 7[%]"]) / 2
    df["Li FT avg"] = (df["Li FT 6[%]"] + df["Li FT 7[%]"]) / 2
    df["Mg/Li FT"] = df["Mg FT avg"] / df["Li FT avg"]
    df["Mg/Li FP"] = df["Mg QQ FP [%]"] / df["Li QQ FP [%]"]
    df["Factor Producción"] = df["Producción "] / df["Carga E3 [m3/h]"]

    washing_cols = [f"Washing C{i}" for i in range(1, 11)]
    if all(c in df.columns for c in washing_cols):
        df["Washing Mediana"] = df[washing_cols].median(axis=1)
        df["Washing Avg"] = df[washing_cols].mean(axis=1)

    df["Flujo Recup Total"] = (
        df["Flujo PRS"] + df["Flujo SX1-2"] + df["Flujo SX3 "] + df["Flujo NF"]
    )

    # Caudal total entrada E1 y tiempo de residencia
    df["Q total E1"] = (
        df["Flujo Bref"] + df["Flujo SCS"] + df["Flujo LM "]
        + df["Flujo PRS"] + df["Flujo SX1-2"] + df["Flujo SX3 "]
        + df["Flujo NF"] + df["Flujo 605 rec."]
    )

    if "V inventario reactores 1E" in df.columns:
        df["Tres (h)"] = df["V inventario reactores 1E"] / df["Q total E1"]
    else:
        # Fallback con volumen referencial
        df["Tres (h)"] = 1000.0 / df["Q total E1"]

    return df

try:
    df = procesar(uploaded.getvalue(), meta)
except Exception as e:
    st.error(f"❌ Error al leer el Excel: {e}")
    st.markdown("Verifica que el archivo tiene la hoja **Hoja2** con las columnas estándar.")
    st.stop()

n_dias = len(df)
sobre = df[df["Grupo"] == "Sobre meta"]
bajo = df[df["Grupo"] == "Bajo meta"]
n_sobre = len(sobre); n_bajo = len(bajo)

if n_sobre == 0 or n_bajo == 0:
    st.warning(f"⚠️ Todos los días están en el mismo grupo (meta = {meta*100:.1f} %). "
               "Ajusta la meta en la barra lateral para hacer la comparación.")
    st.stop()

# ===================== KPIS PRINCIPALES =====================
st.markdown("## 📊 Resumen del periodo analizado")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.metric("Días analizados", n_dias)
with c2:
    st.metric("Sobre meta", f"{n_sobre} días",
              f"{sobre['Rendimiento AQ'].mean()*100:.1f}% promedio")
with c3:
    st.metric("Bajo meta", f"{n_bajo} días",
              f"{bajo['Rendimiento AQ'].mean()*100:.1f}% promedio",
              delta_color="inverse")
with c4:
    delta_rend = (sobre["Rendimiento AQ"].mean() - bajo["Rendimiento AQ"].mean()) * 100
    st.metric("Brecha de rendimiento", f"{delta_rend:.2f} pts")


# ===================== RENDIMIENTO DIARIO =====================
st.markdown("## 📈 Rendimiento diario")

fig_rend = go.Figure()
colores = [COLOR_SOBRE if g == "Sobre meta" else COLOR_BAJO for g in df["Grupo"]]
fig_rend.add_trace(go.Bar(
    x=df["Dia"], y=df["Rendimiento AQ"] * 100,
    marker_color=colores,
    text=[f"{v*100:.1f}%" for v in df["Rendimiento AQ"]],
    textposition="outside",
    name="Rendimiento",
    hovertemplate="<b>%{x}</b><br>Rendimiento: %{y:.2f}%<extra></extra>",
))
fig_rend.add_hline(y=meta * 100, line_dash="dash", line_color=COLOR_META,
                   annotation_text=f"Meta {meta*100:.0f}%",
                   annotation_position="top right")
fig_rend.update_layout(
    height=400, plot_bgcolor="white",
    yaxis=dict(title="Rendimiento AQ [%]", gridcolor="#EEE",
               range=[max(74, df["Rendimiento AQ"].min()*100 - 2),
                      min(85, df["Rendimiento AQ"].max()*100 + 2)]),
    xaxis=dict(title=""),
    showlegend=False, margin=dict(t=20, b=40, l=40, r=20),
)
st.plotly_chart(fig_rend, use_container_width=True)


# ===================== HELPERS DE GRÁFICOS =====================
def comparar_bar(variable_dict, n_cols=3):
    """variable_dict: {label: {'sobre': v1, 'bajo': v2, 'unit': '%', 'fmt': '.2f'}}"""
    items = list(variable_dict.items())
    n = len(items)
    n_rows = int(np.ceil(n / n_cols))
    cols_tpl = st.columns(n_cols)

    for i, (label, vals) in enumerate(items):
        with cols_tpl[i % n_cols]:
            fmt = vals.get("fmt", ".2f")
            unit = vals.get("unit", "")
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=["Sobre meta", "Bajo meta"],
                y=[vals["sobre"], vals["bajo"]],
                marker_color=[COLOR_SOBRE, COLOR_BAJO],
                text=[f"{vals['sobre']:{fmt}}", f"{vals['bajo']:{fmt}}"],
                textposition="outside",
                hovertemplate="<b>%{x}</b><br>"+f"{label}: %{{y:{fmt}}} {unit}<extra></extra>",
            ))
            fig.update_layout(
                title=dict(text=label, font=dict(size=13, color=COLOR_DARK)),
                height=280, plot_bgcolor="white", showlegend=False,
                margin=dict(t=40, b=20, l=40, r=20),
                yaxis=dict(title=unit, gridcolor="#EEE"),
            )
            st.plotly_chart(fig, use_container_width=True)
        if (i + 1) % n_cols == 0 and i + 1 < n:
            cols_tpl = st.columns(n_cols)


def get_avg(g, col, mult=1.0):
    if col not in df.columns:
        return None, None
    s_v = sobre[col].mean() * mult
    b_v = bajo[col].mean() * mult
    return s_v, b_v


# ===================== TABS =====================
tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "🎯 Enfoque",
    "🔬 E1 — Control",
    "📐 E1 — Desempeño",
    "🔄 TK605",
    "⚗️ E2 y E3",
    "🧂 Na LM",
    "✅ Setpoints",
])

# -------------------- TAB 1: ENFOQUE --------------------
with tab1:
    st.markdown("### Enfoque del análisis")
    st.markdown("Separación entre **parámetros de control (entrada)** e "
                "**indicadores de desempeño (salida)**.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
<div class="nova-dark">
<h4>Parámetros de control (entrada)</h4>
<i>Lo que definimos como punto de operación</i>
<ul style="margin-top:0.5rem;">
<li><b>E1</b> — Flujos: Bref, SCS, LM, PRS, SX1-2, SX3, NF, TK605</li>
<li><b>E1</b> — Temperatura AFT</li>
<li><b>E1</b> — Tiempo de residencia (V/Q)</li>
<li><b>E1</b> — Li/Mg Bref</li>
<li><b>E2</b> — pH BP, T BP2</li>
<li><b>E3</b> — Carga E3, washing centrífugas</li>
</ul>
</div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
<div style="background-color:#1FB8A6; color:white; padding:1rem 1.25rem; border-radius:0.4rem; margin:1rem 0;">
<h4 style="color:white; margin-top:0;">Indicadores de desempeño (salida)</h4>
<i>Resultado del ejercicio operacional</i>
<ul style="margin-top:0.5rem;">
<li><b>E1</b> — KPI clave: <b>Li Bfil</b> y <b>Mg/Li FT</b></li>
<li><b>E1</b> — Li AFT, Mg Bfil, Pérdida E1</li>
<li><b>E2</b> — Li BP, Mg BP, Mg/Li FP, Pérdida E2</li>
<li><b>E3</b> — Li LM, Na LM, Pérdida E3</li>
<li>Global — Factor producción, Rendimiento AQ</li>
</ul>
</div>
        """, unsafe_allow_html=True)

    # Mapa de correlaciones
    st.markdown("### Mapa de correlaciones con Rendimiento AQ")
    cols_corr = ["Flujo Bref", "Flujo SCS", "Flujo LM ",
                 "Flujo PRS", "Flujo SX1-2", "Flujo SX3 ", "Flujo NF",
                 "Flujo 605 rec.", "Q total E1", "Tres (h)",
                 "Li/Mg Bref", "T AFT [°C]", "Li AFT [%]", "Li Bfil [%]",
                 "Mg Bfil [%]", "Mg/Li FT", "Perdida E1 [%]",
                 "pH BP ", "Li BP [%]", "Mg BP [ppm]", "Mg/Li FP",
                 "Perdida E2 [%]", "Carga E3 [m3/h]", "Li LM [%]",
                 "Na LM [%]", "Factor Producción", "Perdida E3 [%]"]
    cols_corr = [c for c in cols_corr if c in df.columns]
    corr = (df[cols_corr + ["Rendimiento AQ"]]
            .corr()["Rendimiento AQ"].drop("Rendimiento AQ").sort_values())

    fig_corr = go.Figure()
    fig_corr.add_trace(go.Bar(
        x=corr.values, y=corr.index, orientation="h",
        marker_color=[COLOR_BAJO if v < 0 else COLOR_SOBRE for v in corr.values],
        text=[f"{v:+.2f}" for v in corr.values],
        textposition="outside",
        hovertemplate="<b>%{y}</b><br>Correlación: %{x:.3f}<extra></extra>",
    ))
    fig_corr.update_layout(
        height=600, plot_bgcolor="white",
        xaxis=dict(title="Correlación de Pearson", range=[-1, 1], gridcolor="#EEE",
                   zeroline=True, zerolinecolor="black", zerolinewidth=1),
        yaxis=dict(title=""), margin=dict(t=20, b=40, l=180, r=20),
    )
    st.plotly_chart(fig_corr, use_container_width=True)

# -------------------- TAB 2: E1 CONTROL --------------------
with tab2:
    st.markdown("### Primera etapa (E1) — Parámetros de control")
    st.markdown("Flujos individuales de entrada, T AFT, Li/Mg Bref y tiempo de residencia.")

    # Flujos individuales
    st.markdown("#### Flujos individuales")
    flujos_dict = {}
    for col, label, fmt in [
        ("Flujo Bref", "Flujo Bref", ".1f"),
        ("Flujo SCS", "Flujo SCS (ceniza)", ".1f"),
        ("Flujo LM ", "Flujo LM", ".1f"),
        ("Flujo PRS", "Flujo PRS", ".1f"),
        ("Flujo SX1-2", "Flujo SX1-2", ".1f"),
        ("Flujo SX3 ", "Flujo SX3", ".1f"),
        ("Flujo NF", "Flujo NF", ".2f"),
        ("Flujo 605 rec.", "Flujo TK605", ".2f"),
    ]:
        s_v, b_v = get_avg("Sobre meta", col)
        if s_v is not None:
            flujos_dict[label] = {"sobre": s_v, "bajo": b_v, "unit": "m³/h", "fmt": fmt}
    comparar_bar(flujos_dict, n_cols=4)

    # Otros controles E1
    st.markdown("#### Otros parámetros de control")
    otros = {}
    for col, label, fmt, unit in [
        ("T AFT [°C]", "T AFT", ".2f", "°C"),
        ("Li/Mg Bref", "Li/Mg Bref", ".3f", "[-]"),
        ("Flujo Recup Total", "Recuperaciones total", ".1f", "m³/h"),
    ]:
        s_v, b_v = get_avg("Sobre meta", col)
        if s_v is not None:
            otros[label] = {"sobre": s_v, "bajo": b_v, "unit": unit, "fmt": fmt}
    comparar_bar(otros, n_cols=3)

    # Tiempo de residencia
    st.markdown("#### Tiempo de residencia (V real / Q total)")
    tres = {}
    for col, label, fmt, unit in [
        ("Q total E1", "Caudal total E1", ".1f", "m³/h"),
        ("V inventario reactores 1E", "Volumen reactores 1E", ".0f", "m³"),
        ("Tres (h)", "Tiempo de residencia", ".3f", "h"),
    ]:
        s_v, b_v = get_avg("Sobre meta", col)
        if s_v is not None:
            tres[label] = {"sobre": s_v, "bajo": b_v, "unit": unit, "fmt": fmt}
    comparar_bar(tres, n_cols=3)

    s_v, b_v = get_avg("Sobre meta", "Tres (h)")
    if s_v is not None and b_v is not None:
        st.markdown(f"""
<div class="nova-callout">
<b>Lectura:</b> tiempo de residencia es {s_v:.2f} h sobre meta y {b_v:.2f} h bajo meta.
{'Sube' if b_v > s_v else 'Baja'} en el periodo bajo meta — el tiempo de reacción
{'no es' if b_v >= s_v else 'podría ser'} el factor limitante.
</div>
        """, unsafe_allow_html=True)

# -------------------- TAB 3: E1 DESEMPEÑO --------------------
with tab3:
    st.markdown("### Primera etapa (E1) — Indicadores de desempeño")
    st.markdown("**KPI clave según feedback R&D: Li Bfil y Mg/Li FT.**")

    desemp = {}
    for col, label, fmt, unit, mult in [
        ("Li AFT [%]", "Li AFT", ".3f", "%", 1.0),
        ("Li Bfil [%]", "Li Bfil ⭐", ".3f", "%", 1.0),
        ("Mg Bfil [%]", "Mg Bfil", ".3f", "%", 1.0),
        ("Mg/Li FT", "Mg/Li FT ⭐", ".3f", "[-]", 1.0),
        ("Perdida E1 [%]", "Pérdida E1", ".2f", "%", 100.0),
    ]:
        s_v, b_v = get_avg("Sobre meta", col, mult)
        if s_v is not None:
            desemp[label] = {"sobre": s_v, "bajo": b_v, "unit": unit, "fmt": fmt}
    comparar_bar(desemp, n_cols=3)

    s_mgli, b_mgli = get_avg("Sobre meta", "Mg/Li FT")
    s_aft, b_aft = get_avg("Sobre meta", "Li AFT [%]")
    s_p1, b_p1 = get_avg("Sobre meta", "Perdida E1 [%]", 100)
    if all(v is not None for v in [s_mgli, b_mgli, s_aft, b_aft, s_p1, b_p1]):
        st.markdown(f"""
<div class="nova-callout">
<b>Lectura del desempeño de E1:</b><br>
• <b>Mg/Li FT</b>: {s_mgli:.3f} → {b_mgli:.3f}
({(b_mgli-s_mgli)/s_mgli*100:+.1f} %): el sólido sale con
{'menos' if b_mgli < s_mgli else 'más'} Mg y
{'más' if b_mgli < s_mgli else 'menos'} Li → menor selectividad.<br>
• <b>Li AFT</b>: {s_aft:.3f} → {b_aft:.3f} %: mayor Li remanente soluble.<br>
• <b>Pérdida E1</b>: {s_p1:.2f} → {b_p1:.2f} % ({b_p1-s_p1:+.2f} pts).
</div>
        """, unsafe_allow_html=True)

# -------------------- TAB 4: TK605 --------------------
with tab4:
    st.markdown("### Recirculación interna TK605")
    st.markdown("Análisis de la recirculación interna de E1 sobre el rendimiento.")

    if "Flujo 605 rec." in df.columns:
        # Evolución diaria + scatter
        c1, c2 = st.columns(2)
        with c1:
            fig = go.Figure()
            colores = [COLOR_SOBRE if g == "Sobre meta" else COLOR_BAJO for g in df["Grupo"]]
            fig.add_trace(go.Bar(
                x=df["Dia"], y=df["Flujo 605 rec."],
                marker_color=colores,
                text=[f"{v:.1f}" for v in df["Flujo 605 rec."]],
                textposition="outside",
            ))
            fig.add_hline(y=df["Flujo 605 rec."].mean(),
                          line_dash="dot", line_color="gray",
                          annotation_text=f"Promedio = {df['Flujo 605 rec.'].mean():.1f}")
            fig.update_layout(
                title="Flujo TK605 por día", height=380,
                plot_bgcolor="white", showlegend=False,
                yaxis=dict(title="Flujo TK605 [m³/h]"),
                margin=dict(t=40, b=40, l=40, r=20),
            )
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            fig2 = go.Figure()
            for g, color in [("Sobre meta", COLOR_SOBRE), ("Bajo meta", COLOR_BAJO)]:
                sub = df[df["Grupo"] == g]
                fig2.add_trace(go.Scatter(
                    x=sub["Flujo 605 rec."], y=sub["Rendimiento AQ"]*100,
                    mode="markers+text",
                    marker=dict(size=14, color=color, line=dict(color="white", width=2)),
                    text=sub["Dia"], textposition="top center",
                    name=g, hovertemplate="<b>%{text}</b><br>"
                          "TK605: %{x:.2f}<br>Rendimiento: %{y:.2f}%<extra></extra>",
                ))
            corr_v = df["Flujo 605 rec."].corr(df["Rendimiento AQ"])
            z = np.polyfit(df["Flujo 605 rec."], df["Rendimiento AQ"]*100, 1)
            xr = np.linspace(df["Flujo 605 rec."].min(), df["Flujo 605 rec."].max(), 50)
            fig2.add_trace(go.Scatter(
                x=xr, y=np.polyval(z, xr), mode="lines",
                line=dict(dash="dash", color="gray"),
                name=f"Tendencia r = {corr_v:+.2f}",
            ))
            fig2.add_hline(y=meta*100, line_dash="dash", line_color=COLOR_META,
                           annotation_text=f"Meta {meta*100:.0f}%")
            fig2.update_layout(
                title="TK605 vs Rendimiento", height=380, plot_bgcolor="white",
                xaxis=dict(title="Flujo TK605 [m³/h]"),
                yaxis=dict(title="Rendimiento AQ [%]"),
                margin=dict(t=40, b=40, l=40, r=20),
            )
            st.plotly_chart(fig2, use_container_width=True)

        # Correlaciones específicas
        st.markdown("#### Impacto del flujo TK605 sobre variables de E1")
        corrs_tk = {}
        for col, label in [("Li AFT [%]", "Li AFT"),
                           ("Perdida E1 [%]", "Pérdida E1"),
                           ("Mg/Li FT", "Mg/Li FT"),
                           ("Mg Bfil [%]", "Mg Bfil"),
                           ("Rendimiento AQ", "Rendimiento AQ")]:
            if col in df.columns:
                corrs_tk[label] = df["Flujo 605 rec."].corr(df[col])

        fig_c = go.Figure()
        labels = list(corrs_tk.keys()); vals = list(corrs_tk.values())
        fig_c.add_trace(go.Bar(
            x=vals, y=labels, orientation="h",
            marker_color=[COLOR_BAJO if v < 0 else COLOR_SOBRE for v in vals],
            text=[f"{v:+.2f}" for v in vals], textposition="outside",
        ))
        fig_c.update_layout(
            height=300, plot_bgcolor="white", showlegend=False,
            xaxis=dict(title="Correlación con Flujo TK605", range=[-1, 1],
                       zeroline=True, zerolinecolor="black"),
            margin=dict(t=20, b=40, l=120, r=20),
        )
        st.plotly_chart(fig_c, use_container_width=True)

        s_tk, b_tk = get_avg("Sobre meta", "Flujo 605 rec.")
        st.markdown(f"""
<div class="nova-callout">
<b>Mecanismo:</b> mayor TK605 aporta semilla de Li₂CO₃ al reactor, favorece el
crecimiento de MgCO₃ y reduce la co-precipitación de Li → Mg/Li FT sube,
Li AFT baja, pérdida E1 baja.<br><br>
<b>Hallazgo:</b> promedio sobre meta = {s_tk:.2f} m³/h, bajo meta = {b_tk:.2f} m³/h
({(b_tk-s_tk)/s_tk*100:+.1f} %).
</div>
        """, unsafe_allow_html=True)
    else:
        st.warning("Columna 'Flujo 605 rec.' no encontrada en el Excel.")

# -------------------- TAB 5: E2 Y E3 --------------------
with tab5:
    st.markdown("### Segunda etapa (E2)")
    e2 = {}
    for col, label, fmt, unit, mult in [
        ("pH BP ", "pH BP", ".2f", "", 1.0),
        ("Li BP [%]", "Li BP", ".3f", "%", 1.0),
        ("Mg BP [ppm]", "Mg BP", ".2f", "ppm", 1.0),
        ("Mg/Li FP", "Mg/Li FP", ".2f", "[-]", 1.0),
        ("Perdida E2 [%]", "Pérdida E2", ".2f", "%", 100.0),
        ("T BP2 [°C]", "T BP2", ".2f", "°C", 1.0),
    ]:
        s_v, b_v = get_avg("Sobre meta", col, mult)
        if s_v is not None:
            e2[label] = {"sobre": s_v, "bajo": b_v, "unit": unit, "fmt": fmt}
    comparar_bar(e2, n_cols=3)

    st.markdown("### Tercera etapa (E3)")
    e3 = {}
    for col, label, fmt, unit, mult in [
        ("Carga E3 [m3/h]", "Carga E3", ".1f", "m³/h", 1.0),
        ("Li LM [%]", "Li LM", ".4f", "%", 1.0),
        ("Na LM [%]", "Na LM", ".2f", "%", 1.0),
        ("Factor Producción", "Factor producción", ".3f", "Ton/m³", 1.0),
        ("Perdida E3 [%]", "Pérdida E3", ".2f", "%", 100.0),
        ("Producción ", "Producción", ".1f", "Ton LCE", 1.0),
    ]:
        s_v, b_v = get_avg("Sobre meta", col, mult)
        if s_v is not None:
            e3[label] = {"sobre": s_v, "bajo": b_v, "unit": unit, "fmt": fmt}
    comparar_bar(e3, n_cols=3)

# -------------------- TAB 6: NA LM --------------------
with tab6:
    st.markdown("### Impacto del Na en licor madre (LM)")
    st.markdown("El Na LM es uno de los marcadores más fuertes de calidad y rendimiento.")

    if "Na LM [%]" in df.columns:
        c1, c2 = st.columns(2)
        with c1:
            fig = go.Figure()
            colores = [COLOR_SOBRE if g == "Sobre meta" else COLOR_BAJO for g in df["Grupo"]]
            fig.add_trace(go.Bar(
                x=df["Dia"], y=df["Na LM [%]"],
                marker_color=colores,
                text=[f"{v:.2f}" for v in df["Na LM [%]"]],
                textposition="outside",
            ))
            fig.add_hline(y=7.5, line_dash="dash", line_color=COLOR_META,
                          annotation_text="Objetivo ≤ 7.5%")
            fig.add_hline(y=df["Na LM [%]"].mean(),
                          line_dash="dot", line_color="gray",
                          annotation_text=f"Promedio = {df['Na LM [%]'].mean():.2f}%")
            fig.update_layout(
                title="Na LM por día", height=380, plot_bgcolor="white",
                showlegend=False, yaxis=dict(title="Na LM [%]"),
                margin=dict(t=40, b=40, l=40, r=20),
            )
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            fig2 = go.Figure()
            for g, color in [("Sobre meta", COLOR_SOBRE), ("Bajo meta", COLOR_BAJO)]:
                sub = df[df["Grupo"] == g]
                fig2.add_trace(go.Scatter(
                    x=sub["Na LM [%]"], y=sub["Rendimiento AQ"]*100,
                    mode="markers+text",
                    marker=dict(size=14, color=color, line=dict(color="white", width=2)),
                    text=sub["Dia"], textposition="top center",
                    name=g,
                ))
            corr_v = df["Na LM [%]"].corr(df["Rendimiento AQ"])
            z = np.polyfit(df["Na LM [%]"], df["Rendimiento AQ"]*100, 1)
            xr = np.linspace(df["Na LM [%]"].min(), df["Na LM [%]"].max(), 50)
            fig2.add_trace(go.Scatter(
                x=xr, y=np.polyval(z, xr), mode="lines",
                line=dict(dash="dash", color="gray"),
                name=f"Tendencia r = {corr_v:+.2f}",
            ))
            fig2.add_hline(y=meta*100, line_dash="dash", line_color=COLOR_META)
            fig2.update_layout(
                title="Na LM vs Rendimiento", height=380, plot_bgcolor="white",
                xaxis=dict(title="Na LM [%]"),
                yaxis=dict(title="Rendimiento AQ [%]"),
                margin=dict(t=40, b=40, l=40, r=20),
            )
            st.plotly_chart(fig2, use_container_width=True)

        st.markdown(f"""
<div class="nova-callout">
<b>¿Por qué importa el Na LM?</b><br>
• Indica mayor impregnación salina (NaCl) en la torta de Li₂CO₃ tras las centrífugas
→ riesgo directo sobre la calidad del producto final (Cl, Na).<br>
• Un LM con más Na implica menos Li y CO3 disponibles en la recirculación interna,
reforzando la pérdida de rendimiento global.<br>
<b>Correlación con rendimiento: r = {corr_v:+.2f}</b>
</div>
        """, unsafe_allow_html=True)

# -------------------- TAB 7: SETPOINTS --------------------
with tab7:
    st.markdown("### Setpoints recomendados de E1 para optimizar rendimiento")
    st.markdown(f"Basados en el promedio de los **{n_sobre} días sobre meta**, "
                "con rango min–max para control diario.")

    flujos_setpoints = [
        ("Flujo Bref", "Bref"),
        ("Flujo SCS", "SCS"),
        ("Flujo LM ", "LM"),
        ("Flujo PRS", "PRS"),
        ("Flujo SX1-2", "SX1-2"),
        ("Flujo SX3 ", "SX3"),
        ("Flujo NF", "NF"),
        ("Flujo 605 rec.", "TK605"),
    ]

    rows_table = []
    means, mins, maxs, labels_g = [], [], [], []
    for col, lab in flujos_setpoints:
        if col in sobre.columns:
            mn = sobre[col].min(); mx = sobre[col].max(); avg = sobre[col].mean()
            rows_table.append({
                "Flujo": lab,
                "Setpoint [m³/h]": f"{avg:.1f}",
                "Mínimo": f"{mn:.1f}",
                "Máximo": f"{mx:.1f}",
                "Rango": f"{mn:.1f} – {mx:.1f}",
            })
            means.append(avg); mins.append(mn); maxs.append(mx); labels_g.append(lab)

    if rows_table:
        c1, c2 = st.columns([2, 1])
        with c1:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=labels_g, y=means,
                marker_color=COLOR_SOBRE,
                text=[f"{v:.1f}" for v in means],
                textposition="outside",
                error_y=dict(
                    type="data", symmetric=False,
                    array=[mx - m for mx, m in zip(maxs, means)],
                    arrayminus=[m - mn for m, mn in zip(means, mins)],
                    color="#2C2C4E", thickness=2,
                ),
                hovertemplate="<b>%{x}</b><br>Setpoint: %{y:.1f} m³/h<extra></extra>",
            ))
            fig.update_layout(
                title="Setpoints recomendados con rango min–max", height=420,
                plot_bgcolor="white", showlegend=False,
                yaxis=dict(title="Flujo [m³/h]"),
                margin=dict(t=40, b=40, l=40, r=20),
            )
            st.plotly_chart(fig, use_container_width=True)

        with c2:
            df_sp = pd.DataFrame(rows_table)
            st.dataframe(df_sp[["Flujo", "Setpoint [m³/h]", "Rango"]],
                         hide_index=True, use_container_width=True)

        # T AFT y Mg/Li FT como complemento
        s_taft, _ = get_avg("Sobre meta", "T AFT [°C]")
        s_mgli, _ = get_avg("Sobre meta", "Mg/Li FT")
        st.markdown(f"""
<div class="nova-dark">
<h4>Lineamientos complementarios</h4>
• Mantener <b>T AFT ≥ {s_taft:.1f} °C</b> (promedio sobre meta).<br>
• Monitorear diariamente <b>Mg/Li FT ≥ {s_mgli:.2f}</b> como KPI de selectividad.<br>
• Vigilar <b>Na LM ≤ 7.5 %</b> como KPI de calidad y eficiencia de recirculación.<br>
• Mantener <b>TK605 ≥ {sobre['Flujo 605 rec.'].min():.1f} m³/h</b> como umbral mínimo.
</div>
        """, unsafe_allow_html=True)

        # CSV descargable
        csv = df_sp.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Descargar setpoints (CSV)",
            data=csv,
            file_name="setpoints_recomendados_E1.csv",
            mime="text/csv",
        )


# ===================== TABLA RESUMEN + DESCARGA =====================
st.markdown("---")
st.markdown("## 📋 Tabla resumen consolidada")

resumen_rows = []
def add_row(label, col, tipo, etapa, fmt=".2f", mult=1.0):
    if col not in df.columns:
        return
    s_v = sobre[col].mean() * mult
    b_v = bajo[col].mean() * mult
    resumen_rows.append({
        "Variable": label, "Tipo": tipo, "Etapa": etapa,
        "Sobre meta": f"{s_v:{fmt}}",
        "Bajo meta": f"{b_v:{fmt}}",
        "Δ": f"{(b_v - s_v):+{fmt}}",
    })

# E1 control
add_row("Flujo Bref [m³/h]", "Flujo Bref", "Control", "E1", ".1f")
add_row("Flujo SCS [m³/h]", "Flujo SCS", "Control", "E1", ".1f")
add_row("Flujo LM [m³/h]", "Flujo LM ", "Control", "E1", ".1f")
add_row("Flujo PRS [m³/h]", "Flujo PRS", "Control", "E1", ".1f")
add_row("Flujo SX1-2 [m³/h]", "Flujo SX1-2", "Control", "E1", ".1f")
add_row("Flujo SX3 [m³/h]", "Flujo SX3 ", "Control", "E1", ".1f")
add_row("Flujo NF [m³/h]", "Flujo NF", "Control", "E1", ".2f")
add_row("Flujo TK605 [m³/h]", "Flujo 605 rec.", "Control", "E1", ".2f")
add_row("T AFT [°C]", "T AFT [°C]", "Control", "E1", ".2f")
add_row("Li/Mg Bref [-]", "Li/Mg Bref", "Control", "E1", ".3f")
add_row("Tres real [h]", "Tres (h)", "Control", "E1", ".3f")
# E1 desempeño
add_row("Li AFT [%]", "Li AFT [%]", "Desempeño", "E1", ".3f")
add_row("Li Bfil [%]", "Li Bfil [%]", "Desempeño", "E1", ".3f")
add_row("Mg Bfil [%]", "Mg Bfil [%]", "Desempeño", "E1", ".3f")
add_row("Mg/Li FT [-]", "Mg/Li FT", "Desempeño", "E1", ".3f")
add_row("Pérdida E1 [%]", "Perdida E1 [%]", "Desempeño", "E1", ".2f", 100.0)
# E2
add_row("pH BP", "pH BP ", "Control", "E2", ".2f")
add_row("Mg BP [ppm]", "Mg BP [ppm]", "Desempeño", "E2", ".2f")
add_row("Mg/Li FP [-]", "Mg/Li FP", "Desempeño", "E2", ".2f")
add_row("Pérdida E2 [%]", "Perdida E2 [%]", "Desempeño", "E2", ".2f", 100.0)
# E3
add_row("Carga E3 [m³/h]", "Carga E3 [m3/h]", "Control", "E3", ".1f")
add_row("Li LM [%]", "Li LM [%]", "Desempeño", "E3", ".4f")
add_row("Na LM [%]", "Na LM [%]", "Desempeño", "E3", ".2f")
add_row("Pérdida E3 [%]", "Perdida E3 [%]", "Desempeño", "E3", ".2f", 100.0)
# Global
add_row("Factor producción", "Factor Producción", "Desempeño", "Global", ".3f")
add_row("Rendimiento AQ [%]", "Rendimiento AQ", "Desempeño", "Global", ".2f", 100.0)

df_resumen = pd.DataFrame(resumen_rows)
st.dataframe(df_resumen, hide_index=True, use_container_width=True, height=400)

# Descargas
st.markdown("### 📥 Descargas")
c1, c2 = st.columns(2)
with c1:
    csv_full = df.to_csv(index=False).encode("utf-8")
    st.download_button("Datos procesados (CSV)", data=csv_full,
                       file_name="datos_procesados.csv", mime="text/csv")
with c2:
    csv_resumen = df_resumen.to_csv(index=False).encode("utf-8")
    st.download_button("Tabla resumen (CSV)", data=csv_resumen,
                       file_name="tabla_resumen.csv", mime="text/csv")

st.markdown("---")
st.markdown("<div style='text-align:center; color:#7A6AB8; font-size:0.85em;'>"
            "Novandino Litio © 2026 — #SomosLitioSomosFuturo</div>",
            unsafe_allow_html=True)
