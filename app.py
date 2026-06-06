import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# ─────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Optimización de Cartera de Proveedores",
    page_icon="🏭",
    layout="wide",
)

# ─────────────────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────────────────
st.markdown("""
<style>
    html, body, [class*="css"] { font-family: 'Segoe UI', sans-serif; }
    .main-title { font-size: 2.1rem; font-weight: 700; color: #001d3d; margin-bottom: 0.1rem; }
    .subtitle   { font-size: 1.05rem; color: #555; margin-bottom: 0.5rem; }
    .section-header {
        font-size: 1.2rem; font-weight: 700; color: #001d3d;
        border-left: 5px solid #f15b2b;
        padding: 4px 0 4px 12px; margin: 2rem 0 0.8rem 0;
        background: linear-gradient(90deg, #f8f9fa 0%, transparent 100%);
    }
    .box-info    { background:#eef4fb; border-radius:8px; padding:11px 16px; font-size:0.88rem; color:#333; margin:0.5rem 0; border-left:4px solid #0d7377; }
    .box-warning { background:#fff8e1; border-radius:8px; padding:11px 16px; font-size:0.88rem; color:#6d4c00; margin:0.4rem 0; border-left:4px solid #f15b2b; }
    .box-ok      { background:#e8f5e9; border-radius:8px; padding:11px 16px; font-size:0.88rem; color:#1b5e20; margin:0.4rem 0; border-left:4px solid #43a047; }
    .box-danger  { background:#fdecea; border-radius:8px; padding:11px 16px; font-size:0.88rem; color:#7f1d1d; margin:0.4rem 0; border-left:4px solid #e53935; }
    div[data-testid="metric-container"] { background:#f8f9fa; border-radius:8px; padding:10px 16px; border:1px solid #e0e0e0; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
# FUNCIONES CACHEADAS
# ─────────────────────────────────────────────────────────

@st.cache_data(show_spinner="📂 Leyendo archivo Excel...")
def cargar_datos(archivo_bytes, hoja_d, hoja_g):
    """Carga y valida ambas hojas. Se ejecuta solo cuando cambia el archivo o los nombres de hoja."""
    import io
    xl = pd.ExcelFile(io.BytesIO(archivo_bytes))
    hojas = xl.sheet_names

    if hoja_d not in hojas:
        return None, None, [], hojas, f"No se encontró la hoja '{hoja_d}'. Disponibles: {hojas}"

    df_raw = pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=hoja_d)
    df_raw.rename(columns={df_raw.columns[0]: "Mes"}, inplace=True)
    proveedores = [c for c in df_raw.columns[1:] if not str(c).startswith("Unnamed")]

    if len(proveedores) < 2:
        return None, None, [], hojas, "Se necesitan al menos 2 proveedores."

    for col in proveedores:
        df_raw[col] = pd.to_numeric(df_raw[col], errors="coerce")
    df_raw.dropna(subset=proveedores, inplace=True)

    if len(df_raw) < 3:
        return None, None, [], hojas, "Se requieren al menos 3 observaciones históricas."

    score_df = df_raw.set_index("Mes")[proveedores]

    gasto_df = None
    if hoja_g in hojas:
        try:
            df_g = pd.read_excel(io.BytesIO(archivo_bytes), sheet_name=hoja_g)
            df_g.rename(columns={df_g.columns[0]: "Mes"}, inplace=True)
            provs_g = [c for c in proveedores if c in df_g.columns]
            for col in provs_g:
                df_g[col] = pd.to_numeric(df_g[col], errors="coerce")
            df_g.dropna(subset=provs_g, inplace=True)
            gasto_df = df_g.set_index("Mes")[provs_g]
        except Exception:
            gasto_df = None

    return score_df, gasto_df, proveedores, hojas, None


@st.cache_data(show_spinner="🎲 Ejecutando simulación Monte Carlo...")
def ejecutar_monte_carlo(mu_tuple, cov_flat, n_prov, n_sim, proveedores_tuple):
    """Monte Carlo cacheado — solo se recalcula si cambian mu, cov o n_sim."""
    mu_arr  = np.array(mu_tuple)
    cov_mat = np.array(cov_flat).reshape(n_prov, n_prov)
    proveedores = list(proveedores_tuple)

    np.random.seed(42)
    resultados = np.zeros((n_sim, 3 + n_prov))

    for i in range(n_sim):
        w  = np.random.random(n_prov)
        w /= w.sum()
        s  = float(np.dot(w, mu_arr))
        r  = float(np.sqrt(w @ cov_mat @ w))
        resultados[i, 0] = s
        resultados[i, 1] = r
        resultados[i, 2] = s / r if r > 0 else 0.0
        resultados[i, 3:] = w

    cols   = ["Score", "Riesgo", "Sharpe"] + proveedores
    df_sim = pd.DataFrame(resultados, columns=cols)
    opt    = df_sim.loc[df_sim["Sharpe"].idxmax()]
    minr   = df_sim.loc[df_sim["Riesgo"].idxmin()]
    return df_sim, opt, minr


# ─────────────────────────────────────────────────────────
# ENCABEZADO
# ─────────────────────────────────────────────────────────
st.markdown('<div class="main-title">🏭 Optimización de Cartera de Proveedores</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="subtitle">Lógica de Markowitz aplicada al abastecimiento · '
    'Maximiza desempeño · Minimiza concentración · Reduce vulnerabilidad operativa</div>',
    unsafe_allow_html=True
)

with st.expander("📖 ¿Cómo funciona esta aplicación?", expanded=False):
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown("""
        **El problema clásico de abastecimiento:**

        Muchas empresas concentran sus compras en pocos proveedores. Si ese
        proveedor falla en precio, calidad, entrega o disponibilidad, toda
        la operación se ve afectada.

        **La solución de Markowitz:**

        Distribuir el gasto entre varias alternativas para maximizar el
        desempeño esperado y minimizar el riesgo operativo.
        """)
    with col_e2:
        st.markdown("""
        | Concepto Markowitz | En esta app |
        |---|---|
        | Activo | Proveedor |
        | Retorno esperado | Score de desempeño promedio (0–1) |
        | Riesgo | Volatilidad del score (σ) |
        | Correlación | Sincronía de problemas entre proveedores |
        | Cartera óptima | % óptimo de gasto por proveedor |
        | Índice Sharpe | Score esperado / Riesgo del portafolio |
        """)
    st.info("**Score** = Calidad×0.30 + Puntualidad×0.25 + Precio×0.20 + Disponibilidad×0.15 + Logística×0.10")

# ─────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    archivo = st.file_uploader("📂 Archivo Excel (.xlsx)", type=["xlsx"])

    st.markdown("---")
    gasto_total = st.number_input(
        "💰 Gasto mensual total en proveedores ($)",
        min_value=10_000, max_value=100_000_000,
        value=7_000_000, step=100_000,
        help="Monto total que se redistribuirá según la cartera óptima"
    )
    umbral = st.slider(
        "⚠️ Umbral de alerta por concentración (%)",
        min_value=15, max_value=60, value=35, step=5,
        help="Proveedor que supere este % se marca como riesgo"
    )
    n_sim = st.select_slider(
        "🎲 Simulaciones Monte Carlo",
        options=[1_000, 3_000, 5_000, 10_000, 20_000],
        value=5_000,
        help="Más simulaciones = mayor precisión, pero más tiempo de cálculo"
    )

    st.markdown("---")
    hoja_d = st.text_input(
        "📋 Hoja de desempeño",
        value="Desempeño Proveedores",
        help="Nombre exacto de la hoja con scores históricos (0–1)"
    )
    hoja_g = st.text_input(
        "📋 Hoja de gasto",
        value="Gasto Mensual",
        help="Nombre exacto de la hoja con gasto mensual por proveedor"
    )
    st.markdown("---")
    st.caption("Lógica de Markowitz aplicada a proveedores\n\nUniversidad Panamericana · IA para el Análisis Financiero")

# ─────────────────────────────────────────────────────────
# PANTALLA INICIAL
# ─────────────────────────────────────────────────────────
if archivo is None:
    st.info("👈 Carga tu archivo Excel desde el panel lateral para comenzar.")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.markdown("**📋 Hoja: Desempeño Proveedores** *(requerida)*")
        st.markdown("""
| Mes | Prov. A | Prov. B | Prov. C |
|---|---|---|---|
| Ene-2024 | 0.825 | 0.761 | 0.602 |
| Feb-2024 | 0.810 | 0.774 | 0.615 |

> Score 0–1. Fila 1 = encabezados, sin filas de título.
        """)
    with col_f2:
        st.markdown("**📋 Hoja: Gasto Mensual** *(recomendada)*")
        st.markdown("""
| Mes | Prov. A | Prov. B | Prov. C |
|---|---|---|---|
| Ene-2024 | 3,784 | 996 | 470 |
| Feb-2024 | 3,650 | 1,020 | 490 |

> Gasto real en $. Fila 1 = encabezados.
        """)
    st.stop()

# ─────────────────────────────────────────────────────────
# CARGA DE DATOS (cacheada por archivo + nombres de hoja)
# ─────────────────────────────────────────────────────────
archivo_bytes = archivo.read()

score_df, gasto_df, proveedores, hojas_disponibles, error_carga = cargar_datos(
    archivo_bytes, hoja_d, hoja_g
)

if error_carga:
    st.error(f"❌ {error_carga}")
    st.stop()

tiene_gasto = gasto_df is not None
if not tiene_gasto:
    st.sidebar.warning(f"Hoja '{hoja_g}' no encontrada. Diagnóstico de concentración no disponible.")

# ─────────────────────────────────────────────────────────
# PRE-CÁLCULOS (estadísticos — instantáneos)
# ─────────────────────────────────────────────────────────
mu      = score_df.mean()
sigma   = score_df.std()
cov_mat = score_df.cov().values
mu_arr  = mu.values
n_prov  = len(proveedores)

if tiene_gasto:
    gasto_prom       = gasto_df.mean().reindex(proveedores).fillna(0)
    gasto_total_real = gasto_prom.sum()
    participacion    = (gasto_prom / gasto_total_real * 100)
    pesos_actuales   = (gasto_prom / gasto_total_real).values

# ─────────────────────────────────────────────────────────
# SECCIÓN 0 — DIAGNÓSTICO DE CONCENTRACIÓN
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">0 · Diagnóstico de Concentración Actual</div>', unsafe_allow_html=True)

if tiene_gasto:
    col_d1, col_d2 = st.columns([1.1, 1])
    with col_d1:
        part_sorted = participacion.sort_values(ascending=False)
        df_conc = pd.DataFrame({
            "Proveedor":         part_sorted.index,
            "Gasto Prom. ($)":   gasto_prom[part_sorted.index].round(0),
            "Participación (%)": part_sorted.values.round(1),
        }).reset_index(drop=True)

        def highlight_conc(val):
            return "background-color: #fdecea; font-weight: bold" if val > umbral else ""

        st.dataframe(
            df_conc.style
                .map(highlight_conc, subset=["Participación (%)"])
                .format({"Gasto Prom. ($)": "${:,.0f}", "Participación (%)": "{:.1f}%"}),
            use_container_width=True, hide_index=True
        )

        en_riesgo = participacion[participacion > umbral]
        if len(en_riesgo) > 0:
            for prov, pct in en_riesgo.sort_values(ascending=False).items():
                st.markdown(
                    f'<div class="box-warning">⚠️ <b>{prov}</b> concentra el <b>{pct:.1f}%</b> '
                    f'del gasto — supera el umbral de {umbral}%.</div>',
                    unsafe_allow_html=True
                )
        else:
            st.markdown(
                f'<div class="box-ok">✅ Ningún proveedor supera el umbral de concentración del {umbral}%.</div>',
                unsafe_allow_html=True
            )

    with col_d2:
        colors_dona = ["#e53935" if participacion[p] > umbral else "#001d3d" for p in part_sorted.index]
        fig_dona_act = go.Figure(go.Pie(
            labels=part_sorted.index, values=part_sorted.values,
            hole=0.50, textinfo="label+percent",
            marker=dict(colors=colors_dona),
            hovertemplate="<b>%{label}</b><br>%{percent}<extra></extra>"
        ))
        fig_dona_act.update_layout(
            title="Distribución actual del gasto",
            height=340, showlegend=False,
            margin=dict(t=50, b=20, l=20, r=20)
        )
        st.plotly_chart(fig_dona_act, use_container_width=True)
else:
    st.info("ℹ️ Carga la hoja de gasto mensual para ver el diagnóstico de concentración.")

# ─────────────────────────────────────────────────────────
# SECCIÓN 1 — DESEMPEÑO HISTÓRICO
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">1 · Desempeño Histórico por Proveedor</div>', unsafe_allow_html=True)

col_t1, col_t2 = st.columns([2, 1])
with col_t1:
    st.dataframe(
        score_df.style
            .format("{:.3f}")
            .background_gradient(cmap="RdYlGn", axis=None, vmin=0.4, vmax=1.0),
        use_container_width=True
    )
with col_t2:
    resumen = pd.DataFrame({
        "Score Prom.": mu.round(3),
        "Riesgo (σ)":  sigma.round(4),
        "Mín.":        score_df.min().round(3),
        "Máx.":        score_df.max().round(3),
    })
    st.dataframe(
        resumen.style
            .format("{:.3f}")
            .background_gradient(subset=["Score Prom."], cmap="Greens")
            .background_gradient(subset=["Riesgo (σ)"],  cmap="Reds_r"),
        use_container_width=True
    )

# Evolución histórica
palette = px.colors.qualitative.D3
fig_hist = go.Figure()
for i, prov in enumerate(proveedores):
    fig_hist.add_trace(go.Scatter(
        x=score_df.index.astype(str), y=score_df[prov],
        mode="lines+markers", name=prov,
        line=dict(width=2, color=palette[i % len(palette)]),
        marker=dict(size=5)
    ))
fig_hist.add_hline(y=0.70, line_dash="dot", line_color="#f15b2b", line_width=1.5,
    annotation_text="Umbral mínimo 0.70", annotation_position="bottom right",
    annotation_font=dict(color="#f15b2b", size=10))
fig_hist.update_layout(
    title="Evolución del Score de Desempeño por Proveedor",
    xaxis_title="Mes", yaxis_title="Score (0–1)",
    yaxis=dict(range=[0.35, 1.02]), height=370,
    hovermode="x unified", legend=dict(orientation="h", y=-0.28),
    margin=dict(t=50, b=20)
)
st.plotly_chart(fig_hist, use_container_width=True)

# Mapa riesgo–desempeño
fig_mapa = go.Figure()
med_sigma = sigma.median()
med_mu    = mu.median()
for i, prov in enumerate(proveedores):
    fig_mapa.add_trace(go.Scatter(
        x=[sigma[prov]], y=[mu[prov]],
        mode="markers+text", text=[prov], textposition="top center",
        marker=dict(size=18, color=palette[i % len(palette)], line=dict(color="white", width=1.5)),
        name=prov,
        hovertemplate=f"<b>{prov}</b><br>Score: {mu[prov]:.3f}<br>σ: {sigma[prov]:.4f}<extra></extra>"
    ))
fig_mapa.add_vline(x=med_sigma, line_dash="dot", line_color="#aaa", line_width=1)
fig_mapa.add_hline(y=med_mu,    line_dash="dot", line_color="#aaa", line_width=1)
fig_mapa.update_layout(
    title="Mapa Riesgo–Desempeño  (esquina superior izquierda = ideal ⭐)",
    xaxis_title="Riesgo — Volatilidad del Score (σ)",
    yaxis_title="Desempeño Esperado — Score Promedio",
    height=400, showlegend=False, margin=dict(t=50, b=30)
)
st.plotly_chart(fig_mapa, use_container_width=True)
st.markdown(
    '<div class="box-info">💡 <b>Esquina superior izquierda</b> = alto desempeño + baja volatilidad = proveedor ideal. '
    '<b>Esquina inferior derecha</b> = bajo desempeño + alta variabilidad = candidato a reducir o reemplazar.</div>',
    unsafe_allow_html=True
)

# ─────────────────────────────────────────────────────────
# SECCIÓN 2 — CORRELACIÓN Y COVARIANZA
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">2 · Correlación y Covarianza entre Proveedores</div>', unsafe_allow_html=True)

col_c1, col_c2 = st.columns(2)
mat_corr = score_df.corr().round(3)
mat_cov  = score_df.cov().round(5)
with col_c1:
    fig_corr = px.imshow(mat_corr, text_auto=True, color_continuous_scale="RdBu_r",
        zmin=-1, zmax=1, aspect="auto", title="Correlación entre scores de desempeño")
    fig_corr.update_layout(height=380, margin=dict(t=50, b=10))
    st.plotly_chart(fig_corr, use_container_width=True)
with col_c2:
    fig_cov = px.imshow(mat_cov, text_auto=True, color_continuous_scale="Blues",
        aspect="auto", title="Covarianza entre proveedores")
    fig_cov.update_layout(height=380, margin=dict(t=50, b=10))
    st.plotly_chart(fig_cov, use_container_width=True)
st.markdown(
    '<div class="box-info">💡 <b>Correlación baja</b> entre dos proveedores significa que cuando uno falla, '
    'el otro no necesariamente también falla — eso es diversificación real.</div>',
    unsafe_allow_html=True
)

# ─────────────────────────────────────────────────────────
# SECCIÓN 3 — MONTE CARLO (cacheado por n_sim + datos)
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">3 · Simulación Monte Carlo — Frontera Eficiente</div>', unsafe_allow_html=True)

# Pasamos mu y cov como tuplas para que sean hasheables por el caché
df_sim, cartera_opt, cartera_minr = ejecutar_monte_carlo(
    tuple(mu_arr),
    tuple(cov_mat.flatten()),
    n_prov,
    n_sim,
    tuple(proveedores)
)

fig_front = go.Figure()
fig_front.add_trace(go.Scatter(
    x=df_sim["Riesgo"], y=df_sim["Score"], mode="markers",
    marker=dict(color=df_sim["Sharpe"], colorscale="Viridis", size=3, opacity=0.55,
                colorbar=dict(title="Sharpe", thickness=12, len=0.7)),
    name="Carteras simuladas",
    hovertemplate="Riesgo: %{x:.5f}<br>Score: %{y:.4f}<extra></extra>"
))
fig_front.add_trace(go.Scatter(
    x=[cartera_minr["Riesgo"]], y=[cartera_minr["Score"]], mode="markers",
    marker=dict(symbol="diamond", size=18, color="#1976d2", line=dict(color="white", width=1.5)),
    name="🔵 Mínimo Riesgo",
    hovertemplate=f"Score: {cartera_minr['Score']:.4f}<br>Riesgo: {cartera_minr['Riesgo']:.5f}<extra></extra>"
))
fig_front.add_trace(go.Scatter(
    x=[cartera_opt["Riesgo"]], y=[cartera_opt["Score"]], mode="markers",
    marker=dict(symbol="star", size=24, color="#e53935", line=dict(color="white", width=1.5)),
    name="⭐ Cartera Óptima",
    hovertemplate=f"Score: {cartera_opt['Score']:.4f}<br>Riesgo: {cartera_opt['Riesgo']:.5f}<br>Sharpe: {cartera_opt['Sharpe']:.2f}<extra></extra>"
))
if tiene_gasto:
    s_act_front = float(np.dot(pesos_actuales, mu_arr))
    r_act_front = float(np.sqrt(pesos_actuales @ cov_mat @ pesos_actuales))
    fig_front.add_trace(go.Scatter(
        x=[r_act_front], y=[s_act_front], mode="markers",
        marker=dict(symbol="square", size=16, color="#f15b2b", line=dict(color="white", width=1.5)),
        name="🟠 Distribución Actual",
        hovertemplate=f"Score actual: {s_act_front:.4f}<br>Riesgo actual: {r_act_front:.5f}<extra></extra>"
    ))
fig_front.update_layout(
    title=f"Frontera Eficiente de Abastecimiento  ({n_sim:,} simulaciones)",
    xaxis_title="Riesgo del Portafolio — σ del Score",
    yaxis_title="Score de Desempeño Esperado",
    height=500, hovermode="closest",
    legend=dict(orientation="h", y=-0.18), margin=dict(t=55, b=20)
)
st.plotly_chart(fig_front, use_container_width=True)

col_le1, col_le2 = st.columns(2)
with col_le1:
    st.markdown('<div class="box-info">⭐ <b>Cartera Óptima (estrella roja):</b> maximiza el score por unidad de riesgo.</div>', unsafe_allow_html=True)
with col_le2:
    st.markdown('<div class="box-info">🔵 <b>Mínimo Riesgo (rombo azul):</b> minimiza la variabilidad operativa.</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────
# SECCIÓN 4 — ASIGNACIÓN ÓPTIMA
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">4 · Asignación Óptima del Gasto</div>', unsafe_allow_html=True)

m1, m2, m3 = st.columns(3)
m1.metric("📈 Score Esperado",         f"{cartera_opt['Score']:.3f}")
m2.metric("⚡ Riesgo del Portafolio (σ)", f"{cartera_opt['Riesgo']:.5f}")
m3.metric("🏆 Índice Sharpe",           f"{cartera_opt['Sharpe']:.1f}")

st.markdown("---")

pesos_opt = {p: cartera_opt[p] for p in proveedores}
df_asig = pd.DataFrame({
    "Proveedor":         proveedores,
    "Peso Óptimo (%)":   [pesos_opt[p] * 100        for p in proveedores],
    "Gasto Asignado ($)":[pesos_opt[p] * gasto_total for p in proveedores],
    "Score Prom.":       [mu[p]   for p in proveedores],
    "Riesgo (σ)":        [sigma[p] for p in proveedores],
}).sort_values("Peso Óptimo (%)", ascending=False).reset_index(drop=True)

if tiene_gasto:
    df_asig["% Actual"]    = df_asig["Proveedor"].map(participacion).round(1)
    df_asig["Cambio (pp)"] = (df_asig["Peso Óptimo (%)"] - df_asig["% Actual"]).round(1)

fmt = {
    "Peso Óptimo (%)":    "{:.1f}%",
    "Gasto Asignado ($)": "${:,.0f}",
    "Score Prom.":        "{:.3f}",
    "Riesgo (σ)":         "{:.4f}",
}
if tiene_gasto:
    fmt["% Actual"]    = "{:.1f}%"
    fmt["Cambio (pp)"] = "{:+.1f}"

col_a1, col_a2 = st.columns([1.1, 1])
with col_a1:
    st.dataframe(
        df_asig.style
            .format(fmt)
            .background_gradient(subset=["Peso Óptimo (%)"], cmap="Greens")
            .background_gradient(subset=["Score Prom."],     cmap="Blues"),
        use_container_width=True, hide_index=True
    )
with col_a2:
    fig_dona = go.Figure(go.Pie(
        labels=df_asig["Proveedor"],
        values=df_asig["Peso Óptimo (%)"],
        hole=0.50, textinfo="label+percent",
        marker=dict(colors=px.colors.qualitative.D3),
        hovertemplate="%{label}: %{percent}<extra></extra>"
    ))
    fig_dona.update_layout(
        title=f"Distribución óptima — Total: ${gasto_total:,.0f}",
        height=360, showlegend=False,
        margin=dict(t=50, b=10, l=10, r=10)
    )
    st.plotly_chart(fig_dona, use_container_width=True)

# ─────────────────────────────────────────────────────────
# SECCIÓN 5 — COMPARATIVA ACTUAL VS. ÓPTIMA
# ─────────────────────────────────────────────────────────
if tiene_gasto:
    st.markdown('<div class="section-header">5 · Comparativa: Distribución Actual vs. Óptima</div>', unsafe_allow_html=True)

    s_act  = float(np.dot(pesos_actuales, mu_arr))
    r_act  = float(np.sqrt(pesos_actuales @ cov_mat @ pesos_actuales))
    sh_act = s_act / r_act if r_act > 0 else 0

    df_comp = pd.DataFrame({
        "Estrategia":     ["Distribución Actual", "Cartera Mínimo Riesgo", "Cartera Óptima (Máx. Eficiencia)"],
        "Score Esperado": [s_act, cartera_minr["Score"], cartera_opt["Score"]],
        "Riesgo (σ)":     [r_act, cartera_minr["Riesgo"], cartera_opt["Riesgo"]],
        "Índice Sharpe":  [sh_act, cartera_minr["Sharpe"], cartera_opt["Sharpe"]],
    })

    mejora_score  = (cartera_opt["Score"] - s_act) / s_act * 100
    reduce_riesgo = (r_act - cartera_opt["Riesgo"]) / r_act * 100

    col_c1, col_c2 = st.columns([1.6, 1])
    with col_c1:
        st.dataframe(
            df_comp.style
                .format({"Score Esperado": "{:.4f}", "Riesgo (σ)": "{:.5f}", "Índice Sharpe": "{:.1f}"})
                .highlight_max(subset=["Score Esperado", "Índice Sharpe"], color="#d4edda")
                .highlight_min(subset=["Riesgo (σ)"], color="#d4edda"),
            use_container_width=True, hide_index=True
        )
    with col_c2:
        st.metric("📈 Mejora en Score vs. actual",     f"{mejora_score:+.2f}%")
        st.metric("🛡️ Reducción de riesgo vs. actual", f"{reduce_riesgo:+.1f}%")

    # Barras comparativas
    fig_barras = go.Figure()
    fig_barras.add_trace(go.Bar(
        name="% Actual", x=proveedores,
        y=[participacion.get(p, 0) for p in proveedores],
        marker_color="#001d3d",
        text=[f"{participacion.get(p,0):.1f}%" for p in proveedores],
        textposition="outside"
    ))
    fig_barras.add_trace(go.Bar(
        name="% Óptimo", x=proveedores,
        y=[pesos_opt[p] * 100 for p in proveedores],
        marker_color="#f15b2b",
        text=[f"{pesos_opt[p]*100:.1f}%" for p in proveedores],
        textposition="outside"
    ))
    fig_barras.add_hline(y=umbral, line_dash="dot", line_color="red", line_width=1.5,
        annotation_text=f"Umbral de alerta ({umbral}%)",
        annotation_position="top right",
        annotation_font=dict(color="red", size=10))
    fig_barras.update_layout(
        title="Distribución actual vs. óptima por proveedor",
        xaxis_title="Proveedor", yaxis_title="% del Gasto Total",
        barmode="group", height=420,
        legend=dict(orientation="h", y=-0.22), margin=dict(t=55, b=20)
    )
    st.plotly_chart(fig_barras, use_container_width=True)

# ─────────────────────────────────────────────────────────
# SECCIÓN 6 — SEÑALES DE ACCIÓN
# ─────────────────────────────────────────────────────────
st.markdown('<div class="section-header">6 · Señales de Acción por Proveedor</div>', unsafe_allow_html=True)
st.markdown("Clasificación basada en score promedio, volatilidad y peso óptimo del modelo:")
st.markdown("")

for prov in proveedores:
    sc   = mu[prov]
    vol  = sigma[prov]
    peso = pesos_opt[prov] * 100

    cambio_txt = ""
    if tiene_gasto:
        delta = peso - participacion.get(prov, 0)
        cambio_txt = f"  |  Cambio sugerido: **{'+'if delta>0 else ''}{delta:.1f} pp**"

    if sc >= 0.75 and vol < 0.04:
        cat    = "🟢 CONSOLIDAR"
        cls    = "box-ok"
        accion = f"Proveedor altamente confiable y estable. Mantener o incrementar participación hasta **{peso:.1f}%**."
    elif sc >= 0.75 and vol >= 0.04:
        cat    = "🟡 MONITOREAR"
        cls    = "box-warning"
        accion = f"Buen desempeño pero variable. Implementar SLAs y seguimiento mensual. Peso óptimo: **{peso:.1f}%**."
    elif sc >= 0.65:
        cat    = "🔵 DESARROLLAR"
        cls    = "box-info"
        accion = f"Desempeño aceptable con potencial. Invertir en la relación comercial. Peso óptimo: **{peso:.1f}%**."
    else:
        cat    = "🔴 REDUCIR / REEMPLAZAR"
        cls    = "box-danger"
        accion = f"Desempeño bajo el umbral aceptable. Reducir participación o buscar alternativas. Peso óptimo: **{peso:.1f}%**."

    st.markdown(
        f'<div class="{cls}">'
        f'<b>{prov}</b> &nbsp;·&nbsp; {cat} &nbsp;·&nbsp; '
        f'Score: <b>{sc:.3f}</b> &nbsp;·&nbsp; σ: <b>{vol:.4f}</b>{cambio_txt}<br>'
        f'<small>{accion}</small>'
        f'</div>',
        unsafe_allow_html=True
    )

st.markdown("---")
st.caption(
    "Aplicación desarrollada con fines didácticos | "
    "Lógica de Markowitz aplicada a la gestión de proveedores | "
    "Universidad Panamericana · IA para el Análisis Financiero"
)
