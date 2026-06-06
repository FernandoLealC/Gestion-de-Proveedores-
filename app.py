import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px

# ─────────────────────────────────────────────
# CONFIGURACIÓN DE PÁGINA
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Optimización de Cartera de Proveedores",
    page_icon="🏭",
    layout="wide",
)

# ─────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    .main-title {
        font-size: 2.2rem; font-weight: 700;
        color: #001d3d; margin-bottom: 0.2rem;
    }
    .subtitle {
        font-size: 1.05rem; color: #555; margin-bottom: 1.5rem;
    }
    .section-header {
        font-size: 1.25rem; font-weight: 700; color: #001d3d;
        border-left: 5px solid #f15b2b; padding-left: 12px;
        margin: 2rem 0 0.8rem 0;
    }
    .alert-box {
        background: #fff3cd; border-left: 5px solid #f15b2b;
        border-radius: 6px; padding: 12px 16px;
        font-size: 0.9rem; color: #333; margin: 0.4rem 0;
    }
    .ok-box {
        background: #d4edda; border-left: 5px solid #28a745;
        border-radius: 6px; padding: 12px 16px;
        font-size: 0.9rem; color: #155724; margin: 0.4rem 0;
    }
    .insight-box {
        background: #eef4fb; border-radius: 8px;
        padding: 12px 16px; font-size: 0.9rem;
        color: #333; margin: 0.4rem 0;
    }
    .formula-box {
        background: #1e1e1e; border-radius: 8px;
        padding: 14px 18px; font-size: 0.95rem;
        color: #a8d8a8; font-family: monospace;
        margin: 0.6rem 0; letter-spacing: 0.02em;
    }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ENCABEZADO
# ─────────────────────────────────────────────
st.markdown('<div class="main-title">🏭 Optimización de Cartera de Proveedores</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Lógica de Markowitz aplicada al abastecimiento — Reduce concentración, maximiza desempeño, minimiza vulnerabilidad operativa</div>', unsafe_allow_html=True)

with st.expander("ℹ️ ¿Cómo funciona esta aplicación?"):
    col_e1, col_e2 = st.columns(2)
    with col_e1:
        st.markdown("""
        **El problema que resuelve:**
        Muchas empresas concentran sus compras en pocos proveedores.
        Si ese proveedor falla en precio, calidad o entrega, toda la operación se ve afectada.

        **La solución de Markowitz:**
        Tratar cada proveedor como un *activo* de un portafolio financiero y encontrar
        la combinación de gasto que **maximiza el desempeño esperado** con el **menor riesgo posible**.
        """)
    with col_e2:
        st.markdown("""
        | Concepto Markowitz | En proveedores |
        |---|---|
        | Activo | Proveedor |
        | Retorno esperado | Score de desempeño promedio |
        | Riesgo | Volatilidad del score (σ) |
        | Correlación | Sincronía entre fallas de 2 proveedores |
        | Cartera óptima | % de gasto óptimo por proveedor |
        """)
    st.markdown('<div class="formula-box">Score = Calidad×0.30 + Puntualidad×0.25 + Precio×0.20 + Disponibilidad×0.15 + Logística×0.10</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Configuración")
    archivo = st.file_uploader("📂 Cargar archivo Excel (.xlsx)", type=["xlsx"])
    st.markdown("---")

    gasto_total = st.number_input(
        "💰 Gasto total mensual en proveedores ($)",
        min_value=10_000, max_value=100_000_000,
        value=7_071_000, step=10_000,
        help="Monto que se redistribuirá según la cartera óptima"
    )
    umbral_conc = st.slider(
        "⚠️ Umbral de alerta por concentración (%)",
        min_value=20, max_value=60, value=35, step=5,
        help="Proveedores que superen este % del gasto serán marcados como riesgo"
    )
    n_sim = st.select_slider(
        "🎲 Simulaciones Monte Carlo",
        options=[1000, 3000, 5000, 10000, 20000], value=5000
    )
    st.markdown("---")
    hoja_desemp = st.text_input(
        "📋 Hoja de desempeño",
        value="Desempeño Proveedores"
    )
    hoja_gasto_nombre = st.text_input(
        "📋 Hoja de gasto mensual",
        value="Gasto Mensual"
    )
    st.markdown("---")
    st.caption("Lógica de Markowitz aplicada a proveedores\nUniversidad Panamericana · IA para el Análisis Financiero")

# ─────────────────────────────────────────────
# ESTADO INICIAL — sin archivo
# ─────────────────────────────────────────────
if archivo is None:
    st.info("👈 Carga tu archivo Excel desde el panel lateral para comenzar.")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("""
        **Hoja principal — Desempeño Proveedores:**

        | Mes | Prov. A | Prov. B | Prov. C |
        |---|---|---|---|
        | Ene-2024 | 0.85 | 0.72 | 0.60 |
        | Feb-2024 | 0.80 | 0.74 | 0.62 |

        > Score entre 0 y 1 por proveedor por mes
        """)
    with c2:
        st.markdown("""
        **Hoja principal — Gasto Mensual:**

        | Mes | Prov. A | Prov. B | Prov. C |
        |---|---|---|---|
        | Ene-2024 | 450,000 | 280,000 | 150,000 |
        | Feb-2024 | 430,000 | 295,000 | 160,000 |

        > Gasto real en $ por proveedor por mes
        """)
    st.stop()

# ─────────────────────────────────────────────
# CARGA Y VALIDACIÓN
# ─────────────────────────────────────────────
try:
    xl = pd.ExcelFile(archivo)
    hojas = xl.sheet_names

    if hoja_desemp not in hojas:
        st.error(f"❌ No se encontró la hoja **'{hoja_desemp}'**. Hojas disponibles: {hojas}")
        st.stop()

    df_d = pd.read_excel(archivo, sheet_name=hoja_desemp)
    df_d.rename(columns={df_d.columns[0]: "Mes"}, inplace=True)
    proveedores = df_d.columns[1:].tolist()

    if len(proveedores) < 2:
        st.error("❌ Se requieren al menos 2 proveedores.")
        st.stop()

    for col in proveedores:
        df_d[col] = pd.to_numeric(df_d[col], errors="coerce")
    df_d.dropna(inplace=True)

    if len(df_d) < 3:
        st.error("❌ Se requieren al menos 3 observaciones históricas válidas.")
        st.stop()

    score_df = df_d.set_index("Mes")[proveedores]

    # Hoja de gasto (opcional)
    tiene_gasto = hoja_gasto_nombre in hojas
    if tiene_gasto:
        df_g = pd.read_excel(archivo, sheet_name=hoja_gasto_nombre)
        df_g.rename(columns={df_g.columns[0]: "Mes"}, inplace=True)
        cols_gasto = [c for c in proveedores if c in df_g.columns]
        for col in cols_gasto:
            df_g[col] = pd.to_numeric(df_g[col], errors="coerce")
        df_g.dropna(inplace=True)
        gasto_df = df_g.set_index("Mes")[cols_gasto]
    else:
        st.warning(f"⚠️ No se encontró la hoja **'{hoja_gasto_nombre}'**. El diagnóstico de concentración no estará disponible.")
        tiene_gasto = False

except Exception as e:
    st.error(f"❌ Error al leer el archivo: {e}")
    st.stop()

n_prov = len(proveedores)
mu = score_df.mean().values
cov_matrix = score_df.cov().values

# ─────────────────────────────────────────────
# SECCIÓN 0 — DIAGNÓSTICO DE CONCENTRACIÓN
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">0 · Diagnóstico de Concentración Actual</div>', unsafe_allow_html=True)

if tiene_gasto:
    gasto_prom = gasto_df.mean()
    gasto_total_real = gasto_prom.sum()
    participacion = (gasto_prom / gasto_total_real * 100).sort_values(ascending=False)

    col_d1, col_d2 = st.columns([1, 1])

    with col_d1:
        df_conc = pd.DataFrame({
            "Proveedor": participacion.index,
            "Gasto Promedio ($)": gasto_prom[participacion.index].round(0),
            "Participación (%)": participacion.values.round(1),
        }).reset_index(drop=True)

        def highlight_conc(val):
            if isinstance(val, (int, float)) and val > umbral_conc:
                return "background-color:#ffcccc; font-weight:bold; color:#7d0000"
            return ""

        st.dataframe(
            df_conc.style
                .format({"Gasto Promedio ($)": "${:,.0f}", "Participación (%)": "{:.1f}%"})
                .applymap(highlight_conc, subset=["Participación (%)"]),
            use_container_width=True, height=390
        )

        en_riesgo = participacion[participacion > umbral_conc]
        if len(en_riesgo) > 0:
            for prov, pct in en_riesgo.items():
                st.markdown(
                    f'<div class="alert-box">⚠️ <b>{prov}</b> concentra el <b>{pct:.1f}%</b> del gasto total — supera el umbral de {umbral_conc}%.</div>',
                    unsafe_allow_html=True
                )
        else:
            st.markdown(
                f'<div class="ok-box">✅ Ningún proveedor supera el umbral de concentración del {umbral_conc}%.</div>',
                unsafe_allow_html=True
            )

    with col_d2:
        colors = ["#e74c3c" if v > umbral_conc else "#001d3d" for v in participacion.values]
        fig_dona_actual = go.Figure(go.Pie(
            labels=participacion.index,
            values=participacion.values.round(1),
            hole=0.48,
            textinfo="label+percent",
            marker=dict(colors=colors),
            hovertemplate="<b>%{label}</b><br>%{value:.1f}% del gasto<extra></extra>"
        ))
        fig_dona_actual.update_layout(
            title=dict(text="Concentración actual del gasto por proveedor", font=dict(size=14)),
            height=420, showlegend=False,
            margin=dict(t=50, b=20, l=20, r=20)
        )
        st.plotly_chart(fig_dona_actual, use_container_width=True)
else:
    st.info("ℹ️ Carga la hoja de gasto mensual para ver el diagnóstico de concentración.")

# ─────────────────────────────────────────────
# SECCIÓN 1 — DESEMPEÑO HISTÓRICO
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">1 · Desempeño Histórico por Proveedor</div>', unsafe_allow_html=True)

resumen = pd.DataFrame({
    "Score Promedio": score_df.mean().round(3),
    "Volatilidad (σ)": score_df.std().round(4),
    "Score Mínimo": score_df.min().round(3),
    "Score Máximo": score_df.max().round(3),
    "Rango": (score_df.max() - score_df.min()).round(3),
})

col_t1, col_t2 = st.columns([1, 1])
with col_t1:
    st.markdown("**Tabla de métricas por proveedor:**")
    st.dataframe(
        resumen.style
            .format("{:.3f}")
            .background_gradient(subset=["Score Promedio"], cmap="Greens")
            .background_gradient(subset=["Volatilidad (σ)"], cmap="Reds_r")
            .background_gradient(subset=["Rango"], cmap="Oranges_r"),
        use_container_width=True, height=400
    )

with col_t2:
    # Mapa riesgo–desempeño
    mu_s = score_df.mean()
    sig_s = score_df.std()
    colores_scatter = ["#e74c3c" if mu_s[p] < 0.65 else "#f39c12" if mu_s[p] < 0.75 else "#27ae60" for p in proveedores]

    fig_mapa = go.Figure()
    for i, prov in enumerate(proveedores):
        fig_mapa.add_trace(go.Scatter(
            x=[sig_s[prov]], y=[mu_s[prov]],
            mode="markers+text",
            name=prov,
            text=[prov], textposition="top center",
            marker=dict(size=18, color=colores_scatter[i],
                        line=dict(width=1.5, color="white")),
            hovertemplate=f"<b>{prov}</b><br>Score: {mu_s[prov]:.3f}<br>σ: {sig_s[prov]:.4f}<extra></extra>"
        ))

    fig_mapa.add_hline(y=0.75, line_dash="dot", line_color="#f39c12",
                       annotation_text="Umbral bueno (0.75)", annotation_position="right")
    fig_mapa.add_hline(y=0.65, line_dash="dot", line_color="#e74c3c",
                       annotation_text="Umbral mínimo (0.65)", annotation_position="right")

    fig_mapa.update_layout(
        title="Mapa Riesgo–Desempeño (ideal: arriba–izquierda)",
        xaxis_title="Riesgo — Volatilidad del Score (σ)",
        yaxis_title="Desempeño Esperado (Score Promedio)",
        height=400, showlegend=False, hovermode="closest"
    )
    st.plotly_chart(fig_mapa, use_container_width=True)

st.markdown('<div class="insight-box">💡 El proveedor ideal está en la <b>esquina superior izquierda</b>: alto desempeño y baja variabilidad. Verde = buen desempeño, amarillo = aceptable, rojo = bajo el umbral mínimo.</div>', unsafe_allow_html=True)

# Gráfica de evolución histórica
fig_hist = go.Figure()
for prov in proveedores:
    fig_hist.add_trace(go.Scatter(
        x=score_df.index.astype(str), y=score_df[prov],
        mode="lines+markers", name=prov,
        line=dict(width=2),
        hovertemplate=f"<b>{prov}</b><br>%{{x}}: %{{y:.3f}}<extra></extra>"
    ))
fig_hist.add_hline(y=0.70, line_dash="dot", line_color="#e74c3c",
                   annotation_text="Umbral mínimo aceptable (0.70)")
fig_hist.update_layout(
    title="Evolución Histórica del Score de Desempeño por Proveedor",
    xaxis_title="Mes", yaxis_title="Score (0–1)",
    yaxis=dict(range=[0.35, 1.02]),
    height=380, hovermode="x unified",
    legend=dict(orientation="h", y=-0.3, x=0)
)
st.plotly_chart(fig_hist, use_container_width=True)

# ─────────────────────────────────────────────
# SECCIÓN 2 — CORRELACIÓN Y COVARIANZA
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">2 · Correlación y Covarianza entre Proveedores</div>', unsafe_allow_html=True)

col_c1, col_c2 = st.columns(2)
mat_corr = score_df.corr().round(3)
mat_cov  = score_df.cov().round(5)

with col_c1:
    fig_corr = px.imshow(
        mat_corr, text_auto=True,
        color_continuous_scale="RdBu_r", zmin=-1, zmax=1,
        aspect="auto", title="Matriz de Correlación entre Proveedores"
    )
    fig_corr.update_layout(height=400)
    st.plotly_chart(fig_corr, use_container_width=True)

with col_c2:
    fig_cov = px.imshow(
        mat_cov, text_auto=".4f",
        color_continuous_scale="Blues",
        aspect="auto", title="Matriz de Covarianza entre Proveedores"
    )
    fig_cov.update_layout(height=400)
    st.plotly_chart(fig_cov, use_container_width=True)

st.markdown('<div class="insight-box">💡 <b>Correlación baja o negativa entre proveedores</b> significa que cuando uno falla, el otro no necesariamente también falla — eso es diversificación real. Proveedores del mismo sector o región tienden a correlación alta (riesgo sistémico).</div>', unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SECCIÓN 3 — SIMULACIÓN MONTE CARLO
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">3 · Frontera Eficiente de Abastecimiento — Monte Carlo</div>', unsafe_allow_html=True)

try:
    np.random.seed(42)
    resultados = np.zeros((n_sim, 3 + n_prov))

    for i in range(n_sim):
        w = np.random.random(n_prov)
        w /= w.sum()
        score_p  = np.dot(w, mu)
        riesgo_p = np.sqrt(w @ cov_matrix @ w)
        sharpe_p = score_p / riesgo_p if riesgo_p > 0 else 0
        resultados[i] = [score_p, riesgo_p, sharpe_p] + list(w)

    cols_r = ["Score", "Riesgo", "Sharpe"] + proveedores
    df_sim = pd.DataFrame(resultados, columns=cols_r)

    idx_opt  = df_sim["Sharpe"].idxmax()
    idx_minr = df_sim["Riesgo"].idxmin()
    cartera_opt  = df_sim.loc[idx_opt]
    cartera_minr = df_sim.loc[idx_minr]

    fig_fe = go.Figure()

    # Carteras simuladas coloreadas por Sharpe
    fig_fe.add_trace(go.Scatter(
        x=df_sim["Riesgo"], y=df_sim["Score"],
        mode="markers",
        marker=dict(
            color=df_sim["Sharpe"], colorscale="Viridis",
            size=4, opacity=0.5,
            colorbar=dict(title="Sharpe", thickness=12, len=0.7)
        ),
        name="Carteras simuladas",
        hovertemplate="Riesgo: %{x:.4f}<br>Score: %{y:.3f}<extra></extra>"
    ))

    # Mínimo riesgo
    fig_fe.add_trace(go.Scatter(
        x=[cartera_minr["Riesgo"]], y=[cartera_minr["Score"]],
        mode="markers",
        marker=dict(symbol="diamond", size=18, color="#3498db",
                    line=dict(color="white", width=1.5)),
        name="🔵 Mínimo Riesgo",
        hovertemplate=f"Score: {cartera_minr['Score']:.3f}<br>Riesgo: {cartera_minr['Riesgo']:.5f}<br>Sharpe: {cartera_minr['Sharpe']:.2f}<extra></extra>"
    ))

    # Cartera óptima
    fig_fe.add_trace(go.Scatter(
        x=[cartera_opt["Riesgo"]], y=[cartera_opt["Score"]],
        mode="markers",
        marker=dict(symbol="star", size=24, color="#e74c3c",
                    line=dict(color="white", width=1.5)),
        name="⭐ Cartera Óptima",
        hovertemplate=f"Score: {cartera_opt['Score']:.3f}<br>Riesgo: {cartera_opt['Riesgo']:.5f}<br>Sharpe: {cartera_opt['Sharpe']:.2f}<extra></extra>"
    ))

    fig_fe.update_layout(
        title=f"Frontera Eficiente de Abastecimiento — {n_sim:,} combinaciones simuladas",
        xaxis_title="Riesgo del Portafolio (σ del Score)",
        yaxis_title="Score de Desempeño Esperado",
        height=500, hovermode="closest",
        legend=dict(orientation="h", y=-0.15, x=0)
    )
    st.plotly_chart(fig_fe, use_container_width=True)

    ci1, ci2 = st.columns(2)
    with ci1:
        st.markdown('<div class="insight-box">⭐ <b>Cartera Óptima (estrella roja):</b> Maximiza el score de desempeño por unidad de riesgo — la combinación más eficiente de proveedores.</div>', unsafe_allow_html=True)
    with ci2:
        st.markdown('<div class="insight-box">🔵 <b>Mínimo Riesgo (rombo azul):</b> La combinación que minimiza la variabilidad operativa — útil cuando la estabilidad es prioritaria sobre el desempeño.</div>', unsafe_allow_html=True)

except Exception as e:
    st.error(f"❌ Error en simulación Monte Carlo: {e}")
    st.stop()

# ─────────────────────────────────────────────
# SECCIÓN 4 — ASIGNACIÓN ÓPTIMA
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">4 · Asignación Óptima del Gasto en Proveedores</div>', unsafe_allow_html=True)

m1, m2, m3 = st.columns(3)
with m1:
    st.metric("📈 Score Esperado del Portafolio", f"{cartera_opt['Score']:.4f}")
with m2:
    st.metric("⚡ Riesgo del Portafolio (σ)", f"{cartera_opt['Riesgo']:.5f}")
with m3:
    st.metric("🏆 Índice Sharpe (Score / Riesgo)", f"{cartera_opt['Sharpe']:.2f}")

st.markdown("---")

pesos_opt = {p: cartera_opt[p] for p in proveedores}
df_asig = pd.DataFrame({
    "Proveedor": proveedores,
    "Peso Óptimo (%)": [pesos_opt[p] * 100 for p in proveedores],
    "Gasto Asignado ($)": [pesos_opt[p] * gasto_total for p in proveedores],
    "Score Promedio": [score_df[p].mean() for p in proveedores],
    "Volatilidad (σ)": [score_df[p].std() for p in proveedores],
}).sort_values("Peso Óptimo (%)", ascending=False).reset_index(drop=True)

# Agregar % actual si hay datos de gasto
if tiene_gasto:
    part_actual = (gasto_df.mean() / gasto_df.mean().sum() * 100)
    df_asig["% Actual"] = df_asig["Proveedor"].map(part_actual).round(1)
    df_asig["Cambio (pp)"] = (df_asig["Peso Óptimo (%)"] - df_asig["% Actual"]).round(1)

col_ta, col_do = st.columns([1, 1])

with col_ta:
    fmt = {
        "Peso Óptimo (%)": "{:.1f}%",
        "Gasto Asignado ($)": "${:,.0f}",
        "Score Promedio": "{:.3f}",
        "Volatilidad (σ)": "{:.4f}",
    }
    if tiene_gasto:
        fmt["% Actual"] = "{:.1f}%"
        fmt["Cambio (pp)"] = "{:+.1f}"

    st.dataframe(
        df_asig.style
            .format(fmt)
            .background_gradient(subset=["Peso Óptimo (%)"], cmap="Greens")
            .background_gradient(subset=["Score Promedio"], cmap="Blues"),
        use_container_width=True, height=400
    )

with col_do:
    fig_dona = go.Figure(go.Pie(
        labels=df_asig["Proveedor"],
        values=df_asig["Peso Óptimo (%)"].round(1),
        hole=0.50,
        textinfo="label+percent",
        hovertemplate="<b>%{label}</b><br>%{value:.1f}%<br>${:,.0f}<extra></extra>".format(gasto_total)
    ))
    fig_dona.update_layout(
        title=f"Distribución óptima — Total: ${gasto_total:,.0f}",
        height=400, showlegend=False,
        margin=dict(t=50, b=20, l=20, r=20)
    )
    st.plotly_chart(fig_dona, use_container_width=True)

# ─────────────────────────────────────────────
# SECCIÓN 5 — COMPARATIVA ACTUAL VS. ÓPTIMA
# ─────────────────────────────────────────────
if tiene_gasto:
    st.markdown('<div class="section-header">5 · Comparativa: Distribución Actual vs. Óptima</div>', unsafe_allow_html=True)

    pesos_actuales = part_actual.reindex(proveedores).fillna(0).values / 100
    score_actual  = np.dot(pesos_actuales, mu)
    riesgo_actual = np.sqrt(pesos_actuales @ cov_matrix @ pesos_actuales)
    sharpe_actual = score_actual / riesgo_actual if riesgo_actual > 0 else 0

    pesos_minr = np.array([cartera_minr[p] for p in proveedores])

    df_comp = pd.DataFrame({
        "Estrategia": ["Distribución Actual", "Cartera Mínimo Riesgo", "Cartera Óptima (Máx. Sharpe)"],
        "Score Esperado": [score_actual, cartera_minr["Score"], cartera_opt["Score"]],
        "Riesgo (σ)":    [riesgo_actual, cartera_minr["Riesgo"], cartera_opt["Riesgo"]],
        "Índice Sharpe": [sharpe_actual, cartera_minr["Sharpe"], cartera_opt["Sharpe"]],
    })

    mejora_score = ((cartera_opt["Score"] - score_actual) / score_actual * 100)
    reduccion_r  = ((riesgo_actual - cartera_opt["Riesgo"]) / riesgo_actual * 100)

    col_ct, col_cm = st.columns([3, 2])
    with col_ct:
        st.dataframe(
            df_comp.style
                .format({"Score Esperado":"{:.4f}", "Riesgo (σ)":"{:.5f}", "Índice Sharpe":"{:.2f}"})
                .highlight_max(subset=["Score Esperado","Índice Sharpe"], color="#d4edda")
                .highlight_min(subset=["Riesgo (σ)"], color="#d4edda"),
            use_container_width=True, height=160
        )

    with col_cm:
        st.metric("📈 Mejora en Score", f"+{mejora_score:.2f}%" if mejora_score > 0 else f"{mejora_score:.2f}%")
        st.metric("🛡️ Reducción de Riesgo", f"-{reduccion_r:.2f}%" if reduccion_r > 0 else f"{reduccion_r:.2f}%")

    # Barras agrupadas actual vs óptima
    fig_barras = go.Figure()
    fig_barras.add_trace(go.Bar(
        name="% Actual", x=proveedores,
        y=[part_actual.get(p, 0) for p in proveedores],
        marker_color="#001d3d",
        text=[f"{part_actual.get(p,0):.1f}%" for p in proveedores],
        textposition="outside"
    ))
    fig_barras.add_trace(go.Bar(
        name="% Óptimo", x=proveedores,
        y=[pesos_opt[p] * 100 for p in proveedores],
        marker_color="#f15b2b",
        text=[f"{pesos_opt[p]*100:.1f}%" for p in proveedores],
        textposition="outside"
    ))
    fig_barras.add_hline(
        y=umbral_conc, line_dash="dot", line_color="red",
        annotation_text=f"Umbral de concentración ({umbral_conc}%)",
        annotation_position="top right"
    )
    fig_barras.update_layout(
        title="Distribución Actual vs. Óptima del Gasto por Proveedor",
        xaxis_title="Proveedor", yaxis_title="% del Gasto Total",
        barmode="group", height=420,
        legend=dict(orientation="h", y=-0.2, x=0)
    )
    st.plotly_chart(fig_barras, use_container_width=True)

# ─────────────────────────────────────────────
# SECCIÓN 6 — SEÑALES DE ACCIÓN
# ─────────────────────────────────────────────
st.markdown('<div class="section-header">6 · Señales de Acción por Proveedor</div>', unsafe_allow_html=True)
st.markdown("Clasificación basada en score promedio, volatilidad y peso óptimo asignado por el modelo:")

for prov in proveedores:
    score_p = score_df[prov].mean()
    vol_p   = score_df[prov].std()
    peso_p  = pesos_opt[prov] * 100

    if score_p >= 0.75 and vol_p < 0.07:
        cat   = "🟢 CONSOLIDAR"
        estilo = "ok-box"
        accion = f"Proveedor confiable y estable. Mantener o incrementar participación. Peso óptimo: <b>{peso_p:.1f}%</b>."
    elif score_p >= 0.75 and vol_p >= 0.07:
        cat   = "🟡 MONITOREAR"
        estilo = "alert-box"
        accion = f"Buen desempeño pero variable. Implementar SLAs y seguimiento mensual. Peso óptimo: <b>{peso_p:.1f}%</b>."
    elif score_p >= 0.65:
        cat   = "🔵 DESARROLLAR"
        estilo = "insight-box"
        accion = f"Desempeño aceptable con potencial de mejora. Invertir en la relación comercial. Peso óptimo: <b>{peso_p:.1f}%</b>."
    else:
        cat   = "🔴 REDUCIR / REEMPLAZAR"
        estilo = "alert-box"
        accion = f"Score por debajo del umbral mínimo. Reducir participación o buscar alternativas. Peso óptimo: <b>{peso_p:.1f}%</b>."

    cambio_txt = ""
    if tiene_gasto:
        actual_pct = part_actual.get(prov, 0)
        delta = peso_p - actual_pct
        cambio_txt = f"&nbsp;|&nbsp; Cambio sugerido: <b>{'+'if delta>0 else ''}{delta:.1f} pp</b>"

    st.markdown(
        f'<div class="{estilo}"><b>{prov}</b> — {cat}'
        f'&nbsp;|&nbsp; Score: <b>{score_p:.3f}</b>'
        f'&nbsp;|&nbsp; σ: <b>{vol_p:.4f}</b>'
        f'{cambio_txt}<br><small>{accion}</small></div>',
        unsafe_allow_html=True
    )

st.markdown("---")
st.caption("Desarrollado con lógica de Markowitz aplicada a la gestión de proveedores · Universidad Panamericana — IA para el Análisis Financiero")
