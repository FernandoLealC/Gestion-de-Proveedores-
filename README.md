# 🏭 Optimización de Cartera de Proveedores

Aplicación Streamlit que aplica la **lógica de Markowitz** a la gestión estratégica de proveedores.

**Pregunta central:** ¿Cómo distribuir el gasto entre proveedores para maximizar el desempeño y minimizar la vulnerabilidad operativa?

## 🚀 Demo

[![Open in Streamlit](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://share.streamlit.io)

## 🎯 ¿Qué hace la app?

| Concepto Markowitz | Equivalente en proveedores |
|---|---|
| Activo | Proveedor |
| Retorno esperado | Score de desempeño promedio (calidad, precio, puntualidad…) |
| Riesgo | Volatilidad histórica del score (σ) |
| Correlación | Sincronía entre los problemas de dos proveedores |
| Cartera óptima | % de gasto óptimo asignado a cada proveedor |

## 📋 Secciones

| # | Sección | Contenido |
|---|---|---|
| 0 | Diagnóstico de concentración | Alerta automática por umbral configurable + dona del gasto actual |
| 1 | Desempeño histórico | Tabla de métricas, mapa Riesgo–Desempeño, evolución histórica |
| 2 | Correlación y covarianza | Heatmaps con interpretación didáctica |
| 3 | Frontera eficiente | Monte Carlo (1K–20K sim), cartera óptima ⭐ y mínimo riesgo 🔵 |
| 4 | Asignación óptima | Tabla peso % y $, dona de distribución, métricas Sharpe |
| 5 | Comparativa actual vs. óptima | Barras agrupadas, mejora en score y reducción de riesgo |
| 6 | Señales de acción | CONSOLIDAR / MONITOREAR / DESARROLLAR / REDUCIR por proveedor |

## 📁 Archivo Excel requerido — 2 hojas principales

**Hoja: `Desempeño Proveedores`**

| Mes | Proveedor A | Proveedor B | ... |
|---|---|---|---|
| Ene-2024 | 0.85 | 0.72 | ... |

> Score entre 0 y 1 (combinación de calidad, puntualidad, precio, disponibilidad, logística)

**Hoja: `Gasto Mensual`**

| Mes | Proveedor A | Proveedor B | ... |
|---|---|---|---|
| Ene-2024 | 450000 | 280000 | ... |

> Gasto real en $ por proveedor por mes

📎 Se incluye **`base_datos_proveedores_v2.xlsx`** con datos simulados de 10 proveedores durante 24 meses listos para usar.

## ⚙️ Instalación local

```bash
git clone https://github.com/FernandoLealC/markowitz-proveedores
cd markowitz-proveedores
pip install -r requirements.txt
streamlit run app.py
```

## 🛠️ Tecnologías

Python · Streamlit · Pandas · NumPy · Plotly · OpenPyXL

---
Desarrollado con fines didácticos | Universidad Panamericana · IA para el Análisis Financiero
