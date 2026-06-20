import plotly.graph_objects as go
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, empty_state

inject_css()
page_header("B5 · Branch-and-Bound", "Métriques de résolution de l'arbre Branch-and-Bound")

df_bb = dl.cached_load_sheet("branch_and_bound")
if df_bb is None or df_bb.empty:
    empty_state()
    st.stop()

m = dict(zip(df_bb["Métrique"], df_bb["Valeur"]))

def _to_float(s, default=0.0):
    try:
        return float(str(s).replace(",", "").replace("%", "").strip())
    except (TypeError, ValueError):
        return default

gap_pct = _to_float(m.get("Gap B&B (%)", 0))

kpi_row([
    {"label": "Statut résolution", "value": m.get("Statut résolution", "—"),
     "tone": "pos" if str(m.get("Statut résolution", "")).lower() == "optimal" else "neutral"},
    {"label": "Marge optimale", "value": f"{m.get('Marge optimale (MAD)', '—')} MAD", "tone": "pos"},
    {"label": "Nœuds explorés", "value": m.get("Nœuds explorés", "—")},
    {"label": "Variables binaires", "value": m.get("Variables binaires totales", "—")},
    {"label": "Temps de résolution", "value": f"{m.get('Temps de résolution (s)', '—')} s"},
])

col1, col2 = st.columns([1, 1.4])

with col1:
    st.subheader("Gap d'optimalité")
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=gap_pct,
        number={"suffix": " %"},
        gauge={
            "axis": {"range": [0, max(5, gap_pct * 1.5)]},
            "bar": {"color": "#C8102E"},
            "steps": [
                {"range": [0, 1], "color": "#3FA66A"},
                {"range": [1, max(5, gap_pct * 1.5)], "color": "#343840"},
            ],
        },
    ))
    st.plotly_chart(style_fig(fig, height=300), use_container_width=True)
    st.caption(m.get("Interprétation gap", ""))

with col2:
    st.subheader("Toutes les métriques")
    st.dataframe(df_bb, use_container_width=True, hide_index=True)
