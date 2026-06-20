import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("Plan de marche", "Production planifiée par ligne, famille et semaine")

df_plan = dl.cached_load_sheet("plan_marche")
if df_plan is None or df_plan.empty:
    empty_state()
    st.stop()

fam_top = df_plan.groupby("Famille")["Tonnage_traite_T"].sum().idxmax()

kpi_row([
    {"label": "Tonnage total traité", "value": f"{df_plan['Tonnage_traite_T'].sum():,.0f} T"},
    {"label": "Lignes actives", "value": df_plan["Ligne"].nunique()},
    {"label": "Familles traitées", "value": df_plan["Famille"].nunique()},
    {"label": "Famille dominante", "value": fam_top, "tone": "neutral"},
])

st.subheader("Tonnage traité par ligne et semaine (par famille)")
fig = px.bar(
    df_plan, x="Ligne", y="Tonnage_traite_T", color="Famille",
    facet_col="Semaine", facet_col_wrap=4,
)
st.plotly_chart(style_fig(fig, height=420), use_container_width=True)

st.subheader("Détail du plan de marche")
render_filtered_table(df_plan, filter_cols=["Ligne", "Famille", "Semaine"], key_prefix="plan")
