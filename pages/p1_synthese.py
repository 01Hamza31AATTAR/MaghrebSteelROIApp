import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, empty_state

inject_css()
page_header("Synthèse", "Vue d'ensemble du scénario baseline (modèle LP)")

df_synth = dl.cached_load_sheet("synthese")
df_res = dl.cached_load_sheet("commandes")

if df_synth is None or df_res is None:
    empty_state()
    st.stop()

val = dict(zip(df_synth["Indicateur"], df_synth["Valeur"]))

kpi_row([
    {"label": "Marge totale", "value": f"{val.get('Marge totale (MAD)', '—')} MAD", "tone": "pos"},
    {"label": "Taux de service", "value": val.get("Taux de service (%)", "—"), "tone": "neutral"},
    {"label": "Commandes honorées", "value": val.get("Commandes honorées", "—"), "tone": "pos"},
    {"label": "Commandes partielles", "value": val.get("Commandes partielles", "—"), "tone": "neutral"},
    {"label": "Commandes refusées", "value": val.get("Commandes refusées", "—"), "tone": "neg"},
])

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("Répartition des commandes")
    counts = df_res["Statut"].value_counts().reset_index()
    counts.columns = ["Statut", "Nombre"]
    fig = px.pie(counts, names="Statut", values="Nombre", hole=0.55)
    st.plotly_chart(style_fig(fig, height=340), use_container_width=True)

with col2:
    st.subheader("Tonnage par famille")
    fam = df_res.groupby("Famille")[["Tonnage", "Livree"]].sum().reset_index()
    fig = px.bar(fam, x="Famille", y=["Tonnage", "Livree"], barmode="group")
    fig.update_layout(legend_title_text="")
    st.plotly_chart(style_fig(fig, height=340), use_container_width=True)

st.subheader("Indicateurs détaillés")
st.dataframe(df_synth, use_container_width=True, hide_index=True)
