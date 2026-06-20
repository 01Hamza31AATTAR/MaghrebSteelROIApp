import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("B1 · Tarification d'opportunité", "Prix plancher et signal de renégociation pour les commandes non honorées")

df_t = dl.cached_load_sheet("tarification")
if df_t is None or df_t.empty:
    st.info("Aucune commande refusée ou partielle sur ce scénario : pas de tarification d'opportunité à calculer.")
    st.stop()

nb_sous_plancher = (df_t["Signal"].str.startswith("Sous le plancher")).sum()

kpi_row([
    {"label": "Commandes ciblées", "value": len(df_t)},
    {"label": "Reliquat total", "value": f"{df_t['Reliquat_T'].sum():,.0f} T"},
    {"label": "Prix d'opportunité moyen", "value": f"{df_t['Prix_opportunite_MAD_T'].mean():,.0f} MAD/T"},
    {"label": "Sous le prix plancher", "value": int(nb_sous_plancher), "tone": "neg" if nb_sous_plancher else "pos"},
])

col1, col2 = st.columns([1.2, 1])

with col1:
    st.subheader("Prix de vente vs prix plancher par commande")
    fig = px.scatter(
        df_t, x="Prix_plancher_MAD_T", y="PrixVente_MAD_T",
        color="Signal", size="Reliquat_T", hover_data=["ID", "Famille"],
    )
    max_v = max(df_t["Prix_plancher_MAD_T"].max(), df_t["PrixVente_MAD_T"].max())
    fig.add_shape(type="line", x0=0, y0=0, x1=max_v, y1=max_v, line=dict(color="#A9ADB3", dash="dot"))
    st.plotly_chart(style_fig(fig, height=400), use_container_width=True)

with col2:
    st.subheader("Répartition par signal")
    counts = df_t["Signal"].value_counts().reset_index()
    counts.columns = ["Signal", "Nombre"]
    fig = px.bar(counts, x="Nombre", y="Signal", orientation="h")
    st.plotly_chart(style_fig(fig, height=400), use_container_width=True)

st.subheader("Détail des commandes")
render_filtered_table(df_t, filter_cols=["Famille", "Statut", "Signal"], key_prefix="tarif")
