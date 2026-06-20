import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("Commandes", "Détail des commandes : honorées, partielles, refusées")

df_res = dl.cached_load_sheet("commandes")
df_refus = dl.cached_load_sheet("commandes_refusees")

if df_res is None:
    empty_state()
    st.stop()

taux_service = df_res["Livree"].sum() / df_res["Tonnage"].sum() * 100 if df_res["Tonnage"].sum() else 0

kpi_row([
    {"label": "Commandes totales", "value": len(df_res)},
    {"label": "Tonnage demandé", "value": f"{df_res['Tonnage'].sum():,.0f} T"},
    {"label": "Tonnage livré", "value": f"{df_res['Livree'].sum():,.0f} T", "tone": "pos"},
    {"label": "Taux de service", "value": f"{taux_service:.1f} %", "tone": "neutral"},
    {"label": "Commandes refusées", "value": len(df_refus) if df_refus is not None else 0, "tone": "neg"},
])

tab_all, tab_refus = st.tabs(["📦 Toutes les commandes", "🚫 Commandes refusées"])

with tab_all:
    st.subheader("Tonnage demandé vs livré par famille")
    fam = df_res.groupby("Famille")[["Tonnage", "Livree"]].sum().reset_index()
    fig = px.bar(fam, x="Famille", y=["Tonnage", "Livree"], barmode="group")
    fig.update_layout(legend_title_text="")
    st.plotly_chart(style_fig(fig, height=340), use_container_width=True)

    st.subheader("Détail des commandes")
    render_filtered_table(df_res, filter_cols=["Famille", "Statut", "Priorite"], key_prefix="cmd")

with tab_refus:
    if df_refus is None or df_refus.empty:
        st.success("Aucune commande refusée — taux de service maximal sur ce scénario.")
    else:
        st.subheader("Tonnage perdu par famille")
        fam_r = df_refus.groupby("Famille")["Tonnage"].sum().reset_index()
        fig = px.bar(fam_r, x="Famille", y="Tonnage")
        st.plotly_chart(style_fig(fig, height=320), use_container_width=True)
        st.subheader("Liste des commandes refusées")
        render_filtered_table(df_refus, filter_cols=["Famille", "Priorite"], key_prefix="refus")
