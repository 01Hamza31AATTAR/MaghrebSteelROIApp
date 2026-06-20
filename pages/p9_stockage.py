import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("B3 · Coûts de stockage", "Valorisation du stock produit fini et impact sur la marge nette")

df_res_st = dl.cached_load_sheet("stockage_commandes")
df_stock_t = dl.cached_load_sheet("stockage_detail")
df_comp = dl.cached_load_sheet("stockage_synthese")
df_h = dl.cached_load_sheet("stockage_taux_hf")

if df_comp is None or df_comp.empty:
    empty_state()
    st.stop()

comp_map = dict(zip(df_comp["Indicateur"], df_comp["Valeur"]))

kpi_row([
    {"label": "Marge nette totale", "value": f"{comp_map.get('Marge nette totale (MAD)', '—')} MAD", "tone": "pos"},
    {"label": "Coût stockage total", "value": f"{comp_map.get('dont Coût stockage total (MAD)', '—')} MAD", "tone": "neg"},
    {"label": "Taux de service", "value": comp_map.get("Taux de service (%)", "—")},
    {"label": "Mode de valorisation", "value": comp_map.get("Mode valorisation stock", "—"), "tone": "neutral"},
])

tab_detail, tab_hf, tab_cmd, tab_synth = st.tabs([
    "🏬 Stock & coût par semaine", "📐 Taux h_f", "📦 Commandes", "📋 Synthèse",
])

with tab_detail:
    if df_stock_t is None or df_stock_t.empty:
        empty_state()
    else:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Évolution du stock par famille")
            fig = px.area(df_stock_t, x="Semaine", y="Stock_T", color="Famille")
            st.plotly_chart(style_fig(fig, height=360), use_container_width=True)
        with col2:
            st.subheader("Coût de stockage cumulé par famille")
            cout_fam = df_stock_t.groupby("Famille")["Cout_stockage_MAD"].sum().reset_index()
            fig = px.bar(cout_fam, x="Famille", y="Cout_stockage_MAD")
            st.plotly_chart(style_fig(fig, height=360), use_container_width=True)
        render_filtered_table(df_stock_t, filter_cols=["Famille", "Semaine"], key_prefix="stock")

with tab_hf:
    if df_h is None or df_h.empty:
        empty_state()
    else:
        st.subheader("Taux de stockage h_f par famille (MAD/T/semaine)")
        fig = px.bar(df_h, x="Famille", y="h_total_MAD_T_sem", color="Famille")
        st.plotly_chart(style_fig(fig, height=340), use_container_width=True)
        st.dataframe(df_h, use_container_width=True, hide_index=True)

with tab_cmd:
    if df_res_st is None or df_res_st.empty:
        empty_state()
    else:
        render_filtered_table(df_res_st, filter_cols=["Famille", "Statut"], key_prefix="rst")

with tab_synth:
    st.dataframe(df_comp, use_container_width=True, hide_index=True)
