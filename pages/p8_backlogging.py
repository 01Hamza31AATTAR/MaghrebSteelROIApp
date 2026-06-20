import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("B2 · Backlogging", "Modélisation des retards de livraison autorisés")

df_bl = dl.cached_load_sheet("backlogging_commandes")
df_comp = dl.cached_load_sheet("backlogging_synthese")

if df_bl is None or df_bl.empty:
    empty_state()
    st.stop()

comp_map = dict(zip(df_comp["Indicateur"], df_comp["Backlogging"])) if df_comp is not None else {}

kpi_row([
    {"label": "Marge totale", "value": f"{comp_map.get('Marge totale (MAD)', '—')} MAD", "tone": "pos"},
    {"label": "Taux de service global", "value": comp_map.get("Taux de service global (%)", "—")},
    {"label": "Tonnes livrées en retard", "value": comp_map.get("Tonnes livrées en retard", "—"), "tone": "neutral"},
    {"label": "Pénalités retard totales", "value": f"{comp_map.get('Pénalités retard totales (MAD)', '—')} MAD", "tone": "neg"},
])

tab_detail, tab_synth = st.tabs(["⏳ Détail des commandes", "📋 Synthèse comparative"])

with tab_detail:
    st.subheader("Tonnage livré à temps vs en retard, par famille")
    fam = df_bl.groupby("Famille")[["Livre_a_temps_T", "Livre_en_retard_T"]].sum().reset_index()
    fig = px.bar(fam, x="Famille", y=["Livre_a_temps_T", "Livre_en_retard_T"], barmode="stack")
    fig.update_layout(legend_title_text="")
    st.plotly_chart(style_fig(fig, height=360), use_container_width=True)

    st.subheader("Détail par commande")
    render_filtered_table(df_bl, filter_cols=["Famille", "Statut", "Priorite"], key_prefix="bl")

with tab_synth:
    if df_comp is None or df_comp.empty:
        empty_state()
    else:
        st.subheader("Indicateurs de synthèse — scénario Backlogging")
        st.dataframe(df_comp, use_container_width=True, hide_index=True)
