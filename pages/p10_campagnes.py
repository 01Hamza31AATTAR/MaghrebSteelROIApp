import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("B4 · Campagnes (MILP)", "Plan de campagnes par ligne et coûts de changement de famille")

df_cmd = dl.cached_load_sheet("campagnes_commandes")
df_plan = dl.cached_load_sheet("campagnes_plan")
df_comp = dl.cached_load_sheet("campagnes_synthese")

if df_comp is None or df_comp.empty:
    empty_state()
    st.stop()

comp_map = dict(zip(df_comp["Indicateur"], df_comp["Valeur"]))

kpi_row([
    {"label": "Marge totale", "value": f"{comp_map.get('Marge totale (MAD)', '—')} MAD", "tone": "pos"},
    {"label": "Coût changements campagne", "value": f"{comp_map.get('Coût total changements campagne (MAD)', '—')} MAD", "tone": "neg"},
    {"label": "Nombre de changements", "value": comp_map.get("Nombre de changements de campagne", "—"), "tone": "neutral"},
    {"label": "Taux de service", "value": comp_map.get("Taux de service (%)", "—")},
])

tab_plan, tab_cmd, tab_synth = st.tabs([
    "🔁 Campagnes_plan", "📦 Campagnes_commandes", "📋 Campagnes_synthese",
])

with tab_plan:
    if df_plan is None or df_plan.empty:
        empty_state()
    else:
        st.subheader("Famille active par ligne et semaine")
        pivot = df_plan.pivot_table(
            index="Ligne", columns="Semaine", values="Famille_active", aggfunc="first"
        )
        st.dataframe(pivot, use_container_width=True)

        st.subheader("Tonnage traité et changements par ligne")
        fig = px.bar(
            df_plan, x="Semaine", y="Tonnes_input_T", color="Famille_active",
            facet_col="Ligne", facet_col_wrap=4,
        )
        st.plotly_chart(style_fig(fig, height=420), use_container_width=True)

        st.subheader("Détail du plan de campagnes")
        render_filtered_table(df_plan, filter_cols=["Ligne", "Famille_active", "Changement"], key_prefix="camp_plan")

with tab_cmd:
    if df_cmd is None or df_cmd.empty:
        empty_state()
    else:
        st.subheader("Statut des commandes — scénario Campagnes")
        counts = df_cmd["Statut"].value_counts().reset_index()
        counts.columns = ["Statut", "Nombre"]
        fig = px.pie(counts, names="Statut", values="Nombre", hole=0.5)
        st.plotly_chart(style_fig(fig, height=340), use_container_width=True)
        render_filtered_table(df_cmd, filter_cols=["Famille", "Statut"], key_prefix="camp_cmd")

with tab_synth:
    st.dataframe(df_comp, use_container_width=True, hide_index=True)
