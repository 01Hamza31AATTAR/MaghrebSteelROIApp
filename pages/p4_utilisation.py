import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("Utilisation des lignes", "Taux de charge hebdomadaire par ligne de production")

df_util = dl.cached_load_sheet("utilisation")
if df_util is None or df_util.empty:
    empty_state()
    st.stop()

taux_moyen = df_util["Taux_pct"].mean()
ligne_max = df_util.loc[df_util["Taux_pct"].idxmax()]
ligne_min = df_util.loc[df_util["Taux_pct"].idxmin()]
nb_saturees = (df_util["Taux_pct"] >= 90).sum()

kpi_row([
    {"label": "Taux moyen d'utilisation", "value": f"{taux_moyen:.1f} %"},
    {"label": "Ligne la plus chargée", "value": f"{ligne_max['Ligne']} (S{ligne_max['Semaine']})", "tone": "neg"},
    {"label": "Ligne la moins chargée", "value": f"{ligne_min['Ligne']} (S{ligne_min['Semaine']})", "tone": "pos"},
    {"label": "Créneaux ≥ 90 % charge", "value": int(nb_saturees), "tone": "neg" if nb_saturees else "pos"},
])

st.subheader("Carte de charge — Ligne × Semaine")
pivot = df_util.pivot(index="Ligne", columns="Semaine", values="Taux_pct")
fig = px.imshow(
    pivot, text_auto=".0f", aspect="auto",
    color_continuous_scale=["#24272C", "#6B7178", "#C8102E"],
    labels=dict(color="Taux %"),
)
fig.update_xaxes(title="Semaine")
fig.update_yaxes(title="Ligne")
st.plotly_chart(style_fig(fig, height=380), use_container_width=True)

st.subheader("Détail par ligne et semaine")
render_filtered_table(df_util, filter_cols=["Ligne", "Semaine"], key_prefix="util")
