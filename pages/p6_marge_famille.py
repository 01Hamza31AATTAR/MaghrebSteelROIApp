import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, empty_state

inject_css()
page_header("Marge par famille", "Décomposition du compte de résultat par famille de produit")

df_m = dl.cached_load_sheet("marge_famille")
if df_m is None or df_m.empty:
    empty_state()
    st.stop()

meilleure = df_m.loc[df_m["Marge_pct"].idxmax()]
pire = df_m.loc[df_m["Marge_pct"].idxmin()]

kpi_row([
    {"label": "Marge nette totale", "value": f"{df_m['Marge_Nette_MAD'].sum():,.0f} MAD", "tone": "pos"},
    {"label": "Marge moyenne", "value": f"{df_m['Marge_pct'].mean():.1f} %"},
    {"label": "Famille la plus rentable", "value": f"{meilleure['Famille']} ({meilleure['Marge_pct']:.1f}%)", "tone": "pos"},
    {"label": "Famille la moins rentable", "value": f"{pire['Famille']} ({pire['Marge_pct']:.1f}%)", "tone": "neg"},
])

col1, col2 = st.columns([1.3, 1])

with col1:
    st.subheader("CA, coûts et marge nette par famille")
    fig = px.bar(
        df_m, x="Famille", y=["CA_MAD", "Cout_Total_MAD", "Marge_Nette_MAD"],
        barmode="group",
    )
    fig.update_layout(legend_title_text="")
    st.plotly_chart(style_fig(fig, height=380), use_container_width=True)

with col2:
    st.subheader("Répartition de la marge nette")
    fig = px.pie(df_m, names="Famille", values="Marge_Nette_MAD", hole=0.5)
    st.plotly_chart(style_fig(fig, height=380), use_container_width=True)

st.subheader("Détail complet par famille")
st.dataframe(df_m, use_container_width=True, hide_index=True)
