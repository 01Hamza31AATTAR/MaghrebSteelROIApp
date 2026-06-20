import plotly.express as px
import streamlit as st

import data_layer as dl
from components import inject_css, page_header, kpi_row, style_fig, render_filtered_table, empty_state

inject_css()
page_header("Prix fictifs & sensibilité", "Goulots d'étranglement structurels et robustesse de la solution")

df_sp = dl.cached_load_sheet("shadow_prices")
df_sensi = dl.cached_load_sheet("sensibilite")

tab_sp, tab_sensi = st.tabs(["🎯 Prix fictifs (shadow prices)", "🧪 Analyse de sensibilité"])

with tab_sp:
    if df_sp is None or df_sp.empty:
        empty_state()
    else:
        contrainte_max = df_sp.loc[df_sp["Shadow_price"].idxmax()]
        kpi_row([
            {"label": "Contraintes actives", "value": len(df_sp)},
            {"label": "Contrainte la plus chère", "value": contrainte_max["Contrainte"], "tone": "neg"},
            {"label": "Valeur max (MAD/unité)", "value": f"{contrainte_max['Shadow_price']:,.0f}"},
        ])
        st.subheader("Top contraintes par prix fictif")
        top = df_sp.head(15).sort_values("Shadow_price")
        fig = px.bar(top, x="Shadow_price", y="Contrainte", orientation="h")
        st.plotly_chart(style_fig(fig, height=420), use_container_width=True)
        st.subheader("Détail")
        render_filtered_table(df_sp, key_prefix="sp")

with tab_sensi:
    if df_sensi is None or df_sensi.empty:
        empty_state()
    else:
        pire = df_sensi.loc[df_sensi["Delta_vs_baseline"].idxmin()]
        kpi_row([
            {"label": "Scénarios testés", "value": len(df_sensi)},
            {"label": "Scénario le plus impactant", "value": pire["Scenario"], "tone": "neg"},
            {"label": "Impact (MAD)", "value": f"{pire['Delta_vs_baseline']:,.0f}", "tone": "neg"},
        ])
        st.subheader("Impact de chaque scénario sur la marge")
        fig = px.bar(
            df_sensi.sort_values("Delta_vs_baseline"),
            x="Delta_vs_baseline", y="Scenario", orientation="h",
            color="Delta_vs_baseline", color_continuous_scale=["#C8102E", "#6B7178"],
        )
        fig.update_layout(coloraxis_showscale=False)
        st.plotly_chart(style_fig(fig, height=380), use_container_width=True)
        st.subheader("Détail des scénarios")
        st.dataframe(df_sensi, use_container_width=True, hide_index=True)
