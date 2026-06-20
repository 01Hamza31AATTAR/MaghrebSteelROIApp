"""
app.py — Point d'entrée Streamlit — Maghreb Steel
══════════════════════════════════════════════════
Lance avec :  streamlit run app.py
"""
from datetime import datetime
from pathlib import Path

import streamlit as st

import data_layer as dl
from components import inject_css, sidebar_brand

st.set_page_config(
    page_title="Maghreb Steel — Pilotage Capacité-Commande",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)
inject_css()

# ═══════════════════════════════════════════
# Sidebar — statut des données + lancement optimisation
# ═══════════════════════════════════════════
with st.sidebar:
    sidebar_brand()
    st.divider()

    results_path = dl.get_results_path()
    if dl.results_available(results_path):
        ts = datetime.fromtimestamp(dl.last_updated(results_path))
        st.success(f"Résultats disponibles\n\nMAJ : {ts:%d/%m/%Y %H:%M}")
    else:
        st.warning("Aucun résultat — lancez une optimisation.")

    st.divider()
    st.subheader("⚙️ Optimisation")
    input_path = st.text_input("Fichier de données source", value=dl.get_source_path())

    if st.button("🚀 Lancer l'optimisation", use_container_width=True, type="primary"):
        if not Path(input_path).exists():
            st.error(f"Fichier introuvable : `{input_path}`")
        else:
            progress = st.progress(0, text="Démarrage…")

            def _cb(frac: float, label: str):
                progress.progress(min(frac, 1.0), text=label)

            try:
                dl.run_optimization(
                    input_path=input_path,
                    output_path=results_path,
                    progress_callback=_cb,
                )
                dl.clear_cache()
                st.success("Optimisation terminée — résultats sauvegardés.")
                st.rerun()
            except Exception as exc:  # noqa: BLE001
                st.error(f"Échec de l'optimisation : {exc}")

    st.caption(f"Source des résultats : `{results_path}`")

# ═══════════════════════════════════════════
# Navigation — groupes de pages
# ═══════════════════════════════════════════
pages_vue_ensemble = [
    st.Page("pages/p1_synthese.py",      title="Synthèse",            icon="📊", default=True),
    st.Page("pages/p2_commandes.py",     title="Commandes",           icon="📦"),
    st.Page("pages/p3_plan_marche.py",   title="Plan de marche",      icon="🗓️"),
    st.Page("pages/p4_utilisation.py",   title="Utilisation lignes",  icon="⚙️"),
    st.Page("pages/p5_shadow_prices.py", title="Prix fictifs & sensi", icon="🎯"),
    st.Page("pages/p6_marge_famille.py", title="Marge par famille",   icon="💰"),
]

pages_bonus = [
    st.Page("pages/p7_tarification.py", title="B1 · Tarification d'opportunité", icon="🏷️"),
    st.Page("pages/p8_backlogging.py",  title="B2 · Backlogging",                icon="⏳"),
    st.Page("pages/p9_stockage.py",     title="B3 · Coûts de stockage",          icon="🏬"),
    st.Page("pages/p10_campagnes.py",   title="B4 · Campagnes (MILP)",           icon="🔁"),
    st.Page("pages/p11_bb.py",          title="B5 · Branch-and-Bound",           icon="🌳"),
]

nav = st.navigation({
    "Vue d'ensemble":     pages_vue_ensemble,
    "Bonus — Extensions": pages_bonus,
})
nav.run()
