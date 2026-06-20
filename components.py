"""
components.py — Composants d'interface réutilisables — Maghreb Steel
══════════════════════════════════════════════════════════════════
Tout le HTML brut est défini dans templates/*.html (pas de string HTML
en dur dans le code Python) afin de garder les fichiers de page lisibles.
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

TEMPLATES_DIR = Path(__file__).parent / "templates"

# Palette du thème — partagée par tous les graphiques Plotly
MS_RED        = "#C8102E"
MS_RED_DARK   = "#8F0B20"
MS_GREY_700   = "#343840"
MS_GREY_500   = "#6B7178"
MS_GREY_300   = "#A9ADB3"
MS_WHITE      = "#F2F3F4"
MS_GREEN      = "#3FA66A"

DISCRETE_PALETTE = [MS_RED, MS_GREY_500, "#D94F4F", "#8B9094", MS_RED_DARK, "#C7CBCF"]


def load_template(name: str) -> str:
    return (TEMPLATES_DIR / name).read_text(encoding="utf-8")


def inject_css():
    """Injecte le CSS du thème une seule fois par session."""
    if st.session_state.get("_ms_css_injected"):
        return
    st.markdown(load_template("styles.html"), unsafe_allow_html=True)
    st.session_state["_ms_css_injected"] = True


def page_header(title: str, subtitle: str = ""):
    st.markdown(
        load_template("header.html").format(title=title, subtitle=subtitle),
        unsafe_allow_html=True,
    )


def sidebar_brand():
    st.markdown(load_template("sidebar_brand.html"), unsafe_allow_html=True)


def kpi_row(items: list[dict]):
    """
    items: liste de dicts {'label', 'value', 'delta' (optionnel), 'tone' (optionnel: pos/neg/neutral)}
    """
    card_tpl = load_template("kpi_card.html")
    cards_html = "".join(
        card_tpl.format(
            label=it.get("label", ""),
            value=it.get("value", ""),
            delta=it.get("delta", ""),
            tone=it.get("tone", "neutral"),
        )
        for it in items
    )
    st.markdown(load_template("kpi_row.html").format(cards=cards_html), unsafe_allow_html=True)


def empty_state(message: str = "Aucun résultat disponible — lancez une optimisation depuis le menu de gauche."):
    st.info(message)


def style_fig(fig: go.Figure, height: int = 380) -> go.Figure:
    """Applique le thème visuel Maghreb Steel à une figure Plotly."""
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=40, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter, sans-serif", color=MS_WHITE, size=13),
        title_font=dict(family="Oswald, sans-serif", size=15, color=MS_WHITE),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
        colorway=DISCRETE_PALETTE,
    )
    fig.update_xaxes(gridcolor=MS_GREY_700, zerolinecolor=MS_GREY_700)
    fig.update_yaxes(gridcolor=MS_GREY_700, zerolinecolor=MS_GREY_700)
    return fig


def render_filtered_table(
    df: pd.DataFrame,
    filter_cols: Optional[list[str]] = None,
    key_prefix: str = "tbl",
):
    """Affiche une table avec filtres multiselect optionnels au-dessus."""
    if df is None or df.empty:
        empty_state("Aucune donnée à afficher pour cette sélection.")
        return df

    filtered = df.copy()
    filter_cols = filter_cols or []
    if filter_cols:
        cols = st.columns(len(filter_cols))
        for c, colname in zip(cols, filter_cols):
            if colname not in df.columns:
                continue
            options = sorted(df[colname].dropna().unique().tolist())
            chosen = c.multiselect(colname, options, key=f"{key_prefix}_{colname}")
            if chosen:
                filtered = filtered[filtered[colname].isin(chosen)]

    st.dataframe(filtered, use_container_width=True, hide_index=True)
    return filtered
