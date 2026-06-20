"""
data_layer.py — Couche d'accès aux données — Maghreb Steel
═══════════════════════════════════════════════════════════
Point unique par lequel l'interface lit et écrit les données.

Aujourd'hui   : backend Excel (openpyxl).
Demain        : remplacer le CORPS des fonctions ci-dessous par des
                requêtes PostgreSQL — les pages Streamlit n'ont
                RIEN à changer car elles n'appellent que :
                    - load_sheet(name) / load_all_results()
                    - save_results(dict_de_dataframes)
                    - run_optimization(...)

Aucune autre partie du projet ne doit lire/écrire un fichier Excel
directement : tout passe par ce module.
"""
from __future__ import annotations

import os
import time
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
import streamlit as st

import main as opti  # le moteur de calcul LP/MILP — inchangé

# ═══════════════════════════════════════════
# 0. CONFIGURATION DES CHEMINS
# ═══════════════════════════════════════════
# Ces deux constantes sont les SEULS endroits à modifier pour changer
# l'emplacement des fichiers. Pour passer à Postgres, elles seraient
# remplacées par une chaîne de connexion (ex: variable d'env DATABASE_URL).

DEFAULT_SOURCE_PATH  = "Donnees_MaghrebSteel.xlsx"   # données d'entrée (commandes, cadences...)
DEFAULT_RESULTS_PATH = "outputs/resultats.xlsx"      # résultats d'optimisation


def get_source_path() -> str:
    return os.environ.get("GEEX_OPTI_SOURCE_PATH", DEFAULT_SOURCE_PATH)


def get_results_path() -> str:
    return os.environ.get("GEEX_OPTI_RESULTS_PATH", DEFAULT_RESULTS_PATH)


# ═══════════════════════════════════════════
# 1. NOMS DE FEUILLES / TABLES LOGIQUES
# ═══════════════════════════════════════════
# Clé logique -> nom de feuille Excel (deviendrait nom de table Postgres).
# Les pages Streamlit n'utilisent JAMAIS de chaîne en dur : elles passent
# toujours par SHEETS['ma_cle'].

SHEETS = {
    "synthese":               "Synthese",
    "commandes":               "Commandes",
    "commandes_refusees":      "Commandes_refusees",
    "plan_marche":              "Plan_de_marche",
    "utilisation":              "Utilisation",
    "shadow_prices":            "Shadow_prices",
    "sensibilite":              "Sensibilite",
    "marge_famille":            "Marge_par_famille",
    "tarification":             "Tarification_opportunite",
    "backlogging_commandes":    "Backlogging_commandes",
    "backlogging_synthese":     "Backlogging_synthese",
    "stockage_commandes":       "Stockage_commandes",
    "stockage_detail":          "Stockage_detail",
    "stockage_synthese":        "Stockage_synthese",
    "stockage_taux_hf":         "Stockage_taux_hf",
    "campagnes_commandes":      "Campagnes_commandes",
    "campagnes_plan":           "Campagnes_plan",
    "campagnes_synthese":       "Campagnes_synthese",
    "branch_and_bound":         "BranchAndBound",
}


# ═══════════════════════════════════════════
# 2. LECTURE DES RÉSULTATS (swappable)
# ═══════════════════════════════════════════

def load_all_results(path: Optional[str] = None) -> dict[str, pd.DataFrame]:
    """
    Charge toutes les feuilles de résultats.
    Backend actuel : Excel. À remplacer par un SELECT * par table Postgres.
    """
    path = path or get_results_path()
    if not Path(path).exists():
        return {}
    xl = pd.ExcelFile(path)
    return {name: xl.parse(name) for name in xl.sheet_names}


def load_sheet(sheet_key: str, path: Optional[str] = None) -> Optional[pd.DataFrame]:
    """
    Charge une feuille unique par sa CLÉ LOGIQUE (voir SHEETS).
    Retourne None si le fichier ou la feuille n'existe pas encore
    (ex : aucune optimisation n'a encore été lancée).
    Backend actuel : Excel. À remplacer par `SELECT * FROM <table>`.
    """
    path = path or get_results_path()
    sheet_name = SHEETS.get(sheet_key, sheet_key)
    if not Path(path).exists():
        return None
    try:
        return pd.read_excel(path, sheet_name=sheet_name)
    except ValueError:
        return None  # feuille absente de ce fichier


def results_available(path: Optional[str] = None) -> bool:
    path = path or get_results_path()
    return Path(path).exists()


def last_updated(path: Optional[str] = None) -> Optional[float]:
    path = path or get_results_path()
    return Path(path).stat().st_mtime if Path(path).exists() else None


# --- Versions mises en cache pour l'interface Streamlit ---------------
# La mise en cache reste ici (et non dans les pages) pour que les pages
# n'aient jamais à savoir COMMENT les données sont chargées.

@st.cache_data(show_spinner=False)
def cached_load_sheet(sheet_key: str, path: Optional[str] = None) -> Optional[pd.DataFrame]:
    return load_sheet(sheet_key, path)


@st.cache_data(show_spinner=False)
def cached_load_all(path: Optional[str] = None) -> dict[str, pd.DataFrame]:
    return load_all_results(path)


def clear_cache():
    st.cache_data.clear()


# ═══════════════════════════════════════════
# 3. SAUVEGARDE DES RÉSULTATS (swappable)
# ═══════════════════════════════════════════

def save_results(sheets: dict[str, pd.DataFrame], path: Optional[str] = None) -> str:
    """
    Sauvegarde un dictionnaire {cle_logique: DataFrame}.
    Backend actuel : un classeur Excel, une feuille par clé.
    Pour Postgres : remplacer le corps par une boucle de
    `df.to_sql(table_name, engine, if_exists="replace")`.
    """
    path = path or get_results_path()
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for key, df in sheets.items():
            sheet_name = SHEETS.get(key, key)[:31]  # limite Excel
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    return path


# ═══════════════════════════════════════════
# 4. PIPELINE D'OPTIMISATION COMPLET
# ═══════════════════════════════════════════
# Reproduit fidèlement la séquence du bloc `if __name__ == "__main__"`
# de main.py (qui reste inchangé), étape par étape, avec callback de
# progression pour la barre de l'interface.

PIPELINE_STEPS = [
    "Chargement des données",
    "Résolution baseline (LP)",
    "Extraction des résultats",
    "Marge nette par famille",
    "Tarification d'opportunité (B1)",
    "Validation a posteriori",
    "Backlogging (B2)",
    "Coûts de stockage (B3)",
    "Campagnes MILP (B4)",
    "Branch-and-Bound (B5)",
    "Analyse de sensibilité",
    "Sauvegarde des résultats",
]


def run_optimization(
    input_path: Optional[str] = None,
    output_path: Optional[str] = None,
    progress_callback: Optional[Callable[[float, str], None]] = None,
) -> dict[str, pd.DataFrame]:
    """
    Exécute tout le pipeline de calcul (main.py) et sauvegarde les
    résultats via save_results(). Retourne le dictionnaire des
    DataFrames produits, clé identique à SHEETS.
    """
    input_path = input_path or get_source_path()
    output_path = output_path or get_results_path()

    n_steps = len(PIPELINE_STEPS)

    def _tick(i: int):
        if progress_callback:
            progress_callback(i / n_steps, PIPELINE_STEPS[i])

    _tick(0)
    (df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc, dispo,
     stk_pk_init, stk_pk_min, stk_pk_max,
     stk_pf_init, stk_pf_min, stk_pf_max, get_arret) = opti.load_data(input_path)

    _tick(1)
    model, x, liv, S_pf, _elapsed = opti.build_and_solve(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc, dispo,
        stk_pk_init, stk_pk_min, stk_pf_init, stk_pf_min, stk_pf_max, get_arret)
    baseline_obj = opti.value(model.objective)

    _tick(2)
    df_res, df_plan, df_util, df_sp = opti.extract_results(
        model, x, liv, df_main, cad_dict, rho, get_arret)

    _tick(3)
    df_marge_fam = opti.compute_marge_par_famille(
        model, x, liv, df_main, rho, tau_chute, get_cv, get_prix_hrc)

    _tick(4)
    df_tarif = opti.compute_tarification_opportunite(
        model, x, df_res, df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc)

    _tick(5)
    opti.validate_solution(model, x, liv, S_pf, df_main, cad_dict, rho, get_arret,
                            dispo, stk_pk_init, stk_pk_min, stk_pf_min, stk_pf_max)

    _tick(6)
    _model_bl, df_res_bl, df_comp_bl = opti.compute_backlogging(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret)

    _tick(7)
    _model_st, df_res_st, df_stock_t, df_comp_st, df_h = opti.compute_cout_stockage(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
        valeur_mode="cout_revient", df_marge_fam=df_marge_fam)

    _tick(8)
    _model_milp, df_res_milp, df_campagnes, df_comp_milp = opti.compute_campagnes_milp(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret)

    _tick(9)
    df_bb, _log_lines = opti.analyse_branch_and_bound(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret)

    _tick(10)
    df_sensi = opti.compute_sensibilite(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
        df_sp, baseline_obj)

    _tick(11)
    df_synthese = pd.DataFrame({
        "Indicateur": ["Marge totale (MAD)", "Taux de service (%)",
                       "Commandes honorées", "Commandes partielles", "Commandes refusées"],
        "Valeur": [f"{baseline_obj:,.0f}",
                   f"{df_res['Livree'].sum() / df_res['Tonnage'].sum() * 100:.1f}%",
                   (df_res["Statut"] == "Honorée").sum(),
                   (df_res["Statut"] == "Partielle").sum(),
                   (df_res["Statut"] == "Refusée").sum()],
    })

    sheets = {
        "synthese":             df_synthese,
        "commandes":            df_res,
        "commandes_refusees":   df_res[df_res["Statut"] == "Refusée"],
        "plan_marche":          df_plan,
        "utilisation":          df_util,
        "shadow_prices":        df_sp,
        "sensibilite":          df_sensi,
        "marge_famille":        df_marge_fam,
        "tarification":         df_tarif,
        "backlogging_commandes": df_res_bl,
        "backlogging_synthese": df_comp_bl,
        "stockage_commandes":   df_res_st,
        "stockage_detail":      df_stock_t,
        "stockage_synthese":    df_comp_st,
        "stockage_taux_hf":     df_h,
        "campagnes_commandes":  df_res_milp,
        "campagnes_plan":       df_campagnes,
        "campagnes_synthese":   df_comp_milp,
        "branch_and_bound":     df_bb,
    }

    save_results(sheets, output_path)
    if progress_callback:
        progress_callback(1.0, "Terminé")
    return sheets
