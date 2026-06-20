"""
main.py — Simulateur Capacité-Commande — Maghreb Steel
Résolution LP avec PuLP (CBC), export des résultats vers Excel
"""
import time as tm
import numpy as np
import pandas as pd
from pulp import *

# ═══════════════════════════════════════════
# 1. CHARGEMENT DES DONNÉES
# ═══════════════════════════════════════════

def load_data(path="Donnees_MaghrebSteel.xlsx"):
    xl = pd.ExcelFile(path)
    
    # Commandes
    df_cmd = xl.parse("Commandes", header=1)
    df_cmd.columns = ['ID','Client','Famille','Grade','Epaisseur','Largeur','Tonnage','PrixVente','SemaineLiv','Priorite']
    df_cmd = df_cmd[df_cmd['ID'].astype(str).str.startswith('CMD')].reset_index(drop=True)
    for c in ['Tonnage','PrixVente','SemaineLiv','Epaisseur','Largeur']:
        df_cmd[c] = pd.to_numeric(df_cmd[c], errors='coerce')
    df_cmd = df_cmd[df_cmd['Famille'].isin(['CRC','HDG','PPGI','BACR','HRC DEC'])].reset_index(drop=True)
    
    # Cadences
    df_cad = xl.parse("Cadences", header=1)
    df_cad.columns = ['Ligne','HRC_DEC','CRC','HDG','PPGI','BACR'] + list(df_cad.columns[6:])
    df_cad = df_cad[df_cad['Ligne'].isin(['PK','CRMA','CRMB','BAF','SKP','LGA','LGB'])].reset_index(drop=True)
    for c in ['HRC_DEC','CRC','HDG','PPGI','BACR']:
        df_cad[c] = pd.to_numeric(df_cad[c], errors='coerce')
    
    cad_dict = {}
    for _, row in df_cad.iterrows():
        for fam, col in [('HRC DEC','HRC_DEC'),('CRC','CRC'),('HDG','HDG'),('PPGI','PPGI'),('BACR','BACR')]:
            v = row[col]
            if pd.notna(v) and v > 0:
                cad_dict[(row['Ligne'], fam)] = float(v)
    
    # Rendements
    df_rend = xl.parse("Rendements", header=1)
    df_rend.columns = ['Process','Rendement','Chute','Declasse','NonConf'] + list(df_rend.columns[5:])
    df_rend = df_rend[df_rend['Process'].isin(['PK','CRMA','CRMB','BAF','SKP','LGA','LGB'])].reset_index(drop=True)
    def to_float(v):
        try: return float(str(v).replace('%','').strip())
        except: return np.nan
    df_rend['Rendement'] = df_rend['Rendement'].apply(to_float)
    df_rend['Chute']     = df_rend['Chute'].apply(to_float)
    if df_rend['Rendement'].max() > 1:
        df_rend['Rendement'] /= 100
        df_rend['Chute'] /= 100
    rho       = dict(zip(df_rend['Process'], df_rend['Rendement']))
    tau_chute = dict(zip(df_rend['Process'], df_rend['Chute']))
    
    # Coûts variables
    df_cv = xl.parse("Couts_Variables", header=1)
    ep_cols = ['<0.3','0.3-0.4','0.4-0.5','0.5-0.7','0.7-1.0','1.0-1.5','>1.5']
    df_cv.columns = ['Process'] + ep_cols + list(df_cv.columns[8:])
    proc_cv = ['PK','CRMA','CRMB','BAF','SKP','LGA-HDG','LGA-PPGI','LGA-BACR','LGB-HDG','LGB-BACR']
    df_cv = df_cv[df_cv['Process'].isin(proc_cv)].reset_index(drop=True)
    for c in ep_cols:
        df_cv[c] = pd.to_numeric(df_cv[c], errors='coerce')
    
    def get_cv(process, epaisseur):
        ep = float(epaisseur)
        if ep < 0.3:   col = '<0.3'
        elif ep < 0.4: col = '0.3-0.4'
        elif ep < 0.5: col = '0.4-0.5'
        elif ep < 0.7: col = '0.5-0.7'
        elif ep < 1.0: col = '0.7-1.0'
        elif ep < 1.5: col = '1.0-1.5'
        else:          col = '>1.5'
        row = df_cv[df_cv['Process'] == process]
        return float(row[col].values[0]) if not row.empty else 0.0
    
    # Prix HRC
    df_hrc_raw = xl.parse("Prix_HRC", header=None)
    grades  = ['DC01','DD13','DX51','DX52','S320']
    widths  = [1020, 1090, 1100, 1140, 1250, 1280, 1320]
    price_rows = {}
    dispo      = {}
    for _, row in df_hrc_raw.iterrows():
        key = str(row[0]).strip() if pd.notna(row[0]) else ''
        if key in grades:
            vals = [v for v in row[1:] if pd.notna(v)]
            if len(vals) > 1:
                price_rows[key] = {str(widths[i]): float(vals[i]) for i in range(min(len(widths), len(vals)))}
            elif len(vals) == 1:
                dispo[key] = float(vals[0])
    
    def get_prix_hrc(grade, largeur):
        row = price_rows.get(grade, {})
        return row.get(str(int(float(largeur))), 6000.0)
    
    # Stocks initiaux
    df_stk  = xl.parse("Stocks_Initiaux", header=None)
    stk_pk_init = {}; stk_pk_min = {}; stk_pk_max = {}
    stk_pf_init = {}; stk_pf_min = {}; stk_pf_max = {}
    fams_pf = ['HRC DEC','CRC','HDG','PPGI','BACR']
    for _, row in df_stk.iterrows():
        key = str(row[0]).strip() if pd.notna(row[0]) else ''
        try: v1, v2, v3 = float(row[1]), float(row[2]), float(row[3])
        except: continue
        if key in grades:
            stk_pk_init[key] = v1; stk_pk_min[key] = v2; stk_pk_max[key] = v3
        if key in fams_pf:
            stk_pf_init[key] = v1; stk_pf_min[key] = v2; stk_pf_max[key] = v3
    
    # Arrêts planifiés
    df_arr = xl.parse("Arrets_Planifies", header=1)
    df_arr.columns = ['Ligne','S1','S2','S3','S4'] + list(df_arr.columns[5:])
    df_arr = df_arr[df_arr['Ligne'].isin(['PK','CRMA','CRMB','BAF','SKP','LGA','LGB'])].reset_index(drop=True)
    for c in ['S1','S2','S3','S4']:
        df_arr[c] = pd.to_numeric(df_arr[c], errors='coerce').fillna(0)
    
    def get_arret(ligne, semaine):
        row = df_arr[df_arr['Ligne'] == ligne]
        return float(row[f'S{semaine}'].values[0]) if not row.empty else 0
    
    return (df_cmd, cad_dict, rho, tau_chute, get_cv, get_prix_hrc, dispo,
            stk_pk_init, stk_pk_min, stk_pk_max,
            stk_pf_init, stk_pf_min, stk_pf_max, get_arret)

# ═══════════════════════════════════════════
# 2. ROUTES MÉTALLURGIQUES
# ═══════════════════════════════════════════

ROUTES = {
    'CRC': {
        'CRC_CRMB': [('PK','PK'),('CRMB','CRMB'),('BAF','BAF'),('SKP','SKP')]
    },
    'HDG': {
        'HDG_CRMA_LGA': [('PK','PK'),('CRMA','CRMA'),('LGA','LGA-HDG')],
        'HDG_CRMA_LGB': [('PK','PK'),('CRMA','CRMA'),('LGB','LGB-HDG')],
        'HDG_CRMB_LGA': [('PK','PK'),('CRMB','CRMB'),('LGA','LGA-HDG')],
        'HDG_CRMB_LGB': [('PK','PK'),('CRMB','CRMB'),('LGB','LGB-HDG')],
    },
    'PPGI': {
        'PPGI_CRMA_LGA': [('PK','PK'),('CRMA','CRMA'),('LGA','LGA-PPGI')],
        'PPGI_CRMB_LGA': [('PK','PK'),('CRMB','CRMB'),('LGA','LGA-PPGI')],
    },
    'BACR': {
        'BACR_A':          [('PK','PK'),('CRMB','CRMB'),('BAF','BAF'),('LGB','LGB-BACR')],
        'BACR_B_CRMA_LGA': [('PK','PK'),('CRMA','CRMA'),('LGA','LGA-BACR')],
        'BACR_B_CRMA_LGB': [('PK','PK'),('CRMA','CRMA'),('LGB','LGB-BACR')],
        'BACR_B_CRMB_LGA': [('PK','PK'),('CRMB','CRMB'),('LGA','LGA-BACR')],
        'BACR_B_CRMB_LGB': [('PK','PK'),('CRMB','CRMB'),('LGB','LGB-BACR')],
    },
    'HRC DEC': {
        'HRC_DEC_PK': [('PK','PK')]
    }
}

V_CHUTE    = 1800.0
C_ZINC     = 18000.0;  A_ZINC     = 0.025
C_PEINTURE = 12000.0;  A_PEINTURE = 0.010

# Coûts physiques différenciés par famille (MAD/T/semaine)
# Décomposition : espace + manutention + protection
COUT_PHYSIQUE_STOCK = {
    'HRC DEC': 15,   # 3 espace + 5 manutention + 7 protection
    'CRC':     20,   # 3 espace + 8 manutention + 9 protection
    'HDG':     30,   # 3 espace + 12 manutention + 15 protection
    'PPGI':    45,   # 3 espace + 15 manutention + 25 protection (UV + mousse)
    'BACR':    40,   # 3 espace + 15 manutention + 22 protection (anti-rayures)
}

# Taux d'immobilisation capital : 10%/an → 10/52 ≈ 0.1923%/semaine
TAUX_IMMO_SEMAINE = 0.10 / 52

# Lignes concernées par les campagnes + familles qu'elles peuvent traiter
LIGNES_CAMPAGNE = {
    'LGA': ['HDG', 'PPGI', 'BACR'],
    'LGB': ['HDG', 'BACR'],
}

# Coût de changement de campagne par ligne (MAD)
# Justification : 1 jour de production perdu = cadence × valeur journée
# LGA : cadence HDG ≈ 350 T/j × prix moyen HDG ≈ 10 000 MAD/T × 1j = ~350 000 MAD
# LGB : cadence HDG ≈ 300 T/j × même logique = ~300 000 MAD
# On prend une estimation conservative à 50% (nettoyage partiel possible)
COUT_CHANGEMENT = {
    'LGA': 175_000,   # MAD par changement de famille
    'LGB': 150_000,   # MAD par changement de famille
}

PRIO_MAP = {
    'Haute':   1, 'haute':   1,
    'Normale': 2, 'normale': 2,
    'Basse':   3, 'basse':   3,
    'Moyenne': 2, 'moyenne': 2,  # sécurité au cas où
}

def parse_prio(val):
    if pd.isna(val):
        return 2
    s = str(val).strip()
    if s in PRIO_MAP:
        return PRIO_MAP[s]
    try:
        return int(float(s))
    except:
        return 2

def rho_route(route_steps, rho):
    r = 1.0
    for (l, _) in route_steps:
        r *= rho[l]
    return r

def gamma_lr(route_steps, target_ligne, rho):
    idx = [i for i, (l, _) in enumerate(route_steps) if l == target_ligne]
    if not idx: return 1.0
    g = 1.0
    for (l, _) in route_steps[idx[0] + 1:]:
        g *= rho[l]
    return g

def ct_route(route_steps, epaisseur, get_cv, rho):
    total = 0.0
    for (ligne, cv_key) in route_steps:
        g = gamma_lr(route_steps, ligne, rho)
        total += get_cv(cv_key, epaisseur) / g
    return total

def vc_route(route_steps, epaisseur, rho, tau_chute):
    total = 0.0
    for (ligne, _) in route_steps:
        g   = gamma_lr(route_steps, ligne, rho)
        inp = 1.0 / (rho[ligne] * g)
        total += V_CHUTE * tau_chute[ligne] * inp
    return total

def cz_fam(fam):
    return C_ZINC * A_ZINC if fam in ['HDG', 'PPGI'] else 0.0

def cp_fam(fam):
    return C_PEINTURE * A_PEINTURE if fam == 'PPGI' else 0.0

# ═══════════════════════════════════════════
# 3. CONSTRUCTION ET RÉSOLUTION DU MODÈLE LP
# ═══════════════════════════════════════════

def build_and_solve(df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
                    dispo, stk_pk_init, stk_pk_min,
                    stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
                    J=7, T_periods=(1, 2, 3, 4),
                    arret_delta=None, hrc_factor=1.0):
    fams_main  = ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']
    lignes     = ['PK', 'CRMA', 'CRMB', 'BAF', 'SKP', 'LGA', 'LGB']
    grades_all = ['DC01', 'DD13', 'DX51', 'DX52', 'S320']

    model = LpProblem("MaghrebSteel_CapComm", LpMaximize)

    x     = {}
    liv   = {}
    S_pf  = {}

    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']
        for rname in ROUTES[fam]:
            for t in T_periods:
                x[(cid, rname, t)] = LpVariable(f"x_{cid}_{rname}_{t}", lowBound=0)
        for t in T_periods:
            liv[(cid, t)] = LpVariable(f"l_{cid}_{t}", lowBound=0)
    for fam in fams_main:
        for t in T_periods:
            S_pf[(fam, t)] = LpVariable(f"S_{fam}_{t}", lowBound=0)

    # Objectif
    ot = []
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; pv = float(row['PrixVente'])
        for t in T_periods:
            ot.append(pv * liv[(cid, t)])
        for rname, rsteps in ROUTES[fam].items():
            for t in T_periods:
                ep = float(row['Epaisseur']); g = row['Grade']; w = int(float(row['Largeur']))
                rr  = rho_route(rsteps, rho)
                cm  = get_prix_hrc(g, w) * hrc_factor / rr
                c_t = ct_route(rsteps, ep, get_cv, rho)
                v_c = vc_route(rsteps, ep, rho, tau_chute)
                ot.append(-(cm + c_t + cz_fam(fam) + cp_fam(fam) - v_c) * x[(cid, rname, t)])
    model += lpSum(ot), "Marge_totale"

    # Capacité
    for ligne in lignes:
        for t in T_periods:
            extra = (arret_delta or {}).get((ligne, t), 0)
            arr   = get_arret(ligne, t) + extra
            terms = []
            for _, row in df_main.iterrows():
                cid = row['ID']; fam = row['Famille']
                for rname, rsteps in ROUTES[fam].items():
                    if ligne not in [l for (l, _) in rsteps]: continue
                    cad = cad_dict.get((ligne, fam))
                    if not cad: continue
                    gl = gamma_lr(rsteps, ligne, rho)
                    terms.append((1.0 / (gl * cad)) * x[(cid, rname, t)])
            if terms:
                model += lpSum(terms) <= max(0, J - arr), f"Cap_{ligne}_S{t}"

    # Stock PF
    for fam in fams_main:
        fam_cmds = df_main[df_main['Famille'] == fam]['ID'].tolist()
        for t in T_periods:
            prod_t = [x[(cid, rname, t)] for cid in fam_cmds for rname in ROUTES[fam]]
            del_t  = [liv[(cid, t)] for cid in fam_cmds]
            S0 = stk_pf_init.get(fam, 0) if t == 1 else S_pf[(fam, t - 1)]
            model += S_pf[(fam, t)] == S0 + lpSum(prod_t) - lpSum(del_t), f"StockBal_{fam}_S{t}"
            model += S_pf[(fam, t)] >= stk_pf_min.get(fam, 0),   f"StockMin_{fam}_S{t}"
            model += S_pf[(fam, t)] <= stk_pf_max.get(fam, 9999), f"StockMax_{fam}_S{t}"

    # Livraison
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; D = float(row['Tonnage'])
        model += lpSum(liv[(cid, t)] for t in T_periods) <= D, f"MaxDeliv_{cid}"
        for t in T_periods:
            cum_p = [x[(cid, rname, tau)] for rname in ROUTES[fam] for tau in range(1, t + 1)]
            cum_l = [liv[(cid, tau)] for tau in range(1, t + 1)]
            model += lpSum(cum_l) <= lpSum(cum_p), f"ProdBefore_{cid}_S{t}"

    # HRC
    for g in grades_all:
        terms = []
        for _, row in df_main.iterrows():
            if row['Grade'] != g: continue
            cid = row['ID']; fam = row['Famille']
            for rname, rsteps in ROUTES[fam].items():
                rr = rho_route(rsteps, rho)
                for t in T_periods:
                    terms.append((1.0 / rr) * x[(cid, rname, t)])
        if terms:
            avail = dispo[g] + stk_pk_init.get(g, 0) - stk_pk_min.get(g, 0)
            model += lpSum(terms) <= avail, f"HRC_{g}"

    # Résolution
    t0 = tm.time()
    model.solve(PULP_CBC_CMD(msg=0, timeLimit=300))
    elapsed = tm.time() - t0

    print(f"Statut : {LpStatus[model.status]}")
    print(f"Marge totale : {value(model.objective):,.0f} MAD")
    print(f"Temps de résolution : {elapsed:.3f} s")

    return model, x, liv, S_pf, elapsed

# ═══════════════════════════════════════════
# 4. EXTRACTION DES RÉSULTATS
# ═══════════════════════════════════════════

def extract_results(model, x, liv, df_main, cad_dict, rho, get_arret,
                    J=7, T_periods=(1, 2, 3, 4)):
    fams_main = ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']
    lignes    = ['PK', 'CRMA', 'CRMB', 'BAF', 'SKP', 'LGA', 'LGB']

    rows = []
    for _, row in df_main.iterrows():
        cid = row['ID']; D = float(row['Tonnage'])
        ltot = sum(value(liv[(cid, t)]) or 0 for t in T_periods)
        statut = 'Honorée' if ltot >= D - 0.01 else ('Partielle' if ltot > 0.01 else 'Refusée')
        rows.append({'ID': cid, 'Famille': row['Famille'], 'Grade': row['Grade'],
                     'Tonnage': D, 'Livree': round(ltot, 1),
                     'Priorite': row['Priorite'], 'Statut': statut})
    df_res = pd.DataFrame(rows)

    plan = []
    for ligne in lignes:
        for fam in fams_main:
            for t in T_periods:
                t_prod = 0.0
                for _, row in df_main.iterrows():
                    if row['Famille'] != fam: continue
                    cid = row['ID']
                    for rname, rsteps in ROUTES[fam].items():
                        if ligne not in [l for (l, _) in rsteps]: continue
                        xval = value(x[(cid, rname, t)]) or 0
                        if xval > 0.01:
                            gl = gamma_lr(rsteps, ligne, rho)
                            t_prod += xval / gl
                if t_prod > 0.01:
                    cad   = cad_dict.get((ligne, fam))
                    avail = (J - get_arret(ligne, t)) * cad if cad else 0
                    plan.append({'Ligne': ligne, 'Famille': fam, 'Semaine': t,
                                 'Tonnage_traite_T': round(t_prod, 1),
                                 'Capacite_nette_T': round(avail, 1)})
    df_plan = pd.DataFrame(plan)

    util_rows = []
    for ligne in lignes:
        for t in T_periods:
            used = 0.0
            for fam in fams_main:
                for _, row in df_main.iterrows():
                    if row['Famille'] != fam: continue
                    cid = row['ID']
                    for rname, rsteps in ROUTES[fam].items():
                        if ligne not in [l for (l, _) in rsteps]: continue
                        cad = cad_dict.get((ligne, fam))
                        if not cad: continue
                        xval = value(x[(cid, rname, t)]) or 0
                        if xval > 0.01:
                            gl = gamma_lr(rsteps, ligne, rho)
                            used += (xval / gl) / cad
            avail = J - get_arret(ligne, t)
            util_rows.append({'Ligne': ligne, 'Semaine': t,
                              'Jours_utilises': round(used, 2),
                              'Jours_disponibles': avail,
                              'Taux_pct': round(used / avail * 100, 1) if avail > 0 else 0})
    df_util = pd.DataFrame(util_rows)

    sp = [{'Contrainte': n, 'Shadow_price': round(c.pi, 2)}
          for n, c in model.constraints.items()
          if c.pi and abs(c.pi) > 0.5 and (n.startswith('Cap_') or n.startswith('HRC_'))]
    df_sp = pd.DataFrame(sp).sort_values('Shadow_price', ascending=False)

    return df_res, df_plan, df_util, df_sp

# ═══════════════════════════════════════════
# 4b. MARGE NETTE PAR FAMILLE
# ═══════════════════════════════════════════

def compute_marge_par_famille(model, x, liv, df_main, rho, tau_chute,
                              get_cv, get_prix_hrc, T_periods=(1, 2, 3, 4)):
    """
    Calcule la marge nette réelle par famille en reconstituant, pour chaque
    tonne produite, le coût complet : HRC + coûts variables de transformation
    + zinc/peinture - valeur chute.
    Marge nette = CA livré - coût total de production.
    """
    rows = []
    for fam in ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']:
        ca_fam       = 0.0
        cout_hrc     = 0.0
        cout_transfo = 0.0
        cout_zinc_p  = 0.0
        val_chute    = 0.0
        tonnage_livre = 0.0
        tonnage_prod  = 0.0

        for _, row in df_main.iterrows():
            if row['Famille'] != fam:
                continue
            cid = row['ID']
            pv  = float(row['PrixVente'])
            ep  = float(row['Epaisseur'])
            g   = row['Grade']
            w   = int(float(row['Largeur']))

            # CA réalisé (somme des livraisons × prix de vente)
            for t in T_periods:
                lval = value(liv[(cid, t)]) or 0.0
                ca_fam        += pv * lval
                tonnage_livre += lval

            # Coûts de production par route et période
            for rname, rsteps in ROUTES[fam].items():
                rr = rho_route(rsteps, rho)
                c_t = ct_route(rsteps, ep, get_cv, rho)
                v_c = vc_route(rsteps, ep, rho, tau_chute)
                # coût HRC par tonne produite finie = prix_HRC / rendement_route
                cm = get_prix_hrc(g, w) / rr

                for t in T_periods:
                    xval = value(x[(cid, rname, t)]) or 0.0
                    if xval < 1e-6:
                        continue
                    tonnage_prod  += xval
                    cout_hrc      += cm  * xval
                    cout_transfo  += c_t * xval
                    cout_zinc_p   += (cz_fam(fam) + cp_fam(fam)) * xval
                    val_chute     += v_c * xval   # récupération (signe +)

        cout_total = cout_hrc + cout_transfo + cout_zinc_p - val_chute
        marge_nette = ca_fam - cout_total

        rows.append({
            'Famille':          fam,
            'Tonnage_produit_T': round(tonnage_prod,  1),
            'Tonnage_livre_T':  round(tonnage_livre, 1),
            'CA_MAD':           round(ca_fam,        0),
            'Cout_HRC_MAD':     round(cout_hrc,      0),
            'Cout_Transfo_MAD': round(cout_transfo,  0),
            'Cout_ZincPeinture_MAD': round(cout_zinc_p, 0),
            'Val_Chute_MAD':    round(val_chute,     0),
            'Cout_Total_MAD':   round(cout_total,    0),
            'Marge_Nette_MAD':  round(marge_nette,   0),
            'Marge_pct':        round(marge_nette / ca_fam * 100, 1) if ca_fam > 0 else 0.0,
        })

    return pd.DataFrame(rows)

# ═══════════════════════════════════════════
# 5. ANALYSE DE SENSIBILITÉ
# ═══════════════════════════════════════════

def compute_sensibilite(df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
                        dispo, stk_pk_init, stk_pk_min,
                        stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
                        df_sp, baseline_obj):
    """
    Re-résout le modèle pour chaque scénario de perturbation :
      - Arrêt +1 jour sur chaque ligne goulot identifiée par les shadow prices
      - Prix HRC +10%
    Retourne un DataFrame avec la marge obtenue et le delta vs baseline.
    """
    sensi_rows = []

    # --- Scénarios : arrêt +1j sur les lignes goulot ---
    lignes_goulot_caps = [
        n for n in df_sp['Contrainte'].tolist()
        if n.startswith('Cap_')
    ]
    # Dédoublonner : on perturbe chaque (ligne, semaine) une seule fois
    perturbations_arret = set()
    for cap_name in lignes_goulot_caps:
        parts = cap_name.split('_')   # ex: ['Cap', 'CRMA', 'S2']
        if len(parts) < 3:
            continue
        ligne = parts[1]
        t     = int(parts[2].replace('S', ''))
        perturbations_arret.add((ligne, t))

    for (ligne, t) in sorted(perturbations_arret):
        print(f"  Sensibilité : arrêt +1j {ligne} S{t}…")
        model_s, x_s, liv_s, S_pf_s, _ = build_and_solve(
            df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc, dispo,
            stk_pk_init, stk_pk_min,
            stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
            arret_delta={(ligne, t): 1}
        )
        marge_s = value(model_s.objective) or 0.0
        sensi_rows.append({
            'Scenario':          f'Arrêt +1j {ligne} S{t}',
            'Marge_MAD':         round(marge_s),
            'Delta_vs_baseline': round(marge_s - baseline_obj),
            'Delta_pct':         round((marge_s - baseline_obj) / baseline_obj * 100, 2)
                                 if baseline_obj else 0.0,
        })

    # --- Scénario : prix HRC +10% ---
    print("  Sensibilité : prix HRC +10%…")
    model_hrc, x_hrc, liv_hrc, S_pf_hrc, _ = build_and_solve(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc, dispo,
        stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
        hrc_factor=1.10
    )
    marge_hrc = value(model_hrc.objective) or 0.0
    sensi_rows.append({
        'Scenario':          'Prix HRC +10%',
        'Marge_MAD':         round(marge_hrc),
        'Delta_vs_baseline': round(marge_hrc - baseline_obj),
        'Delta_pct':         round((marge_hrc - baseline_obj) / baseline_obj * 100, 2)
                             if baseline_obj else 0.0,
    })

    return pd.DataFrame(sensi_rows)

# ═══════════════════════════════════════════
# B1. TARIFICATION D'OPPORTUNITÉ COMMERCIALE
# ═══════════════════════════════════════════

def compute_tarification_opportunite(model, x, df_res, df_main,
                                     cad_dict, rho, tau_chute,
                                     get_cv, get_prix_hrc,
                                     T_periods=(1, 2, 3, 4)):
    """
    Pour chaque commande refusée ou partielle, calcule :
      - Cout_complet       : coût de revient complet (HRC + transfo + zinc/peinture - chute)
      - Shadow_cost_capa   : pression des goulots machine sur cette commande
      - Shadow_cost_hrc    : pression du stock HRC rare sur cette commande
      - Shadow_cost_total  : somme des deux
      - Prix_opportunite   : PrixVente - shadow_cost_total
      - Prix_plancher      : cout_complet + shadow_cost_total → prix min de renégociation
      - Marge_nette_th     : PrixVente - cout_complet (sans pression ressources)
      - Signal             : interprétation actionnable
    """

    # --- Shadow prices capacité : π[(ligne, t)] en MAD / jour libéré ---
    pi = {}
    for name, cstr in model.constraints.items():
        if not name.startswith('Cap_'):
            continue
        parts = name.split('_')        # 'Cap_CRMA_S2' → ['Cap','CRMA','S2']
        if len(parts) < 3:
            continue
        ligne = parts[1]
        t     = int(parts[2].replace('S', ''))
        pi[(ligne, t)] = cstr.pi or 0.0

    # --- Shadow prices HRC : π_hrc[grade] en MAD / tonne HRC libérée ---
    pi_hrc = {}
    for name, cstr in model.constraints.items():
        if not name.startswith('HRC_'):
            continue
        grade = name.replace('HRC_', '')   # 'HRC_S320' → 'S320'
        pi_hrc[grade] = cstr.pi or 0.0

    rows = []
    cibles = df_res[df_res['Statut'].isin(['Refusée', 'Partielle'])]

    for _, res_row in cibles.iterrows():
        cid = res_row['ID']

        # Données commande
        cmd = df_main[df_main['ID'] == cid].iloc[0]
        fam = cmd['Famille']
        pv  = float(cmd['PrixVente'])
        ep  = float(cmd['Epaisseur'])
        g   = cmd['Grade']
        w   = int(float(cmd['Largeur']))
        D   = float(cmd['Tonnage'])
        liv = res_row['Livree']

        # --- Sélection de la route ---
        if res_row['Statut'] == 'Partielle':
            # Route réelle : celle où le LP a effectivement produit du tonnage
            best_route_name = None
            for rname in ROUTES[fam]:
                if any((value(x[(cid, rname, t)]) or 0) > 0.01 for t in T_periods):
                    best_route_name = rname
                    break
            # Fallback analytique si rien trouvé (sécurité)
            if best_route_name is None:
                best_route_name = min(
                    ROUTES[fam],
                    key=lambda rname: ct_route(ROUTES[fam][rname], ep, get_cv, rho)
                )
        else:
            # Refusée : le LP n'a rien produit → on choisit analytiquement
            # la route la moins coûteuse en transformation
            best_route_name = min(
                ROUTES[fam],
                key=lambda rname: ct_route(ROUTES[fam][rname], ep, get_cv, rho)
            )

        rsteps_best = ROUTES[fam][best_route_name]

        # --- Shadow cost capacité sur la route retenue ---
        best_shadow_cost_capa = 0.0
        for (ligne, _) in rsteps_best:
            cad = cad_dict.get((ligne, fam))
            if not cad:
                continue
            gl     = gamma_lr(rsteps_best, ligne, rho)
            pi_moy = sum(pi.get((ligne, t), 0.0) for t in T_periods) / len(T_periods)
            best_shadow_cost_capa += pi_moy / (cad * gl)

        # --- Shadow cost HRC sur la route retenue ---
        rr = rho_route(rsteps_best, rho)
        best_shadow_cost_hrc = pi_hrc.get(g, 0.0) / rr

        # Total shadow cost
        best_shadow_cost = best_shadow_cost_capa + best_shadow_cost_hrc

        # --- Coût complet de production (hors pression ressources) ---
        cm           = get_prix_hrc(g, w) / rr
        c_t          = ct_route(rsteps_best, ep, get_cv, rho)
        v_c          = vc_route(rsteps_best, ep, rho, tau_chute)
        cout_complet = cm + c_t + cz_fam(fam) + cp_fam(fam) - v_c

        # --- Métriques commerciales ---
        prix_opport    = pv - best_shadow_cost
        prix_plancher  = cout_complet + best_shadow_cost
        marge_nette_th = pv - cout_complet

        # --- Signal actionnable ---
        if pv < prix_plancher:
            signal = "Sous le plancher — renégocier le prix ou abandonner"
        elif prix_opport < 0:
            signal = "Opportunité négative — ressources rares coûtent plus que la marge"
        elif res_row['Statut'] == 'Refusée':
            signal = "Rentable — déblocage capacité ou sous-traitance à étudier"
        else:
            signal = "Partielle — augmenter capacité pour honorer le reliquat"

        rows.append({
            'ID':                     cid,
            'Famille':                fam,
            'Grade':                  g,
            'Statut':                 res_row['Statut'],
            'Tonnage_demande_T':      D,
            'Tonnage_livre_T':        liv,
            'Reliquat_T':             round(D - liv, 1),
            'PrixVente_MAD_T':        round(pv, 0),
            'Cout_complet_MAD_T':     round(cout_complet, 0),
            'Shadow_cost_capa_MAD_T': round(best_shadow_cost_capa, 0),
            'Shadow_cost_hrc_MAD_T':  round(best_shadow_cost_hrc, 0),
            'Shadow_cost_total_MAD_T':round(best_shadow_cost, 0),
            'Prix_opportunite_MAD_T': round(prix_opport, 0),
            'Prix_plancher_MAD_T':    round(prix_plancher, 0),
            'Marge_nette_th_MAD_T':   round(marge_nette_th, 0),
            'Route_retenue':          best_route_name,
            'Signal':                 signal,
        })

    return pd.DataFrame(rows)

# ═══════════════════════════════════════════
# 6. VALIDATION A POSTERIORI
# ═══════════════════════════════════════════

def validate_solution(model, x, liv, S_pf, df_main, cad_dict, rho, get_arret,
                      dispo, stk_pk_init, stk_pk_min, stk_pf_min, stk_pf_max,
                      J=7, T_periods=(1, 2, 3, 4)):
    violations = []
    lignes     = ['PK', 'CRMA', 'CRMB', 'BAF', 'SKP', 'LGA', 'LGB']
    fams_main  = ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']
    grades_all = ['DC01', 'DD13', 'DX51', 'DX52', 'S320']

    for ligne in lignes:
        for t in T_periods:
            used = 0.0
            for _, row in df_main.iterrows():
                cid = row['ID']; fam = row['Famille']
                for rname, rsteps in ROUTES[fam].items():
                    if ligne not in [l for (l, _) in rsteps]: continue
                    cad = cad_dict.get((ligne, fam))
                    if not cad: continue
                    xval = value(x[(cid, rname, t)]) or 0
                    gl   = gamma_lr(rsteps, ligne, rho)
                    used += (xval / gl) / cad
            avail = J - get_arret(ligne, t)
            if used > avail + 0.01:
                violations.append(f"CAPACITE {ligne} S{t}: {used:.2f} > {avail}")

    for g in grades_all:
        total = 0.0
        for _, row in df_main.iterrows():
            if row['Grade'] != g: continue
            cid = row['ID']; fam = row['Famille']
            for rname, rsteps in ROUTES[fam].items():
                rr = rho_route(rsteps, rho)
                for t in T_periods:
                    total += (value(x[(cid, rname, t)]) or 0) / rr
        avail = dispo[g] + stk_pk_init.get(g, 0) - stk_pk_min.get(g, 0)
        if total > avail + 0.01:
            violations.append(f"HRC {g}: {total:.1f} > {avail:.1f}")

    for _, row in df_main.iterrows():
        cid = row['ID']; D = float(row['Tonnage'])
        ltot = sum(value(liv[(cid, t)]) or 0 for t in T_periods)
        if ltot > D + 0.01:
            violations.append(f"LIVRAISON {cid}: {ltot:.1f} > {D:.1f}")

    for fam in fams_main:
        for t in T_periods:
            sv   = value(S_pf[(fam, t)]) or 0
            smin = stk_pf_min.get(fam, 0)
            smax = stk_pf_max.get(fam, 9999)
            if sv < smin - 0.01:
                violations.append(f"STOCK MIN {fam} S{t}: {sv:.1f} < {smin}")
            if sv > smax + 0.01:
                violations.append(f"STOCK MAX {fam} S{t}: {sv:.1f} > {smax}")

    if violations:
        print(f"VALIDATION : {len(violations)} violation(s) détectée(s) !")
        for v in violations:
            print(f"  ❌ {v}")
    else:
        n = len(lignes) * 4 + len(grades_all) + len(df_main) + len(fams_main) * 4
        print(f"VALIDATION : ✅ Aucune violation sur {n} contraintes vérifiées.")

    return violations

# ═══════════════════════════════════════════
# B2. FORMULATION DES RETARDS (BACKLOGGING)
# ═══════════════════════════════════════════

def compute_backlogging(df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
                        dispo, stk_pk_init, stk_pk_min,
                        stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
                        alpha=0.005, beta_map=None,
                        J=7, T_periods=(1, 2, 3, 4)):
    """
    Ré-résout le modèle LP en autorisant des livraisons en retard.

    Pénalité de retard par commande i et semaine t :
        p_i = alpha × PrixVente_i × beta[Priorite_i]
    Coût du retard dans l'objectif :
        - p_i × (t - t_liv_i) × r[i, t]   pour t > t_liv_i

    Retard maximum autorisé : jusqu'à la dernière semaine de l'horizon
    (jamais au-delà de max(T_periods)).

    Paramètres
    ----------
    alpha     : taux de pénalité par semaine (défaut 0.5% du PV)
    beta_map  : dict {priorité → coefficient multiplicateur}
                défaut {1: 3.0, 2: 1.5, 3: 0.5}

    Retourne
    --------
    model_bl  : modèle PuLP résolu
    df_res_bl : DataFrame résultats commandes (avec colonnes retard)
    df_comp   : DataFrame comparaison baseline vs backlogging
    """

    if beta_map is None:
        beta_map = {1: 3.0, 2: 1.5, 3: 0.5}

    T_max     = max(T_periods)
    fams_main = ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']
    lignes    = ['PK', 'CRMA', 'CRMB', 'BAF', 'SKP', 'LGA', 'LGB']
    grades_all= ['DC01', 'DD13', 'DX51', 'DX52', 'S320']

    model_bl = LpProblem("MaghrebSteel_Backlogging", LpMaximize)

    x    = {}
    liv  = {}
    r_bl = {}   # tonnes livrées en retard
    S_pf = {}

    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']
        t_liv = int(row['SemaineLiv']) if pd.notna(row['SemaineLiv']) else T_max

        for rname in ROUTES[fam]:
            for t in T_periods:
                x[(cid, rname, t)] = LpVariable(f"x_{cid}_{rname}_{t}", lowBound=0)
        for t in T_periods:
            liv[(cid, t)] = LpVariable(f"l_{cid}_{t}", lowBound=0)
            # retard uniquement si t > t_liv et t <= T_max
            if t > t_liv:
                r_bl[(cid, t)] = LpVariable(f"r_{cid}_{t}", lowBound=0)

    for fam in fams_main:
        for t in T_periods:
            S_pf[(fam, t)] = LpVariable(f"S_{fam}_{t}", lowBound=0)

    # ── Objectif ──────────────────────────────────────────────────────────
    ot = []
    for _, row in df_main.iterrows():
        cid   = row['ID']; fam = row['Famille']
        pv    = float(row['PrixVente'])
        prio  = parse_prio(row['Priorite'])
        t_liv = int(row['SemaineLiv']) if pd.notna(row['SemaineLiv']) else T_max
        beta  = beta_map.get(prio, 1.0)
        p_i   = alpha * pv * beta   # MAD / T / semaine de retard

        # CA livraisons à temps
        for t in T_periods:
            ot.append(pv * liv[(cid, t)])

        # Pénalité retards
        for t in T_periods:
            if (cid, t) in r_bl:
                delai = t - t_liv          # nombre de semaines de retard
                ot.append(-p_i * delai * r_bl[(cid, t)])

        # Coûts de production (identiques au baseline)
        for rname, rsteps in ROUTES[fam].items():
            for t in T_periods:
                ep  = float(row['Epaisseur'])
                g   = row['Grade']
                w   = int(float(row['Largeur']))
                rr  = rho_route(rsteps, rho)
                cm  = get_prix_hrc(g, w) / rr
                c_t = ct_route(rsteps, ep, get_cv, rho)
                v_c = vc_route(rsteps, ep, rho, tau_chute)
                ot.append(-(cm + c_t + cz_fam(fam) + cp_fam(fam) - v_c) * x[(cid, rname, t)])

    model_bl += lpSum(ot), "Marge_totale_backlog"

    # ── Capacité (identique au baseline) ──────────────────────────────────
    for ligne in lignes:
        for t in T_periods:
            arr   = get_arret(ligne, t)
            terms = []
            for _, row in df_main.iterrows():
                cid = row['ID']; fam = row['Famille']
                for rname, rsteps in ROUTES[fam].items():
                    if ligne not in [l for (l, _) in rsteps]: continue
                    cad = cad_dict.get((ligne, fam))
                    if not cad: continue
                    gl = gamma_lr(rsteps, ligne, rho)
                    terms.append((1.0 / (gl * cad)) * x[(cid, rname, t)])
            if terms:
                model_bl += lpSum(terms) <= max(0, J - arr), f"Cap_{ligne}_S{t}"

    # ── Stock PF (identique au baseline) ──────────────────────────────────
    for fam in fams_main:
        fam_cmds = df_main[df_main['Famille'] == fam]['ID'].tolist()
        for t in T_periods:
            prod_t = [x[(cid, rname, t)] for cid in fam_cmds for rname in ROUTES[fam]]
            del_t  = [liv[(cid, t)] for cid in fam_cmds]
            S0 = stk_pf_init.get(fam, 0) if t == 1 else S_pf[(fam, t - 1)]
            model_bl += S_pf[(fam, t)] == S0 + lpSum(prod_t) - lpSum(del_t), f"StockBal_{fam}_S{t}"
            model_bl += S_pf[(fam, t)] >= stk_pf_min.get(fam, 0),            f"StockMin_{fam}_S{t}"
            model_bl += S_pf[(fam, t)] <= stk_pf_max.get(fam, 9999),         f"StockMax_{fam}_S{t}"

    # ── Livraison + retard ────────────────────────────────────────────────
    for _, row in df_main.iterrows():
        cid   = row['ID']; fam = row['Famille']
        D     = float(row['Tonnage'])

        # Total livré (à temps + en retard) ≤ demande
        all_liv  = [liv[(cid, t)] for t in T_periods]
        all_retard = [r_bl[(cid, t)] for t in T_periods if (cid, t) in r_bl]
        model_bl += lpSum(all_liv) + lpSum(all_retard) <= D, f"MaxDeliv_{cid}"

        # Cumul livraisons + retards ≤ cumul production (on ne livre que ce qu'on a produit)
        for t in T_periods:
            cum_p   = [x[(cid, rname, tau)] for rname in ROUTES[fam] for tau in range(1, t + 1)]
            cum_l   = [liv[(cid, tau)] for tau in range(1, t + 1)]
            cum_r   = [r_bl[(cid, tau)] for tau in range(1, t + 1) if (cid, tau) in r_bl]
            model_bl += lpSum(cum_l) + lpSum(cum_r) <= lpSum(cum_p), f"ProdBefore_{cid}_S{t}"

    # ── HRC (identique au baseline) ───────────────────────────────────────
    for g in grades_all:
        terms = []
        for _, row in df_main.iterrows():
            if row['Grade'] != g: continue
            cid = row['ID']; fam = row['Famille']
            for rname, rsteps in ROUTES[fam].items():
                rr = rho_route(rsteps, rho)
                for t in T_periods:
                    terms.append((1.0 / rr) * x[(cid, rname, t)])
        if terms:
            avail = dispo[g] + stk_pk_init.get(g, 0) - stk_pk_min.get(g, 0)
            model_bl += lpSum(terms) <= avail, f"HRC_{g}"

    # ── Résolution ────────────────────────────────────────────────────────
    print("  Résolution modèle backlogging...")
    t0 = tm.time()
    model_bl.solve(PULP_CBC_CMD(msg=0, timeLimit=300))
    elapsed = tm.time() - t0
    print(f"  Statut : {LpStatus[model_bl.status]}")
    print(f"  Marge totale (avec retards) : {value(model_bl.objective):,.0f} MAD")
    print(f"  Temps : {elapsed:.3f} s")

    # ── Extraction résultats ──────────────────────────────────────────────
    rows = []
    for _, row in df_main.iterrows():
        cid   = row['ID']
        D     = float(row['Tonnage'])
        pv    = float(row['PrixVente'])
        prio  = parse_prio(row['Priorite'])
        t_liv = int(row['SemaineLiv']) if pd.notna(row['SemaineLiv']) else T_max
        beta  = beta_map.get(prio, 1.0)
        p_i   = alpha * pv * beta

        ltot  = sum(value(liv[(cid, t)]) or 0 for t in T_periods)
        rtot  = sum(value(r_bl[(cid, t)]) or 0 for t in T_periods if (cid, t) in r_bl)

        # Semaines réelles de retard (pondérées par les tonnes)
        retard_moy = 0.0
        if rtot > 0.01:
            retard_moy = sum(
                (t - t_liv) * (value(r_bl[(cid, t)]) or 0)
                for t in T_periods if (cid, t) in r_bl
            ) / rtot

        penalite_tot = sum(
            p_i * (t - t_liv) * (value(r_bl[(cid, t)]) or 0)
            for t in T_periods if (cid, t) in r_bl
        )

        total_honore = ltot + rtot
        statut = ('Honorée' if total_honore >= D - 0.01
                  else ('Partielle' if total_honore > 0.01 else 'Refusée'))

        rows.append({
            'ID':                cid,
            'Famille':           row['Famille'],
            'Grade':             row['Grade'],
            'Priorite':          prio,
            'Tonnage_demande_T': D,
            'Livre_a_temps_T':   round(ltot, 1),
            'Livre_en_retard_T': round(rtot, 1),
            'Total_honore_T':    round(total_honore, 1),
            'Retard_moy_sem':    round(retard_moy, 2),
            'Penalite_MAD':      round(penalite_tot, 0),
            'Statut':            statut,
        })

    df_res_bl = pd.DataFrame(rows)

    # ── Synthèse comparaison ──────────────────────────────────────────────
    df_comp = pd.DataFrame([
        {
            'Indicateur': 'Marge totale (MAD)',
            'Backlogging': f"{value(model_bl.objective):,.0f}",
        },
        {
            'Indicateur': 'Taux de service global (%)',
            'Backlogging': f"{df_res_bl['Total_honore_T'].sum() / df_res_bl['Tonnage_demande_T'].sum() * 100:.1f}%",
        },
        {
            'Indicateur': 'Tonnes livrées à temps',
            'Backlogging': f"{df_res_bl['Livre_a_temps_T'].sum():,.0f}",
        },
        {
            'Indicateur': 'Tonnes livrées en retard',
            'Backlogging': f"{df_res_bl['Livre_en_retard_T'].sum():,.0f}",
        },
        {
            'Indicateur': 'Commandes honorées (total)',
            'Backlogging': (df_res_bl['Statut'] == 'Honorée').sum(),
        },
        {
            'Indicateur': 'Commandes partielles',
            'Backlogging': (df_res_bl['Statut'] == 'Partielle').sum(),
        },
        {
            'Indicateur': 'Commandes refusées',
            'Backlogging': (df_res_bl['Statut'] == 'Refusée').sum(),
        },
        {
            'Indicateur': 'Pénalités retard totales (MAD)',
            'Backlogging': f"{df_res_bl['Penalite_MAD'].sum():,.0f}",
        },
    ])

    return model_bl, df_res_bl, df_comp

# ═══════════════════════════════════════════
# B3. MODÉLISATION DES COÛTS DE STOCKAGE
# ═══════════════════════════════════════════

def compute_cout_stockage(df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
                          dispo, stk_pk_init, stk_pk_min,
                          stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
                          valeur_mode='cout_revient',
                          valeur_fixe_par_fam=None,
                          df_marge_fam=None,
                          J=7, T_periods=(1, 2, 3, 4)):
    """
    Ré-résout le modèle LP en intégrant les coûts de stockage dans l'objectif.

    Coût de stockage par famille f et semaine t :
        CS(f, t) = h_f × S_pf(f, t)

    Avec :
        h_f = taux_immo × valeur_stock_f + cout_physique_f

    Paramètres
    ----------
    valeur_mode : 'cout_revient' → valeur stock = coût de revient depuis df_marge_fam
                  'fixe'        → valeur stock = valeur_fixe_par_fam (dict manuel)
    valeur_fixe_par_fam : dict {famille → MAD/T}, utilisé si valeur_mode='fixe'
    df_marge_fam : DataFrame issu de compute_marge_par_famille, requis si mode='cout_revient'

    Retourne
    --------
    model_st   : modèle PuLP résolu
    df_res_st  : DataFrame résultats commandes
    df_stock_t : DataFrame stocks par famille et semaine avec coût associé
    df_comp    : DataFrame comparaison baseline vs stockage
    df_h       : DataFrame des taux h_f utilisés (traçabilité)
    """

    fams_main  = ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']
    lignes     = ['PK', 'CRMA', 'CRMB', 'BAF', 'SKP', 'LGA', 'LGB']
    grades_all = ['DC01', 'DD13', 'DX51', 'DX52', 'S320']

    # ── Calcul des taux h_f ───────────────────────────────────────────────
    # valeur_stock_f selon le mode choisi
    if valeur_mode == 'cout_revient':
        if df_marge_fam is None:
            raise ValueError("df_marge_fam requis pour valeur_mode='cout_revient'")
        # Coût de revient moyen par tonne produite = Cout_Total / Tonnage_produit
        valeur_stock = {}
        for _, row in df_marge_fam.iterrows():
            fam = row['Famille']
            if row['Tonnage_produit_T'] > 0:
                valeur_stock[fam] = row['Cout_Total_MAD'] / row['Tonnage_produit_T']
            else:
                # Fallback : prix de vente moyen des commandes de cette famille
                pv_moy = df_main[df_main['Famille'] == fam]['PrixVente'].mean()
                valeur_stock[fam] = float(pv_moy) if pd.notna(pv_moy) else 8000.0

    elif valeur_mode == 'fixe':
        if valeur_fixe_par_fam is None:
            # Valeurs par défaut raisonnables si non fournies
            valeur_fixe_par_fam = {
                'HRC DEC': 6000,
                'CRC':     7500,
                'HDG':     9000,
                'PPGI':   11000,
                'BACR':   10500,
            }
        valeur_stock = valeur_fixe_par_fam.copy()

    else:
        raise ValueError(f"valeur_mode inconnu : {valeur_mode}. Choisir 'cout_revient' ou 'fixe'.")

    # h_f = taux_immo × valeur_stock_f + cout_physique_f
    h = {}
    h_rows = []
    for fam in fams_main:
        vs  = valeur_stock.get(fam, 8000.0)
        cp  = COUT_PHYSIQUE_STOCK.get(fam, 20)
        hf  = TAUX_IMMO_SEMAINE * vs + cp
        h[fam] = hf
        h_rows.append({
            'Famille':              fam,
            'Valeur_stock_MAD_T':   round(vs, 0),
            'Taux_immo_semaine_pct':round(TAUX_IMMO_SEMAINE * 100, 4),
            'Cout_immo_MAD_T_sem':  round(TAUX_IMMO_SEMAINE * vs, 2),
            'Cout_physique_MAD_T_sem': cp,
            'h_total_MAD_T_sem':    round(hf, 2),
        })
    df_h = pd.DataFrame(h_rows)

    print(f"\n  Taux de stockage h_f utilisés (mode={valeur_mode}) :")
    print(df_h[['Famille','Valeur_stock_MAD_T','Cout_immo_MAD_T_sem',
                 'Cout_physique_MAD_T_sem','h_total_MAD_T_sem']].to_string(index=False))

    # ── Construction du modèle ────────────────────────────────────────────
    model_st = LpProblem("MaghrebSteel_Stockage", LpMaximize)

    x    = {}
    liv  = {}
    S_pf = {}

    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']
        for rname in ROUTES[fam]:
            for t in T_periods:
                x[(cid, rname, t)] = LpVariable(f"x_{cid}_{rname}_{t}", lowBound=0)
        for t in T_periods:
            liv[(cid, t)] = LpVariable(f"l_{cid}_{t}", lowBound=0)
    for fam in fams_main:
        for t in T_periods:
            S_pf[(fam, t)] = LpVariable(f"S_{fam}_{t}", lowBound=0)

    # ── Objectif : marge - coûts production - coûts stockage ─────────────
    ot = []

    # CA et coûts production (identiques au baseline)
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; pv = float(row['PrixVente'])
        for t in T_periods:
            ot.append(pv * liv[(cid, t)])
        for rname, rsteps in ROUTES[fam].items():
            for t in T_periods:
                ep  = float(row['Epaisseur'])
                g   = row['Grade']
                w   = int(float(row['Largeur']))
                rr  = rho_route(rsteps, rho)
                cm  = get_prix_hrc(g, w) / rr
                c_t = ct_route(rsteps, ep, get_cv, rho)
                v_c = vc_route(rsteps, ep, rho, tau_chute)
                ot.append(-(cm + c_t + cz_fam(fam) + cp_fam(fam) - v_c) * x[(cid, rname, t)])

    # Coûts de stockage : h_f × S_pf(f, t) pour chaque famille et semaine
    for fam in fams_main:
        for t in T_periods:
            ot.append(-h[fam] * S_pf[(fam, t)])

    model_st += lpSum(ot), "Marge_nette_stockage"

    # ── Capacité ──────────────────────────────────────────────────────────
    for ligne in lignes:
        for t in T_periods:
            arr   = get_arret(ligne, t)
            terms = []
            for _, row in df_main.iterrows():
                cid = row['ID']; fam = row['Famille']
                for rname, rsteps in ROUTES[fam].items():
                    if ligne not in [l for (l, _) in rsteps]: continue
                    cad = cad_dict.get((ligne, fam))
                    if not cad: continue
                    gl = gamma_lr(rsteps, ligne, rho)
                    terms.append((1.0 / (gl * cad)) * x[(cid, rname, t)])
            if terms:
                model_st += lpSum(terms) <= max(0, J - arr), f"Cap_{ligne}_S{t}"

    # ── Stock PF ──────────────────────────────────────────────────────────
    for fam in fams_main:
        fam_cmds = df_main[df_main['Famille'] == fam]['ID'].tolist()
        for t in T_periods:
            prod_t = [x[(cid, rname, t)] for cid in fam_cmds for rname in ROUTES[fam]]
            del_t  = [liv[(cid, t)] for cid in fam_cmds]
            S0 = stk_pf_init.get(fam, 0) if t == 1 else S_pf[(fam, t - 1)]
            model_st += S_pf[(fam, t)] == S0 + lpSum(prod_t) - lpSum(del_t), f"StockBal_{fam}_S{t}"
            model_st += S_pf[(fam, t)] >= stk_pf_min.get(fam, 0),            f"StockMin_{fam}_S{t}"
            model_st += S_pf[(fam, t)] <= stk_pf_max.get(fam, 9999),         f"StockMax_{fam}_S{t}"

    # ── Livraison ─────────────────────────────────────────────────────────
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; D = float(row['Tonnage'])
        model_st += lpSum(liv[(cid, t)] for t in T_periods) <= D, f"MaxDeliv_{cid}"
        for t in T_periods:
            cum_p = [x[(cid, rname, tau)] for rname in ROUTES[fam] for tau in range(1, t + 1)]
            cum_l = [liv[(cid, tau)] for tau in range(1, t + 1)]
            model_st += lpSum(cum_l) <= lpSum(cum_p), f"ProdBefore_{cid}_S{t}"

    # ── HRC ───────────────────────────────────────────────────────────────
    for g in grades_all:
        terms = []
        for _, row in df_main.iterrows():
            if row['Grade'] != g: continue
            cid = row['ID']; fam = row['Famille']
            for rname, rsteps in ROUTES[fam].items():
                rr = rho_route(rsteps, rho)
                for t in T_periods:
                    terms.append((1.0 / rr) * x[(cid, rname, t)])
        if terms:
            avail = dispo[g] + stk_pk_init.get(g, 0) - stk_pk_min.get(g, 0)
            model_st += lpSum(terms) <= avail, f"HRC_{g}"

    # ── Résolution ────────────────────────────────────────────────────────
    print("\n  Résolution modèle avec coûts de stockage...")
    t0 = tm.time()
    model_st.solve(PULP_CBC_CMD(msg=0, timeLimit=300))
    elapsed = tm.time() - t0
    print(f"  Statut     : {LpStatus[model_st.status]}")
    print(f"  Marge nette (après stockage) : {value(model_st.objective):,.0f} MAD")
    print(f"  Temps      : {elapsed:.3f} s")

    # ── Extraction résultats commandes ────────────────────────────────────
    rows = []
    for _, row in df_main.iterrows():
        cid = row['ID']; D = float(row['Tonnage'])
        ltot   = sum(value(liv[(cid, t)]) or 0 for t in T_periods)
        statut = ('Honorée' if ltot >= D - 0.01
                  else ('Partielle' if ltot > 0.01 else 'Refusée'))
        rows.append({
            'ID':      cid,
            'Famille': row['Famille'],
            'Tonnage': D,
            'Livree':  round(ltot, 1),
            'Statut':  statut,
        })
    df_res_st = pd.DataFrame(rows)

    # ── Stocks par famille et semaine avec coût ───────────────────────────
    stock_rows = []
    cout_stockage_total = 0.0
    for fam in fams_main:
        for t in T_periods:
            sv   = value(S_pf[(fam, t)]) or 0.0
            cout = h[fam] * sv
            cout_stockage_total += cout
            stock_rows.append({
                'Famille':          fam,
                'Semaine':          t,
                'Stock_T':          round(sv, 1),
                'h_MAD_T_sem':      round(h[fam], 2),
                'Cout_stockage_MAD':round(cout, 0),
            })
    df_stock_t = pd.DataFrame(stock_rows)

    # ── Comparaison synthétique ───────────────────────────────────────────
    marge_brute = sum(
        float(row['PrixVente']) * (value(liv[(row['ID'], t)]) or 0)
        for _, row in df_main.iterrows()
        for t in T_periods
    )

    df_comp = pd.DataFrame([
        {'Indicateur': 'Marge nette totale (MAD)',
         'Valeur': f"{value(model_st.objective):,.0f}"},
        {'Indicateur': 'dont CA livraisons (MAD)',
         'Valeur': f"{marge_brute:,.0f}"},
        {'Indicateur': 'dont Coût stockage total (MAD)',
         'Valeur': f"{cout_stockage_total:,.0f}"},
        {'Indicateur': 'Taux de service (%)',
         'Valeur': f"{df_res_st['Livree'].sum() / df_res_st['Tonnage'].sum() * 100:.1f}%"},
        {'Indicateur': 'Commandes honorées',
         'Valeur': (df_res_st['Statut'] == 'Honorée').sum()},
        {'Indicateur': 'Commandes partielles',
         'Valeur': (df_res_st['Statut'] == 'Partielle').sum()},
        {'Indicateur': 'Commandes refusées',
         'Valeur': (df_res_st['Statut'] == 'Refusée').sum()},
        {'Indicateur': 'Mode valorisation stock',
         'Valeur': valeur_mode},
    ])

    return model_st, df_res_st, df_stock_t, df_comp, df_h

# ═══════════════════════════════════════════
# B4. MODÉLISATION DES CAMPAGNES (MILP)
# ═══════════════════════════════════════════

def compute_campagnes_milp(df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
                           dispo, stk_pk_init, stk_pk_min,
                           stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
                           cout_changement=None,
                           J=7, T_periods=(1, 2, 3, 4)):
    """
    Ré-résout le modèle en MILP avec contraintes de campagne sur LGA et LGB.

    Variables binaires :
        y[(ligne, fam, t)] ∈ {0,1} — campagne famille fam active sur ligne à semaine t

    Contraintes campagne :
        1. Une seule famille active par ligne et semaine :
               Σ_f y[(l,f,t)] ≤ 1
        2. Production liée à l'activation (Big-M) :
               Σ_i x[(i,r,t)] ≤ M_(l,f,t) × y[(l,f,t)]
        3. Coût de changement entre semaines consécutives :
               z[(l,f,t)] ≥ y[(l,f,t-1)] - y[(l,f,t)]   z ∈ {0,1}
               Coût = Σ_(l,f,t) C_l × z[(l,f,t)]

    Paramètres
    ----------
    cout_changement : dict {ligne → MAD/changement}, défaut COUT_CHANGEMENT

    Retourne
    --------
    model_milp   : modèle PuLP résolu
    df_res_milp  : DataFrame résultats commandes
    df_campagnes : DataFrame plan de campagnes (quelle famille sur quelle ligne/semaine)
    df_comp      : DataFrame comparaison baseline vs MILP
    """

    if cout_changement is None:
        cout_changement = COUT_CHANGEMENT

    fams_main  = ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']
    lignes     = ['PK', 'CRMA', 'CRMB', 'BAF', 'SKP', 'LGA', 'LGB']
    grades_all = ['DC01', 'DD13', 'DX51', 'DX52', 'S320']

    model_milp = LpProblem("MaghrebSteel_Campagnes", LpMaximize)

    x    = {}
    liv  = {}
    S_pf = {}
    y    = {}   # binaires campagne : y[(ligne, fam, t)]
    z    = {}   # binaires changement : z[(ligne, fam, t)] = 1 si changement à t

    # ── Variables continues (identiques au baseline) ──────────────────────
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']
        for rname in ROUTES[fam]:
            for t in T_periods:
                x[(cid, rname, t)] = LpVariable(f"x_{cid}_{rname}_{t}", lowBound=0)
        for t in T_periods:
            liv[(cid, t)] = LpVariable(f"l_{cid}_{t}", lowBound=0)
    for fam in fams_main:
        for t in T_periods:
            S_pf[(fam, t)] = LpVariable(f"S_{fam}_{t}", lowBound=0)

    # ── Variables binaires campagne ───────────────────────────────────────
    for ligne, fams_ligne in LIGNES_CAMPAGNE.items():
        for fam in fams_ligne:
            for t in T_periods:
                y[(ligne, fam, t)] = LpVariable(f"y_{ligne}_{fam}_{t}", cat='Binary')
            # z[(ligne, fam, t)] = 1 si on quitte la campagne fam après semaine t-1
            for t in T_periods[1:]:   # pas de changement avant S1
                z[(ligne, fam, t)] = LpVariable(f"z_{ligne}_{fam}_{t}", cat='Binary')

    # ── Objectif ──────────────────────────────────────────────────────────
    ot = []

    # CA et coûts production
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; pv = float(row['PrixVente'])
        for t in T_periods:
            ot.append(pv * liv[(cid, t)])
        for rname, rsteps in ROUTES[fam].items():
            for t in T_periods:
                ep  = float(row['Epaisseur'])
                g   = row['Grade']
                w   = int(float(row['Largeur']))
                rr  = rho_route(rsteps, rho)
                cm  = get_prix_hrc(g, w) / rr
                c_t = ct_route(rsteps, ep, get_cv, rho)
                v_c = vc_route(rsteps, ep, rho, tau_chute)
                ot.append(-(cm + c_t + cz_fam(fam) + cp_fam(fam) - v_c) * x[(cid, rname, t)])

    # Coûts de changement de campagne
    for ligne in LIGNES_CAMPAGNE:
        cl = cout_changement.get(ligne, 0)
        for fam in LIGNES_CAMPAGNE[ligne]:
            for t in T_periods[1:]:
                ot.append(-cl * z[(ligne, fam, t)])

    model_milp += lpSum(ot), "Marge_campagnes"

    # ── Contraintes campagne ──────────────────────────────────────────────

    # 1. Une seule famille active par ligne et semaine
    for ligne, fams_ligne in LIGNES_CAMPAGNE.items():
        for t in T_periods:
            model_milp += (
                lpSum(y[(ligne, fam, t)] for fam in fams_ligne) <= 1,
                f"UneFamille_{ligne}_S{t}"
            )

    # 2. Big-M : production possible seulement si campagne activée
    #    M = capacité max de la ligne sur la semaine (J jours × cadence max)
    for ligne, fams_ligne in LIGNES_CAMPAGNE.items():
        for fam in fams_ligne:
            for t in T_periods:
                # Big-M = jours dispo × cadence de cette famille sur cette ligne
                cad = cad_dict.get((ligne, fam), 0)
                arr = get_arret(ligne, t)
                M   = max(0, J - arr) * cad if cad else 0

                if M == 0:
                    continue

                # Somme des tonnes INPUT sur cette ligne pour cette famille/semaine
                # (on somme sur toutes les routes qui passent par cette ligne)
                terms_prod = []
                for _, row in df_main.iterrows():
                    if row['Famille'] != fam:
                        continue
                    cid = row['ID']
                    for rname, rsteps in ROUTES[fam].items():
                        if ligne not in [l for (l, _) in rsteps]:
                            continue
                        gl = gamma_lr(rsteps, ligne, rho)
                        # tonnes input sur la ligne = x / gamma
                        terms_prod.append((1.0 / gl) * x[(cid, rname, t)])

                if terms_prod:
                    model_milp += (
                        lpSum(terms_prod) <= M * y[(ligne, fam, t)],
                        f"BigM_{ligne}_{fam}_S{t}"
                    )

    # 3. Détection des changements de campagne
    #    z[(l,f,t)] ≥ y[(l,f,t-1)] - y[(l,f,t)]
    #    → z = 1 si on était en campagne f à t-1 et on ne l'est plus à t
    for ligne in LIGNES_CAMPAGNE:
        for fam in LIGNES_CAMPAGNE[ligne]:
            for t in T_periods[1:]:
                t_prev = t - 1
                model_milp += (
                    z[(ligne, fam, t)] >= y[(ligne, fam, t_prev)] - y[(ligne, fam, t)],
                    f"Changement_{ligne}_{fam}_S{t}"
                )

    # ── Capacité (identique au baseline) ──────────────────────────────────
    for ligne in lignes:
        for t in T_periods:
            arr   = get_arret(ligne, t)
            terms = []
            for _, row in df_main.iterrows():
                cid = row['ID']; fam = row['Famille']
                for rname, rsteps in ROUTES[fam].items():
                    if ligne not in [l for (l, _) in rsteps]: continue
                    cad = cad_dict.get((ligne, fam))
                    if not cad: continue
                    gl = gamma_lr(rsteps, ligne, rho)
                    terms.append((1.0 / (gl * cad)) * x[(cid, rname, t)])
            if terms:
                model_milp += lpSum(terms) <= max(0, J - arr), f"Cap_{ligne}_S{t}"

    # ── Stock PF ──────────────────────────────────────────────────────────
    for fam in fams_main:
        fam_cmds = df_main[df_main['Famille'] == fam]['ID'].tolist()
        for t in T_periods:
            prod_t = [x[(cid, rname, t)] for cid in fam_cmds for rname in ROUTES[fam]]
            del_t  = [liv[(cid, t)] for cid in fam_cmds]
            S0 = stk_pf_init.get(fam, 0) if t == 1 else S_pf[(fam, t - 1)]
            model_milp += S_pf[(fam, t)] == S0 + lpSum(prod_t) - lpSum(del_t), f"StockBal_{fam}_S{t}"
            model_milp += S_pf[(fam, t)] >= stk_pf_min.get(fam, 0),            f"StockMin_{fam}_S{t}"
            model_milp += S_pf[(fam, t)] <= stk_pf_max.get(fam, 9999),         f"StockMax_{fam}_S{t}"

    # ── Livraison ─────────────────────────────────────────────────────────
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; D = float(row['Tonnage'])
        model_milp += lpSum(liv[(cid, t)] for t in T_periods) <= D, f"MaxDeliv_{cid}"
        for t in T_periods:
            cum_p = [x[(cid, rname, tau)] for rname in ROUTES[fam] for tau in range(1, t + 1)]
            cum_l = [liv[(cid, tau)] for tau in range(1, t + 1)]
            model_milp += lpSum(cum_l) <= lpSum(cum_p), f"ProdBefore_{cid}_S{t}"

    # ── HRC ───────────────────────────────────────────────────────────────
    for g in grades_all:
        terms = []
        for _, row in df_main.iterrows():
            if row['Grade'] != g: continue
            cid = row['ID']; fam = row['Famille']
            for rname, rsteps in ROUTES[fam].items():
                rr = rho_route(rsteps, rho)
                for t in T_periods:
                    terms.append((1.0 / rr) * x[(cid, rname, t)])
        if terms:
            avail = dispo[g] + stk_pk_init.get(g, 0) - stk_pk_min.get(g, 0)
            model_milp += lpSum(terms) <= avail, f"HRC_{g}"

    # ── Résolution ────────────────────────────────────────────────────────
    print("\n  Résolution modèle MILP campagnes...")
    t0 = tm.time()
    model_milp.solve(PULP_CBC_CMD(msg=0, timeLimit=300))
    elapsed = tm.time() - t0
    print(f"  Statut  : {LpStatus[model_milp.status]}")
    print(f"  Marge   : {value(model_milp.objective):,.0f} MAD")
    print(f"  Temps   : {elapsed:.3f} s")

    # ── Plan de campagnes ─────────────────────────────────────────────────
    camp_rows = []
    for ligne, fams_ligne in LIGNES_CAMPAGNE.items():
        for t in T_periods:
            fam_active = None
            for fam in fams_ligne:
                if (value(y[(ligne, fam, t)]) or 0) > 0.5:
                    fam_active = fam
                    break
            # Tonnes produites sur cette ligne cette semaine
            tonnes = 0.0
            if fam_active:
                for _, row in df_main.iterrows():
                    if row['Famille'] != fam_active: continue
                    cid = row['ID']
                    for rname, rsteps in ROUTES[fam_active].items():
                        if ligne not in [l for (l, _) in rsteps]: continue
                        gl = gamma_lr(rsteps, ligne, rho)
                        tonnes += (value(x[(cid, rname, t)]) or 0) / gl

            # Changement vs semaine précédente
            changement = False
            if t > 1:
                fam_prev = next(
                    (f for f in fams_ligne
                     if (value(y[(ligne, f, t - 1)]) or 0) > 0.5),
                    None
                )
                changement = (fam_prev != fam_active)

            camp_rows.append({
                'Ligne':         ligne,
                'Semaine':       t,
                'Famille_active':fam_active if fam_active else '—',
                'Tonnes_input_T':round(tonnes, 1),
                'Changement':    '⚠ OUI' if changement else 'non',
                'Cout_chgt_MAD': cout_changement.get(ligne, 0) if changement else 0,
            })
    df_campagnes = pd.DataFrame(camp_rows)

    # Coût total changements
    cout_chgt_total = df_campagnes['Cout_chgt_MAD'].sum()

    # ── Résultats commandes ───────────────────────────────────────────────
    rows = []
    for _, row in df_main.iterrows():
        cid = row['ID']; D = float(row['Tonnage'])
        ltot   = sum(value(liv[(cid, t)]) or 0 for t in T_periods)
        statut = ('Honorée' if ltot >= D - 0.01
                  else ('Partielle' if ltot > 0.01 else 'Refusée'))
        rows.append({
            'ID':      cid,
            'Famille': row['Famille'],
            'Tonnage': D,
            'Livree':  round(ltot, 1),
            'Statut':  statut,
        })
    df_res_milp = pd.DataFrame(rows)

    # ── Synthèse comparaison ──────────────────────────────────────────────
    df_comp = pd.DataFrame([
        {'Indicateur': 'Marge totale (MAD)',
         'Valeur': f"{value(model_milp.objective):,.0f}"},
        {'Indicateur': 'Coût total changements campagne (MAD)',
         'Valeur': f"{cout_chgt_total:,.0f}"},
        {'Indicateur': 'Nombre de changements de campagne',
         'Valeur': (df_campagnes['Changement'] == '⚠ OUI').sum()},
        {'Indicateur': 'Taux de service (%)',
         'Valeur': f"{df_res_milp['Livree'].sum() / df_res_milp['Tonnage'].sum() * 100:.1f}%"},
        {'Indicateur': 'Commandes honorées',
         'Valeur': (df_res_milp['Statut'] == 'Honorée').sum()},
        {'Indicateur': 'Commandes partielles',
         'Valeur': (df_res_milp['Statut'] == 'Partielle').sum()},
        {'Indicateur': 'Commandes refusées',
         'Valeur': (df_res_milp['Statut'] == 'Refusée').sum()},
    ])

    return model_milp, df_res_milp, df_campagnes, df_comp

# ═══════════════════════════════════════════
# B5. ANALYSE BRANCH-AND-BOUND
# ═══════════════════════════════════════════

def analyse_branch_and_bound(df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
                              dispo, stk_pk_init, stk_pk_min,
                              stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
                              cout_changement=None,
                              J=7, T_periods=(1, 2, 3, 4)):
    """
    Ré-résout le modèle MILP de B4 avec msg=1 pour capturer les journaux CBC,
    puis parse ces journaux pour extraire les métriques Branch-and-Bound.

    Métriques extraites :
      - Nombre de nœuds explorés
      - Nombre de solutions entières trouvées
      - Borne supérieure (relaxation LP racine)
      - Meilleure solution entière (objectif final)
      - Gap relatif (%) entre borne sup et meilleure solution
      - Temps de résolution total

    Retourne
    --------
    df_bb : DataFrame des métriques B&B
    log   : liste brute des lignes de log CBC (pour debug si besoin)
    """

    import io
    import re
    import sys

    if cout_changement is None:
        cout_changement = COUT_CHANGEMENT

    fams_main  = ['CRC', 'HDG', 'PPGI', 'BACR', 'HRC DEC']
    lignes     = ['PK', 'CRMA', 'CRMB', 'BAF', 'SKP', 'LGA', 'LGB']
    grades_all = ['DC01', 'DD13', 'DX51', 'DX52', 'S320']

    # ── Reconstruction du modèle MILP (identique à B4) ───────────────────
    model_bb = LpProblem("MaghrebSteel_BB", LpMaximize)

    x    = {}
    liv  = {}
    S_pf = {}
    y    = {}
    z    = {}

    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']
        for rname in ROUTES[fam]:
            for t in T_periods:
                x[(cid, rname, t)] = LpVariable(f"x_{cid}_{rname}_{t}", lowBound=0)
        for t in T_periods:
            liv[(cid, t)] = LpVariable(f"l_{cid}_{t}", lowBound=0)
    for fam in fams_main:
        for t in T_periods:
            S_pf[(fam, t)] = LpVariable(f"S_{fam}_{t}", lowBound=0)

    for ligne, fams_ligne in LIGNES_CAMPAGNE.items():
        for fam in fams_ligne:
            for t in T_periods:
                y[(ligne, fam, t)] = LpVariable(f"y_{ligne}_{fam}_{t}", cat='Binary')
            for t in T_periods[1:]:
                z[(ligne, fam, t)] = LpVariable(f"z_{ligne}_{fam}_{t}", cat='Binary')

    # Objectif
    ot = []
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; pv = float(row['PrixVente'])
        for t in T_periods:
            ot.append(pv * liv[(cid, t)])
        for rname, rsteps in ROUTES[fam].items():
            for t in T_periods:
                ep  = float(row['Epaisseur']); g = row['Grade']; w = int(float(row['Largeur']))
                rr  = rho_route(rsteps, rho)
                cm  = get_prix_hrc(g, w) / rr
                c_t = ct_route(rsteps, ep, get_cv, rho)
                v_c = vc_route(rsteps, ep, rho, tau_chute)
                ot.append(-(cm + c_t + cz_fam(fam) + cp_fam(fam) - v_c) * x[(cid, rname, t)])
    for ligne in LIGNES_CAMPAGNE:
        cl = cout_changement.get(ligne, 0)
        for fam in LIGNES_CAMPAGNE[ligne]:
            for t in T_periods[1:]:
                ot.append(-cl * z[(ligne, fam, t)])
    model_bb += lpSum(ot), "Marge_BB"

    # Contraintes campagne
    for ligne, fams_ligne in LIGNES_CAMPAGNE.items():
        for t in T_periods:
            model_bb += lpSum(y[(ligne, fam, t)] for fam in fams_ligne) <= 1, f"UneFamille_{ligne}_S{t}"
        for fam in fams_ligne:
            for t in T_periods:
                cad = cad_dict.get((ligne, fam), 0)
                arr = get_arret(ligne, t)
                M   = max(0, J - arr) * cad if cad else 0
                if M == 0: continue
                terms_prod = []
                for _, row in df_main.iterrows():
                    if row['Famille'] != fam: continue
                    cid = row['ID']
                    for rname, rsteps in ROUTES[fam].items():
                        if ligne not in [l for (l, _) in rsteps]: continue
                        gl = gamma_lr(rsteps, ligne, rho)
                        terms_prod.append((1.0 / gl) * x[(cid, rname, t)])
                if terms_prod:
                    model_bb += lpSum(terms_prod) <= M * y[(ligne, fam, t)], f"BigM_{ligne}_{fam}_S{t}"
            for t in T_periods[1:]:
                model_bb += z[(ligne, fam, t)] >= y[(ligne, fam, t-1)] - y[(ligne, fam, t)], f"Chgt_{ligne}_{fam}_S{t}"

    # Capacité
    for ligne in lignes:
        for t in T_periods:
            arr = get_arret(ligne, t)
            terms = []
            for _, row in df_main.iterrows():
                cid = row['ID']; fam = row['Famille']
                for rname, rsteps in ROUTES[fam].items():
                    if ligne not in [l for (l, _) in rsteps]: continue
                    cad = cad_dict.get((ligne, fam))
                    if not cad: continue
                    gl = gamma_lr(rsteps, ligne, rho)
                    terms.append((1.0 / (gl * cad)) * x[(cid, rname, t)])
            if terms:
                model_bb += lpSum(terms) <= max(0, J - arr), f"Cap_{ligne}_S{t}"

    # Stock PF
    for fam in fams_main:
        fam_cmds = df_main[df_main['Famille'] == fam]['ID'].tolist()
        for t in T_periods:
            prod_t = [x[(cid, rname, t)] for cid in fam_cmds for rname in ROUTES[fam]]
            del_t  = [liv[(cid, t)] for cid in fam_cmds]
            S0 = stk_pf_init.get(fam, 0) if t == 1 else S_pf[(fam, t - 1)]
            model_bb += S_pf[(fam, t)] == S0 + lpSum(prod_t) - lpSum(del_t), f"StockBal_{fam}_S{t}"
            model_bb += S_pf[(fam, t)] >= stk_pf_min.get(fam, 0),            f"StockMin_{fam}_S{t}"
            model_bb += S_pf[(fam, t)] <= stk_pf_max.get(fam, 9999),         f"StockMax_{fam}_S{t}"

    # Livraison
    for _, row in df_main.iterrows():
        cid = row['ID']; fam = row['Famille']; D = float(row['Tonnage'])
        model_bb += lpSum(liv[(cid, t)] for t in T_periods) <= D, f"MaxDeliv_{cid}"
        for t in T_periods:
            cum_p = [x[(cid, rname, tau)] for rname in ROUTES[fam] for tau in range(1, t + 1)]
            cum_l = [liv[(cid, tau)] for tau in range(1, t + 1)]
            model_bb += lpSum(cum_l) <= lpSum(cum_p), f"ProdBefore_{cid}_S{t}"

    # HRC
    for g in grades_all:
        terms = []
        for _, row in df_main.iterrows():
            if row['Grade'] != g: continue
            cid = row['ID']; fam = row['Famille']
            for rname, rsteps in ROUTES[fam].items():
                rr = rho_route(rsteps, rho)
                for t in T_periods:
                    terms.append((1.0 / rr) * x[(cid, rname, t)])
        if terms:
            avail = dispo[g] + stk_pk_init.get(g, 0) - stk_pk_min.get(g, 0)
            model_bb += lpSum(terms) <= avail, f"HRC_{g}"

    # ── Résolution avec capture des logs CBC ──────────────────────────────
    print("\n  Résolution B&B avec journaux CBC...")

    # Capturer stdout où CBC écrit ses logs
    log_capture = io.StringIO()
    old_stdout  = sys.stdout
    sys.stdout  = log_capture

    t0 = tm.time()
    model_bb.solve(PULP_CBC_CMD(msg=1, timeLimit=300))
    elapsed = tm.time() - t0

    sys.stdout = old_stdout
    log_lines  = log_capture.getvalue().splitlines()

    # ── Parsing des métriques depuis les logs CBC ─────────────────────────
    # CBC écrit des lignes comme :
    #   "Result - Optimal solution found"
    #   "Enumerated nodes:           42"
    #   "Solutions found:             3"
    #   "Best objective  3.5887675e+07, best bound 3.5887675e+07, gap 0.0000%"

    nb_noeuds       = None
    nb_solutions    = None
    borne_sup       = None
    meilleure_sol   = None
    gap_pct         = None

    for line in log_lines:
        line_s = line.strip()

        # Nœuds explorés
        m = re.search(r'Enumerated nodes[:\s]+(\d+)', line_s, re.IGNORECASE)
        if m:
            nb_noeuds = int(m.group(1))

        # Solutions entières trouvées
        m = re.search(r'Solutions found[:\s]+(\d+)', line_s, re.IGNORECASE)
        if m:
            nb_solutions = int(m.group(1))

        # Meilleure solution + borne + gap (ligne "Best objective ...")
        m = re.search(
            r'Best objective\s+([\d.e+\-]+),\s*best bound\s+([\d.e+\-]+),\s*gap\s+([\d.]+)%',
            line_s, re.IGNORECASE
        )
        if m:
            meilleure_sol = float(m.group(1))
            borne_sup     = float(m.group(2))
            gap_pct       = float(m.group(3))

    # Fallback si CBC n'a pas loggué en format standard
    if meilleure_sol is None:
        meilleure_sol = value(model_bb.objective) or 0.0
    if borne_sup is None:
        borne_sup = meilleure_sol   # gap = 0 si optimal prouvé
    if gap_pct is None:
        gap_pct = abs(borne_sup - meilleure_sol) / abs(borne_sup) * 100 if borne_sup else 0.0
    if nb_noeuds is None:
        nb_noeuds = 0
    if nb_solutions is None:
        nb_solutions = 1 if LpStatus[model_bb.status] == 'Optimal' else 0

    # Nombre de variables binaires dans le modèle
    nb_binaires = len(y) + len(z)

    print(f"  Statut       : {LpStatus[model_bb.status]}")
    print(f"  Marge        : {meilleure_sol:,.0f} MAD")
    print(f"  Temps        : {elapsed:.3f} s")

    # ── DataFrame métriques ───────────────────────────────────────────────
    df_bb = pd.DataFrame([
        {'Métrique': 'Statut résolution',
         'Valeur': LpStatus[model_bb.status]},
        {'Métrique': 'Marge optimale (MAD)',
         'Valeur': f"{meilleure_sol:,.0f}"},
        {'Métrique': 'Borne supérieure LP relaxée (MAD)',
         'Valeur': f"{borne_sup:,.0f}"},
        {'Métrique': 'Gap B&B (%)',
         'Valeur': f"{gap_pct:.4f}%"},
        {'Métrique': 'Nœuds explorés',
         'Valeur': nb_noeuds},
        {'Métrique': 'Solutions entières trouvées',
         'Valeur': nb_solutions},
        {'Métrique': 'Variables binaires totales',
         'Valeur': nb_binaires},
        {'Métrique': 'Temps de résolution (s)',
         'Valeur': f"{elapsed:.3f}"},
        {'Métrique': 'Interprétation gap',
         'Valeur': (
             'Optimal prouvé — gap nul'          if gap_pct < 0.001 else
             'Quasi-optimal — gap < 1%'          if gap_pct < 1.0   else
             'Solution approchée — gap > 1%'
         )},
    ])

    print("\n  Métriques Branch-and-Bound :")
    print(df_bb.to_string(index=False))

    return df_bb, log_lines

# ═══════════════════════════════════════════
# 7. EXPORT EXCEL
# ═══════════════════════════════════════════

def export_excel(df_res, df_plan, df_util, df_sp, obj_value,
                 df_marge_fam, df_sensi, df_tarif, df_res_bl, df_comp_bl,
                 df_res_st, df_stock_t, df_comp_st, df_h,
                 df_res_milp, df_campagnes, df_comp_milp, df_bb,
                 path="outputs/resultats.xlsx"):
    import os; os.makedirs(os.path.dirname(path), exist_ok=True)
    with pd.ExcelWriter(path, engine='openpyxl') as w:
        synth = pd.DataFrame({
            'Indicateur': ['Marge totale (MAD)', 'Taux de service (%)',
                           'Commandes honorées', 'Commandes partielles', 'Commandes refusées'],
            'Valeur': [f"{obj_value:,.0f}",
                       f"{df_res['Livree'].sum()/df_res['Tonnage'].sum()*100:.1f}%",
                       (df_res['Statut'] == 'Honorée').sum(),
                       (df_res['Statut'] == 'Partielle').sum(),
                       (df_res['Statut'] == 'Refusée').sum()]
        })
        synth.to_excel(w, sheet_name='Synthese', index=False)
        df_res.to_excel(w, sheet_name='Commandes', index=False)
        df_plan.to_excel(w, sheet_name='Plan_de_marche', index=False)
        df_util.to_excel(w, sheet_name='Utilisation', index=False)
        df_sp.to_excel(w, sheet_name='Shadow_prices', index=False)
        df_marge_fam.to_excel(w, sheet_name='Marge_par_famille', index=False)
        df_res[df_res['Statut'] == 'Refusée'].to_excel(w, sheet_name='Commandes_refusees', index=False)
        df_sensi.to_excel(w, sheet_name='Sensibilite', index=False)
        df_tarif.to_excel(w, sheet_name='Tarification_opportunite', index=False)
        df_res_bl.to_excel(w, sheet_name='Backlogging_commandes', index=False)
        df_comp_bl.to_excel(w, sheet_name='Backlogging_synthese', index=False)
        df_res_st.to_excel(w,   sheet_name='Stockage_commandes',  index=False)
        df_stock_t.to_excel(w,  sheet_name='Stockage_detail',     index=False)
        df_comp_st.to_excel(w,  sheet_name='Stockage_synthese',   index=False)
        df_h.to_excel(w,        sheet_name='Stockage_taux_hf',    index=False)
        df_res_milp.to_excel(w,  sheet_name='Campagnes_commandes', index=False)
        df_campagnes.to_excel(w, sheet_name='Campagnes_plan',      index=False)
        df_comp_milp.to_excel(w, sheet_name='Campagnes_synthese',  index=False)
        df_bb.to_excel(w, sheet_name='BranchAndBound', index=False)

    print(f"Résultats exportés : {path}")

# ═══════════════════════════════════════════
# 8. MAIN
# ═══════════════════════════════════════════

if __name__ == "__main__":
    print("=" * 55)
    print("  Maghreb Steel — Simulateur Capacité-Commande")
    print("=" * 55)

    (df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc, dispo,
     stk_pk_init, stk_pk_min, stk_pk_max,
     stk_pf_init, stk_pf_min, stk_pf_max, get_arret) = load_data()

    print(f"\n{len(df_main)} commandes chargées | Tonnage total : {df_main['Tonnage'].sum():,.0f} T")

    print("\n[1/4] Résolution du modèle baseline...")
    model, x, liv, S_pf, elapsed = build_and_solve(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc, dispo,
        stk_pk_init, stk_pk_min, stk_pf_init, stk_pf_min, stk_pf_max, get_arret)

    baseline_obj = value(model.objective)

    print("\n[2/4] Extraction des résultats...")
    df_res, df_plan, df_util, df_sp = extract_results(
        model, x, liv, df_main, cad_dict, rho, get_arret)

    print(f"  Taux de service : {df_res['Livree'].sum()/df_res['Tonnage'].sum()*100:.1f}%")
    print(f"  Commandes refusées : {(df_res['Statut']=='Refusée').sum()}")
    print("\n  Shadow prices structurels :")
    print(df_sp.to_string(index=False))

    # Marge nette réelle par famille (note 1 appliquée)
    df_marge_fam = compute_marge_par_famille(
        model, x, liv, df_main, rho, tau_chute, get_cv, get_prix_hrc)
    print("\n  Marge nette par famille :")
    print(df_marge_fam[['Famille','Tonnage_livre_T','CA_MAD','Cout_Total_MAD',
                          'Marge_Nette_MAD','Marge_pct']].to_string(index=False))
    
    print("\n  Tarification d'opportunité (refusées + partielles)...")
    df_tarif = compute_tarification_opportunite(
        model, x, df_res, df_main,
        cad_dict, rho, tau_chute, get_cv, get_prix_hrc)
    print(df_tarif[['ID','Statut','PrixVente_MAD_T','Prix_plancher_MAD_T',
                    'Prix_opportunite_MAD_T','Signal']].to_string(index=False))

    print("\n[3/4] Validation a posteriori...")
    validate_solution(model, x, liv, S_pf, df_main, cad_dict, rho, get_arret,
                      dispo, stk_pk_init, stk_pk_min, stk_pf_min, stk_pf_max)
    
    print("\n[B2] Backlogging...")
    model_bl, df_res_bl, df_comp_bl = compute_backlogging(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret)
    print(df_comp_bl.to_string(index=False))

    print("\n[B3] Coûts de stockage — mode coût de revient...")
    model_st, df_res_st, df_stock_t, df_comp_st, df_h = compute_cout_stockage(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
        valeur_mode='cout_revient',
        df_marge_fam=df_marge_fam)   # déjà calculé plus haut
    print(df_comp_st.to_string(index=False))

    print("\n[B4] Campagnes MILP...")
    model_milp, df_res_milp, df_campagnes, df_comp_milp = compute_campagnes_milp(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret)
    print(df_campagnes.to_string(index=False))
    print(df_comp_milp.to_string(index=False))

    print("\n[B5] Analyse Branch-and-Bound...")
    df_bb, _ = analyse_branch_and_bound(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret)

    # Analyse de sensibilité (note 2 appliquée : unpacking correct)
    print("\n[4/4] Analyse de sensibilité...")
    df_sensi = compute_sensibilite(
        df_main, cad_dict, rho, tau_chute, get_cv, get_prix_hrc,
        dispo, stk_pk_init, stk_pk_min,
        stk_pf_init, stk_pf_min, stk_pf_max, get_arret,
        df_sp, baseline_obj)
    print(df_sensi.to_string(index=False))

    export_excel(df_res, df_plan, df_util, df_sp, baseline_obj,
                 df_marge_fam, df_sensi, df_tarif, df_res_bl, df_comp_bl,
                 df_res_st, df_stock_t, df_comp_st, df_h,
                 df_res_milp, df_campagnes, df_comp_milp, df_bb)
    print("\nTerminé.")