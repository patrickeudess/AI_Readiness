# ============================================================
# AgriAI-Readiness — Application Streamlit
# Mémoire DU Données — UCAD/EBAD — Patrick-Eudess Zatty ALLA
# Évaluation automatique de l'AI-Readiness des données agricoles
# Inspiré de : Wilkinson et al. (2016), Wang & Strong (1996),
#              Sambasivan et al. (2021), Andrew Ng (Data-Centric AI)
# ============================================================

import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import io
import re
import os
from datetime import datetime

# ── Configuration page ──────────────────────────────────────
st.set_page_config(
    page_title="AgriAI-Readiness",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── CSS personnalisé ─────────────────────────────────────────
st.markdown("""
<style>
/* Palette académique */
:root {
    --bleu:    #1F3864;
    --bleu2:   #2E5496;
    --vert:    #1F6B3A;
    --orange:  #C55A11;
    --rouge:   #C00000;
    --gris:    #F5F7FA;
}

/* En-tête principal */
.main-header {
    background: linear-gradient(135deg, #1F3864 0%, #2E5496 100%);
    color: white;
    padding: 1.5rem 2rem;
    border-radius: 10px;
    margin-bottom: 1.5rem;
    text-align: center;
}
.main-header h1 { font-size: 1.8rem; margin: 0; font-weight: 700; }
.main-header p  { font-size: 0.9rem; margin: 0.3rem 0 0; opacity: 0.85; }

/* Cartes score */
.score-card {
    background: white;
    border-radius: 10px;
    padding: 1.2rem;
    text-align: center;
    border-left: 5px solid #2E5496;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    margin-bottom: 1rem;
}
.score-number { font-size: 2.2rem; font-weight: 800; color: #1F3864; }
.score-label  { font-size: 0.8rem; color: #666; text-transform: uppercase; letter-spacing: 0.05em; }
.score-dim    { font-size: 1rem;   color: #2E5496; font-weight: 600; margin-bottom: 0.3rem; }

/* Badges niveau */
.badge-faible  { background:#FDECEA; color:#C00000; padding:4px 14px; border-radius:20px; font-weight:700; font-size:0.85rem; }
.badge-moyen   { background:#FFF3E0; color:#C55A11; padding:4px 14px; border-radius:20px; font-weight:700; font-size:0.85rem; }
.badge-eleve   { background:#E8F5E9; color:#1F6B3A; padding:4px 14px; border-radius:20px; font-weight:700; font-size:0.85rem; }

/* Recommandations */
.reco-card {
    background: #F0F7EE;
    border-left: 4px solid #1F6B3A;
    padding: 0.8rem 1rem;
    border-radius: 0 8px 8px 0;
    margin: 0.4rem 0;
    font-size: 0.88rem;
}
.reco-urgent {
    background: #FDECEA;
    border-left: 4px solid #C00000;
}
.reco-info {
    background: #EEF4FB;
    border-left: 4px solid #2E5496;
}

/* Info box */
.info-box {
    background: #EEF4FB;
    border: 1px solid #2E5496;
    border-radius: 8px;
    padding: 0.8rem 1rem;
    font-size: 0.85rem;
    margin: 0.5rem 0;
}

/* Séparateur section */
.section-title {
    font-size: 1.1rem;
    font-weight: 700;
    color: #1F3864;
    border-bottom: 2px solid #2E5496;
    padding-bottom: 0.3rem;
    margin: 1.2rem 0 0.8rem;
}
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# MODULE 1 — CHARGEMENT DES DONNÉES
# ══════════════════════════════════════════════════════════════

def load_data(uploaded_file):
    """
    Charge un fichier Excel (mono ou multi-feuilles) ou CSV.
    Retourne un dictionnaire {nom_feuille: DataFrame} pour Excel,
    ou {"Données": DataFrame} pour CSV.
    """
    name = uploaded_file.name.lower()
    sheets = {}

    if name.endswith(".csv"):
        # Essai de plusieurs encodages courants
        for enc in ["utf-8", "latin-1", "cp1252"]:
            try:
                df = pd.read_csv(uploaded_file, encoding=enc)
                sheets["Données"] = df
                break
            except Exception:
                uploaded_file.seek(0)
        if not sheets:
            st.error("Impossible de lire le fichier CSV. Vérifiez l'encodage.")

    elif name.endswith((".xlsx", ".xls")):
        try:
            xl = pd.ExcelFile(uploaded_file)
            for sheet in xl.sheet_names:
                df = xl.parse(sheet)
                if not df.empty:
                    sheets[sheet] = df
        except Exception as e:
            st.error(f"Erreur lecture Excel : {e}")

    return sheets


# ══════════════════════════════════════════════════════════════
# MODULE 2 — DIMENSION 1 : COMPLÉTUDE (/25)
# Références : Wang & Strong (1996), Batini & Scannapieco (2016)
# ══════════════════════════════════════════════════════════════

def evaluate_completude(df):
    """
    Évalue la complétude du DataFrame selon 3 sous-critères :
    - C1 : Taux global de valeurs manquantes (/10)
    - C2 : Proportion de colonnes incomplètes > 5% (/8)
    - C3 : Absence de colonnes entièrement vides (/7)
    Score total : /25
    """
    details = {}
    n_cells = df.shape[0] * df.shape[1]
    n_missing = df.isnull().sum().sum()

    # C1 — Taux global valeurs manquantes
    taux_global = n_missing / n_cells if n_cells > 0 else 1
    if   taux_global == 0:      c1 = 10
    elif taux_global <= 0.02:   c1 = 9
    elif taux_global <= 0.05:   c1 = 8
    elif taux_global <= 0.10:   c1 = 6
    elif taux_global <= 0.20:   c1 = 4
    elif taux_global <= 0.35:   c1 = 2
    else:                       c1 = 0

    # C2 — Colonnes avec > 5% de manquants
    miss_par_col = df.isnull().mean()
    cols_incomplete = (miss_par_col > 0.05).sum()
    ratio_incomp = cols_incomplete / df.shape[1] if df.shape[1] > 0 else 1
    if   ratio_incomp == 0:     c2 = 8
    elif ratio_incomp <= 0.10:  c2 = 7
    elif ratio_incomp <= 0.25:  c2 = 5
    elif ratio_incomp <= 0.50:  c2 = 3
    else:                       c2 = 0

    # C3 — Colonnes entièrement vides
    cols_vides = (miss_par_col == 1.0).sum()
    c3 = 0 if cols_vides > 0 else 7

    score = c1 + c2 + c3

    # Colonnes les plus problématiques (top 5)
    top_missing = miss_par_col[miss_par_col > 0].sort_values(ascending=False).head(5)

    details = {
        "score": score,
        "taux_global_pct": round(taux_global * 100, 1),
        "n_missing": int(n_missing),
        "n_cells": n_cells,
        "cols_incomplete": int(cols_incomplete),
        "cols_vides": int(cols_vides),
        "top_missing": top_missing,
        "sous_scores": {"C1 — Taux global": (c1, 10),
                        "C2 — Colonnes incomplètes": (c2, 8),
                        "C3 — Colonnes vides": (c3, 7)},
    }
    return details


# ══════════════════════════════════════════════════════════════
# MODULE 3 — DIMENSION 2 : COHÉRENCE (/25)
# Références : Andrew Ng (Data-Centric AI), Sambasivan et al. (2021)
# ══════════════════════════════════════════════════════════════

def detect_incoherence_binaire(serie):
    """
    Détecte les colonnes binaires avec codages mixtes.
    Ex : Oui/oui/1/Non/0/non → incohérent.
    """
    vals = serie.dropna().astype(str).str.strip().str.lower().unique()
    codes_oui = {"oui","yes","1","true","o","vrai"}
    codes_non = {"non","no","0","false","n","faux"}
    found = set(vals)
    has_oui = bool(found & codes_oui)
    has_non = bool(found & codes_non)
    if has_oui and has_non:
        # incohérent si plusieurs représentations de la même valeur
        repr_oui = found & codes_oui
        repr_non = found & codes_non
        return len(repr_oui) > 1 or len(repr_non) > 1
    return False

def detect_date_formats(serie):
    """
    Détecte les colonnes date avec formats mixtes.
    Patterns reconnus : DD/MM/YYYY, YYYY-MM-DD, mois texte, année seule.
    """
    patterns = [
        r'\d{2}/\d{2}/\d{4}',       # DD/MM/YYYY
        r'\d{4}-\d{2}-\d{2}',       # ISO
        r'\d{2}-\d{2}-\d{4}',       # DD-MM-YYYY
        r'[a-zé]{3,9}\.?\s+\d{4}',  # mois texte
        r'^\d{4}$',                  # année seule
    ]
    formats_detectes = set()
    for val in serie.dropna().astype(str).head(50):
        for i, pat in enumerate(patterns):
            if re.search(pat, val.strip(), re.IGNORECASE):
                formats_detectes.add(i)
    return len(formats_detectes) > 1

def detect_texte_numerique(serie):
    """
    Détecte les colonnes numériques avec valeurs texte parasites.
    Ex : "150 000 f", "3 ha", "n/a".
    """
    if pd.api.types.is_numeric_dtype(serie):
        return False
    sample = serie.dropna().astype(str).head(100)
    n_numerique = sum(1 for v in sample if re.search(r'\d', v) and not v.replace('.','').replace(',','').isdigit())
    return n_numerique > len(sample) * 0.3

def evaluate_coherence(df):
    """
    Évalue la cohérence du DataFrame selon 4 sous-critères :
    - CO1 : Colonnes binaires avec codages multiples (/7)
    - CO2 : Colonnes date avec formats mixtes (/6)
    - CO3 : Texte dans colonnes numériques (/6)
    - CO4 : Cohérence des types de colonnes (/6)
    Score total : /25
    """
    cols_binaires_incoherentes = []
    cols_dates_mixtes = []
    cols_texte_num = []

    for col in df.columns:
        serie = df[col]
        n_unique = serie.dropna().nunique()

        # Détection binaire incohérente
        if 2 <= n_unique <= 6:
            if detect_incoherence_binaire(serie):
                cols_binaires_incoherentes.append(col)

        # Détection date mixte (colonnes dont le nom suggère une date)
        col_lower = col.lower()
        if any(k in col_lower for k in ["date","annee","année","mois","year","period"]):
            if detect_date_formats(serie):
                cols_dates_mixtes.append(col)

        # Détection texte dans numérique
        if detect_texte_numerique(serie):
            cols_texte_num.append(col)

    # CO1 — Binaire incohérent
    co1 = 7 if not cols_binaires_incoherentes else max(0, 7 - len(cols_binaires_incoherentes) * 2)

    # CO2 — Dates mixtes
    co2 = 6 if not cols_dates_mixtes else max(0, 6 - len(cols_dates_mixtes) * 3)

    # CO3 — Texte dans numérique
    co3 = 6 if not cols_texte_num else max(0, 6 - len(cols_texte_num) * 2)

    # CO4 — Stabilité des types (colonnes objet qui devraient être numériques)
    cols_object = df.select_dtypes(include=["object","string"]).columns
    n_suspect = 0
    for col in cols_object:
        # Tenter conversion numérique
        try:
            converted = pd.to_numeric(df[col].dropna().astype(str).str.replace(r'[^\d\.\-]','',regex=True), errors='coerce')
            taux_conv = converted.notna().mean()
            if taux_conv > 0.7:  # > 70% convertibles → devrait être numérique
                n_suspect += 1
        except Exception:
            pass
    co4 = max(0, 6 - n_suspect * 2)

    score = min(25, co1 + co2 + co3 + co4)

    return {
        "score": score,
        "cols_binaires_incoherentes": cols_binaires_incoherentes,
        "cols_dates_mixtes": cols_dates_mixtes,
        "cols_texte_num": cols_texte_num,
        "n_cols_type_suspect": n_suspect,
        "sous_scores": {
            "CO1 — Codes binaires": (co1, 7),
            "CO2 — Formats dates": (co2, 6),
            "CO3 — Texte/numérique": (co3, 6),
            "CO4 — Types de colonnes": (co4, 6),
        }
    }


# ══════════════════════════════════════════════════════════════
# MODULE 4 — DIMENSION 3 : DOCUMENTATION (/25)
# Référence : Wilkinson et al. (2016) — FAIR Principles
# ══════════════════════════════════════════════════════════════

def evaluate_documentation(df, all_sheets, sheet_name):
    """
    Évalue la documentation selon 4 sous-critères :
    - D1 : Noms de colonnes explicites (/7)
    - D2 : Présence d'un dictionnaire de données (/7)
    - D3 : Présence d'un README (/6)
    - D4 : Unités détectables dans les noms de colonnes (/5)
    Score total : /25
    """
    col_names = list(df.columns)

    # D1 — Noms de colonnes explicites (longueur > 3, pas "col1", pas "Unnamed")
    noms_explicites = sum(
        1 for c in col_names
        if len(str(c)) > 3
        and not str(c).lower().startswith("unnamed")
        and not re.match(r'^col\d+$', str(c).lower())
        and not re.match(r'^[a-z]\d+$', str(c).lower())  # A1, B2...
    )
    ratio_explicites = noms_explicites / len(col_names) if col_names else 0
    d1 = int(round(ratio_explicites * 7))

    # D2 — Dictionnaire de données dans une feuille dédiée
    dict_keywords = ["dictionnaire","dictionary","codebook","metadata","metadonnee",
                     "metadonnées","variables","légende","legende","definitions"]
    has_dict = any(
        any(kw in s.lower() for kw in dict_keywords)
        for s in all_sheets.keys()
    )
    d2 = 7 if has_dict else 0

    # D3 — README ou Notes dans une feuille
    readme_keywords = ["readme","notes","note","documentation","about","info","lisez"]
    has_readme = any(
        any(kw in s.lower() for kw in readme_keywords)
        for s in all_sheets.keys()
    )
    d3 = 6 if has_readme else 0

    # D4 — Unités détectables dans les noms de colonnes
    unit_patterns = [r'_kg', r'_ha', r'_pct', r'_m2', r'_km', r'_l$',
                     r'_t$', r'\(kg\)', r'\(ha\)', r'\(%\)', r'_nb', r'_num']
    has_units = sum(
        1 for c in col_names
        if any(re.search(pat, str(c).lower()) for pat in unit_patterns)
    )
    d4 = min(5, has_units)

    score = d1 + d2 + d3 + d4

    return {
        "score": score,
        "ratio_noms_explicites": round(ratio_explicites * 100, 1),
        "has_dict": has_dict,
        "has_readme": has_readme,
        "n_cols_avec_unites": has_units,
        "n_cols_total": len(col_names),
        "sous_scores": {
            "D1 — Noms de colonnes": (d1, 7),
            "D2 — Dictionnaire données": (d2, 7),
            "D3 — README / Notes": (d3, 6),
            "D4 — Unités dans colonnes": (d4, 5),
        }
    }


# ══════════════════════════════════════════════════════════════
# MODULE 5 — DIMENSION 4 : INTEROPÉRABILITÉ (/25)
# Références : Wilkinson et al. (2016), Lawrence (2017)
# ══════════════════════════════════════════════════════════════

def evaluate_interoperabilite(df, all_sheets):
    """
    Évalue l'interopérabilité analytique selon 4 sous-critères :
    - IN1 : Structure tabulaire propre (/7)
    - IN2 : Types de données stables et exploitables (/7)
    - IN3 : Identifiants cohérents (/5)
    - IN4 : Compatibilité ML (features exploitables) (/6)
    Score total : /25
    """

    # IN1 — Structure tabulaire propre
    # Vérifier : pas de lignes entièrement vides, pas d'en-têtes dupliqués
    lignes_vides = df.isnull().all(axis=1).sum()
    cols_dupliquees = len(col_names := list(df.columns)) - len(set(col_names))
    has_merged_like = any(str(c).startswith("Unnamed") for c in df.columns)

    in1_penalite = lignes_vides * 1 + cols_dupliquees * 2 + (3 if has_merged_like else 0)
    in1 = max(0, 7 - in1_penalite)

    # IN2 — Types stables
    n_cols = df.shape[1]
    n_numeric  = df.select_dtypes(include=["number"]).shape[1]
    n_datetime = df.select_dtypes(include=["datetime"]).shape[1]
    n_object   = df.select_dtypes(include=["object"]).shape[1]
    # Pénaliser les colonnes object qui devraient être numériques ou datetime
    n_suspect = 0
    for col in df.select_dtypes(include=["object","string"]).columns:
        try:
            conv = pd.to_numeric(df[col].dropna().astype(str).str.replace(r'[^\d\.\-]','',regex=True), errors='coerce')
            if conv.notna().mean() > 0.7:
                n_suspect += 1
        except Exception:
            pass
    in2 = max(0, 7 - n_suspect * 2)

    # IN3 — Identifiants cohérents
    # Chercher colonnes ID/clé potentielles
    id_cols = [c for c in df.columns if any(k in str(c).lower() for k in ["id","code","identif","parc","num","ref"])]
    in3 = 5
    for col in id_cols:
        serie = df[col].dropna().astype(str)
        # Détecter incohérences de format dans les IDs
        patterns_trouves = set()
        for val in serie.head(50):
            if re.match(r'^[A-Z]-\d+$', val):        patterns_trouves.add("PREFIX-NUM")
            elif re.match(r'^\d+$', val):              patterns_trouves.add("NUM")
            elif re.match(r'^[A-Za-z]+ \d+$', val):   patterns_trouves.add("TEXT NUM")
            elif re.match(r'^\d{3}$', val):            patterns_trouves.add("ZERO-PAD")
        if len(patterns_trouves) > 1:
            in3 = max(0, in3 - 2)

    # IN4 — Compatibilité ML
    # Vérifier : au moins 1 variable numérique cible potentielle + features exploitables
    n_numeric_cols = df.select_dtypes(include=["number"]).shape[1]
    n_features_ok  = n_numeric_cols  # colonnes directement utilisables
    n_total        = df.shape[1]
    ratio_features = n_features_ok / n_total if n_total > 0 else 0

    if   ratio_features >= 0.7: in4 = 6
    elif ratio_features >= 0.5: in4 = 5
    elif ratio_features >= 0.3: in4 = 3
    elif ratio_features >= 0.1: in4 = 1
    else:                       in4 = 0

    score = min(25, in1 + in2 + in3 + in4)

    return {
        "score": score,
        "lignes_vides": int(lignes_vides),
        "cols_dupliquees": int(cols_dupliquees),
        "n_numeric": int(n_numeric),
        "n_object": int(n_object),
        "n_suspect_type": int(n_suspect),
        "id_cols": id_cols,
        "ratio_features_pct": round(ratio_features * 100, 1),
        "sous_scores": {
            "IN1 — Structure tabulaire": (in1, 7),
            "IN2 — Types de données": (in2, 7),
            "IN3 — Identifiants": (in3, 5),
            "IN4 — Compatibilité ML": (in4, 6),
        }
    }


# ══════════════════════════════════════════════════════════════
# MODULE 6 — SCORE GLOBAL & NIVEAU
# ══════════════════════════════════════════════════════════════

def compute_score_global(comp, cohe, docu, inte):
    total = comp["score"] + cohe["score"] + docu["score"] + inte["score"]
    if total <= 39:   niveau = ("Faible",  "badge-faible",  "🔴")
    elif total <= 69: niveau = ("Moyen",   "badge-moyen",   "🟡")
    else:             niveau = ("Élevé",   "badge-eleve",   "🟢")
    return total, niveau


# ══════════════════════════════════════════════════════════════
# MODULE 7 — RECOMMANDATIONS AUTOMATIQUES
# ══════════════════════════════════════════════════════════════

def generate_recommendations(comp, cohe, docu, inte):
    """
    Génère des recommandations hiérarchisées selon les scores obtenus.
    Priorité URGENTE (rouge) si score < 40% du max.
    Priorité IMPORTANTE (orange) si 40-70%.
    Priorité AMÉLIORATION (verte) si > 70%.
    """
    recos = []

    # ── Complétude ──
    if comp["taux_global_pct"] > 20:
        recos.append(("URGENT", "Complétude",
            f"Taux de valeurs manquantes élevé ({comp['taux_global_pct']}%). "
            "Appliquer une stratégie d'imputation (moyenne, médiane, mode) "
            "ou documenter explicitement les raisons des données manquantes."))
    elif comp["taux_global_pct"] > 5:
        recos.append(("IMPORTANT", "Complétude",
            f"{comp['cols_incomplete']} colonnes présentent plus de 5% de valeurs manquantes. "
            "Vérifier les colonnes critiques et envisager une imputation ciblée."))
    if comp["cols_vides"] > 0:
        recos.append(("URGENT", "Complétude",
            f"{comp['cols_vides']} colonne(s) entièrement vide(s) détectée(s). "
            "Supprimer ces colonnes ou les renseigner avant tout usage analytique."))

    # ── Cohérence ──
    if cohe["cols_binaires_incoherentes"]:
        recos.append(("URGENT", "Cohérence",
            f"Codes binaires incohérents détectés dans : {', '.join(cohe['cols_binaires_incoherentes'][:3])}. "
            "Standardiser en 0/1 ou Oui/Non de façon uniforme."))
    if cohe["cols_dates_mixtes"]:
        recos.append(("URGENT", "Cohérence",
            f"Formats de dates mixtes dans : {', '.join(cohe['cols_dates_mixtes'])}. "
            "Normaliser au format ISO 8601 (YYYY-MM-DD)."))
    if cohe["cols_texte_num"]:
        recos.append(("IMPORTANT", "Cohérence",
            f"Colonnes numériques contenant du texte : {', '.join(cohe['cols_texte_num'][:3])}. "
            "Nettoyer avec pandas (str.replace, to_numeric) ou OpenRefine."))
    if cohe["n_cols_type_suspect"] > 0:
        recos.append(("IMPORTANT", "Cohérence",
            f"{cohe['n_cols_type_suspect']} colonne(s) de type texte devraient être numériques. "
            "Convertir avec pd.to_numeric() après nettoyage."))

    # ── Documentation ──
    if not docu["has_dict"]:
        recos.append(("URGENT", "Documentation",
            "Dictionnaire de données absent. Créer une feuille 'Dictionnaire' avec : "
            "nom, type, unité, valeurs admissibles et définition de chaque variable. "
            "Standard recommandé : Frictionless Data Package."))
    if not docu["has_readme"]:
        recos.append(("IMPORTANT", "Documentation",
            "README absent. Ajouter une feuille 'README' décrivant : source, "
            "méthode de collecte, zone géographique, campagne, contact et licence."))
    if docu["ratio_noms_explicites"] < 80:
        recos.append(("IMPORTANT", "Documentation",
            f"Seulement {docu['ratio_noms_explicites']}% des colonnes ont des noms explicites. "
            "Renommer les colonnes abrégées en noms complets (ex : rdt_kg → rendement_kg_ha)."))
    if docu["n_cols_avec_unites"] == 0:
        recos.append(("AMÉLIORATION", "Documentation",
            "Aucune unité de mesure détectée dans les noms de colonnes. "
            "Ajouter l'unité en suffixe : rendement_kg_ha, surface_m2, poids_frais_kg."))

    # ── Interopérabilité ──
    if inte["n_suspect_type"] > 0:
        recos.append(("URGENT", "Interopérabilité",
            f"{inte['n_suspect_type']} colonne(s) doivent être converties en types numériques "
            "pour être exploitables dans Python/R/Power BI."))
    if inte["ratio_features_pct"] < 40:
        recos.append(("IMPORTANT", "Interopérabilité",
            f"Seulement {inte['ratio_features_pct']}% des colonnes sont directement exploitables "
            "pour la modélisation ML. Encoder les variables catégorielles (LabelEncoder, OneHot)."))
    if inte["lignes_vides"] > 0:
        recos.append(("IMPORTANT", "Interopérabilité",
            f"{inte['lignes_vides']} ligne(s) entièrement vide(s). "
            "Supprimer avec df.dropna(how='all')."))

    # Recommandation systématique FAIR
    recos.append(("AMÉLIORATION", "FAIR & Dépôt",
        "Pour améliorer la conformité FAIR : déposer le dataset sur Zenodo ou DataSuds (IRD) "
        "pour obtenir un DOI pérenne. Référence : Wilkinson et al. (2016)."))
    recos.append(("AMÉLIORATION", "FAIR & Dépôt",
        "Attribuer une licence explicite (CC-BY 4.0 recommandée) pour autoriser "
        "la réutilisation scientifique et commerciale du dataset."))

    # Tri par priorité
    ordre = {"URGENT": 0, "IMPORTANT": 1, "AMÉLIORATION": 2}
    recos.sort(key=lambda x: ordre[x[0]])
    return recos


# ══════════════════════════════════════════════════════════════
# MODULE 8 — GRAPHIQUES
# ══════════════════════════════════════════════════════════════

def plot_radar(scores_dims, labels):
    """Graphique radar des 4 dimensions AI-Readiness."""
    N = 4
    angles = np.linspace(0, 2 * np.pi, N, endpoint=False).tolist()
    angles += angles[:1]
    vals = scores_dims + scores_dims[:1]
    max_vals = [25] * N + [25]

    fig, ax = plt.subplots(figsize=(5, 5), subplot_kw=dict(polar=True))
    ax.set_facecolor("#F5F7FA")
    fig.patch.set_facecolor("#F5F7FA")

    # Grille
    for r in [5, 10, 15, 20, 25]:
        circle_vals = [r/25] * N + [r/25]
        ax.plot(angles, circle_vals, color="#CCCCCC", linewidth=0.5, linestyle="--")

    # Zone remplie
    norm_vals = [v/25 for v in vals]
    ax.fill(angles, norm_vals, alpha=0.25, color="#2E5496")
    ax.plot(angles, norm_vals, color="#1F3864", linewidth=2)

    # Points
    for angle, val in zip(angles[:-1], [v/25 for v in scores_dims]):
        ax.plot(angle, val, "o", color="#1F3864", markersize=7)

    # Labels
    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, size=9, color="#1F3864", fontweight="bold")
    ax.set_yticks([])
    ax.set_ylim(0, 1)
    ax.set_title("Profil AI-Readiness", size=11, color="#1F3864", fontweight="bold", pad=15)
    plt.tight_layout()
    return fig

def plot_barres(scores, labels, maxs):
    """Graphique barres horizontales des sous-critères."""
    fig, ax = plt.subplots(figsize=(7, max(3, len(scores) * 0.45)))
    fig.patch.set_facecolor("#F5F7FA")
    ax.set_facecolor("#F5F7FA")

    colors = []
    for s, m in zip(scores, maxs):
        ratio = s / m if m > 0 else 0
        if ratio < 0.4:   colors.append("#C00000")
        elif ratio < 0.7: colors.append("#C55A11")
        else:              colors.append("#1F6B3A")

    bars = ax.barh(labels, scores, color=colors, height=0.6, zorder=3)
    ax.barh(labels, maxs, color="#E0E0E0", height=0.6, zorder=2)

    for bar, score, max_v in zip(bars, scores, maxs):
        ax.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height()/2,
                f"{score}/{max_v}", va="center", ha="left", fontsize=8.5,
                color="#1F3864", fontweight="bold")

    ax.set_xlim(0, max(maxs) * 1.25)
    ax.set_xlabel("Score", color="#555555", fontsize=9)
    ax.tick_params(axis="y", labelsize=8.5)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.grid(axis="x", alpha=0.3, zorder=1)
    plt.tight_layout()
    return fig

def plot_gauge(score_total):
    """Jauge circulaire du score global."""
    fig, ax = plt.subplots(figsize=(4, 2.5))
    fig.patch.set_facecolor("#F5F7FA")
    ax.set_facecolor("#F5F7FA")
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 5.5)
    ax.axis("off")

    # Fond demi-cercle
    theta = np.linspace(np.pi, 0, 200)
    x_bg = 5 + 4 * np.cos(theta)
    y_bg = 0 + 4 * np.sin(theta)
    ax.fill_between(x_bg, y_bg, 0, color="#E0E0E0", zorder=1)

    # Arc coloré selon score
    ratio = score_total / 100
    theta_score = np.linspace(np.pi, np.pi - ratio * np.pi, 200)
    if score_total <= 39:   col = "#C00000"
    elif score_total <= 69: col = "#C55A11"
    else:                   col = "#1F6B3A"

    x_sc = 5 + 4 * np.cos(theta_score)
    y_sc = 0 + 4 * np.sin(theta_score)
    ax.fill_between(x_sc, y_sc, 0, color=col, alpha=0.85, zorder=2)

    # Cercle blanc central
    circle = plt.Circle((5, 0), 2.5, color="#F5F7FA", zorder=3)
    ax.add_patch(circle)

    # Score
    ax.text(5, 0.8, f"{score_total}", ha="center", va="center",
            fontsize=26, fontweight="bold", color="#1F3864", zorder=4)
    ax.text(5, 0.1, "/100", ha="center", va="center",
            fontsize=10, color="#666666", zorder=4)

    # Légende
    ax.text(1.2, -0.3, "0\nFaible", ha="center", fontsize=7, color="#C00000")
    ax.text(5,   -0.3, "40–69\nMoyen",  ha="center", fontsize=7, color="#C55A11")
    ax.text(8.8, -0.3, "100\nÉlevé",   ha="center", fontsize=7, color="#1F6B3A")
    plt.tight_layout()
    return fig


# ══════════════════════════════════════════════════════════════
# MODULE 9 — EXPORT RAPPORT EXCEL
# ══════════════════════════════════════════════════════════════

def export_rapport_excel(sheet_name, df, comp, cohe, docu, inte, score_total, niveau, recos):
    """
    Génère un rapport Excel structuré avec 3 feuilles :
    - Résumé exécutif
    - Scores détaillés
    - Recommandations
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    BLEU_F  = PatternFill("solid", start_color="1F3864")
    BLEU2_F = PatternFill("solid", start_color="2E5496")
    VERT_F  = PatternFill("solid", start_color="E8F5E9")
    ROUGE_F = PatternFill("solid", start_color="FDECEA")
    ORANGE_F= PatternFill("solid", start_color="FFF3E0")
    GRIS_F  = PatternFill("solid", start_color="F5F7FA")

    H_FONT  = Font(name="Arial", bold=True, color="FFFFFF", size=11)
    B_FONT  = Font(name="Arial", size=10)
    T_FONT  = Font(name="Arial", bold=True, size=14, color="1F3864")

    thin = Side(style="thin", color="CCCCCC")
    STD_B = Border(top=thin, bottom=thin, left=thin, right=thin)

    def set_cell(ws, row, col, val, font=None, fill=None, align="left", border=None):
        cell = ws.cell(row=row, column=col, value=val)
        if font:   cell.font      = font
        if fill:   cell.fill      = fill
        if border: cell.border    = border
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        return cell

    # ── Feuille 1 : Résumé ──────────────────────────────────
    ws1 = wb.active
    ws1.title = "Résumé_Exécutif"
    ws1.row_dimensions[1].height = 40

    set_cell(ws1, 1, 1, "RAPPORT AgriAI-Readiness", T_FONT, BLEU_F, "center")
    ws1.merge_cells("A1:C1")
    ws1.cell(1,1).font = Font(name="Arial", bold=True, size=14, color="FFFFFF")

    infos = [
        ("Dataset analysé",   sheet_name),
        ("Date d'analyse",    datetime.now().strftime("%d/%m/%Y %H:%M")),
        ("Lignes",            df.shape[0]),
        ("Colonnes",          df.shape[1]),
        ("Score AI-Readiness",f"{score_total}/100"),
        ("Niveau",            niveau[0]),
    ]
    for i, (k, v) in enumerate(infos, 3):
        set_cell(ws1, i, 1, k, Font(name="Arial", bold=True, size=10), GRIS_F, border=STD_B)
        set_cell(ws1, i, 2, v, B_FONT, border=STD_B)

    # Scores dimensions
    set_cell(ws1, 10, 1, "SCORES PAR DIMENSION", Font(name="Arial", bold=True, size=11, color="FFFFFF"), BLEU2_F, "center")
    ws1.merge_cells("A10:C10")
    ws1.cell(10,1).font = Font(name="Arial", bold=True, size=11, color="FFFFFF")

    dims = [
        ("1 — Complétude",          comp["score"], 25),
        ("2 — Cohérence",           cohe["score"], 25),
        ("3 — Documentation",       docu["score"], 25),
        ("4 — Interopérabilité",    inte["score"], 25),
        ("TOTAL AI-Readiness",      score_total,   100),
    ]
    for i, (dim, s, m) in enumerate(dims, 11):
        is_total = i == 15
        fill = BLEU_F if is_total else GRIS_F
        font_d = Font(name="Arial", bold=True, size=10, color="FFFFFF" if is_total else "1F3864")
        set_cell(ws1, i, 1, dim, font_d, fill, border=STD_B)
        pct = s/m
        score_fill = (ROUGE_F if pct < 0.4 else ORANGE_F if pct < 0.7 else VERT_F)
        set_cell(ws1, i, 2, f"{s}/{m}", Font(name="Arial", bold=True, size=10),
                 BLEU_F if is_total else score_fill, "center", STD_B)
        pct_cell = set_cell(ws1, i, 3, f"{round(pct*100)}%", B_FONT,
                            BLEU_F if is_total else None, "center", STD_B)
        if is_total:
            pct_cell.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")

    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 16
    ws1.column_dimensions["C"].width = 12

    # ── Feuille 2 : Scores détaillés ─────────────────────────
    ws2 = wb.create_sheet("Scores_Détaillés")
    ws2.row_dimensions[1].height = 30
    for col, hdr in enumerate(["Dimension", "Sous-critère", "Score obtenu", "Score max", "Taux (%)"], 1):
        set_cell(ws2, 1, col, hdr, H_FONT, BLEU_F, "center", STD_B)

    row = 2
    for dim_name, dim_data in [
        ("Complétude", comp), ("Cohérence", cohe),
        ("Documentation", docu), ("Interopérabilité", inte)
    ]:
        for sous_crit, (s, m) in dim_data["sous_scores"].items():
            pct = round(s/m*100) if m > 0 else 0
            fill = ROUGE_F if pct < 40 else ORANGE_F if pct < 70 else VERT_F
            set_cell(ws2, row, 1, dim_name, B_FONT, GRIS_F if row%2==0 else None, border=STD_B)
            set_cell(ws2, row, 2, sous_crit, B_FONT, GRIS_F if row%2==0 else None, border=STD_B)
            set_cell(ws2, row, 3, s,    Font(name="Arial", bold=True, size=10), fill, "center", STD_B)
            set_cell(ws2, row, 4, m,    B_FONT, fill, "center", STD_B)
            set_cell(ws2, row, 5, f"{pct}%", B_FONT, fill, "center", STD_B)
            row += 1

    for i, w in enumerate([18, 22, 14, 12, 10], 1):
        ws2.column_dimensions[get_column_letter(i)].width = w

    # ── Feuille 3 : Recommandations ───────────────────────────
    ws3 = wb.create_sheet("Recommandations")
    ws3.row_dimensions[1].height = 30
    for col, hdr in enumerate(["Priorité", "Dimension", "Recommandation"], 1):
        set_cell(ws3, 1, col, hdr, H_FONT, BLEU_F, "center", STD_B)

    prio_fills = {"URGENT": ROUGE_F, "IMPORTANT": ORANGE_F, "AMÉLIORATION": VERT_F}
    for i, (prio, dim, reco) in enumerate(recos, 2):
        fill = prio_fills.get(prio, GRIS_F)
        set_cell(ws3, i, 1, prio, Font(name="Arial", bold=True, size=10), fill, "center", STD_B)
        set_cell(ws3, i, 2, dim,  B_FONT, fill, border=STD_B)
        set_cell(ws3, i, 3, reco, B_FONT, border=STD_B)
        ws3.row_dimensions[i].height = 40

    ws3.column_dimensions["A"].width = 14
    ws3.column_dimensions["B"].width = 18
    ws3.column_dimensions["C"].width = 70

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════
# INTERFACE PRINCIPALE STREAMLIT
# ══════════════════════════════════════════════════════════════

def main():

    # ── En-tête ──────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
        <h1>🌱 AgriAI-Readiness</h1>
        <p>Évaluation automatique de la maturité des données agricoles pour l'intelligence artificielle</p>
        <p style="font-size:0.78rem; opacity:0.7;">
            Basé sur : Wilkinson et al. (2016) · Wang & Strong (1996) · Sambasivan et al. (2021) · Andrew Ng (Data-Centric AI)
        </p>
    </div>
    """, unsafe_allow_html=True)

    # ── Sidebar ───────────────────────────────────────────────
    with st.sidebar:
        st.markdown("## ⚙️ Configuration")
        st.markdown("---")
        st.markdown("### 📂 Chargement du fichier")
        uploaded_file = st.file_uploader(
            "Fichier Excel (.xlsx) ou CSV",
            type=["xlsx", "xls", "csv"],
            help="Taille max recommandée : 50 Mo"
        )

        st.markdown("---")
        st.markdown("### ℹ️ À propos")
        st.markdown("""
        **AgriAI-Readiness** évalue automatiquement la qualité des données agricoles
        selon 4 dimensions scientifiques.

        **Score /100 :**
        - 🔴 0–39 : Faible
        - 🟡 40–69 : Moyen
        - 🟢 70–100 : Élevé

        **Références :**
        - Wilkinson et al. (2016)
        - Wang & Strong (1996)
        - Sambasivan et al. (2021)
        - Andrew Ng — Data-Centric AI
        """)
        st.markdown("---")
        st.caption("Mémoire DU Données — UCAD/EBAD\nPatrick-Eudess Zatty ALLA — 2026")

    # ── Page d'accueil si pas de fichier ─────────────────────
    if uploaded_file is None:
        col1, col2, col3, col4 = st.columns(4)
        dims_info = [
            ("📊", "Complétude", "Valeurs manquantes\net taux de renseignement", "/25"),
            ("🔁", "Cohérence",  "Formats, codes\net types de données", "/25"),
            ("📄", "Documentation", "Dictionnaire, README\net noms de colonnes", "/25"),
            ("🔗", "Interopérabilité", "Compatibilité ML,\nidentifiants, structure", "/25"),
        ]
        for col, (icon, title, desc, pts) in zip([col1, col2, col3, col4], dims_info):
            with col:
                st.markdown(f"""
                <div class="score-card">
                    <div style="font-size:2rem">{icon}</div>
                    <div class="score-dim">{title}</div>
                    <div style="font-size:0.8rem; color:#666; margin:0.3rem 0">{desc}</div>
                    <div class="score-number">{pts}</div>
                </div>
                """, unsafe_allow_html=True)

        st.info("👆 Chargez un fichier Excel ou CSV dans la barre latérale pour démarrer l'analyse.")
        return

    # ── Chargement ────────────────────────────────────────────
    sheets = load_data(uploaded_file)
    if not sheets:
        st.error("Fichier invalide ou vide.")
        return

    # Sélection de la feuille si Excel multi-feuilles
    sheet_names = list(sheets.keys())
    if len(sheet_names) > 1:
        st.markdown('<div class="section-title">📋 Sélection de la feuille à analyser</div>', unsafe_allow_html=True)
        selected = st.selectbox("Feuille active :", sheet_names)
        st.info(f"💡 Feuilles disponibles : {', '.join(sheet_names)}. "
                "La présence de feuilles 'README' ou 'Dictionnaire' est prise en compte dans le score Documentation.")
    else:
        selected = sheet_names[0]

    df = sheets[selected]

    # ── Aperçu des données ───────────────────────────────────
    with st.expander("📊 Aperçu des données", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Lignes",    f"{df.shape[0]:,}")
        c2.metric("Colonnes",  f"{df.shape[1]}")
        c3.metric("Feuilles",  f"{len(sheet_names)}")
        c4.metric("Valeurs manquantes", f"{df.isnull().sum().sum():,}")
        st.dataframe(df.head(8), use_container_width=True)

    # ── Calcul des scores ─────────────────────────────────────
    with st.spinner("🔍 Analyse en cours..."):
        comp = evaluate_completude(df)
        cohe = evaluate_coherence(df)
        docu = evaluate_documentation(df, sheets, selected)
        inte = evaluate_interoperabilite(df, sheets)
        score_total, niveau = compute_score_global(comp, cohe, docu, inte)
        recos = generate_recommendations(comp, cohe, docu, inte)

    # ── Score global ─────────────────────────────────────────
    st.markdown('<div class="section-title">🎯 Score AI-Readiness Global</div>', unsafe_allow_html=True)

    col_gauge, col_scores = st.columns([1, 2])

    with col_gauge:
        fig_gauge = plot_gauge(score_total)
        st.pyplot(fig_gauge, use_container_width=True)
        st.markdown(f"""
        <div style="text-align:center; margin-top:-0.5rem">
            <span class="{niveau[1]}">{niveau[2]} Niveau {niveau[0]}</span>
        </div>
        """, unsafe_allow_html=True)

    with col_scores:
        dims_data = [
            ("📊 Complétude",        comp["score"], 25),
            ("🔁 Cohérence",         cohe["score"], 25),
            ("📄 Documentation",     docu["score"], 25),
            ("🔗 Interopérabilité",  inte["score"], 25),
        ]
        for label, score, max_s in dims_data:
            pct = score / max_s
            color = "#C00000" if pct < 0.4 else "#C55A11" if pct < 0.7 else "#1F6B3A"
            st.markdown(f"""
            <div class="score-card" style="border-left-color:{color}">
                <div style="display:flex; justify-content:space-between; align-items:center">
                    <div class="score-dim">{label}</div>
                    <div class="score-number" style="color:{color}">{score}<span style="font-size:1rem;color:#999">/{max_s}</span></div>
                </div>
                <div style="background:#E0E0E0;border-radius:10px;height:8px;margin-top:0.4rem">
                    <div style="background:{color};width:{pct*100:.0f}%;height:8px;border-radius:10px"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

    # ── Graphique radar ──────────────────────────────────────
    st.markdown('<div class="section-title">📈 Visualisations</div>', unsafe_allow_html=True)
    col_r, col_b = st.columns([1, 2])

    with col_r:
        scores_dims = [comp["score"], cohe["score"], docu["score"], inte["score"]]
        fig_radar = plot_radar(scores_dims, ["Complétude","Cohérence","Documentation","Interopérabilité"])
        st.pyplot(fig_radar, use_container_width=True)

    with col_b:
        # Tous les sous-critères
        all_labels, all_scores, all_maxs = [], [], []
        for dim_data in [comp, cohe, docu, inte]:
            for k, (s, m) in dim_data["sous_scores"].items():
                all_labels.append(k)
                all_scores.append(s)
                all_maxs.append(m)
        fig_barres = plot_barres(all_scores, all_labels, all_maxs)
        st.pyplot(fig_barres, use_container_width=True)

    # ── Détails par dimension ─────────────────────────────────
    st.markdown('<div class="section-title">🔬 Diagnostic détaillé</div>', unsafe_allow_html=True)

    tab1, tab2, tab3, tab4 = st.tabs(["📊 Complétude", "🔁 Cohérence", "📄 Documentation", "🔗 Interopérabilité"])

    with tab1:
        c1, c2, c3 = st.columns(3)
        c1.metric("Taux manquants", f"{comp['taux_global_pct']}%")
        c2.metric("Colonnes incomplètes (>5%)", comp['cols_incomplete'])
        c3.metric("Colonnes entièrement vides", comp['cols_vides'])
        if not comp["top_missing"].empty:
            st.markdown("**Colonnes les plus lacunaires :**")
            for col, pct in comp["top_missing"].items():
                color = "red" if pct > 0.2 else "orange"
                st.markdown(f"- `{col}` : **:{color}[{pct*100:.1f}% manquant]**")

    with tab2:
        c1, c2, c3 = st.columns(3)
        c1.metric("Codes binaires incohérents", len(cohe['cols_binaires_incoherentes']))
        c2.metric("Formats dates mixtes", len(cohe['cols_dates_mixtes']))
        c3.metric("Colonnes texte/numérique", len(cohe['cols_texte_num']))
        if cohe['cols_binaires_incoherentes']:
            st.markdown(f"**Colonnes binaires incohérentes :** `{'`, `'.join(cohe['cols_binaires_incoherentes'])}`")
        if cohe['cols_dates_mixtes']:
            st.markdown(f"**Colonnes dates mixtes :** `{'`, `'.join(cohe['cols_dates_mixtes'])}`")

    with tab3:
        c1, c2, c3 = st.columns(3)
        c1.metric("Noms explicites", f"{docu['ratio_noms_explicites']}%")
        c2.metric("Dictionnaire présent", "✅ Oui" if docu['has_dict'] else "❌ Non")
        c3.metric("README présent", "✅ Oui" if docu['has_readme'] else "❌ Non")
        st.metric("Colonnes avec unités", f"{docu['n_cols_avec_unites']} / {docu['n_cols_total']}")

    with tab4:
        c1, c2, c3 = st.columns(3)
        c1.metric("Lignes vides", inte['lignes_vides'])
        c2.metric("Features ML directes", f"{inte['ratio_features_pct']}%")
        c3.metric("Colonnes types suspects", inte['n_suspect_type'])
        if inte['id_cols']:
            st.markdown(f"**Colonnes ID détectées :** `{'`, `'.join(inte['id_cols'])}`")

    # ── Recommandations ──────────────────────────────────────
    st.markdown('<div class="section-title">💡 Recommandations</div>', unsafe_allow_html=True)

    prio_styles = {
        "URGENT":       ("reco-card reco-urgent", "🔴 URGENT"),
        "IMPORTANT":    ("reco-card",              "🟡 IMPORTANT"),
        "AMÉLIORATION": ("reco-card reco-info",    "🟢 AMÉLIORATION"),
    }
    for prio, dim, reco in recos:
        css_class, label = prio_styles[prio]
        st.markdown(f"""
        <div class="{css_class}">
            <strong>{label} — {dim}</strong><br>{reco}
        </div>
        """, unsafe_allow_html=True)

    # ── Export ───────────────────────────────────────────────
    st.markdown('<div class="section-title">📥 Export du rapport</div>', unsafe_allow_html=True)

    col_exp1, col_exp2 = st.columns(2)

    with col_exp1:
        excel_buf = export_rapport_excel(
            selected, df, comp, cohe, docu, inte, score_total, niveau, recos
        )
        st.download_button(
            label="📊 Télécharger le rapport Excel",
            data=excel_buf,
            file_name=f"rapport_agriai_{selected}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary"
        )

    with col_exp2:
        # Export CSV nettoyé (colonnes numériques forcées)
        df_export = df.copy()
        csv_buf = io.StringIO()
        df_export.to_csv(csv_buf, index=False, encoding="utf-8-sig")
        st.download_button(
            label="📄 Télécharger les données (CSV)",
            data=csv_buf.getvalue().encode("utf-8-sig"),
            file_name=f"data_{selected}_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
            use_container_width=True
        )

    # ── Pied de page ─────────────────────────────────────────
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; font-size:0.78rem; color:#888; padding:0.5rem">
        AgriAI-Readiness · Mémoire DU Données · UCAD/EBAD · Patrick-Eudess Zatty ALLA · 2026<br>
        Références : Wilkinson et al. (2016) · Wang & Strong (1996) · Sambasivan et al. (2021) · Andrew Ng (Data-Centric AI)
    </div>
    """, unsafe_allow_html=True)


# ── Point d'entrée ───────────────────────────────────────────
if __name__ == "__main__":
    main()
