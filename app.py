# =========================
# BDD bdd_CG-JL.xlsx helpers
# =========================
import re
import unicodedata
import pandas as pd
import streamlit as st

# Colonnes "code" (identifiants produit) par feuille POUR TA BDD
CODE_COL = {
    "CABINES": "c_cabine",
    "MOTEURS": "m_moteur",
    "CHASSIS": "ch_chassis",              # <= IMPORTANT (pas c_chassis)
    "CAISSES": "cf_caisse",
    "FRIGO": "gf_groupe_frigo",
    "HAYONS": "hl_hayon_elevateur",
}

_BAD_TOKENS = {"na", "-", "_", ""}

def _pick_header_label(raw: object) -> str:
    """
    Tes headers Excel sont souvent sur 2-3 lignes (avec 'NA' en 1ère ligne).
    On prend la 1ère ligne "utile" (≠ NA / - / _).
    """
    s = "" if raw is None else str(raw)
    lines = [ln.strip() for ln in s.splitlines() if ln.strip()]
    for ln in lines:
        if ln.lower() not in _BAD_TOKENS:
            return ln
    return lines[0] if lines else ""

def _clean_colname(raw: object) -> str:
    label = _pick_header_label(raw)

    # minuscules + suppression accents
    label = "".join(
        ch for ch in unicodedata.normalize("NFD", label)
        if unicodedata.category(ch) != "Mn"
    ).lower().strip()

    # nettoyage + espaces -> underscore
    label = re.sub(r"[^\w\s-]", "", label)
    label = re.sub(r"\s+", " ", label).strip().replace(" ", "_")
    return label

def normalize_df_cols(df: pd.DataFrame) -> pd.DataFrame:
    """
    - normalise les noms de colonnes (multi-lignes, accents, espaces)
    - supprime la grosse colonne "concat" (celle qui répète toutes les colonnes à la fin)
    - garantit l'unicité des noms (suffixes _2, _3...)
    """
    df = df.copy()

    new_cols = []
    seen = {}

    keep_cols = []
    for raw in df.columns:
        raw_s = "" if raw is None else str(raw)

        # Drop colonne "concat" (énorme header qui recopie tout)
        if raw_s.count("\n") > 20 and len(raw_s) > 200:
            continue

        col = _clean_colname(raw)

        # si vide, on la garde mais on lui donne un nom technique
        if not col:
            col = "col"

        if col in seen:
            seen[col] += 1
            col = f"{col}_{seen[col]}"
        else:
            seen[col] = 1

        keep_cols.append(raw)
        new_cols.append(col)

    df = df[keep_cols]
    df.columns = new_cols
    return df

def _norm_code(x: object) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    # évite les "nan" pandas
    if s.lower() in {"nan", "none"}:
        return ""
    return s

def find_row(df: pd.DataFrame, code: object, code_col: str):
    """
    Retourne la 1ère ligne matching (Series) ou None.
    Erreur explicite si la colonne code_col n'existe pas.
    """
    target = _norm_code(code)
    if not target:
        return None

    if code_col not in df.columns:
        raise KeyError(
            f"Colonne '{code_col}' introuvable. Colonnes dispo: {list(df.columns)}"
        )

    col = df[code_col].map(_norm_code)
    cand = df[col == target]
    if cand.empty:
        return None
    return cand.iloc[0]

@st.cache_data(show_spinner=False)
def load_bdd(bdd_source) -> dict[str, pd.DataFrame]:
    """
    bdd_source = chemin str OU fichier upload Streamlit.
    """
    out = {}
    for sheet in CODE_COL.keys():
        df = pd.read_excel(bdd_source, sheet_name=sheet)
        df = normalize_df_cols(df)
        out[sheet] = df
    return out
