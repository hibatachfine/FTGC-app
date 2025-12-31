# ---- COMPOSANTS (ILLIMITÉ + DÉCALAGE) ----

# Anchors du template (cellule de départ de chaque zone)
anchors = {
    "CAB_START": cell_to_rc("B18"),
    "MOT_START": cell_to_rc("F18"),
    "CH_START":  cell_to_rc("H18"),

    "CAB_OPT": cell_to_rc("B38"),
    "MOT_OPT": cell_to_rc("F38"),
    "CH_OPT":  cell_to_rc("H38"),

    "CAISSE_START": cell_to_rc("B40"),
    "CAISSE_OPT":   cell_to_rc("B47"),

    "GF_START": cell_to_rc("B50"),
    "GF_OPT":   cell_to_rc("B58"),

    "HAY_START": cell_to_rc("B61"),
    "HAY_OPT":   cell_to_rc("B68"),
}

# Hauteurs "de base" du modèle (ce qu’il y a déjà dans ton template)
BASE = {
    "CAB_MAIN": 17,
    "CAB_OPT":  3,
    "MOT_MAIN": 17,
    "MOT_OPT":  3,
    "CH_MAIN":  17,
    "CH_OPT":   3,
    "CAISSE_MAIN": 5,
    "CAISSE_OPT":  2,
    "GF_MAIN": 6,
    "GF_OPT":  2,
    "HAY_MAIN": 5,
    "HAY_OPT":  3,
}

def ensure_space(section_key: str, start_anchor_key: str, base_rows: int, needed_rows: int):
    """
    Insère des lignes si needed_rows > base_rows, juste après le bloc,
    et décale toutes les ancres en dessous.
    """
    nonlocal anchors
    extra = max(0, needed_rows - base_rows)
    if extra > 0:
        start_col, start_row = anchors[start_anchor_key]
        insert_at = start_row + base_rows  # juste après la zone de base
        anchors = insert_rows_and_shift(anchors, insert_at, extra)

# ---- build values (illimité : on ne coupe plus) ----
cab_vals = build_values(cab_prod_row, cab_codecol)
cab_opt_vals = build_values(cab_opt_row, cab_codecol)

mot_vals = build_values(mot_prod_row, mot_codecol)
mot_opt_vals = build_values(mot_opt_row, mot_codecol)

ch_vals = build_values(ch_prod_row, ch_codecol)
ch_opt_vals = build_values(ch_opt_row, ch_codecol)

caisse_vals = build_values(caisse_prod_row, caisse_codecol)
caisse_opt_vals = build_values(caisse_opt_row, caisse_codecol)

gf_vals = build_values(gf_prod_row, gf_codecol)
gf_opt_vals = build_values(gf_opt_row, gf_codecol)

hay_vals = build_values(hay_prod_row, hay_codecol)
hay_opt_vals = build_values(hay_opt_row, hay_codecol)

# ---- 1) zone du haut (CAB/MOT/CH) : on prend le max des 3, et on décale tout ce qui est dessous ----
top_needed = max(len(cab_vals), len(mot_vals), len(ch_vals), 1)
ensure_space("TOP_MAIN", "CAB_START", BASE["CAB_MAIN"], top_needed)

# écrire (n_rows = top_needed, pas de limite)
write_block_merged_safe("B18", cab_vals, top_needed)
write_block_merged_safe("F18", mot_vals, top_needed)
write_block_merged_safe("H18", ch_vals,  top_needed)

# ---- 2) options du haut ----
opt_needed = max(len(cab_opt_vals), len(mot_opt_vals), len(ch_opt_vals), 1)
ensure_space("TOP_OPT", "CAB_OPT", BASE["CAB_OPT"], opt_needed)

write_block_merged_safe("B38", cab_opt_vals, opt_needed)
write_block_merged_safe("F38", mot_opt_vals, opt_needed)
write_block_merged_safe("H38", ch_opt_vals,  opt_needed)

# ---- 3) caisse ----
caisse_needed = max(len(caisse_vals), 1)
ensure_space("CAISSE_MAIN", "CAISSE_START", BASE["CAISSE_MAIN"], caisse_needed)
write_block_merged_safe("B40", caisse_vals, caisse_needed)

caisse_opt_needed = max(len(caisse_opt_vals), 1)
ensure_space("CAISSE_OPT", "CAISSE_OPT", BASE["CAISSE_OPT"], caisse_opt_needed)
write_block_merged_safe("B47", caisse_opt_vals, caisse_opt_needed)

# ---- 4) frigo ----
gf_needed = max(len(gf_vals), 1)
ensure_space("GF_MAIN", "GF_START", BASE["GF_MAIN"], gf_needed)
write_block_merged_safe("B50", gf_vals, gf_needed)

gf_opt_needed = max(len(gf_opt_vals), 1)
ensure_space("GF_OPT", "GF_OPT", BASE["GF_OPT"], gf_opt_needed)
write_block_merged_safe("B58", gf_opt_vals, gf_opt_needed)

# ---- 5) hayon ----
hay_needed = max(len(hay_vals), 1)
ensure_space("HAY_MAIN", "HAY_START", BASE["HAY_MAIN"], hay_needed)
write_block_merged_safe("B61", hay_vals, hay_needed)

hay_opt_needed = max(len(hay_opt_vals), 1)
ensure_space("HAY_OPT", "HAY_OPT", BASE["HAY_OPT"], hay_opt_needed)
write_block_merged_safe("B68", hay_opt_vals, hay_opt_needed)
