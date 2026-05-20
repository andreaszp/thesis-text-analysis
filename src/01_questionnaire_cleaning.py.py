"""
01_questionnaire_cleaning.py
Block 1 — Step 1: Load raw data, clean, encode Likert scales.
Saves: outputs/df_clean.json, outputs/df_fl21.json, outputs/df_fl22.json
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from config import (OUTPUT_DIR, ALL_Q_COLS, NUM_COLS, EXCEL_LETTERS,
                    Q_LABELS, COL_SCALE, LIKERT_7, CAPABLE_7, AMOUNT_7, CONSTRUCTS)
from loader import load_clean_data

def run():
    df_clean, df_fl21, df_fl22 = load_clean_data()

    # ── Scale name map ────────────────────────────────────
    scale_name = {
        c: ("Likert"  if m is LIKERT_7
            else "Capable" if m is CAPABLE_7
            else "Amount")
        for c, m in COL_SCALE.items()
    }

    construct_map = {}
    for constr, cols in CONSTRUCTS.items():
        for c in cols:
            construct_map[c] = constr

    # ── Coding summary ────────────────────────────────────
    coding_rows = []
    for col, el, ql in zip(ALL_Q_COLS, EXCEL_LETTERS, Q_LABELS):
        nc       = col + "_num"
        n_coded  = df_clean[nc].notna().sum()
        n_total  = df_clean[col].notna().sum()
        coding_rows.append({
            "excel_col":   el,
            "variable":    col,
            "scale":       scale_name[col],
            "construct":   construct_map.get(col, ""),
            "label":       ql,
            "n_coded":     int(n_coded),
            "n_total":     int(n_total),
            "pct_coded":   round(n_coded / n_total * 100, 1) if n_total > 0 else 0,
        })
    df_coding = pd.DataFrame(coding_rows)

    # ── Descriptive stats per group ───────────────────────
    desc_rows = []
    for col, nc, el, ql in zip(ALL_Q_COLS, NUM_COLS, EXCEL_LETTERS, Q_LABELS):
        g21 = df_fl21[nc].dropna().values
        g22 = df_fl22[nc].dropna().values
        desc_rows.append({
            "excel_col":   el,
            "variable":    col,
            "label":       ql,
            "scale":       scale_name[col],
            "construct":   construct_map.get(col, ""),
            "n_fl21":      len(g21),
            "mean_fl21":   round(float(np.mean(g21)), 3) if len(g21) else np.nan,
            "sd_fl21":     round(float(np.std(g21, ddof=1)), 3) if len(g21) > 1 else np.nan,
            "median_fl21": round(float(np.median(g21)), 1) if len(g21) else np.nan,
            "n_fl22":      len(g22),
            "mean_fl22":   round(float(np.mean(g22)), 3) if len(g22) else np.nan,
            "sd_fl22":     round(float(np.std(g22, ddof=1)), 3) if len(g22) > 1 else np.nan,
            "median_fl22": round(float(np.median(g22)), 1) if len(g22) else np.nan,
        })
    df_desc = pd.DataFrame(desc_rows)

    # ── Save ──────────────────────────────────────────────
    df_clean.to_json(OUTPUT_DIR / "df_clean.json",  orient="records", force_ascii=False)
    df_fl21.to_json( OUTPUT_DIR / "df_fl21.json",   orient="records", force_ascii=False)
    df_fl22.to_json( OUTPUT_DIR / "df_fl22.json",   orient="records", force_ascii=False)
    df_coding.to_json(OUTPUT_DIR / "df_coding.json",orient="records", force_ascii=False)
    df_desc.to_json(  OUTPUT_DIR / "df_desc.json",  orient="records", force_ascii=False)

    print(f"\nCleaning complete — {len(df_clean)} respondents kept")
    print(f"  FL_21: {len(df_fl21)}  FL_22: {len(df_fl22)}")
    print(f"  Saved: df_clean.json, df_fl21.json, df_fl22.json, df_coding.json, df_desc.json")
    return df_clean, df_fl21, df_fl22, df_coding, df_desc

if __name__ == "__main__":
    run()
