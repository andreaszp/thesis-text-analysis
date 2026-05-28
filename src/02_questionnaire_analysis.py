"""
02_questionnaire_analysis.py
Likert scale descriptive statistics per group.
All comparative analyses (t-tests, constructs, mediations) are handled
by 08_new_analyses.py which feeds into 07_export.py.
Saves: outputs/df_desc.json
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from config import (OUTPUT_DIR, ALL_Q_COLS, NUM_COLS, EXCEL_LETTERS,
                    Q_LABELS, COL_SCALE, LIKERT_7, CAPABLE_7, AMOUNT_7,
                    CONSTRUCTS)
from loader import load_clean_data

def run():
    df_clean, df_fl21, df_fl22 = load_clean_data()

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

    # Descriptive stats per group
    desc_rows = []
    for col, nc, el, ql in zip(ALL_Q_COLS, NUM_COLS, EXCEL_LETTERS, Q_LABELS):
        g21 = df_fl21[nc].dropna().values if nc in df_fl21 else np.array([])
        g22 = df_fl22[nc].dropna().values if nc in df_fl22 else np.array([])
        desc_rows.append({
            "excel_col":   el,
            "variable":    col,
            "label":       ql,
            "scale":       scale_name.get(col, ""),
            "construct":   construct_map.get(col, ""),
            "n_fl21":      len(g21),
            "mean_fl21":   round(float(np.mean(g21)), 3)   if len(g21) else None,
            "sd_fl21":     round(float(np.std(g21, ddof=1)), 3) if len(g21) > 1 else None,
            "median_fl21": round(float(np.median(g21)), 1) if len(g21) else None,
            "n_fl22":      len(g22),
            "mean_fl22":   round(float(np.mean(g22)), 3)   if len(g22) else None,
            "sd_fl22":     round(float(np.std(g22, ddof=1)), 3) if len(g22) > 1 else None,
            "median_fl22": round(float(np.median(g22)), 1) if len(g22) else None,
        })

    df_desc = pd.DataFrame(desc_rows)
    df_desc.to_json(OUTPUT_DIR / "df_desc.json", orient="records", force_ascii=False)

    print(f"Descriptive stats complete — {len(desc_rows)} variables")
    print(f"  FL_21: {len(df_fl21)}  FL_22: {len(df_fl22)}")
    print(f"  Saved: df_desc.json")
    return df_desc

if __name__ == "__main__":
    run()
