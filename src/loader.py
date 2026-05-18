"""
loader.py — Chargement et nettoyage des données.
Partagé par tous les modules.
"""
import pandas as pd
import numpy as np
from config import EXCEL_FILE, SHEET_NAME, MSG_COLS, RESP_COLS

def load_clean_data():
    df_raw  = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=0)
    df      = df_raw.iloc[1:].copy().reset_index(drop=True)
    df["Progress"] = pd.to_numeric(df["Progress"], errors="coerce")
    ay_ok   = df["msg_2"].notna() & (df["msg_2"].astype(str).str.strip() != "")
    mask    = (df["Progress"] == 100) & ay_ok
    df_clean = df[mask].copy().reset_index(drop=True)

    def count_valid(row, cols):
        return sum(1 for c in cols if pd.notna(row.get(c)) and str(row.get(c)).strip())

    df_clean["n_turns"]    = df_clean.apply(lambda r: count_valid(r, MSG_COLS), axis=1)
    df_clean["n_bot_msgs"] = df_clean.apply(lambda r: count_valid(r, RESP_COLS), axis=1)

    df_fl21 = df_clean[df_clean["FL_13_DO"] == "FL_21"].copy().reset_index(drop=True)
    df_fl22 = df_clean[df_clean["FL_13_DO"] == "FL_22"].copy().reset_index(drop=True)

    print(f"Données chargées — n_total={len(df_clean)}, FL_21={len(df_fl21)}, FL_22={len(df_fl22)}")
    return df_clean, df_fl21, df_fl22

def collect_messages(df_group, cols, include_version=True):
    """Retourne une liste de dicts {respondent, version, turn, text}."""
    records = []
    for _, row in df_group.iterrows():
        for i, col in enumerate(cols, 1):
            txt = row.get(col)
            if pd.notna(txt) and str(txt).strip():
                rec = {"respondent": row["ResponseId"], "turn": i, "text": str(txt).strip()}
                if include_version:
                    rec["version"] = row["FL_13_DO"]
                records.append(rec)
    return records

def reconstruct_conversation(row):
    """Retourne la conversation complète sous forme de texte formaté."""
    lines = []
    n = row["n_turns"]
    for i in range(1, n + 2):
        msg  = row.get(f"msg_{i}")
        resp = row.get(f"response_{i}")
        if pd.notna(msg)  and str(msg).strip():
            lines.append(f"[P{i}]: {str(msg).strip()}")
        if pd.notna(resp) and str(resp).strip():
            lines.append(f"[B{i}]: {str(resp).strip()}")
    return "\n".join(lines)

exit code 0
