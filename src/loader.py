"""
loader.py
Data loading, cleaning, and conversation reconstruction.
Shared by all modules.
"""
import pandas as pd
import numpy as np
from config import EXCEL_FILE, SHEET_NAME, MSG_COLS, RESP_COLS, COL_SCALE, NUM_COLS

def load_clean_data():
    """
    Load raw data, apply filters, encode Likert scales.
    Filters:
      - Progress == 100
      - At least 1 valid chatbot response
    Returns: df_clean, df_fl21, df_fl22
    """
    df_raw = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME, header=0)
    df     = df_raw.iloc[1:].copy().reset_index(drop=True)
    df["Progress"] = pd.to_numeric(df["Progress"], errors="coerce")

    def count_valid(row, cols):
        return sum(1 for c in cols if pd.notna(row.get(c)) and str(row.get(c)).strip())

    df["n_turns"]    = df.apply(lambda r: count_valid(r, MSG_COLS),  axis=1)
    df["n_bot_msgs"] = df.apply(lambda r: count_valid(r, RESP_COLS), axis=1)

    mask = (df["Progress"] == 100) & (df["n_bot_msgs"] >= 1)
    df_clean = df[mask].copy().reset_index(drop=True)

    # Encode Likert scales
    for col, mapping in COL_SCALE.items():
        df_clean[col + "_num"] = df_clean[col].map(mapping)

    df_fl21 = df_clean[df_clean["FL_13_DO"] == "FL_21"].copy().reset_index(drop=True)
    df_fl22 = df_clean[df_clean["FL_13_DO"] == "FL_22"].copy().reset_index(drop=True)

    print(f"Data loaded — total={len(df_clean)}, FL_21={len(df_fl21)}, FL_22={len(df_fl22)}")
    print(f"Excluded: {len(df_raw)-1-len(df_clean)} respondents")
    return df_clean, df_fl21, df_fl22

def collect_messages(df_group, cols):
    """Return list of {respondent, version, turn, text} dicts."""
    records = []
    for _, row in df_group.iterrows():
        for i, col in enumerate(cols, 1):
            txt = row.get(col)
            if pd.notna(txt) and str(txt).strip():
                records.append({
                    "respondent": row["ResponseId"],
                    "version":    row["FL_13_DO"],
                    "turn":       i,
                    "text":       str(txt).strip(),
                })
    return records

def reconstruct_conversation(row):
    """Return full conversation as formatted string."""
    lines = []
    for i in range(1, row["n_turns"] + 2):
        msg  = row.get(f"msg_{i}")
        resp = row.get(f"response_{i}")
        if pd.notna(msg)  and str(msg).strip():
            lines.append(f"[P{i}]: {str(msg).strip()}")
        if pd.notna(resp) and str(resp).strip():
            lines.append(f"[B{i}]: {str(resp).strip()}")
    return "\n".join(lines)
