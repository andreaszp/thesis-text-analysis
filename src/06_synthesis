"""
06_synthesis.py
Block 2 — Step 4: Full synthesis of all analyses.
Merges metrics + AI results + Likert scores per participant.
Runs t-tests, chi2, correlations, compliance analysis, dropout analysis.
Saves: outputs/synthesis.json, df_merged.json
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from scipy import stats
from scipy.stats import chi2_contingency, pearsonr

from config import OUTPUT_DIR, NUM_COLS, ALL_Q_COLS, Q_LABELS, EXCEL_LETTERS

# ================================================================
# STATISTICAL HELPERS
# ================================================================

def cohens_d(g1, g2):
    g1, g2 = np.array(g1), np.array(g2)
    n1, n2 = len(g1), len(g2)
    if n1 < 2 or n2 < 2: return np.nan
    pool = np.sqrt(((n1-1)*np.std(g1,ddof=1)**2+(n2-1)*np.std(g2,ddof=1)**2)/(n1+n2-2))
    return (np.mean(g1)-np.mean(g2))/pool if pool > 0 else np.nan

def interpret_d(d):
    if pd.isna(d) or d is None: return "n/a"
    d = abs(d)
    if d < 0.2: return "negligible"
    if d < 0.5: return "small"
    if d < 0.8: return "medium"
    return "large"

def sig_label(p):
    if p is None or (isinstance(p, float) and np.isnan(p)): return "n/a"
    if p < 0.001: return "***"
    if p < 0.01:  return "**"
    if p < 0.05:  return "*"
    return "ns"

def ttest(df, col, label):
    """Independent t-test FL_21 vs FL_22 for a numeric column."""
    g21 = df[df["version"]=="FL_21"][col].dropna().values
    g22 = df[df["version"]=="FL_22"][col].dropna().values
    if len(g21) < 2 or len(g22) < 2: return None
    t_s, p_v = stats.ttest_ind(g21, g22)
    d        = cohens_d(g21, g22)
    return {
        "variable":   col, "label": label,
        "n_fl21":     len(g21),
        "mean_fl21":  round(float(np.mean(g21)), 3),
        "sd_fl21":    round(float(np.std(g21, ddof=1)), 3),
        "median_fl21":round(float(np.median(g21)), 3),
        "n_fl22":     len(g22),
        "mean_fl22":  round(float(np.mean(g22)), 3),
        "sd_fl22":    round(float(np.std(g22, ddof=1)), 3),
        "median_fl22":round(float(np.median(g22)), 3),
        "delta":      round(float(np.mean(g21)-np.mean(g22)), 3),
        "t":          round(float(t_s), 4),
        "p":          round(float(p_v), 4),
        "sig":        sig_label(p_v),
        "cohens_d":   round(float(d), 3) if not np.isnan(d) else None,
        "effect_size":interpret_d(d),
    }

def chi2_test(df, col, label):
    """Chi2 test for a binary column (0/1)."""
    try:
        ct = pd.crosstab(df["version"], df[col])
        if ct.shape[1] < 2: return None
        chi2, p_v, _, _ = chi2_contingency(ct)
        pct21 = df[df["version"]=="FL_21"][col].mean()
        pct22 = df[df["version"]=="FL_22"][col].mean()
        return {
            "variable":  col, "label": label,
            "pct_fl21":  round(pct21*100, 1),
            "pct_fl22":  round(pct22*100, 1),
            "pct_total": round(df[col].mean()*100, 1),
            "delta_pct": round((pct21-pct22)*100, 1),
            "chi2":      round(chi2, 4),
            "p":         round(p_v, 4),
            "sig":       sig_label(p_v),
        }
    except Exception:
        return None

def pearson_corr(df, col_x, col_y, lbl_x, lbl_y):
    """Pearson correlation between two numeric columns."""
    tmp = df[[col_x, col_y]].dropna()
    if len(tmp) < 10: return None
    r, p = pearsonr(tmp[col_x], tmp[col_y])
    return {
        "var_x": col_x, "label_x": lbl_x,
        "var_y": col_y, "label_y": lbl_y,
        "r":     round(r, 3),
        "p":     round(p, 4),
        "sig":   sig_label(p),
        "strength": "strong" if abs(r)>=0.5 else "moderate" if abs(r)>=0.3 else "weak",
        "direction":"positive" if r > 0 else "negative",
    }

# ================================================================
# LOAD & MERGE
# ================================================================

def load_all():
    """Load all intermediate JSON files and merge on one row per participant."""
    df_clean = pd.read_json(OUTPUT_DIR/"df_clean.json")
    df_agg   = pd.read_json(OUTPUT_DIR/"df_agg.json")
    df_ai    = pd.read_json(OUTPUT_DIR/"df_ai.json")   if (OUTPUT_DIR/"df_ai.json").exists()   else pd.DataFrame()
    df_prog  = pd.read_json(OUTPUT_DIR/"df_progression.json") if (OUTPUT_DIR/"df_progression.json").exists() else pd.DataFrame()

    # Keep only needed columns from df_clean
    keep = ["ResponseId","FL_13_DO","n_turns","n_bot_msgs"] + \
           [nc for nc in NUM_COLS if nc in df_clean.columns]
    df_base = df_clean[[c for c in keep if c in df_clean.columns]].copy()
    df_base.rename(columns={"FL_13_DO":"version"}, inplace=True)

    # Merge aggregated metrics
    df_agg2 = df_agg.rename(columns={"respondent":"ResponseId"})
    df = df_base.merge(df_agg2, on=["ResponseId","version"], how="left")

    # Merge AI results
    if not df_ai.empty and "respondent_id" in df_ai.columns:
        df_ai2 = df_ai.rename(columns={"respondent_id":"ResponseId"})
        ai_cols = [c for c in df_ai2.columns if c not in df.columns or c == "ResponseId"]
        df = df.merge(df_ai2[ai_cols], on="ResponseId", how="left")

    print(f"Merged dataset: {len(df)} participants × {len(df.columns)} columns")
    return df, df_prog

# ================================================================
# ANALYSIS FUNCTIONS
# ================================================================

def run_all_ttests(df):
    """T-tests for all continuous DVs by chatbot tone."""
    continuous = [
        # AI quality scores
        ("quality_global",      "AI quality score (global 1-5)"),
        ("quality_precision",   "AI quality — precision (1-5)"),
        ("quality_examples",    "AI quality — examples (1-5)"),
        ("quality_relevance",   "AI quality — relevance (1-5)"),
        ("quality_richness",    "AI quality — richness (1-5)"),
        # AI actionability
        ("action_global",       "AI actionability (global 1-5)"),
        # AI profile
        ("profile_engagement",  "AI engagement score (1-5)"),
        ("profile_expertise",   "AI perceived expertise (1-5)"),
        # Computed metrics
        ("avg_words_per_msg",   "Avg words per message"),
        ("avg_word_len",        "Avg word length"),
        ("avg_ttr",             "Lexical richness (TTR)"),
        ("pct_questions",       "% messages with '?'"),
        ("avg_sentiment",       "Avg VADER sentiment"),
        # Turns
        ("n_turns",             "Number of conversation turns"),
        ("breakpoint_turn",     "Breakpoint turn (if exists)"),
        # Chatbot scores
        ("bot_score_friendly",  "Chatbot friendly score (1-5)"),
        ("bot_score_professional","Chatbot professional score (1-5)"),
        ("bot_coherence_score", "Chatbot tone coherence (1-5)"),
        ("bot_compliance_score","Chatbot brief compliance (1-5)"),
        ("bot_open_q_pct",      "% open questions (chatbot)"),
        ("bot_n_friendly_words","N friendly words detected"),
        ("bot_n_formal_words",  "N formal words detected"),
        ("bot_n_encouragements","N encouragement phrases"),
        ("bot_n_sober",         "N sober phrases"),
    ]
    results = [ttest(df, col, lbl) for col, lbl in continuous
               if col in df.columns]
    return [r for r in results if r]

def run_all_chi2(df):
    """Chi2 tests for all binary DVs by chatbot tone."""
    binary = [
        ("action_concrete_pb",  "Contains concrete problem"),
        ("action_advice",       "Contains applicable advice"),
        ("action_use_case",     "Contains precise use case"),
        ("content_opinion",     "Personal opinion expressed"),
        ("content_concrete_pb", "Concrete vs vague problem"),
        ("content_suggestion",  "Feature request / suggestion"),
        ("content_emotion",     "Emotion / frustration expressed"),
        ("content_competitor",  "Competitor reference"),
        ("breakpoint_exists",   "Breakpoint detected"),
        ("completed_fully",     "Completed conversation fully"),
        ("bot_drift",           "Chatbot tone drift detected"),
        ("bot_compliant",       "Chatbot compliant with brief"),
        ("bot_emojis",          "Emojis used by chatbot"),
        ("bot_informal_addr",   "Informal address used"),
        ("bot_formal_addr",     "Formal address used"),
        ("bot_follow_ups",      "Personalised follow-ups"),
    ]
    results = [chi2_test(df, col, lbl) for col, lbl in binary
               if col in df.columns]
    return [r for r in results if r]

def run_correlations(df):
    """Pearson correlations between relevant pairs of DVs."""
    pairs = [
        # Cross-validation: computed metrics vs AI scores
        ("avg_words_per_msg","quality_global",     "Avg words/msg",      "AI quality global"),
        ("avg_ttr",          "quality_richness",   "Lexical richness",   "AI richness score"),
        ("avg_word_len",     "quality_precision",  "Avg word length",    "AI precision score"),
        ("avg_words_per_msg","profile_engagement", "Avg words/msg",      "AI engagement"),
        # Between text DVs
        ("quality_global",   "action_global",      "Quality global",     "Actionability global"),
        ("profile_engagement","quality_global",    "Engagement",         "Quality global"),
        ("profile_expertise","action_global",      "Perceived expertise","Actionability"),
        ("n_turns",          "quality_global",     "N turns",            "Quality global"),
        ("breakpoint_turn",  "quality_global",     "Breakpoint turn",    "Quality global"),
        # Compliance vs markers
        ("bot_compliance_score","bot_informal_addr","Brief compliance",  "Informal address"),
        ("bot_compliance_score","bot_emojis",       "Brief compliance",  "Emojis used"),
        ("bot_compliance_score","bot_n_friendly_words","Brief compliance","N friendly words"),
        ("bot_compliance_score","bot_n_encouragements","Brief compliance","N encouragements"),
        ("bot_score_friendly","bot_compliance_score","Friendly score",   "Brief compliance"),
        # Chatbot tone vs participant quality
        ("bot_score_friendly","quality_global",    "Chatbot friendly score","Participant quality"),
        ("bot_score_friendly","profile_engagement","Chatbot friendly score","Participant engagement"),
        ("bot_compliance_score","quality_global",  "Brief compliance",   "Participant quality"),
    ]
    # Add Likert correlations with quality (evaluation_1..6_num)
    likert_eval = [nc for nc in NUM_COLS[:6] if nc in df.columns]
    for nc, lbl in zip(likert_eval, Q_LABELS[:6]):
        pairs.append(("quality_global", nc, "AI quality global", lbl))
        pairs.append(("profile_engagement", nc, "AI engagement", lbl))

    results = []
    for cx, cy, lx, ly in pairs:
        if cx in df.columns and cy in df.columns:
            r = pearson_corr(df, cx, cy, lx, ly)
            if r: results.append(r)
    return results

def run_compliance_analysis(df):
    """Compliance analysis: per-version summary + marker correlations."""
    out = {}
    for ver in ["FL_21","FL_22"]:
        g = df[df["version"]==ver]
        out[ver] = {
            "n": len(g),
            "compliance_mean": round(g["bot_compliance_score"].mean(), 3) if "bot_compliance_score" in g else None,
            "compliance_sd":   round(g["bot_compliance_score"].std(),  3) if "bot_compliance_score" in g else None,
            "pct_compliant":   round(g["bot_compliant"].mean()*100, 1)    if "bot_compliant"        in g else None,
            "pct_emojis":      round(g["bot_emojis"].mean()*100, 1)       if "bot_emojis"           in g else None,
            "pct_informal":    round(g["bot_informal_addr"].mean()*100, 1) if "bot_informal_addr"   in g else None,
            "pct_formal":      round(g["bot_formal_addr"].mean()*100, 1)   if "bot_formal_addr"     in g else None,
            "avg_encouragements": round(g["bot_n_encouragements"].mean(), 2) if "bot_n_encouragements" in g else None,
            "avg_sober":          round(g["bot_n_sober"].mean(), 2)          if "bot_n_sober"          in g else None,
            "avg_friendly_words": round(g["bot_n_friendly_words"].mean(), 2) if "bot_n_friendly_words" in g else None,
            "avg_formal_words":   round(g["bot_n_formal_words"].mean(), 2)   if "bot_n_formal_words"   in g else None,
        }
    t_comp = ttest(df, "bot_compliance_score", "Chatbot brief compliance score")
    return {"per_version": out, "ttest_compliance": t_comp}

def run_dropout_analysis(df):
    """Profile comparison: dropouts vs completers."""
    df = df.copy()
    df["dropout"] = (df["completed_fully"] == 0).astype(int) if "completed_fully" in df.columns else 0

    dropouts   = df[df["dropout"]==1]
    completers = df[df["dropout"]==0]

    profile_vars = [
        ("quality_global",    "AI quality global"),
        ("action_global",     "AI actionability"),
        ("profile_engagement","AI engagement"),
        ("profile_expertise", "AI perceived expertise"),
        ("n_turns",           "N turns"),
        ("avg_words_per_msg", "Avg words/message"),
    ]
    comparison = []
    for col, lbl in profile_vars:
        if col not in df.columns: continue
        ga = dropouts[col].dropna().values
        gc = completers[col].dropna().values
        if len(ga) < 2 or len(gc) < 2: continue
        t_s, p_v = stats.ttest_ind(ga, gc)
        d        = cohens_d(ga, gc)
        comparison.append({
            "variable":       col, "label": lbl,
            "mean_dropouts":  round(float(np.mean(ga)), 3),
            "sd_dropouts":    round(float(np.std(ga, ddof=1)), 3),
            "mean_completers":round(float(np.mean(gc)), 3),
            "sd_completers":  round(float(np.std(gc, ddof=1)), 3),
            "t":              round(float(t_s), 4),
            "p":              round(float(p_v), 4),
            "sig":            sig_label(p_v),
            "cohens_d":       round(float(d), 3) if not np.isnan(d) else None,
            "effect_size":    interpret_d(d),
        })

    # End type distribution
    fin_dist = {"global": {}}
    if "end_type" in df.columns:
        fin_dist["global"] = {k: round(v*100,1)
                              for k,v in df["end_type"].value_counts(normalize=True).items()}
        for ver in ["FL_21","FL_22"]:
            g = df[df["version"]==ver]
            fin_dist[ver] = {k: round(v*100,1)
                             for k,v in g["end_type"].value_counts(normalize=True).items()}

    return {
        "n_dropouts":        int(df["dropout"].sum()),
        "n_completers":      int((df["dropout"]==0).sum()),
        "pct_dropout_global":round(df["dropout"].mean()*100, 1),
        "pct_dropout_fl21":  round(df[df["version"]=="FL_21"]["dropout"].mean()*100, 1),
        "pct_dropout_fl22":  round(df[df["version"]=="FL_22"]["dropout"].mean()*100, 1),
        "profile_comparison":comparison,
        "end_type_distribution": fin_dist,
    }

def run_progression_analysis(df_prog):
    """Average quality score per turn and version for curve plotting."""
    if df_prog.empty: return {}
    pivot = df_prog.groupby(["version","turn"])["quality_score"].agg(
        mean="mean", sd="std", n="count"
    ).reset_index()
    out = {"FL_21":{}, "FL_22":{}}
    for _, row in pivot.iterrows():
        ver = row["version"]
        t   = int(row["turn"])
        out[ver][t] = {
            "mean": round(row["mean"], 3),
            "sd":   round(row["sd"], 3) if not pd.isna(row["sd"]) else None,
            "n":    int(row["n"]),
        }
    return out

# ================================================================
# RUN
# ================================================================

def run():
    print("\n" + "="*60)
    print("SYNTHESIS — loading data")
    print("="*60)

    df, df_prog = load_all()

    print("\n[1/6] T-tests — continuous DVs by tone...")
    ttests_cont = run_all_ttests(df)
    print(f"  {len(ttests_cont)} variables tested")

    print("[2/6] Chi2 — binary DVs by tone...")
    chi2s = run_all_chi2(df)
    print(f"  {len(chi2s)} variables tested")

    print("[3/6] Correlations between DVs...")
    corrs = run_correlations(df)
    print(f"  {len(corrs)} pairs analysed")

    print("[4/6] Brief compliance analysis...")
    compliance = run_compliance_analysis(df)

    print("[5/6] Dropout analysis...")
    dropouts = run_dropout_analysis(df)

    print("[6/6] Turn-by-turn progression...")
    progression = run_progression_analysis(df_prog)

    synthesis = {
        "ttests_continuous": ttests_cont,
        "chi2_binary":       chi2s,
        "correlations":      corrs,
        "compliance":        compliance,
        "dropouts":          dropouts,
        "progression":       progression,
    }

    with open(OUTPUT_DIR/"synthesis.json","w",encoding="utf-8") as f:
        json.dump(synthesis, f, ensure_ascii=False, indent=2)
    df.to_json(OUTPUT_DIR/"df_merged.json", orient="records", force_ascii=False)

    # Console summary
    print("\n" + "="*60)
    print("SIGNIFICANT RESULTS (p < .05)")
    print("="*60)
    sig_t = [r for r in ttests_cont if r.get("sig") not in ("ns","n/a")]
    sig_c = [r for r in chi2s       if r.get("sig") not in ("ns","n/a")]
    sig_r = [r for r in corrs       if r.get("sig") not in ("ns","n/a")]

    print(f"\nContinuous DVs ({len(sig_t)} significant):")
    for r in sig_t:
        print(f"  {r['label']:45s}  FL21={r['mean_fl21']:.2f}  FL22={r['mean_fl22']:.2f}  "
              f"p={r['p']} {r['sig']}  d={r['cohens_d']} ({r['effect_size']})")

    print(f"\nBinary DVs ({len(sig_c)} significant):")
    for r in sig_c:
        print(f"  {r['label']:45s}  FL21={r['pct_fl21']}%  FL22={r['pct_fl22']}%  "
              f"p={r['p']} {r['sig']}")

    print(f"\nCorrelations ({len(sig_r)} significant):")
    for r in sig_r:
        print(f"  {r['label_x']:28s} <-> {r['label_y']:28s}  "
              f"r={r['r']}  p={r['p']} {r['sig']}  ({r['strength']})")

    print(f"\nDropouts: {dropouts['pct_dropout_global']}% global  "
          f"FL21={dropouts['pct_dropout_fl21']}%  FL22={dropouts['pct_dropout_fl22']}%")

    return synthesis

if __name__ == "__main__":
    run()
