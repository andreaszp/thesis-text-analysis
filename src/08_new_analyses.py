"""
08_new_analyses.py
Computes all new analyses needed for the restructured Excel export.
Run AFTER 01-06 scripts. Saves JSON files to outputs/.

Outputs:
  correlations_full.json   — full correlation matrix + bloc zooms
  vif_full.json            — VIF for every variable against all others
  tone_comparisons.json    — all t-tests / chi2 by tone (replaces q_ttests + synthesis)
  mediation_tone.json      — all mediations where tone is IV (F11)
  regression_novariable.json — regressions/t-tests without tone (F12)
  mediation_noton.json     — mediations without tone as IV (F13)
  dropout_corrected.json   — dropout with fixed end_type + E1/E2 analysis
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from scipy.stats import pearsonr, ttest_ind, chi2_contingency, f_oneway
from scipy import stats

from config import OUTPUT_DIR

# ================================================================
# HELPERS
# ================================================================

def sl(p):
    if p is None or (isinstance(p, float) and np.isnan(p)): return "n/a"
    if p < 0.001: return "***"
    if p < 0.01:  return "**"
    if p < 0.05:  return "*"
    return "ns"

def cohens_d(g1, g2):
    g1, g2 = np.array(g1), np.array(g2)
    n1, n2 = len(g1), len(g2)
    if n1 < 2 or n2 < 2: return np.nan
    pool = np.sqrt(((n1-1)*np.std(g1,ddof=1)**2+(n2-1)*np.std(g2,ddof=1)**2)/(n1+n2-2))
    return (np.mean(g1)-np.mean(g2))/pool if pool > 0 else np.nan

def interpret_d(d):
    if pd.isna(d) or d is None: return "n/a"
    a = abs(d)
    if a < 0.2: return "negligible"
    if a < 0.5: return "small"
    if a < 0.8: return "medium"
    return "large"

def ols_simple(X_mat, y_vec):
    """OLS with intercept. Returns (coefs, se, t, p, r2)."""
    Xd = np.column_stack([np.ones(len(y_vec)), X_mat])
    b = np.linalg.lstsq(Xd, y_vec, rcond=None)[0]
    resid = y_vec - Xd @ b
    dfr = len(y_vec) - Xd.shape[1]
    mse = np.sum(resid**2) / dfr
    try:
        cov = mse * np.linalg.inv(Xd.T @ Xd)
    except np.linalg.LinAlgError:
        cov = mse * np.linalg.pinv(Xd.T @ Xd)
    se = np.sqrt(np.diag(cov))
    t = b / se
    p = 2 * (1 - stats.t.cdf(np.abs(t), dfr))
    ss_tot = np.sum((y_vec - np.mean(y_vec))**2)
    r2 = 1 - np.sum(resid**2) / ss_tot if ss_tot > 0 else 0
    return b, se, t, p, r2

def run_ttest(df, col, label, iv_col="version", g1="FL_21", g2="FL_22"):
    g21 = df[df[iv_col]==g1][col].dropna().values
    g22 = df[df[iv_col]==g2][col].dropna().values
    if len(g21) < 2 or len(g22) < 2: return None
    t_s, p_v = ttest_ind(g21, g22)
    d = cohens_d(g21, g22)
    return {
        "variable": col, "label": label,
        "n_fl21": len(g21), "mean_fl21": round(float(np.mean(g21)),3),
        "sd_fl21": round(float(np.std(g21,ddof=1)),3),
        "n_fl22": len(g22), "mean_fl22": round(float(np.mean(g22)),3),
        "sd_fl22": round(float(np.std(g22,ddof=1)),3),
        "delta": round(float(np.mean(g21)-np.mean(g22)),3),
        "t": round(float(t_s),4), "p": round(float(p_v),4),
        "sig": sl(p_v), "cohens_d": round(float(d),3) if not np.isnan(d) else None,
        "effect_size": interpret_d(d),
    }

def run_chi2(df, col, label, iv_col="version"):
    try:
        ct = pd.crosstab(df[iv_col], df[col])
        if ct.shape[1] < 2: return None
        chi2, p_v, _, _ = chi2_contingency(ct)
        pct21 = df[df[iv_col]=="FL_21"][col].mean()
        pct22 = df[df[iv_col]=="FL_22"][col].mean()
        return {
            "variable": col, "label": label,
            "pct_fl21": round(pct21*100,1), "pct_fl22": round(pct22*100,1),
            "pct_total": round(df[col].mean()*100,1),
            "delta_pct": round((pct21-pct22)*100,1),
            "chi2": round(chi2,4), "p": round(p_v,4), "sig": sl(p_v),
        }
    except Exception: return None

def run_regression(df, iv_cols, dv_col, labels):
    """OLS regression, returns dict with per-predictor results + R2."""
    tmp = df[iv_cols + [dv_col]].dropna()
    if len(tmp) < 20: return None
    X = tmp[iv_cols].values
    y = tmp[dv_col].values
    b, se, t, p, r2 = ols_simple(X, y)
    result = {"n": len(tmp), "dv": dv_col, "r2": round(r2,3), "predictors": []}
    for i, lbl in enumerate(labels):
        result["predictors"].append({
            "variable": iv_cols[i], "label": lbl,
            "b": round(float(b[i+1]),4), "se": round(float(se[i+1]),4),
            "t": round(float(t[i+1]),3), "p": round(float(p[i+1]),4),
            "sig": sl(p[i+1]),
        })
    return result

def run_mediation(df, iv, m, dv, n_boot=5000, seed=42):
    tmp = df[[iv,m,dv]].dropna()
    n = len(tmp)
    if n < 20: return {"error": f"n={n} insufficient", "n": n}
    X, M, Y = tmp[iv].values, tmp[m].values, tmp[dv].values
    if np.std(X) < 1e-10 or np.std(M) < 1e-10:
        return {"error": "zero variance in IV or M", "n": n}

    b_a, se_a, t_a, p_a, _ = ols_simple(X.reshape(-1,1), M)
    b_c, se_c, t_c, p_c, _ = ols_simple(X.reshape(-1,1), Y)
    b_bc, se_bc, t_bc, p_bc, r2_full = ols_simple(np.column_stack([X,M]), Y)
    a, b, c, cp = b_a[1], b_bc[2], b_c[1], b_bc[1]
    indirect = a * b

    rng = np.random.default_rng(seed)
    boots = []
    for _ in range(n_boot):
        idx = rng.integers(0, n, size=n)
        try:
            ba,*_ = ols_simple(X[idx].reshape(-1,1), M[idx])
            bbc,*_ = ols_simple(np.column_stack([X[idx],M[idx]]), Y[idx])
            boots.append(ba[1]*bbc[2])
        except: continue
    boots = np.array(boots)
    ci_low, ci_up = np.percentile(boots, 2.5), np.percentile(boots, 97.5)
    sig_ind = bool(ci_low > 0 or ci_up < 0)

    # mediation type
    a_sig = p_a[1] < 0.05; b_sig = p_bc[2] < 0.05; cp_sig = p_bc[1] < 0.05
    if not a_sig or not b_sig:
        med_type = "No mediation (path a or b ns)"
    elif sig_ind and not cp_sig:
        med_type = "Full mediation"
    elif sig_ind and cp_sig:
        med_type = "Partial mediation"
    else:
        med_type = "Inconsistent"

    def fmt(x): return round(float(x),4) if not np.isnan(x) else None
    return {
        "n": n, "iv": iv, "mediator": m, "dv": dv,
        "path_a": {"coef":fmt(a),"se":fmt(se_a[1]),"t":fmt(t_a[1]),"p":fmt(p_a[1]),"sig":sl(p_a[1])},
        "path_b": {"coef":fmt(b),"se":fmt(se_bc[2]),"t":fmt(t_bc[2]),"p":fmt(p_bc[2]),"sig":sl(p_bc[2])},
        "path_c": {"coef":fmt(c),"se":fmt(se_c[1]),"t":fmt(t_c[1]),"p":fmt(p_c[1]),"sig":sl(p_c[1])},
        "path_cp": {"coef":fmt(cp),"se":fmt(se_bc[1]),"t":fmt(t_bc[1]),"p":fmt(p_bc[1]),"sig":sl(p_bc[1])},
        "indirect": {"coef":fmt(indirect),"ci_low":fmt(ci_low),"ci_up":fmt(ci_up),
                     "significant":sig_ind,
                     "interpretation":"Significant mediation (CI excludes 0)" if sig_ind else "No mediation (CI includes 0)"},
        "mediation_type": med_type,
    }

# ================================================================
# DATA LOADING
# ================================================================

def load_merged():
    """Load and merge all data sources into one participant-level df."""
    df9 = pd.read_json(OUTPUT_DIR/"df_ai.json") if (OUTPUT_DIR/"df_ai.json").exists() else pd.DataFrame()
    df_clean = pd.read_json(OUTPUT_DIR/"df_clean.json")
    df_agg = pd.read_json(OUTPUT_DIR/"df_agg.json") if (OUTPUT_DIR/"df_agg.json").exists() else pd.DataFrame()

    # Numeric encoding for all questionnaire items
    from config import COL_SCALE, NUM_COLS
    df_base = df_clean.copy()

    # Rename version col
    df_base = df_base.rename(columns={"FL_13_DO": "version"})

    # Merge aggregated text metrics
    if not df_agg.empty:
        df_agg2 = df_agg.rename(columns={"respondent":"ResponseId"})
        df_base = df_base.merge(df_agg2, on=["ResponseId","version"], how="left")

    # Merge AI scores
    if not df9.empty and "respondent_id" in df9.columns:
        df9b = df9.rename(columns={"respondent_id":"ResponseId"})
        ai_keep = [c for c in df9b.columns if c not in df_base.columns or c=="ResponseId"]
        df_base = df_base.merge(df9b[ai_keep], on="ResponseId", how="left")

    # Tone dummy: 1=FL_21(friendly), 0=FL_22(professional)
    df_base["tone"] = (df_base["version"]=="FL_21").astype(int)

    # Numeric conversion for AI scores
    ai_num = ["quality_global","quality_precision","quality_examples","quality_relevance",
              "quality_richness","action_concrete_pb","action_advice","action_use_case",
              "content_suggestion","breakpoint_exists","breakpoint_turn","completed_fully",
              "bot_score_friendly","bot_score_professional","bot_compliance_score",
              "bot_coherence_score","bot_n_friendly_words","bot_n_formal_words",
              "bot_n_encouragements","bot_open_q_pct"]
    for c in ai_num:
        if c in df_base.columns:
            df_base[c] = pd.to_numeric(df_base[c], errors="coerce")

    # Text metrics
    for c in ["avg_words_per_msg","avg_word_len","avg_ttr","pct_questions","avg_sentiment","n_turns"]:
        if c in df_base.columns:
            df_base[c] = pd.to_numeric(df_base[c], errors="coerce")

    # Normalise end_type
    if "end_type" in df_base.columns:
        end_map = {"mineral_response":"minimal_response","minimale_response":"minimal_response",
                   "minimal response":"minimal_response"}
        df_base["end_type"] = df_base["end_type"].replace(end_map)

    # NUM_COLS already in df_base from df_clean encoding
    for nc in NUM_COLS:
        if nc in df_base.columns:
            df_base[nc] = pd.to_numeric(df_base[nc], errors="coerce")

    return df_base

# ================================================================
# 1. FULL CORRELATION MATRIX
# ================================================================

def compute_correlations(df):
    # Define all variable groups
    eval_vars = {
        "evaluation_1_num": "E1 Required effort",
        "evaluation_2_num": "E2 Engagement felt",
        "evaluation_3_num": "E3 Chatbot appreciation",
        "evaluation_4_num": "E4 Conversation utility",
        "evaluation_5_num": "E5 Reuse intention",
        "evaluation_6_num": "E6 Chatbot preference",
    }
    quality_vars = {
        "quality_global":    "Quality global",
        "quality_precision": "Quality precision",
        "quality_examples":  "Quality examples",
        "quality_relevance": "Quality relevance",
        "quality_richness":  "Quality richness",
    }
    metric_vars = {
        "avg_words_per_msg": "Avg words/msg",
        "avg_word_len":      "Avg word length",
        "n_turns":           "N turns",
        "breakpoint_turn":   "Breakpoint turn",
    }
    action_vars = {
        "action_concrete_pb": "Concrete problem",
        "action_advice":      "Applicable advice",
        "action_use_case":    "Precise use case",
        "content_emotion":    "Emotion expressed",
    }
    psy_vars = {
        "Perceived Manipulati_1_num": "PM1 Freedom threat",
        "Perceived Manipulati_2_num": "PM2 Decision override",
        "Perceived Manipulati_3_num": "PM3 Manipulation",
        "Perceived Manipulati_4_num": "PM4 Pressure",
        "Competence_1_num":           "Comp1 Skills judgment",
        "Competence_2_num":           "Comp2 Moral judgment",
        "Moral Responsibility_1_num": "MR1 AI harm wrong",
        "Moral Responsibility_2_num": "MR2 AI responsible",
        "Moral Responsibility_3_num": "MR3 Human harm AI",
        "Moral Responsibility_4_num": "MR4 AI moral concern",
        "Sense of Independenc_1_num": "Ind1 AI plans/goals",
        "Sense of Independenc_2_num": "Ind2 AI self-control",
    }

    all_vars = {**eval_vars, **quality_vars, **metric_vars, **action_vars, **psy_vars}

    # Full matrix
    available = {k:v for k,v in all_vars.items() if k in df.columns}
    cols = list(available.keys())
    labels = list(available.values())
    n_vars = len(cols)

    matrix = {"variables": labels, "columns": cols, "data": []}
    sig_pairs = []

    for i in range(n_vars):
        row_r = []
        row_p = []
        row_s = []
        for j in range(n_vars):
            tmp = df[[cols[i], cols[j]]].dropna()
            if i == j or len(tmp) < 10:
                row_r.append(1.0 if i==j else None)
                row_p.append(None)
                row_s.append("—" if i==j else "n/a")
            else:
                r, p = pearsonr(tmp[cols[i]], tmp[cols[j]])
                row_r.append(round(float(r),3))
                row_p.append(round(float(p),4))
                row_s.append(sl(p))
                if i < j and p < 0.05:
                    sig_pairs.append({
                        "var_x": cols[i], "label_x": labels[i],
                        "var_y": cols[j], "label_y": labels[j],
                        "r": round(float(r),3), "p": round(float(p),4),
                        "sig": sl(p),
                        "strength": "strong" if abs(r)>=0.5 else "moderate" if abs(r)>=0.3 else "weak",
                        "direction": "positive" if r>0 else "negative",
                    })
        matrix["data"].append({"var": cols[i], "label": labels[i],
                                "r": row_r, "p": row_p, "sig": row_s})

    sig_pairs.sort(key=lambda x: abs(x["r"]), reverse=True)

    # Bloc zooms
    blocs = [
        ("Evaluation (E1–E6)", eval_vars),
        ("Quality AI (subscores)", quality_vars),
        ("Text Metrics", metric_vars),
        ("Actionability", action_vars),
        ("Perception AI (12 items)", psy_vars),
    ]
    bloc_matrices = {}
    for bloc_name, bloc_vars in blocs:
        bv_avail = {k:v for k,v in bloc_vars.items() if k in df.columns}
        bc = list(bv_avail.keys()); bl = list(bv_avail.values())
        bdata = []
        for i in range(len(bc)):
            row = []
            for j in range(len(bc)):
                tmp = df[[bc[i],bc[j]]].dropna()
                if i==j or len(tmp)<10:
                    row.append({"r":1.0 if i==j else None,"sig":"—" if i==j else "n/a"})
                else:
                    r,p = pearsonr(tmp[bc[i]],tmp[bc[j]])
                    row.append({"r":round(float(r),3),"p":round(float(p),4),"sig":sl(p)})
            bdata.append({"var":bc[i],"label":bl[i],"cells":row})
        bloc_matrices[bloc_name] = {"cols":bc,"labels":bl,"data":bdata}

    return {"full_matrix": matrix, "significant_pairs": sig_pairs, "bloc_matrices": bloc_matrices}

# ================================================================
# 2. VIF — every variable against all others
# ================================================================

def compute_vif_all(df):
    from config import NUM_COLS
    eval_vars = ["evaluation_1_num","evaluation_2_num","evaluation_3_num",
                 "evaluation_4_num","evaluation_5_num","evaluation_6_num"]
    quality_vars = ["quality_global","quality_precision","quality_examples",
                    "quality_relevance","quality_richness"]
    metric_vars = ["avg_words_per_msg","avg_word_len","n_turns","breakpoint_turn"]
    action_vars = ["action_concrete_pb","action_advice","action_use_case"]
    psy_vars = ["Perceived Manipulati_1_num","Perceived Manipulati_2_num",
                "Perceived Manipulati_3_num","Perceived Manipulati_4_num",
                "Competence_1_num","Competence_2_num",
                "Moral Responsibility_1_num","Moral Responsibility_2_num",
                "Moral Responsibility_3_num","Moral Responsibility_4_num",
                "Sense of Independenc_1_num","Sense of Independenc_2_num"]

    all_vars = eval_vars + quality_vars + metric_vars + action_vars + psy_vars
    available = [v for v in all_vars if v in df.columns]

    labels = {
        "evaluation_1_num":"E1 Required effort","evaluation_2_num":"E2 Engagement felt",
        "evaluation_3_num":"E3 Chatbot appreciation","evaluation_4_num":"E4 Conversation utility",
        "evaluation_5_num":"E5 Reuse intention","evaluation_6_num":"E6 Chatbot preference",
        "quality_global":"Quality global","quality_precision":"Quality precision",
        "quality_examples":"Quality examples","quality_relevance":"Quality relevance",
        "quality_richness":"Quality richness",
        "avg_words_per_msg":"Avg words/msg","avg_word_len":"Avg word length",
        "n_turns":"N turns","breakpoint_turn":"Breakpoint turn",
        "action_concrete_pb":"Concrete problem","action_advice":"Applicable advice",
        "action_use_case":"Precise use case",
        "Perceived Manipulati_1_num":"PM1 Freedom threat","Perceived Manipulati_2_num":"PM2 Decision override",
        "Perceived Manipulati_3_num":"PM3 Manipulation","Perceived Manipulati_4_num":"PM4 Pressure",
        "Competence_1_num":"Comp1 Skills","Competence_2_num":"Comp2 Morality",
        "Moral Responsibility_1_num":"MR1 AI harm wrong","Moral Responsibility_2_num":"MR2 AI responsible",
        "Moral Responsibility_3_num":"MR3 Human harm AI","Moral Responsibility_4_num":"MR4 AI concern",
        "Sense of Independenc_1_num":"Ind1 AI plans","Sense of Independenc_2_num":"Ind2 AI self-control",
    }

    results = []
    tmp_full = df[available].dropna()
    for target in available:
        others = [v for v in available if v != target]
        sub = tmp_full[[target] + others].dropna()
        if len(sub) < 20:
            results.append({"variable": target, "label": labels.get(target,target), "vif": None, "r2": None, "flag": "insufficient data"})
            continue
        X = sub[others].values
        y = sub[target].values
        try:
            b, _, _, _, r2 = ols_simple(X, y)
            vif = round(1/(1-r2),2) if r2 < 0.9999 else 999.0
        except Exception:
            vif, r2 = None, None
        flag = "CRITICAL (>10)" if vif and vif>10 else "HIGH (5-10)" if vif and vif>5 else "MODERATE (3-5)" if vif and vif>3 else "OK (<3)"
        results.append({
            "variable": target, "label": labels.get(target,target),
            "vif": vif, "r2": round(r2,3) if r2 else None, "flag": flag,
        })
    return results

# ================================================================
# 3. TONE COMPARISONS (F10)
# ================================================================

def compute_tone_comparisons(df):
    # Evaluation bloc
    eval_tests = [
        run_ttest(df,"evaluation_1_num","E1 Required effort"),
        run_ttest(df,"evaluation_2_num","E2 Engagement felt"),
        run_ttest(df,"evaluation_3_num","E3 Chatbot appreciation"),
        run_ttest(df,"evaluation_4_num","E4 Conversation utility"),
        run_ttest(df,"evaluation_5_num","E5 Reuse intention"),
        run_ttest(df,"evaluation_6_num","E6 Chatbot preference"),
    ]
    # Quality AI
    quality_tests = [
        run_ttest(df,"quality_global","Quality global (1-5)"),
        run_ttest(df,"quality_precision","Quality precision (1-5)"),
        run_ttest(df,"quality_examples","Quality examples (1-5)"),
        run_ttest(df,"quality_relevance","Quality relevance (1-5)"),
        run_ttest(df,"quality_richness","Quality richness (1-5)"),
    ]
    # Text metrics
    metric_tests = [
        run_ttest(df,"avg_words_per_msg","Avg words per message"),
        run_ttest(df,"avg_word_len","Avg word length"),
        run_ttest(df,"n_turns","Number of turns"),
        run_ttest(df,"breakpoint_turn","Breakpoint turn"),
    ]
    # Actionability (binary → chi2)
    action_tests = [
        run_chi2(df,"action_concrete_pb","Contains concrete problem"),
        run_chi2(df,"action_advice","Contains applicable advice"),
        run_chi2(df,"action_use_case","Contains precise use case"),
        run_chi2(df,"content_emotion","Emotion / frustration expressed"),
    ]
    # Perception IA — 12 items t-tests
    psy_tests = [
        run_ttest(df,"Perceived Manipulati_1_num","PM1 Freedom threat"),
        run_ttest(df,"Perceived Manipulati_2_num","PM2 Decision override"),
        run_ttest(df,"Perceived Manipulati_3_num","PM3 Manipulation"),
        run_ttest(df,"Perceived Manipulati_4_num","PM4 Pressure"),
        run_ttest(df,"Competence_1_num","Comp1 Skills judgment capability"),
        run_ttest(df,"Competence_2_num","Comp2 Moral judgment capability"),
        run_ttest(df,"Moral Responsibility_1_num","MR1 Moral harm AI→human"),
        run_ttest(df,"Moral Responsibility_2_num","MR2 AI moral responsibility"),
        run_ttest(df,"Moral Responsibility_3_num","MR3 Moral harm human→AI"),
        run_ttest(df,"Moral Responsibility_4_num","MR4 AI moral concern"),
        run_ttest(df,"Sense of Independenc_1_num","Ind1 AI plans & goals"),
        run_ttest(df,"Sense of Independenc_2_num","Ind2 AI self-control"),
    ]
    # Conversation / dropout
    conv_tests = [
        run_chi2(df,"breakpoint_exists","Breakpoint detected"),
        run_ttest(df,"bot_score_friendly","Chatbot friendly score"),
        run_ttest(df,"bot_score_professional","Chatbot professional score"),
        run_ttest(df,"bot_compliance_score","Chatbot compliance score"),
        run_ttest(df,"bot_coherence_score","Chatbot coherence score"),
    ]
    # End type distribution by tone (chi2)
    end_type_test = None
    if "end_type" in df.columns:
        try:
            ct = pd.crosstab(df["version"], df["end_type"])
            chi2, p, _, _ = chi2_contingency(ct)
            dist = {}
            for ver in ["FL_21","FL_22"]:
                g = df[df["version"]==ver]
                dist[ver] = g["end_type"].value_counts().to_dict()
            end_type_test = {"chi2": round(chi2,4), "p": round(p,4), "sig": sl(p), "distribution": dist}
        except Exception: pass

    # Dropout by tone
    dropout_by_tone = {}
    if "end_type" in df.columns:
        for ver in ["FL_21","FL_22"]:
            g = df[df["version"]==ver]
            n = len(g)
            proper = (g["end_type"]=="proper_end").sum()
            dropout_by_tone[ver] = {
                "n": n,
                "proper_end_pct": round(proper/n*100,1) if n>0 else None,
                "early_dropout_pct": round((g["end_type"]=="early_dropout").sum()/n*100,1) if n>0 else None,
                "minimal_response_pct": round((g["end_type"]=="minimal_response").sum()/n*100,1) if n>0 else None,
            }

    return {
        "evaluation": [r for r in eval_tests if r],
        "quality_ai": [r for r in quality_tests if r],
        "text_metrics": [r for r in metric_tests if r],
        "actionability": [r for r in action_tests if r],
        "perception_ai": [r for r in psy_tests if r],
        "conversation": [r for r in conv_tests if r],
        "end_type": end_type_test,
        "dropout_by_tone": dropout_by_tone,
    }

# ================================================================
# 4. MEDIATIONS WHERE TONE IS IV (F11)
# ================================================================

def compute_tone_mediations(df):
    results = {}

    # Q2.4 — Ton → E2(engagement) → quality_global
    results["Q2.4_tone_E2_quality"] = {
        "label": "Q2.4 — Tone → E2(Engagement felt) → Quality global",
        "model": run_mediation(df, "tone", "evaluation_2_num", "quality_global"),
    }

    # Q2.5 — Ton → E1(effort) → quality_global
    results["Q2.5_tone_E1_quality"] = {
        "label": "Q2.5 — Tone → E1(Required effort) → Quality global",
        "model": run_mediation(df, "tone", "evaluation_1_num", "quality_global"),
    }

    # Q3.4 — Ton → E3(appreciation) → E6(preference)
    results["Q3.4_tone_E3_E6"] = {
        "label": "Q3.4 — Tone → E3(Chatbot appreciation) → E6(Chatbot preference)",
        "model": run_mediation(df, "tone", "evaluation_3_num", "evaluation_6_num"),
    }

    # Q4.5a — Ton → Competence_1 → E2
    results["Q4.5a_tone_comp1_E2"] = {
        "label": "Q4.5a — Tone → Competence_1(Skills) → E2(Engagement felt)",
        "model": run_mediation(df, "tone", "Competence_1_num", "evaluation_2_num"),
    }

    # Q4.5b — Ton → Competence_2 → E2
    results["Q4.5b_tone_comp2_E2"] = {
        "label": "Q4.5b — Tone → Competence_2(Morality) → E2(Engagement felt)",
        "model": run_mediation(df, "tone", "Competence_2_num", "evaluation_2_num"),
    }
  # Q_emotion_a — Ton → content_emotion → quality (5 subscores + global)
    for q_col, q_lbl in [
        ("quality_global",    "Quality global"),
        ("quality_precision", "Quality precision"),
        ("quality_examples",  "Quality examples"),
        ("quality_relevance", "Quality relevance"),
        ("quality_richness",  "Quality richness"),
    ]:
        key = f"tone_emotion_{q_col}"
        results[key] = {
            "label": f"Tone → Emotion expressed → {q_lbl}",
            "model": run_mediation(df, "tone", "content_emotion", q_col),
        }

    # Q_emotion_b — Ton → content_emotion → psychological + evaluation DVs
    for dv_col, dv_lbl in [
        ("Perceived Manipulati_1_num", "PM1 Freedom threat"),
        ("Perceived Manipulati_2_num", "PM2 Decision override"),
        ("Competence_1_num",           "Comp1 Skills judgment"),
        ("Sense of Independenc_1_num", "Ind1 AI plans & goals"),
        ("evaluation_2_num",           "E2 Engagement felt"),
        ("evaluation_3_num",           "E3 Chatbot appreciation"),
    ]:
        key = f"tone_emotion_{dv_col}"
        results[key] = {
            "label": f"Tone → Emotion expressed → {dv_lbl}",
            "model": run_mediation(df, "tone", "content_emotion", dv_col),
        }
    return results

# ================================================================
# 5. REGRESSIONS WITHOUT TONE (F12)
# ================================================================

def compute_regressions_noton(df):
    results = {}

    # Q2.1 — avg_words → quality
    results["Q2.1_words_quality"] = {
        "label": "Q2.1 — Avg words/msg → Quality global",
        "model": run_regression(df, ["avg_words_per_msg"], "quality_global", ["Avg words/msg"]),
    }
    # Q2.2 — E2 → quality
    results["Q2.2_E2_quality"] = {
        "label": "Q2.2 — E2(Engagement felt) → Quality global",
        "model": run_regression(df, ["evaluation_2_num"], "quality_global", ["E2 Engagement felt"]),
    }
    # Q2.3 — E4 → quality
    results["Q2.3_E4_quality"] = {
        "label": "Q2.3 — E4(Utility) → Quality global",
        "model": run_regression(df, ["evaluation_4_num"], "quality_global", ["E4 Utility"]),
    }
    # Q3.1 — E2+E3+E4 → E5
    results["Q3.1_multi_E5"] = {
        "label": "Q3.1 — E2+E3+E4 → E5(Reuse intention) [multiple regression]",
        "model": run_regression(df, ["evaluation_2_num","evaluation_3_num","evaluation_4_num"],
                                "evaluation_5_num",
                                ["E2 Engagement felt","E3 Appreciation","E4 Utility"]),
    }
    # E1+E2+E3+E4 → E5 simultaneous
    results["Q3.1b_full_multi_E5"] = {
        "label": "E1+E2+E3+E4 → E5(Reuse intention) [full multiple regression]",
        "model": run_regression(df,
                                ["evaluation_1_num","evaluation_2_num","evaluation_3_num","evaluation_4_num"],
                                "evaluation_5_num",
                                ["E1 Effort","E2 Engagement","E3 Appreciation","E4 Utility"]),
    }
    # Q3.2 — quality → E6
    results["Q3.2_quality_E6"] = {
        "label": "Q3.2 — Quality global → E6(Chatbot preference)",
        "model": run_regression(df, ["quality_global"], "evaluation_6_num", ["Quality global"]),
    }
    # E1+E2+quality → E6
    results["Q3.2b_E1_E2_quality_E6"] = {
        "label": "E1+E2+Quality global → E6(Chatbot preference)",
        "model": run_regression(df,
                                ["evaluation_1_num","evaluation_2_num","quality_global"],
                                "evaluation_6_num",
                                ["E1 Effort","E2 Engagement","Quality global"]),
    }
    # Q4.1 — 12 psy vars → E3
    psy_cols = ["Perceived Manipulati_1_num","Perceived Manipulati_2_num",
                "Perceived Manipulati_3_num","Perceived Manipulati_4_num",
                "Competence_1_num","Competence_2_num",
                "Moral Responsibility_1_num","Moral Responsibility_2_num",
                "Moral Responsibility_3_num","Moral Responsibility_4_num",
                "Sense of Independenc_1_num","Sense of Independenc_2_num"]
    psy_labels = ["PM1 Freedom threat","PM2 Decision override","PM3 Manipulation","PM4 Pressure",
                  "Comp1 Skills","Comp2 Morality","MR1 AI harm","MR2 AI resp",
                  "MR3 Human harm","MR4 AI concern","Ind1 Plans","Ind2 Self-ctrl"]
    psy_avail = [(c,l) for c,l in zip(psy_cols,psy_labels) if c in df.columns]
    results["Q4.1_psy_E3"] = {
        "label": "Q4.1 — 12 Perception AI variables → E3(Chatbot appreciation) [multiple regression]",
        "model": run_regression(df, [c for c,_ in psy_avail], "evaluation_3_num", [l for _,l in psy_avail]),
    }
  # content_emotion → quality (5 subscores + global)
    for q_col, q_lbl in [
        ("quality_global",    "Quality global"),
        ("quality_precision", "Quality precision"),
        ("quality_examples",  "Quality examples"),
        ("quality_relevance", "Quality relevance"),
        ("quality_richness",  "Quality richness"),
    ]:
        results[f"emotion_{q_col}"] = {
            "label": f"Content emotion → {q_lbl}",
            "model": run_regression(df, ["content_emotion"], q_col, ["Emotion expressed (0/1)"]),
        }

    # content_emotion → psychological perception variables
    for psy_col, psy_lbl in [
        ("Perceived Manipulati_1_num", "PM1 Freedom threat"),
        ("Perceived Manipulati_2_num", "PM2 Decision override"),
        ("Competence_1_num",           "Comp1 Skills judgment"),
        ("Sense of Independenc_1_num", "Ind1 AI plans & goals"),
        ("evaluation_2_num",           "E2 Engagement felt"),
        ("evaluation_3_num",           "E3 Chatbot appreciation"),
    ]:
        results[f"emotion_{psy_col}"] = {
            "label": f"Content emotion → {psy_lbl}",
            "model": run_regression(df, ["content_emotion"], psy_col, ["Emotion expressed (0/1)"]),
        }
    # Q5.2 — breakpoint_turn → quality
    results["Q5.2_bkpt_quality"] = {
        "label": "Q5.2 — Breakpoint turn → Quality global",
        "model": run_regression(df, ["breakpoint_turn"], "quality_global", ["Breakpoint turn"]),
    }
    # Q5.3 — breakpoint_exists → quality (t-test)
    if "breakpoint_exists" in df.columns:
        g0 = df[df["breakpoint_exists"]==0]["quality_global"].dropna().values
        g1 = df[df["breakpoint_exists"]==1]["quality_global"].dropna().values
        if len(g0)>=2 and len(g1)>=2:
            t_s, p_v = ttest_ind(g0, g1)
            d = cohens_d(g0, g1)
            results["Q5.3_bkpt_exists_quality"] = {
                "label": "Q5.3 — Breakpoint exists (0/1) → Quality global [t-test]",
                "model": {
                    "n_no_breakpoint": len(g0), "mean_no_bkpt": round(float(np.mean(g0)),3),
                    "sd_no_bkpt": round(float(np.std(g0,ddof=1)),3),
                    "n_breakpoint": len(g1), "mean_bkpt": round(float(np.mean(g1)),3),
                    "sd_bkpt": round(float(np.std(g1,ddof=1)),3),
                    "t": round(float(t_s),4), "p": round(float(p_v),4), "sig": sl(p_v),
                    "cohens_d": round(float(d),3) if not np.isnan(d) else None,
                    "effect_size": interpret_d(d),
                },
            }
    # content_emotion → E2(engagement) → quality_global
    results["emotion_E2_quality"] = {
        "label": "Emotion expressed → E2 (Engagement felt) → Quality global",
        "model": run_mediation(df, "content_emotion", "evaluation_2_num", "quality_global"),
    }

    # content_emotion → E3(appreciation) → E6(preference)
    results["emotion_E3_E6"] = {
        "label": "Emotion expressed → E3 (Chatbot appreciation) → E6 (Chatbot preference)",
        "model": run_mediation(df, "content_emotion", "evaluation_3_num", "evaluation_6_num"),
    }
    return results

# ================================================================
# 6. MEDIATIONS WITHOUT TONE (F13)
# ================================================================

def compute_mediations_noton(df):
    results = {}
    psy_cols = ["Perceived Manipulati_1_num","Perceived Manipulati_2_num",
                "Perceived Manipulati_3_num","Perceived Manipulati_4_num",
                "Competence_1_num","Competence_2_num",
                "Moral Responsibility_1_num","Moral Responsibility_2_num",
                "Moral Responsibility_3_num","Moral Responsibility_4_num",
                "Sense of Independenc_1_num","Sense of Independenc_2_num"]
    psy_labels = ["PM1 Freedom threat","PM2 Decision override","PM3 Manipulation","PM4 Pressure",
                  "Comp1 Skills","Comp2 Morality","MR1 AI harm","MR2 AI resp",
                  "MR3 Human harm","MR4 AI concern","Ind1 Plans","Ind2 Self-ctrl"]

    # Q2.6 — 12 psy → E2 → quality_global
    q26 = []
    for col, lbl in zip(psy_cols, psy_labels):
        if col not in df.columns: continue
        r = run_mediation(df, col, "evaluation_2_num", "quality_global")
        q26.append({"iv": col, "iv_label": lbl, "mediator": "E2 Engagement felt",
                    "dv": "quality_global", "dv_label": "Quality global", "result": r})
    results["Q2.6_psy_E2_quality"] = {
        "label": "Q2.6 — 12 Perception AI variables → E2(Engagement) → Quality global",
        "models": q26,
    }

    # Q3.3 — E4 → E3 → E6
    results["Q3.3_E4_E3_E6"] = {
        "label": "Q3.3 — E4(Utility) → E3(Appreciation) → E6(Chatbot preference)",
        "model": run_mediation(df, "evaluation_4_num", "evaluation_3_num", "evaluation_6_num"),
    }

    # Q4.4 — 12 psy → E3 → E6
    q44 = []
    for col, lbl in zip(psy_cols, psy_labels):
        if col not in df.columns: continue
        r = run_mediation(df, col, "evaluation_3_num", "evaluation_6_num")
        q44.append({"iv": col, "iv_label": lbl, "mediator": "E3 Chatbot appreciation",
                    "dv": "evaluation_6_num", "dv_label": "Chatbot preference", "result": r})
    results["Q4.4_psy_E3_E6"] = {
        "label": "Q4.4 — 12 Perception AI variables → E3(Appreciation) → E6(Chatbot preference)",
        "models": q44,
    }

    return results

# ================================================================
# 7. DROPOUT ANALYSIS CORRECTED (F9)
# ================================================================

def compute_dropout_corrected(df):
    if "end_type" not in df.columns:
        return {}

    # Normalise (already done in load_merged but double-check)
    end_map = {"mineral_response":"minimal_response","minimale_response":"minimal_response"}
    df = df.copy()
    df["end_type"] = df["end_type"].replace(end_map)

    categories = ["proper_end","early_dropout","minimal_response"]

    # Distribution
    overall_dist = df["end_type"].value_counts().to_dict()
    tone_dist = {}
    for ver in ["FL_21","FL_22"]:
        g = df[df["version"]==ver]
        tone_dist[ver] = g["end_type"].value_counts().to_dict()

    # Chi2: end_type by tone
    ct = pd.crosstab(df["version"], df["end_type"])
    chi2_val, chi2_p, _, _ = chi2_contingency(ct)

    # E1 and E2 by end_type — for each category: group IN vs OUT
    e1e2_analysis = []
    for cat in categories:
        for eval_col, eval_lbl in [("evaluation_1_num","E1 Required effort"),
                                    ("evaluation_2_num","E2 Engagement felt")]:
            if eval_col not in df.columns: continue
            g_in = df[df["end_type"]==cat][eval_col].dropna().values
            g_out = df[df["end_type"]!=cat][eval_col].dropna().values
            if len(g_in)<2 or len(g_out)<2: continue
            t_s, p_v = ttest_ind(g_in, g_out)
            d = cohens_d(g_in, g_out)
            e1e2_analysis.append({
                "end_type": cat, "variable": eval_col, "label": eval_lbl,
                "n_in": len(g_in), "mean_in": round(float(np.mean(g_in)),3),
                "sd_in": round(float(np.std(g_in,ddof=1)),3),
                "n_out": len(g_out), "mean_out": round(float(np.mean(g_out)),3),
                "sd_out": round(float(np.std(g_out,ddof=1)),3),
                "t": round(float(t_s),4), "p": round(float(p_v),4), "sig": sl(p_v),
                "cohens_d": round(float(d),3) if not np.isnan(d) else None,
                "interpretation": f"{eval_lbl} {'higher' if np.mean(g_in)>np.mean(g_out) else 'lower'} in {cat} group",
            })

    # Dropouts vs completers profile
    df["is_dropout"] = (df["end_type"] != "proper_end").astype(int)
    dropouts = df[df["is_dropout"]==1]
    completers = df[df["is_dropout"]==0]
    profile_vars = [("quality_global","Quality global"),("avg_words_per_msg","Avg words/msg"),
                    ("n_turns","N turns"),("evaluation_1_num","E1 Effort"),
                    ("evaluation_2_num","E2 Engagement felt")]
    profile_comparison = []
    for col, lbl in profile_vars:
        if col not in df.columns: continue
        gd = dropouts[col].dropna().values
        gc = completers[col].dropna().values
        if len(gd)<2 or len(gc)<2: continue
        t_s, p_v = ttest_ind(gd, gc)
        d = cohens_d(gd, gc)
        profile_comparison.append({
            "variable": col, "label": lbl,
            "mean_dropout": round(float(np.mean(gd)),3), "sd_dropout": round(float(np.std(gd,ddof=1)),3),
            "mean_completer": round(float(np.mean(gc)),3), "sd_completer": round(float(np.std(gc,ddof=1)),3),
            "t": round(float(t_s),4), "p": round(float(p_v),4), "sig": sl(p_v),
            "cohens_d": round(float(d),3) if not np.isnan(d) else None, "effect_size": interpret_d(d),
        })

    return {
        "overall_distribution": overall_dist,
        "by_tone": tone_dist,
        "chi2_by_tone": {"chi2":round(chi2_val,4),"p":round(chi2_p,4),"sig":sl(chi2_p)},
        "e1_e2_by_endtype": e1e2_analysis,
        "profile_comparison": profile_comparison,
        "n_total": len(df),
        "n_dropouts": int(df["is_dropout"].sum()),
        "n_completers": int((df["is_dropout"]==0).sum()),
        "pct_dropout": round(df["is_dropout"].mean()*100,1),
        "pct_dropout_fl21": round(df[df["version"]=="FL_21"]["is_dropout"].mean()*100,1),
        "pct_dropout_fl22": round(df[df["version"]=="FL_22"]["is_dropout"].mean()*100,1),
    }

# ================================================================
# RUN ALL
# ================================================================

def run():
    print("Loading merged data...")
    df = load_merged()
    print(f"  {len(df)} participants × {len(df.columns)} columns")

    print("Computing correlations...")
    corr = compute_correlations(df)
    with open(OUTPUT_DIR/"correlations_full.json","w",encoding="utf-8") as f:
        json.dump(corr, f, ensure_ascii=False, indent=2)

    print("Computing VIF...")
    vif = compute_vif_all(df)
    with open(OUTPUT_DIR/"vif_full.json","w",encoding="utf-8") as f:
        json.dump(vif, f, ensure_ascii=False, indent=2)

    print("Computing tone comparisons...")
    tone = compute_tone_comparisons(df)
    with open(OUTPUT_DIR/"tone_comparisons.json","w",encoding="utf-8") as f:
        json.dump(tone, f, ensure_ascii=False, indent=2)

    print("Computing tone mediations...")
    med_tone = compute_tone_mediations(df)
    with open(OUTPUT_DIR/"mediation_tone.json","w",encoding="utf-8") as f:
        json.dump(med_tone, f, ensure_ascii=False, indent=2)

    print("Computing regressions (no tone)...")
    reg = compute_regressions_noton(df)
    with open(OUTPUT_DIR/"regression_noton.json","w",encoding="utf-8") as f:
        json.dump(reg, f, ensure_ascii=False, indent=2)

    print("Computing mediations (no tone)...")
    med_no = compute_mediations_noton(df)
    with open(OUTPUT_DIR/"mediation_noton.json","w",encoding="utf-8") as f:
        json.dump(med_no, f, ensure_ascii=False, indent=2)

    print("Computing dropout analysis...")
    drop = compute_dropout_corrected(df)
    with open(OUTPUT_DIR/"dropout_corrected.json","w",encoding="utf-8") as f:
        json.dump(drop, f, ensure_ascii=False, indent=2)

    print("\nAll analyses complete. Files saved to outputs/")
    return df

if __name__ == "__main__":
    # Load df_ai separately to avoid scope issue
    run()
