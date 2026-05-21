"""
02_questionnaire_analysis.py
Block 1 — Step 2: Comparative analysis of questionnaire DVs by chatbot tone.
Saves: outputs/q_ttests.json, outputs/q_constructs.json
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from scipy import stats
from scipy.stats import levene
from config import OUTPUT_DIR, ALL_Q_COLS, NUM_COLS, EXCEL_LETTERS, Q_LABELS, CONSTRUCTS, COL_SCALE, LIKERT_7, CAPABLE_7, AMOUNT_7

def cohens_d(g1, g2):
    n1, n2 = len(g1), len(g2)
    if n1 < 2 or n2 < 2: return np.nan
    pool = np.sqrt(((n1-1)*np.std(g1,ddof=1)**2+(n2-1)*np.std(g2,ddof=1)**2)/(n1+n2-2))
    return (np.mean(g1)-np.mean(g2))/pool if pool > 0 else np.nan

def interpret_d(d):
    if pd.isna(d): return "n/a"
    d = abs(d)
    if d < 0.2: return "negligible"
    if d < 0.5: return "small"
    if d < 0.8: return "medium"
    return "large"

def sig_label(p):
    if pd.isna(p): return "n/a"
    if p < 0.001: return "***"
    if p < 0.01:  return "**"
    if p < 0.05:  return "*"
    return "ns"

def run_analysis():
    df_fl21 = pd.read_json(OUTPUT_DIR / "df_fl21.json")
    df_fl22 = pd.read_json(OUTPUT_DIR / "df_fl22.json")

    scale_name = {
        c: ("Likert" if m is LIKERT_7 else "Capable" if m is CAPABLE_7 else "Amount")
        for c, m in COL_SCALE.items()
    }
    construct_map = {}
    for constr, cols in CONSTRUCTS.items():
        for c in cols: construct_map[c] = constr

    # ── T-tests per question ──────────────────────────────
    ttest_rows = []
    for col, nc, el, ql in zip(ALL_Q_COLS, NUM_COLS, EXCEL_LETTERS, Q_LABELS):
        g21 = df_fl21[nc].dropna().values if nc in df_fl21 else np.array([])
        g22 = df_fl22[nc].dropna().values if nc in df_fl22 else np.array([])
        if len(g21) > 1 and len(g22) > 1:
            t_s, p_v   = stats.ttest_ind(g21, g22)
            lev_s, lev_p = levene(g21, g22)
            d          = cohens_d(g21, g22)
        else:
            t_s = p_v = lev_s = lev_p = d = np.nan
        ttest_rows.append({
            "excel_col":   el, "variable": col, "label": ql,
            "scale":       scale_name.get(col,""), "construct": construct_map.get(col,""),
            "n_fl21":      len(g21), "mean_fl21": round(float(np.mean(g21)),3) if len(g21) else None,
            "sd_fl21":     round(float(np.std(g21,ddof=1)),3) if len(g21)>1 else None,
            "n_fl22":      len(g22), "mean_fl22": round(float(np.mean(g22)),3) if len(g22) else None,
            "sd_fl22":     round(float(np.std(g22,ddof=1)),3) if len(g22)>1 else None,
            "delta":       round(float(np.mean(g21)-np.mean(g22)),3) if len(g21) and len(g22) else None,
            "t":           round(float(t_s),4) if not np.isnan(t_s) else None,
            "p":           round(float(p_v),4) if not np.isnan(p_v) else None,
            "sig":         sig_label(p_v),
            "cohens_d":    round(float(d),3) if not np.isnan(d) else None,
            "effect_size": interpret_d(d),
            "levene_p":    round(float(lev_p),4) if not np.isnan(lev_p) else None,
        })
    df_ttests = pd.DataFrame(ttest_rows)

    # ── Construct-level summary ───────────────────────────
    construct_rows = []
    for constr, cols in CONSTRUCTS.items():
        nc_list = [c+"_num" for c in cols]
        g21_all = np.concatenate([df_fl21[nc].dropna().values for nc in nc_list if nc in df_fl21.columns])
        g22_all = np.concatenate([df_fl22[nc].dropna().values for nc in nc_list if nc in df_fl22.columns])
        if len(g21_all)>1 and len(g22_all)>1:
            t_s, p_v = stats.ttest_ind(g21_all, g22_all)
            d        = cohens_d(g21_all, g22_all)
        else:
            t_s = p_v = d = np.nan
        construct_rows.append({
            "construct":  constr, "n_items": len(cols),
            "n_fl21":     len(df_fl21), "mean_fl21": round(float(np.mean(g21_all)),3) if len(g21_all) else None,
            "sd_fl21":    round(float(np.std(g21_all,ddof=1)),3) if len(g21_all)>1 else None,
            "n_fl22":     len(df_fl22), "mean_fl22": round(float(np.mean(g22_all)),3) if len(g22_all) else None,
            "sd_fl22":    round(float(np.std(g22_all,ddof=1)),3) if len(g22_all)>1 else None,
            "delta":      round(float(np.mean(g21_all)-np.mean(g22_all)),3) if len(g21_all) and len(g22_all) else None,
            "t":          round(float(t_s),4) if not np.isnan(t_s) else None,
            "p":          round(float(p_v),4) if not np.isnan(p_v) else None,
            "sig":        sig_label(p_v),
            "cohens_d":   round(float(d),3) if not np.isnan(d) else None,
            "effect_size":interpret_d(d),
        })
    df_constructs = pd.DataFrame(construct_rows)

    # ── Save ──────────────────────────────────────────────
    df_ttests.to_json(   OUTPUT_DIR/"q_ttests.json",    orient="records", force_ascii=False)
    df_constructs.to_json(OUTPUT_DIR/"q_constructs.json",orient="records",force_ascii=False)

    print("Questionnaire analysis complete")
    sig = df_ttests[df_ttests["sig"].isin(["*","**","***"])]
    print(f"  Significant DVs (p<.05): {len(sig)}/{len(df_ttests)}")
    for _, r in sig.iterrows():
        print(f"    {r['label']:40s}  FL21={r['mean_fl21']}  FL22={r['mean_fl22']}  p={r['p']} {r['sig']}  d={r['cohens_d']} ({r['effect_size']})")
    return df_ttests, df_constructs

if __name__ == "__main__":
    run()

"""
Block 1 — Step 3: Simple mediation analyses with bootstrap (Hayes/PROCESS style).
IV  : Chatbot tone (FL_21 = +1, FL_22 = -1) — contrast effect coding
M   : Competence_1_num
DV1 : Sense of Independence 1 (Sense_of_Independenc_1_num)
DV2 : Perceived Manipulation 1 (Perceived_Manipulati_1_num)
DV3 : Perceived Manipulation 2 (Perceived_Manipulati_2_num)

Method: OLS regression + percentile bootstrap (5000 samples, 95% CI)
Saves: outputs/mediation.json
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from scipy import stats

from config import OUTPUT_DIR

# ================================================================
# BOOTSTRAP MEDIATION
# ================================================================

def run_mediation(df, iv_col, m_col, dv_col, n_boot=5000, seed=42, alpha=0.05):
    """
    Simple mediation using OLS + percentile bootstrap.

    Model:
      Path a  : IV  → M   (OLS)
      Path b  : M   → DV  controlling for IV (OLS)
      Path c  : IV  → DV  (total effect, OLS)
      Path c' : IV  → DV  controlling for M (direct effect, OLS)
      Indirect: a * b (bootstrapped CI)

    Returns a dict with all paths + bootstrap CI on indirect effect.
    """
    # Drop rows with any missing value in the three variables
    tmp = df[[iv_col, m_col, dv_col]].dropna()
    n   = len(tmp)
    if n < 20:
        return {"error": f"Not enough observations (n={n})", "n": n}

    X  = tmp[iv_col].values
    M  = tmp[m_col].values
    Y  = tmp[dv_col].values

    def ols(x_mat, y_vec):
        """OLS with intercept. Returns (coefs, se, t, p, r2)."""
        X_d = np.column_stack([np.ones(len(y_vec)), x_mat])
        b   = np.linalg.lstsq(X_d, y_vec, rcond=None)[0]
        y_hat  = X_d @ b
        resid  = y_vec - y_hat
        df_res = len(y_vec) - X_d.shape[1]
        mse    = np.sum(resid**2) / df_res
        cov_b  = mse * np.linalg.inv(X_d.T @ X_d)
        se     = np.sqrt(np.diag(cov_b))
        t_stat = b / se
        p_vals = 2 * (1 - stats.t.cdf(np.abs(t_stat), df=df_res))
        ss_tot = np.sum((y_vec - np.mean(y_vec))**2)
        r2     = 1 - np.sum(resid**2)/ss_tot if ss_tot > 0 else 0
        return b, se, t_stat, p_vals, r2

    # ── Path a : X → M ───────────────────────────────────
    b_a, se_a, t_a, p_a, r2_a = ols(X.reshape(-1,1), M)
    a  = b_a[1]   # coefficient of X on M

    # ── Path c : X → Y (total effect) ────────────────────
    b_c, se_c, t_c, p_c, r2_c = ols(X.reshape(-1,1), Y)
    c  = b_c[1]

    # ── Paths b & c' : X + M → Y ─────────────────────────
    b_bc, se_bc, t_bc, p_bc, r2_bc = ols(np.column_stack([X, M]), Y)
    c_prime = b_bc[1]   # direct effect of X
    b       = b_bc[2]   # path b: M → Y controlling for X

    # ── Indirect effect: a * b ────────────────────────────
    indirect = a * b

    # ── Bootstrap CI on indirect effect ──────────────────
    rng   = np.random.default_rng(seed)
    boots = []
    for _ in range(n_boot):
        idx    = rng.integers(0, n, size=n)
        Xs, Ms, Ys = X[idx], M[idx], Y[idx]
        try:
            b_a_b, *_ = ols(Xs.reshape(-1,1), Ms)
            b_bc_b, *_ = ols(np.column_stack([Xs, Ms]), Ys)
            boots.append(b_a_b[1] * b_bc_b[2])
        except Exception:
            continue

    boots  = np.array(boots)
    ci_low = np.percentile(boots, 100 * alpha/2)
    ci_up  = np.percentile(boots, 100 * (1 - alpha/2))
    # Significance: CI does not include 0
    sig_indirect = (ci_low > 0) or (ci_up < 0)

    def sig_label(p):
        if p < 0.001: return "***"
        if p < 0.01:  return "**"
        if p < 0.05:  return "*"
        return "ns"

    def fmt(x): return round(float(x), 4) if x is not None else None

    return {
        "n":           n,
        "iv":          iv_col,
        "mediator":    m_col,
        "dv":          dv_col,
        "n_bootstrap": n_boot,
        "alpha":       alpha,

        "path_a": {
            "label":      "IV → Mediator",
            "coef":       fmt(a),
            "se":         fmt(se_a[1]),
            "t":          fmt(t_a[1]),
            "p":          fmt(p_a[1]),
            "sig":        sig_label(p_a[1]),
            "r2":         fmt(r2_a),
        },
        "path_b": {
            "label":      "Mediator → DV (controlling IV)",
            "coef":       fmt(b),
            "se":         fmt(se_bc[2]),
            "t":          fmt(t_bc[2]),
            "p":          fmt(p_bc[2]),
            "sig":        sig_label(p_bc[2]),
        },
        "path_c": {
            "label":      "IV → DV (total effect)",
            "coef":       fmt(c),
            "se":         fmt(se_c[1]),
            "t":          fmt(t_c[1]),
            "p":          fmt(p_c[1]),
            "sig":        sig_label(p_c[1]),
            "r2":         fmt(r2_c),
        },
        "path_c_prime": {
            "label":      "IV → DV (direct effect, controlling M)",
            "coef":       fmt(c_prime),
            "se":         fmt(se_bc[1]),
            "t":          fmt(t_bc[1]),
            "p":          fmt(p_bc[1]),
            "sig":        sig_label(p_bc[1]),
            "r2":         fmt(r2_bc),
        },
        "indirect_effect": {
            "label":      "Indirect effect (a × b)",
            "coef":       fmt(indirect),
            "boot_mean":  fmt(np.mean(boots)),
            "boot_se":    fmt(np.std(boots, ddof=1)),
            "ci_lower":   fmt(ci_low),
            "ci_upper":   fmt(ci_up),
            "significant":sig_indirect,
            "interpretation": (
                "Significant mediation (CI excludes 0)"
                if sig_indirect
                else "No significant mediation (CI includes 0)"
            ),
        },
        "mediation_type": _mediation_type(
            p_a[1], p_bc[2], p_bc[1], sig_indirect
        ),
    }


def _mediation_type(p_a, p_b, p_c_prime, sig_indirect):
    """
    Classify mediation type based on path significance.
    Full mediation   : a & b significant, c' not significant
    Partial mediation: a, b, c' all significant
    No mediation     : a or b not significant
    Indirect only    : significant indirect but direct not sig (= full mediation)
    """
    a_sig      = p_a      < 0.05
    b_sig      = p_b      < 0.05
    cprime_sig = p_c_prime < 0.05

    if not a_sig or not b_sig:
        return "No mediation (path a or b not significant)"
    if sig_indirect and not cprime_sig:
        return "Full mediation"
    if sig_indirect and cprime_sig:
        return "Partial mediation"
    return "Inconsistent pattern — interpret with caution"


# ================================================================
# RUN
# ================================================================

def run():
    # Load cleaned data
    df_path = OUTPUT_DIR / "df_clean.json"
    if not df_path.exists():
        print("ERROR: df_clean.json not found. Run 01_questionnaire_cleaning.py first.")
        return {}

    df = pd.read_json(df_path)

    # ── Contrast effect coding: FL_21 = +1, FL_22 = -1 ───
    if "FL_13_DO" not in df.columns:
        print("ERROR: FL_13_DO column not found in df_clean.json")
        return {}

    df["tone_contrast"] = df["FL_13_DO"].map({"FL_21": 1, "FL_22": -1})

    # Check required columns exist
    required = ["tone_contrast",
                "Competence_1_num",
                "Sense_of_Independenc_1_num",
                "Perceived_Manipulati_1_num",
                "Perceived_Manipulati_2_num"]

    # Handle column name variants (Qualtrics sometimes truncates)
    col_map = {}
    for col in df.columns:
        col_map[col.lower().replace(" ","_")] = col

    # Map standard names to actual column names
    name_variants = {
        "Competence_1_num":            ["Competence_1_num", "competence_1_num"],
        "Sense_of_Independenc_1_num":  ["Sense of Independenc_1_num",
                                        "Sense_of_Independenc_1_num",
                                        "sense_of_independenc_1_num"],
        "Perceived_Manipulati_1_num":  ["Perceived Manipulati_1_num",
                                        "Perceived_Manipulati_1_num",
                                        "perceived_manipulati_1_num"],
        "Perceived_Manipulati_2_num":  ["Perceived Manipulati_2_num",
                                        "Perceived_Manipulati_2_num",
                                        "perceived_manipulati_2_num"],
    }

    resolved = {"tone_contrast": "tone_contrast"}
    for standard_name, variants in name_variants.items():
        found = None
        for v in variants:
            if v in df.columns:
                found = v; break
        if found:
            resolved[standard_name] = found
        else:
            print(f"WARNING: Could not find column for '{standard_name}'")
            print(f"  Available columns containing 'num': {[c for c in df.columns if '_num' in c]}")

    # ── Define the 3 mediation analyses ──────────────────
    analyses = [
        {
            "name":      "Mediation 1",
            "dv_label":  "Sense of Independence 1",
            "iv":        resolved["tone_contrast"],
            "mediator":  resolved.get("Competence_1_num", "Competence_1_num"),
            "dv":        resolved.get("Sense_of_Independenc_1_num", "Sense_of_Independenc_1_num"),
        },
        {
            "name":      "Mediation 2",
            "dv_label":  "Perceived Manipulation 1",
            "iv":        resolved["tone_contrast"],
            "mediator":  resolved.get("Competence_1_num", "Competence_1_num"),
            "dv":        resolved.get("Perceived_Manipulati_1_num", "Perceived_Manipulati_1_num"),
        },
        {
            "name":      "Mediation 3",
            "dv_label":  "Perceived Manipulation 2",
            "iv":        resolved["tone_contrast"],
            "mediator":  resolved.get("Competence_1_num", "Competence_1_num"),
            "dv":        resolved.get("Perceived_Manipulati_2_num", "Perceived_Manipulati_2_num"),
        },
    ]

    print("\n" + "="*65)
    print("MEDIATION ANALYSES — Bootstrap (5000 samples, 95% CI)")
    print("IV: Chatbot tone (FL_21=+1, FL_22=-1)")
    print("M:  AI Competence 1")
    print("="*65)

    results = []
    for spec in analyses:
        print(f"\n{spec['name']}: Tone → Competence_1 → {spec['dv_label']}")
        print("-"*55)

        if spec["mediator"] not in df.columns or spec["dv"] not in df.columns:
            print(f"  ERROR: column not found — skipping")
            results.append({"name": spec["name"], "error": "column not found"})
            continue

        res = run_mediation(df, spec["iv"], spec["mediator"], spec["dv"])
        res["name"]     = spec["name"]
        res["dv_label"] = spec["dv_label"]

        if "error" in res:
            print(f"  ERROR: {res['error']}")
        else:
            print(f"  n = {res['n']}")
            print(f"  Path a  (IV → M):        b={res['path_a']['coef']:>7}  "
                  f"SE={res['path_a']['se']:>6}  t={res['path_a']['t']:>7}  "
                  f"p={res['path_a']['p']:>6}  {res['path_a']['sig']}")
            print(f"  Path b  (M → DV|IV):     b={res['path_b']['coef']:>7}  "
                  f"SE={res['path_b']['se']:>6}  t={res['path_b']['t']:>7}  "
                  f"p={res['path_b']['p']:>6}  {res['path_b']['sig']}")
            print(f"  Path c  (total):         b={res['path_c']['coef']:>7}  "
                  f"SE={res['path_c']['se']:>6}  t={res['path_c']['t']:>7}  "
                  f"p={res['path_c']['p']:>6}  {res['path_c']['sig']}")
            print(f"  Path c' (direct):        b={res['path_c_prime']['coef']:>7}  "
                  f"SE={res['path_c_prime']['se']:>6}  t={res['path_c_prime']['t']:>7}  "
                  f"p={res['path_c_prime']['p']:>6}  {res['path_c_prime']['sig']}")
            ie = res["indirect_effect"]
            print(f"  Indirect (a×b):          b={ie['coef']:>7}  "
                  f"95% CI [{ie['ci_lower']:>7}, {ie['ci_upper']:>7}]  "
                  f"{'✓ SIG' if ie['significant'] else '✗ ns'}")
            print(f"  → {res['mediation_type']}")

        results.append(res)

    # Save
    out_path = OUTPUT_DIR / "mediation.json"
    with open(out_path, "w", encoding="utf-8") as f:
        class NumpyEncoder(json.JSONEncoder):
            def default(self, obj):
                import numpy as np
                if isinstance(obj, (np.bool_, np.integer)): return int(obj)
                if isinstance(obj, np.floating): return float(obj)
                if isinstance(obj, np.ndarray): return obj.tolist()
                return super().default(obj)
        json.dump(results, f, ensure_ascii=False, indent=2, cls=NumpyEncoder)
    print(f"\nResults saved → {out_path}")
    
if __name__ == "__main__":
    run_analysis()
    run()
