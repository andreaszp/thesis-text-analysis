"""
08_new_analyses.py  
Computes all analyses for the restructured Excel export.
Run AFTER 01-06 scripts. Saves JSON files to outputs/.
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from scipy.stats import pearsonr, ttest_ind, chi2_contingency, fisher_exact
from scipy import stats
from config import OUTPUT_DIR

# ================================================================
# HELPERS
# ================================================================
def sl(p):
    if p is None or (isinstance(p,float) and np.isnan(p)): return "n/a"
    if p<0.001: return "***"
    if p<0.01:  return "**"
    if p<0.05:  return "*"
    return "ns"

def cohens_d(g1,g2):
    g1,g2=np.array(g1),np.array(g2)
    n1,n2=len(g1),len(g2)
    if n1<2 or n2<2: return np.nan
    pool=np.sqrt(((n1-1)*np.std(g1,ddof=1)**2+(n2-1)*np.std(g2,ddof=1)**2)/(n1+n2-2))
    return (np.mean(g1)-np.mean(g2))/pool if pool>0 else np.nan

def interpret_d(d):
    if pd.isna(d) or d is None: return "n/a"
    a=abs(d)
    if a<0.2: return "negligible"
    if a<0.5: return "small"
    if a<0.8: return "medium"
    return "large"

def ols(X_mat, y_vec):
    Xd=np.column_stack([np.ones(len(y_vec)),X_mat])
    b=np.linalg.lstsq(Xd,y_vec,rcond=None)[0]
    resid=y_vec-Xd@b; dfr=len(y_vec)-Xd.shape[1]; mse=np.sum(resid**2)/dfr
    try: cov=mse*np.linalg.inv(Xd.T@Xd)
    except: cov=mse*np.linalg.pinv(Xd.T@Xd)
    se=np.sqrt(np.diag(cov)); t=b/se
    p=2*(1-stats.t.cdf(np.abs(t),dfr))
    ss_tot=np.sum((y_vec-np.mean(y_vec))**2)
    r2=1-np.sum(resid**2)/ss_tot if ss_tot>0 else 0
    return b,se,t,p,r2

def run_ttest(df,col,label,g1="FL_21",g2="FL_22"):
    a=df[df["version"]==g1][col].dropna().values
    b_=df[df["version"]==g2][col].dropna().values
    if len(a)<2 or len(b_)<2: return None
    t_s,p_v=ttest_ind(a,b_); d=cohens_d(a,b_)
    return {"variable":col,"label":label,
            "n_fl21":len(a),"mean_fl21":round(float(np.mean(a)),3),"sd_fl21":round(float(np.std(a,ddof=1)),3),
            "n_fl22":len(b_),"mean_fl22":round(float(np.mean(b_)),3),"sd_fl22":round(float(np.std(b_,ddof=1)),3),
            "delta":round(float(np.mean(a)-np.mean(b_)),3),
            "t":round(float(t_s),4),"p":round(float(p_v),4),"sig":sl(p_v),
            "cohens_d":round(float(d),3) if not np.isnan(d) else None,"effect_size":interpret_d(d)}

def run_chi2(df,col,label):
    try:
        ct=pd.crosstab(df["version"],df[col])
        if ct.shape[1]<2: return None
        chi2,p_v,_,_=chi2_contingency(ct)
        p21=df[df["version"]=="FL_21"][col].mean(); p22=df[df["version"]=="FL_22"][col].mean()
        return {"variable":col,"label":label,
                "pct_fl21":round(p21*100,1),"pct_fl22":round(p22*100,1),
                "pct_total":round(df[col].mean()*100,1),"delta_pct":round((p21-p22)*100,1),
                "chi2":round(chi2,4),"p":round(p_v,4),"sig":sl(p_v)}
    except: return None

def run_regression(df,iv_cols,dv_col,labels):
    tmp=df[iv_cols+[dv_col]].dropna()
    if len(tmp)<20: return None
    X=tmp[iv_cols].values; y=tmp[dv_col].values
    b,se,t,p,r2=ols(X,y)
    result={"n":len(tmp),"dv":dv_col,"r2":round(r2,3),"predictors":[]}
    for i,lbl in enumerate(labels):
        result["predictors"].append({"variable":iv_cols[i],"label":lbl,
            "b":round(float(b[i+1]),4),"se":round(float(se[i+1]),4),
            "t":round(float(t[i+1]),3),"p":round(float(p[i+1]),4),"sig":sl(p[i+1])})
    return result

def run_mediation(df,iv,m,dv,n_boot=5000,seed=42):
    tmp=df[[iv,m,dv]].dropna(); n=len(tmp)
    if n<20: return {"error":f"n={n} insufficient","n":n}
    X,M,Y=tmp[iv].values,tmp[m].values,tmp[dv].values
    if np.std(X)<1e-10 or np.std(M)<1e-10: return {"error":"zero variance","n":n}
    b_a,se_a,t_a,p_a,_=ols(X.reshape(-1,1),M)
    b_c,se_c,t_c,p_c,_=ols(X.reshape(-1,1),Y)
    b_bc,se_bc,t_bc,p_bc,_=ols(np.column_stack([X,M]),Y)
    a,b,c,cp=b_a[1],b_bc[2],b_c[1],b_bc[1]
    rng=np.random.default_rng(seed); boots=[]
    for _ in range(n_boot):
        idx=rng.integers(0,n,size=n)
        try:
            ba,*_=ols(X[idx].reshape(-1,1),M[idx])
            bbc,*_=ols(np.column_stack([X[idx],M[idx]]),Y[idx])
            boots.append(ba[1]*bbc[2])
        except: continue
    boots=np.array(boots); ci_low,ci_up=np.percentile(boots,2.5),np.percentile(boots,97.5)
    sig_ind=bool(ci_low>0 or ci_up<0)
    a_sig=p_a[1]<0.05; b_sig=p_bc[2]<0.05; cp_sig=p_bc[1]<0.05
    if not a_sig or not b_sig: med_type="No mediation (path a or b ns)"
    elif sig_ind and not cp_sig: med_type="Full mediation"
    elif sig_ind and cp_sig: med_type="Partial mediation"
    else: med_type="Inconsistent"
    def fmt(x): return round(float(x),4) if not np.isnan(x) else None
    return {"n":n,"iv":iv,"mediator":m,"dv":dv,
            "path_a":{"coef":fmt(a),"se":fmt(se_a[1]),"t":fmt(t_a[1]),"p":fmt(p_a[1]),"sig":sl(p_a[1])},
            "path_b":{"coef":fmt(b),"se":fmt(se_bc[2]),"t":fmt(t_bc[2]),"p":fmt(p_bc[2]),"sig":sl(p_bc[2])},
            "path_c":{"coef":fmt(c),"se":fmt(se_c[1]),"t":fmt(t_c[1]),"p":fmt(p_c[1]),"sig":sl(p_c[1])},
            "path_cp":{"coef":fmt(cp),"se":fmt(se_bc[1]),"t":fmt(t_bc[1]),"p":fmt(p_bc[1]),"sig":sl(p_bc[1])},
            "indirect":{"coef":fmt(a*b),"ci_low":fmt(ci_low),"ci_up":fmt(ci_up),
                        "significant":sig_ind,
                        "interpretation":"Significant (CI excludes 0)" if sig_ind else "Not significant (CI includes 0)"},
            "mediation_type":med_type}

# ================================================================
# DATA LOADING
# ================================================================
def load_merged():
    df_ai=pd.read_json(OUTPUT_DIR/"df_ai.json") if (OUTPUT_DIR/"df_ai.json").exists() else pd.DataFrame()
    df_clean=pd.read_json(OUTPUT_DIR/"df_clean.json")
    df_agg=pd.read_json(OUTPUT_DIR/"df_agg.json") if (OUTPUT_DIR/"df_agg.json").exists() else pd.DataFrame()
    from config import NUM_COLS
    df_base=df_clean.copy().rename(columns={"FL_13_DO":"version"})
    if not df_agg.empty:
        df_agg2=df_agg.rename(columns={"respondent":"ResponseId"})
        df_base=df_base.merge(df_agg2,on=["ResponseId","version"],how="left")
    if not df_ai.empty and "respondent_id" in df_ai.columns:
        df_ai2=df_ai.rename(columns={"respondent_id":"ResponseId"})
        ai_keep=[c for c in df_ai2.columns if c not in df_base.columns or c=="ResponseId"]
        df_base=df_base.merge(df_ai2[ai_keep],on="ResponseId",how="left")
    df_base["tone"]=(df_base["version"]=="FL_21").astype(int)
    ai_num=["quality_global","quality_precision","quality_examples","quality_relevance",
            "quality_richness","action_concrete_pb","action_advice","action_use_case",
            "content_emotion","content_suggestion","breakpoint_exists","breakpoint_turn",
            "completed_fully","bot_score_friendly","bot_score_professional","bot_compliance_score",
            "bot_coherence_score","bot_n_friendly_words","bot_n_formal_words",
            "bot_n_encouragements","bot_open_q_pct","bot_emojis","bot_informal_addr",
            "bot_formal_addr","bot_drift","bot_follow_ups"]
    for c in ai_num:
        if c in df_base.columns: df_base[c]=pd.to_numeric(df_base[c],errors="coerce")
    for c in ["avg_words_per_msg","avg_word_len","avg_ttr","pct_questions","avg_sentiment","n_turns"]:
        if c in df_base.columns: df_base[c]=pd.to_numeric(df_base[c],errors="coerce")
    if "end_type" in df_base.columns:
        end_map={"mineral_response":"minimal_response","minimale_response":"minimal_response",
                 "minimal response":"minimal_response"}
        df_base["end_type"]=df_base["end_type"].replace(end_map)
    for nc in NUM_COLS:
        if nc in df_base.columns: df_base[nc]=pd.to_numeric(df_base[nc],errors="coerce")
    return df_base

# ================================================================
# 1. CORRELATIONS
# ================================================================
def compute_correlations(df):
    eval_vars={"evaluation_1_num":"E1 Required effort","evaluation_2_num":"E2 Engagement felt",
               "evaluation_3_num":"E3 Chatbot appreciation","evaluation_4_num":"E4 Conversation utility",
               "evaluation_5_num":"E5 Reuse intention","evaluation_6_num":"E6 Chatbot preference"}
    quality_vars={"quality_global":"Quality global","quality_precision":"Quality precision",
                  "quality_examples":"Quality examples","quality_relevance":"Quality relevance",
                  "quality_richness":"Quality richness"}
    metric_vars={"avg_words_per_msg":"Avg words/msg","avg_word_len":"Avg word length",
                 "n_turns":"N turns","breakpoint_turn":"Breakpoint turn"}
    action_vars={"action_concrete_pb":"Concrete problem","action_advice":"Applicable advice",
                 "action_use_case":"Precise use case","content_emotion":"Emotion expressed"}
    psy_vars={"Perceived Manipulati_1_num":"PM1 Freedom threat","Perceived Manipulati_2_num":"PM2 Decision override",
              "Perceived Manipulati_3_num":"PM3 Manipulation","Perceived Manipulati_4_num":"PM4 Pressure",
              "Competence_1_num":"Comp1 Skills","Competence_2_num":"Comp2 Morality",
              "Moral Responsibility_1_num":"MR1 AI harm","Moral Responsibility_2_num":"MR2 AI resp",
              "Moral Responsibility_3_num":"MR3 Human harm","Moral Responsibility_4_num":"MR4 AI concern",
              "Sense of Independenc_1_num":"Ind1 AI plans","Sense of Independenc_2_num":"Ind2 AI self-ctrl"}
    all_vars={**eval_vars,**quality_vars,**metric_vars,**action_vars,**psy_vars}
    available={k:v for k,v in all_vars.items() if k in df.columns}
    cols=list(available.keys()); labels=list(available.values()); n=len(cols)
    sig_pairs=[]
    for i in range(n):
        for j in range(i+1,n):
            tmp=df[[cols[i],cols[j]]].dropna()
            if len(tmp)<10: continue
            r,p=pearsonr(tmp[cols[i]],tmp[cols[j]])
            if p<0.05:
                sig_pairs.append({"var_x":cols[i],"label_x":labels[i],"var_y":cols[j],"label_y":labels[j],
                    "r":round(float(r),3),"p":round(float(p),4),"sig":sl(p),
                    "strength":"strong" if abs(r)>=0.5 else "moderate" if abs(r)>=0.3 else "weak",
                    "direction":"positive" if r>0 else "negative"})
    sig_pairs.sort(key=lambda x:abs(x["r"]),reverse=True)
    blocs=[("Evaluation (E1–E6)",eval_vars),("Quality AI",quality_vars),
           ("Text Metrics",metric_vars),("Actionability",action_vars),
           ("Perception AI (12 items)",psy_vars)]
    bloc_matrices={}
    for bname,bvars in blocs:
        bv={k:v for k,v in bvars.items() if k in df.columns}
        bc=list(bv.keys()); bl=list(bv.values())
        bdata=[]
        for i in range(len(bc)):
            row=[]
            for j in range(len(bc)):
                tmp=df[[bc[i],bc[j]]].dropna()
                if i==j or len(tmp)<10: row.append({"r":1.0 if i==j else None,"sig":"—"})
                else:
                    r,p=pearsonr(tmp[bc[i]],tmp[bc[j]])
                    row.append({"r":round(float(r),3),"p":round(float(p),4),"sig":sl(p)})
            bdata.append({"var":bc[i],"label":bl[i],"cells":row})
        bloc_matrices[bname]={"cols":bc,"labels":bl,"data":bdata}
    return {"significant_pairs":sig_pairs,"bloc_matrices":bloc_matrices}

# ================================================================
# 2. TONE COMPARISONS
# ================================================================
def compute_tone_comparisons(df):
    eval_tests=[run_ttest(df,c,l) for c,l in [
        ("evaluation_1_num","E1 Required effort"),("evaluation_2_num","E2 Engagement felt"),
        ("evaluation_3_num","E3 Chatbot appreciation"),("evaluation_4_num","E4 Conversation utility"),
        ("evaluation_5_num","E5 Reuse intention"),("evaluation_6_num","E6 Chatbot preference")]]
    quality_tests=[run_ttest(df,c,l) for c,l in [
        ("quality_global","Quality global (1-5)"),("quality_precision","Quality precision (1-5)"),
        ("quality_examples","Quality examples (1-5)"),("quality_relevance","Quality relevance (1-5)"),
        ("quality_richness","Quality richness (1-5)")]]
    metric_tests=[run_ttest(df,c,l) for c,l in [
        ("avg_words_per_msg","Avg words per message"),("avg_word_len","Avg word length"),
        ("n_turns","Number of turns"),("breakpoint_turn","Breakpoint turn")]]
    action_tests=[run_chi2(df,c,l) for c,l in [
        ("action_concrete_pb","Contains concrete problem"),("action_advice","Contains applicable advice"),
        ("action_use_case","Contains precise use case"),("content_emotion","Emotion / frustration expressed")]]
    psy_tests=[run_ttest(df,c,l) for c,l in [
        ("Perceived Manipulati_1_num","PM1 Freedom threat"),("Perceived Manipulati_2_num","PM2 Decision override"),
        ("Perceived Manipulati_3_num","PM3 Manipulation"),("Perceived Manipulati_4_num","PM4 Pressure"),
        ("Competence_1_num","Comp1 Skills judgment capability"),("Competence_2_num","Comp2 Moral judgment capability"),
        ("Moral Responsibility_1_num","MR1 Moral harm AI→human"),("Moral Responsibility_2_num","MR2 AI moral responsibility"),
        ("Moral Responsibility_3_num","MR3 Moral harm human→AI"),("Moral Responsibility_4_num","MR4 AI moral concern"),
        ("Sense of Independenc_1_num","Ind1 AI plans & goals"),("Sense of Independenc_2_num","Ind2 AI self-control")]]
    conv_tests=[run_ttest(df,c,l) for c,l in [
        ("n_turns","Number of turns"),("breakpoint_turn","Breakpoint turn"),
        ("bot_score_friendly","Chatbot friendly score"),("bot_score_professional","Chatbot professional score"),
        ("bot_compliance_score","Chatbot compliance score"),("bot_coherence_score","Chatbot coherence score")]]

    # Construct composites
    constructs={"Perceived Manipulation":["Perceived Manipulati_1_num","Perceived Manipulati_2_num","Perceived Manipulati_3_num","Perceived Manipulati_4_num"],
                "AI Competence":["Competence_1_num","Competence_2_num"],
                "Moral Responsibility":["Moral Responsibility_1_num","Moral Responsibility_2_num","Moral Responsibility_3_num","Moral Responsibility_4_num"],
                "AI Independence":["Sense of Independenc_1_num","Sense of Independenc_2_num"],
                "General Evaluation":["evaluation_1_num","evaluation_2_num","evaluation_3_num","evaluation_4_num","evaluation_5_num","evaluation_6_num"]}
    construct_tests=[]
    for cname,cols in constructs.items():
        avail=[c for c in cols if c in df.columns]
        if not avail: continue
        df_t=df.copy(); df_t["_comp"]=df_t[avail].mean(axis=1)
        r=run_ttest(df_t,"_comp",f"{cname} (composite)")
        if r: r["variable"]=cname; r["n_items"]=len(avail); construct_tests.append(r)

    # End type detailed table: per category chi2 (binary: in vs not in)
    end_type_detail=[]
    if "end_type" in df.columns:
        for cat in ["proper_end","early_dropout","minimal_response"]:
            df["_cat_bin"]=(df["end_type"]==cat).astype(int)
            n_fl21_in=int((df[df["version"]=="FL_21"]["end_type"]==cat).sum())
            n_fl21_tot=int((df["version"]=="FL_21").sum())
            n_fl22_in=int((df[df["version"]=="FL_22"]["end_type"]==cat).sum())
            n_fl22_tot=int((df["version"]=="FL_22").sum())
            ct=np.array([[n_fl21_in,n_fl21_tot-n_fl21_in],[n_fl22_in,n_fl22_tot-n_fl22_in]])
            try:
                chi2,p,_,_=chi2_contingency(ct)
                # Fisher for small cells
                if ct.min()<5:
                    _,p=fisher_exact(ct)
                    chi2=None
            except: chi2,p=None,None
            end_type_detail.append({
                "category":cat,
                "n_fl21":n_fl21_in,"pct_fl21":round(n_fl21_in/n_fl21_tot*100,1) if n_fl21_tot else None,
                "n_fl22":n_fl22_in,"pct_fl22":round(n_fl22_in/n_fl22_tot*100,1) if n_fl22_tot else None,
                "n_total":n_fl21_in+n_fl22_in,
                "pct_total":round((n_fl21_in+n_fl22_in)/(n_fl21_tot+n_fl22_tot)*100,1),
                "chi2":round(float(chi2),4) if chi2 else None,
                "p":round(float(p),4) if p else None,"sig":sl(p) if p else "n/a",
                "test":"chi2" if chi2 else "Fisher exact"})
        df.drop(columns=["_cat_bin"],inplace=True,errors="ignore")

    # Global end_type chi2
    end_type_global=None
    if "end_type" in df.columns:
        try:
            ct=pd.crosstab(df["version"],df["end_type"])
            chi2_g,p_g,_,_=chi2_contingency(ct)
            end_type_global={"chi2":round(chi2_g,4),"p":round(p_g,4),"sig":sl(p_g),
                             "distribution":{ver:df[df["version"]==ver]["end_type"].value_counts().to_dict()
                                             for ver in ["FL_21","FL_22"]}}
        except: pass

    return {"evaluation":[r for r in eval_tests if r],
            "quality_ai":[r for r in quality_tests if r],
            "text_metrics":[r for r in metric_tests if r],
            "actionability":[r for r in action_tests if r],
            "perception_ai":[r for r in psy_tests if r],
            "conversation":[r for r in conv_tests if r],
            "construct_composites":construct_tests,
            "end_type_detail":end_type_detail,
            "end_type":end_type_global}

# ================================================================
# 3. TONE MEDIATIONS
# ================================================================
def compute_tone_mediations(df):
    results={}
    # Q1.4 — Tone → Competence → Psych DVs (8 models)
    for dv_col,dv_lbl in [("Sense of Independenc_1_num","Ind1 AI plans & goals"),
                           ("Sense of Independenc_2_num","Ind2 AI self-control"),
                           ("Perceived Manipulati_1_num","PM1 Freedom threat"),
                           ("Perceived Manipulati_2_num","PM2 Decision override")]:
        for comp_col,comp_lbl in [("Competence_1_num","Competence_1 (Skills)"),
                                   ("Competence_2_num","Competence_2 (Morality)")]:
            key=f"tone_{comp_col}_{dv_col}"
            results[key]={"label":f"Tone → {comp_lbl} → {dv_lbl}",
                          "bloc":"Q1.4 — Tone → Competence → Perception AI",
                          "model":run_mediation(df,"tone",comp_col,dv_col)}
    # Q2.4/Q2.5
    results["Q2.4_tone_E2_quality"]={"label":"Tone → E2 (Engagement felt) → Quality global",
        "bloc":"Tone → Engagement/Effort → Quality","model":run_mediation(df,"tone","evaluation_2_num","quality_global")}
    results["Q2.5_tone_E1_quality"]={"label":"Tone → E1 (Required effort) → Quality global",
        "bloc":"Tone → Engagement/Effort → Quality","model":run_mediation(df,"tone","evaluation_1_num","quality_global")}
    # Q3.4
    results["Q3.4_tone_E3_E6"]={"label":"Tone → E3 (Chatbot appreciation) → E6 (Chatbot preference)",
        "bloc":"Tone → Appreciation → Preference","model":run_mediation(df,"tone","evaluation_3_num","evaluation_6_num")}
    # Q4.5
    results["Q4.5a_tone_comp1_E2"]={"label":"Tone → Competence_1 (Skills) → E2 (Engagement felt)",
        "bloc":"Tone → Competence → Engagement","model":run_mediation(df,"tone","Competence_1_num","evaluation_2_num")}
    results["Q4.5b_tone_comp2_E2"]={"label":"Tone → Competence_2 (Morality) → E2 (Engagement felt)",
        "bloc":"Tone → Competence → Engagement","model":run_mediation(df,"tone","Competence_2_num","evaluation_2_num")}
    # Tone → emotion → quality (6)
    for q_col,q_lbl in [("quality_global","Quality global"),("quality_precision","Quality precision"),
                         ("quality_examples","Quality examples"),("quality_relevance","Quality relevance"),
                         ("quality_richness","Quality richness")]:
        results[f"tone_emotion_{q_col}"]={"label":f"Tone → Emotion expressed → {q_lbl}",
            "bloc":"Tone → Emotion → Quality","model":run_mediation(df,"tone","content_emotion",q_col)}
    # Tone → emotion → psych/eval DVs
    for dv_col,dv_lbl in [("Perceived Manipulati_1_num","PM1 Freedom threat"),
                           ("Perceived Manipulati_2_num","PM2 Decision override"),
                           ("Competence_1_num","Comp1 Skills judgment"),
                           ("Sense of Independenc_1_num","Ind1 AI plans & goals"),
                           ("evaluation_2_num","E2 Engagement felt"),
                           ("evaluation_3_num","E3 Chatbot appreciation")]:
        results[f"tone_emotion_{dv_col}"]={"label":f"Tone → Emotion expressed → {dv_lbl}",
            "bloc":"Tone → Emotion → Perception/Evaluation","model":run_mediation(df,"tone","content_emotion",dv_col)}
    return results

# ================================================================
# 4. REGRESSIONS WITHOUT TONE
# ================================================================
def compute_regressions_noton(df):
    results={}
    # Quality predictors
    for col,lbl,key in [("avg_words_per_msg","Avg words/msg","Q2.1_words_quality"),
                         ("evaluation_2_num","E2 Engagement felt","Q2.2_E2_quality"),
                         ("evaluation_4_num","E4 Utility","Q2.3_E4_quality")]:
        results[key]={"label":f"{lbl} → Quality global","model":run_regression(df,[col],"quality_global",[lbl])}
    # Emotion → quality
    for q_col,q_lbl in [("quality_global","Quality global"),("quality_precision","Quality precision"),
                         ("quality_examples","Quality examples"),("quality_relevance","Quality relevance"),
                         ("quality_richness","Quality richness")]:
        results[f"emotion_{q_col}"]={"label":f"Emotion expressed → {q_lbl}",
            "model":run_regression(df,["content_emotion"],q_col,["Emotion expressed (0/1)"])}
    # Emotion → psy/eval
    for psy_col,psy_lbl in [("Perceived Manipulati_1_num","PM1 Freedom threat"),
                              ("Perceived Manipulati_2_num","PM2 Decision override"),
                              ("Competence_1_num","Comp1 Skills"),
                              ("Sense of Independenc_1_num","Ind1 AI plans"),
                              ("evaluation_2_num","E2 Engagement felt"),
                              ("evaluation_3_num","E3 Chatbot appreciation")]:
        results[f"emotion_{psy_col}"]={"label":f"Emotion expressed → {psy_lbl}",
            "model":run_regression(df,["content_emotion"],psy_col,["Emotion expressed (0/1)"])}
    # Reuse / preference
    results["Q3.1_multi_E5"]={"label":"E2 + E3 + E4 → E5 (Reuse intention)",
        "model":run_regression(df,["evaluation_2_num","evaluation_3_num","evaluation_4_num"],
                               "evaluation_5_num",["E2 Engagement","E3 Appreciation","E4 Utility"])}
    results["Q3.1b_full_E5"]={"label":"E1 + E2 + E3 + E4 → E5 (Reuse intention)",
        "model":run_regression(df,["evaluation_1_num","evaluation_2_num","evaluation_3_num","evaluation_4_num"],
                               "evaluation_5_num",["E1 Effort","E2 Engagement","E3 Appreciation","E4 Utility"])}
    results["Q3.2_quality_E6"]={"label":"Quality global → E6 (Chatbot preference)",
        "model":run_regression(df,["quality_global"],"evaluation_6_num",["Quality global"])}
    results["Q3.2b_E1E2quality_E6"]={"label":"E1 + E2 + Quality → E6 (Chatbot preference)",
        "model":run_regression(df,["evaluation_1_num","evaluation_2_num","quality_global"],
                               "evaluation_6_num",["E1 Effort","E2 Engagement","Quality global"])}
    # Perception AI → E3
    psy_cols=["Perceived Manipulati_1_num","Perceived Manipulati_2_num","Perceived Manipulati_3_num","Perceived Manipulati_4_num",
              "Competence_1_num","Competence_2_num","Moral Responsibility_1_num","Moral Responsibility_2_num",
              "Moral Responsibility_3_num","Moral Responsibility_4_num","Sense of Independenc_1_num","Sense of Independenc_2_num"]
    psy_labels=["PM1 Freedom threat","PM2 Decision override","PM3 Manipulation","PM4 Pressure",
                "Comp1 Skills","Comp2 Morality","MR1 AI harm","MR2 AI resp",
                "MR3 Human harm","MR4 AI concern","Ind1 Plans","Ind2 Self-ctrl"]
    psy_avail=[(c,l) for c,l in zip(psy_cols,psy_labels) if c in df.columns]
    results["Q4.1_psy_E3"]={"label":"12 Perception AI variables → E3 (Chatbot appreciation)",
        "model":run_regression(df,[c for c,_ in psy_avail],"evaluation_3_num",[l for _,l in psy_avail])}
    # Breakpoint
    results["Q5.2_bkpt_quality"]={"label":"Breakpoint turn → Quality global",
        "model":run_regression(df,["breakpoint_turn"],"quality_global",["Breakpoint turn"])}
    # Q5.3 t-test
    if "breakpoint_exists" in df.columns:
        g0=df[df["breakpoint_exists"]==0]["quality_global"].dropna().values
        g1_=df[df["breakpoint_exists"]==1]["quality_global"].dropna().values
        if len(g0)>=2 and len(g1_)>=2:
            t_s,p_v=ttest_ind(g0,g1_); d=cohens_d(g0,g1_)
            results["Q5.3_bkpt_exists_quality"]={"label":"Breakpoint exists (0/1) → Quality global [t-test]",
                "model":{"n_no_breakpoint":len(g0),"mean_no_bkpt":round(float(np.mean(g0)),3),
                         "sd_no_bkpt":round(float(np.std(g0,ddof=1)),3),
                         "n_breakpoint":len(g1_),"mean_bkpt":round(float(np.mean(g1_)),3),
                         "sd_bkpt":round(float(np.std(g1_,ddof=1)),3),
                         "t":round(float(t_s),4),"p":round(float(p_v),4),"sig":sl(p_v),
                         "cohens_d":round(float(d),3) if not np.isnan(d) else None,
                         "effect_size":interpret_d(d)}}
    return results

# ================================================================
# 5. MEDIATIONS WITHOUT TONE
# ================================================================
def compute_mediations_noton(df):
    results={}
    psy_cols=["Perceived Manipulati_1_num","Perceived Manipulati_2_num","Perceived Manipulati_3_num","Perceived Manipulati_4_num",
              "Competence_1_num","Competence_2_num","Moral Responsibility_1_num","Moral Responsibility_2_num",
              "Moral Responsibility_3_num","Moral Responsibility_4_num","Sense of Independenc_1_num","Sense of Independenc_2_num"]
    psy_labels=["PM1 Freedom threat","PM2 Decision override","PM3 Manipulation","PM4 Pressure",
                "Comp1 Skills","Comp2 Morality","MR1 AI harm","MR2 AI resp",
                "MR3 Human harm","MR4 AI concern","Ind1 Plans","Ind2 Self-ctrl"]
    # Q2.6 — 12 psy → E2 → quality
    q26=[]
    for col,lbl in zip(psy_cols,psy_labels):
        if col not in df.columns: continue
        q26.append({"iv":col,"iv_label":lbl,"mediator":"E2 Engagement felt",
                    "dv":"quality_global","dv_label":"Quality global",
                    "result":run_mediation(df,col,"evaluation_2_num","quality_global")})
    results["Q2.6_psy_E2_quality"]={"label":"12 Perception AI → E2 (Engagement) → Quality global","models":q26}
    # Q3.3
    results["Q3.3_E4_E3_E6"]={"label":"E4 (Utility) → E3 (Appreciation) → E6 (Chatbot preference)",
        "model":run_mediation(df,"evaluation_4_num","evaluation_3_num","evaluation_6_num")}
    # Q4.4 — 12 psy → E3 → E6
    q44=[]
    for col,lbl in zip(psy_cols,psy_labels):
        if col not in df.columns: continue
        q44.append({"iv":col,"iv_label":lbl,"mediator":"E3 Chatbot appreciation",
                    "dv":"evaluation_6_num","dv_label":"Chatbot preference",
                    "result":run_mediation(df,col,"evaluation_3_num","evaluation_6_num")})
    results["Q4.4_psy_E3_E6"]={"label":"12 Perception AI → E3 (Appreciation) → E6 (Chatbot preference)","models":q44}
    # Emotion mediations
    results["emotion_E2_quality"]={"label":"Emotion expressed → E2 (Engagement felt) → Quality global",
        "model":run_mediation(df,"content_emotion","evaluation_2_num","quality_global")}
    results["emotion_E3_E6"]={"label":"Emotion expressed → E3 (Chatbot appreciation) → E6 (Chatbot preference)",
        "model":run_mediation(df,"content_emotion","evaluation_3_num","evaluation_6_num")}
    return results

# ================================================================
# 6. DROPOUT ANALYSIS
# ================================================================
def compute_dropout_corrected(df):
    if "end_type" not in df.columns: return {}
    df=df.copy()
    end_map={"mineral_response":"minimal_response","minimale_response":"minimal_response"}
    df["end_type"]=df["end_type"].replace(end_map)
    overall_dist=df["end_type"].value_counts().to_dict()
    by_tone={ver:df[df["version"]==ver]["end_type"].value_counts().to_dict() for ver in ["FL_21","FL_22"]}
    try:
        ct=pd.crosstab(df["version"],df["end_type"])
        chi2_v,chi2_p,_,_=chi2_contingency(ct)
        chi2_res={"chi2":round(chi2_v,4),"p":round(chi2_p,4),"sig":sl(chi2_p)}
    except: chi2_res={}
    # E1/E2 by end_type
    e1e2=[]
    for cat in ["proper_end","early_dropout","minimal_response"]:
        for ecol,elbl in [("evaluation_1_num","E1 Required effort"),("evaluation_2_num","E2 Engagement felt")]:
            if ecol not in df.columns: continue
            g_in=df[df["end_type"]==cat][ecol].dropna().values
            g_out=df[df["end_type"]!=cat][ecol].dropna().values
            if len(g_in)<2 or len(g_out)<2: continue
            t_s,p_v=ttest_ind(g_in,g_out); d=cohens_d(g_in,g_out)
            e1e2.append({"end_type":cat,"variable":ecol,"label":elbl,
                "n_in":len(g_in),"mean_in":round(float(np.mean(g_in)),3),"sd_in":round(float(np.std(g_in,ddof=1)),3),
                "n_out":len(g_out),"mean_out":round(float(np.mean(g_out)),3),"sd_out":round(float(np.std(g_out,ddof=1)),3),
                "t":round(float(t_s),4),"p":round(float(p_v),4),"sig":sl(p_v),
                "cohens_d":round(float(d),3) if not np.isnan(d) else None,
                "interpretation":f"{elbl} {'higher' if np.mean(g_in)>np.mean(g_out) else 'lower'} in {cat} group"})
    # Profile comparison
    df["is_dropout"]=(df["end_type"]!="proper_end").astype(int)
    dropouts=df[df["is_dropout"]==1]; completers=df[df["is_dropout"]==0]
    profile=[]
    for col,lbl in [("quality_global","Quality global"),("avg_words_per_msg","Avg words/msg"),
                    ("n_turns","N turns"),("evaluation_1_num","E1 Effort"),("evaluation_2_num","E2 Engagement")]:
        if col not in df.columns: continue
        gd=dropouts[col].dropna().values; gc=completers[col].dropna().values
        if len(gd)<2 or len(gc)<2: continue
        t_s,p_v=ttest_ind(gd,gc); d=cohens_d(gd,gc)
        profile.append({"variable":col,"label":lbl,
            "mean_dropout":round(float(np.mean(gd)),3),"sd_dropout":round(float(np.std(gd,ddof=1)),3),
            "mean_completer":round(float(np.mean(gc)),3),"sd_completer":round(float(np.std(gc,ddof=1)),3),
            "t":round(float(t_s),4),"p":round(float(p_v),4),"sig":sl(p_v),
            "cohens_d":round(float(d),3) if not np.isnan(d) else None,"effect_size":interpret_d(d)})
    return {"overall_distribution":overall_dist,"by_tone":by_tone,
            "chi2_by_tone":chi2_res,"e1_e2_by_endtype":e1e2,"profile_comparison":profile,
            "n_total":len(df),"n_dropouts":int(df["is_dropout"].sum()),
            "n_completers":int((df["is_dropout"]==0).sum()),
            "pct_dropout":round(df["is_dropout"].mean()*100,1),
            "pct_dropout_fl21":round(df[df["version"]=="FL_21"]["is_dropout"].mean()*100,1),
            "pct_dropout_fl22":round(df[df["version"]=="FL_22"]["is_dropout"].mean()*100,1)}

# ================================================================
# RUN
# ================================================================
def run():
    print("Loading merged data...")
    df=load_merged()
    print(f"  {len(df)} participants × {len(df.columns)} columns")
    print("Computing correlations..."); corr=compute_correlations(df)
    with open(OUTPUT_DIR/"correlations_full.json","w",encoding="utf-8") as f: json.dump(corr,f,ensure_ascii=False,indent=2)
    print("Computing tone comparisons..."); tone=compute_tone_comparisons(df)
    with open(OUTPUT_DIR/"tone_comparisons.json","w",encoding="utf-8") as f: json.dump(tone,f,ensure_ascii=False,indent=2)
    print("Computing tone mediations..."); med_tone=compute_tone_mediations(df)
    with open(OUTPUT_DIR/"mediation_tone.json","w",encoding="utf-8") as f: json.dump(med_tone,f,ensure_ascii=False,indent=2)
    print("Computing regressions (no tone)..."); reg=compute_regressions_noton(df)
    with open(OUTPUT_DIR/"regression_noton.json","w",encoding="utf-8") as f: json.dump(reg,f,ensure_ascii=False,indent=2)
    print("Computing mediations (no tone)..."); med_no=compute_mediations_noton(df)
    with open(OUTPUT_DIR/"mediation_noton.json","w",encoding="utf-8") as f: json.dump(med_no,f,ensure_ascii=False,indent=2)
    print("Computing dropout analysis..."); drop=compute_dropout_corrected(df)
    with open(OUTPUT_DIR/"dropout_corrected.json","w",encoding="utf-8") as f: json.dump(drop,f,ensure_ascii=False,indent=2)
    print("\nAll analyses complete.")

if __name__=="__main__":
    run()
