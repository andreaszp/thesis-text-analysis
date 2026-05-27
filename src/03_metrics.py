"""
03_metrics.py
Block 2 — Step 1: Computable metrics on participant messages and chatbot responses.
No API call. Saves: outputs/metrics.json, df_msg.json, df_resp.json, df_agg.json
"""
import sys, json, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
from scipy import stats
from scipy.stats import chi2_contingency

import nltk
for pkg in ["punkt","punkt_tab","stopwords","vader_lexicon"]:
    try:
        nltk.data.find(f"tokenizers/{pkg}" if "punkt" in pkg
                       else f"sentiment/{pkg}" if "vader" in pkg
                       else f"corpora/{pkg}")
    except LookupError:
        nltk.download(pkg, quiet=True)

from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.sentiment.vader import SentimentIntensityAnalyzer
from config import MSG_COLS, RESP_COLS, OUTPUT_DIR
from loader import load_clean_data, collect_messages

END_PHRASES = [
    "tu peux cliquer sur la fleche","you can click the arrow",
    "click the arrow at the bottom","<end_of_interview>",
    "avez-vous quelque chose a ajouter","as-tu quelque chose a ajouter",
    "is there anything else","y a-t-il autre chose",
    "anything else to add","anything else you would like to share",
    "do you have anything else","avez-vous autre chose",
]

def detect_proper_end(row):
    n = row["n_turns"]
    for i in range(max(1, n-1), n+2):
        resp = row.get(f"response_{i}")
        if pd.notna(resp):
            txt = str(resp).lower()
            if any(p in txt for p in END_PHRASES):
                return 1
    return 0

def classify_last_msg(row):
    n = row["n_turns"]
    last = str(row.get(f"msg_{n}","") or "").strip()
    nw   = len(last.split())
    if row.get("conv_ended", 0): return "proper_end"
    if n <= 3 and nw <= 5:       return "early_dropout"
    if nw <= 3:                  return "minimal_response"
    return "end_without_closing"

def compute_metrics(records):
    sia  = SentimentIntensityAnalyzer()
    rows = []
    for rec in records:
        txt    = rec["text"]
        tokens = word_tokenize(txt)
        words  = [t for t in tokens if t.isalpha()]
        try:    sents = sent_tokenize(txt)
        except: sents = [txt]
        sc = sia.polarity_scores(txt)
        rows.append({
            "respondent":       rec["respondent"],
            "version":          rec.get("version",""),
            "turn":             rec["turn"],
            "text":             txt,
            "n_words":          len(words),
            "n_chars":          len(txt),
            "n_sentences":      len(sents),
            "avg_word_len":     round(np.mean([len(w) for w in words]),3) if words else 0,
            "type_token_ratio": round(len(set(w.lower() for w in words))/len(words),3) if words else 0,
            "has_question":     int("?" in txt),
            "compound":         round(sc["compound"],3),
            "pos":              round(sc["pos"],3),
            "neg":              round(sc["neg"],3),
            "neu":              round(sc["neu"],3),
        })
    return pd.DataFrame(rows)

def ttest_pair(s21, s22, label):
    g21, g22 = s21.dropna().values, s22.dropna().values
    if len(g21)<2 or len(g22)<2: return None
    t_s, p_v = stats.ttest_ind(g21, g22)
    pool = np.sqrt(((len(g21)-1)*np.std(g21,ddof=1)**2+(len(g22)-1)*np.std(g22,ddof=1)**2)/(len(g21)+len(g22)-2))
    d    = (np.mean(g21)-np.mean(g22))/pool if pool>0 else np.nan
    sig  = "***" if p_v<0.001 else "**" if p_v<0.01 else "*" if p_v<0.05 else "ns"
    return {"label":label,
            "mean_fl21":round(float(np.mean(g21)),3),"sd_fl21":round(float(np.std(g21,ddof=1)),3),
            "mean_fl22":round(float(np.mean(g22)),3),"sd_fl22":round(float(np.std(g22,ddof=1)),3),
            "t":round(float(t_s),4),"p":round(float(p_v),4),"sig":sig,
            "cohens_d":round(float(d),3) if not np.isnan(d) else None}

def run():
    df_clean, _, _ = load_clean_data()
    df_clean["conv_ended"]  = df_clean.apply(detect_proper_end, axis=1)
    df_clean["fin_type"]    = df_clean.apply(classify_last_msg,  axis=1)

    # Turn-level metrics
    msgs_all  = collect_messages(df_clean, MSG_COLS)
    resps_all = collect_messages(df_clean, RESP_COLS)
    df_msg    = compute_metrics(msgs_all)
    df_resp   = compute_metrics(resps_all)

    # Participant-level aggregation
    # NOTE: avg_ttr, avg_sentiment and pct_questions still computed here for
    # descriptive/archival purposes but excluded from main t-test analysis
    df_agg = df_msg.groupby(["respondent","version"]).agg(
        n_messages        =("text","count"),
        total_words       =("n_words","sum"),
        avg_words_per_msg =("n_words","mean"),
        avg_word_len      =("avg_word_len","mean"),
        avg_ttr           =("type_token_ratio","mean"),   # kept for descriptive use only
        pct_questions     =("has_question","mean"),        # kept for descriptive use only
        avg_sentiment     =("compound","mean"),            # kept for descriptive use only
    ).reset_index()

    # ── T-tests on RETAINED metrics only ─────────────────────────────────────
    # Removed from main analysis: avg_ttr (r=-0.788 with words, length artefact)
    # avg_sentiment (r=0.590 with words, length artefact), pct_questions (quasi-zero)
    dm21, dm22 = df_msg[df_msg["version"]=="FL_21"], df_msg[df_msg["version"]=="FL_22"]
    dr21, dr22 = df_resp[df_resp["version"]=="FL_21"], df_resp[df_resp["version"]=="FL_22"]
    am21, am22 = df_agg[df_agg["version"]=="FL_21"], df_agg[df_agg["version"]=="FL_22"]

    # Per-message t-tests: words, chars, sentences, word length (ttr/sentiment REMOVED)
    msg_ttests = [ttest_pair(dm21[c],dm22[c],l) for c,l in [
        ("n_words","Words per message"),
        ("n_chars","Chars per message"),
        ("n_sentences","Sentences per message"),
        ("avg_word_len","Avg word length"),
    ] if ttest_pair(dm21[c],dm22[c],l)]

    resp_ttests = [ttest_pair(dr21[c],dr22[c],l) for c,l in [
        ("n_words","Words per message"),
        ("n_chars","Chars per message"),
        ("n_sentences","Sentences per message"),
        ("avg_word_len","Avg word length"),
    ] if ttest_pair(dr21[c],dr22[c],l)]

    # Per-participant t-tests: words and word_len only (ttr/questions/sentiment REMOVED)
    agg_ttests = [ttest_pair(am21[c],am22[c],l) for c,l in [
        ("avg_words_per_msg","Avg words/msg (per participant)"),
        ("avg_word_len","Avg word length (per participant)"),
    ] if ttest_pair(am21[c],am22[c],l)]

    # Turns stats
    g21_t = df_clean[df_clean["FL_13_DO"]=="FL_21"]["n_turns"].values
    g22_t = df_clean[df_clean["FL_13_DO"]=="FL_22"]["n_turns"].values
    t_s, p_v = stats.ttest_ind(g21_t, g22_t)
    turns_dist = {}
    for ver in ["FL_21","FL_22"]:
        grp = df_clean[df_clean["FL_13_DO"]==ver]
        turns_dist[ver] = {str(k):int(v) for k,v in grp["n_turns"].value_counts().sort_index().items()}

    ct = pd.crosstab(df_clean["FL_13_DO"], df_clean["conv_ended"])
    chi2_v, p_chi, _, _ = chi2_contingency(ct)

    results = {
        "turns": {
            "ttest":       {"label":"Number of turns","mean_fl21":round(float(np.mean(g21_t)),2),
                            "sd_fl21":round(float(np.std(g21_t,ddof=1)),2),
                            "mean_fl22":round(float(np.mean(g22_t)),2),
                            "sd_fl22":round(float(np.std(g22_t,ddof=1)),2),
                            "t":round(float(t_s),4),"p":round(float(p_v),4),
                            "sig":"***" if p_v<0.001 else "**" if p_v<0.01 else "*" if p_v<0.05 else "ns"},
            "chi2_end":    {"chi2":round(chi2_v,4),"p":round(p_chi,4),
                            "sig":"*" if p_chi<0.05 else "ns"},
            "distribution":turns_dist,
            "pct_ended_fl21":round(df_clean[df_clean["FL_13_DO"]=="FL_21"]["conv_ended"].mean()*100,1),
            "pct_ended_fl22":round(df_clean[df_clean["FL_13_DO"]=="FL_22"]["conv_ended"].mean()*100,1),
        },
        "msg_per_message":     [r for r in msg_ttests  if r],
        "resp_per_message":    [r for r in resp_ttests if r],
        "msg_per_participant": [r for r in agg_ttests  if r],
        # Descriptive note for removed variables
        "removed_from_analysis": {
            "avg_ttr":        "Removed: r=-0.788 with avg_words_per_msg (length artefact)",
            "avg_sentiment":  "Removed: r=0.590 with avg_words_per_msg (length artefact); no sig. difference",
            "pct_questions":  "Removed: quasi-zero base rate (1.57% vs 0.45%); no discriminant power",
        }
    }

    with open(OUTPUT_DIR/"metrics.json","w",encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    df_msg.to_json( OUTPUT_DIR/"df_msg.json",  orient="records", force_ascii=False)
    df_resp.to_json(OUTPUT_DIR/"df_resp.json", orient="records", force_ascii=False)
    df_agg.to_json( OUTPUT_DIR/"df_agg.json",  orient="records", force_ascii=False)
    df_clean[["ResponseId","FL_13_DO","n_turns","conv_ended","fin_type"]]\
        .to_json(OUTPUT_DIR/"df_turns.json", orient="records", force_ascii=False)

    print("Metrics complete")
    print(f"  Turns FL21={np.mean(g21_t):.2f}  FL22={np.mean(g22_t):.2f}  p={round(p_v,4)} {results['turns']['ttest']['sig']}")
    return results

if __name__ == "__main__":
    run()
