bash

cat > /home/claude/thesis_project/src/01_metrics.py << 'PYEOF'
"""
01_metrics.py — Métriques calculables sur les messages (longueur, structure, tours).
Aucun appel API. Résultats stockés dans outputs/metrics.json
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
import re
from scipy import stats
from scipy.stats import chi2_contingency

import nltk
for pkg in ["punkt","punkt_tab","stopwords","vader_lexicon"]:
    try: nltk.data.find(f"tokenizers/{pkg}" if "punkt" in pkg else f"corpora/{pkg}" if pkg=="stopwords" else f"sentiment/{pkg}")
    except LookupError: nltk.download(pkg, quiet=True)

from nltk.tokenize import word_tokenize, sent_tokenize
from nltk.sentiment.vader import SentimentIntensityAnalyzer

from config import MSG_COLS, RESP_COLS, OUTPUT_DIR
from loader import load_clean_data, collect_messages

# ── Patterns fin de conversation ──────────────────────────
END_PATTERNS = [
    r"thank you for (sharing|your|taking)",
    r"merci .{0,40}(partag|retour|votre|ton|contribution)",
    r"if you (ever|have|think)",r"si jamais",
    r"have a (great|wonderful|good|nice)",r"bonne (journ|semaine)",
    r"is there anything else",r"y a-t-il (autre chose|quelque chose d)",
    r"anything else you.{0,20}(like|want) to (share|add)",
    r"your (feedback|input) (is|are) (valuable|appreciated|greatly)",
    r"c.était (vraiment |très )?(enrichissant|précieux)",
]

def detect_end(row):
    n = row["n_turns"]
    for i in range(max(1, n-1), n+2):
        resp = row.get(f"response_{i}")
        if pd.notna(resp):
            txt = str(resp).lower()
            for p in END_PATTERNS:
                if re.search(p, txt, re.IGNORECASE): return True
    return False

def compute_msg_metrics(records):
    rows = []
    sia  = SentimentIntensityAnalyzer()
    for rec in records:
        txt    = rec["text"]
        tokens = word_tokenize(txt)
        words  = [t for t in tokens if t.isalpha()]
        try:    sents = sent_tokenize(txt)
        except: sents = [txt]
        sc = sia.polarity_scores(txt)
        rows.append({
            "respondent":   rec["respondent"],
            "version":      rec.get("version",""),
            "turn":         rec["turn"],
            "text":         txt,
            "n_words":      len(words),
            "n_chars":      len(txt),
            "n_sentences":  len(sents),
            "avg_word_len": round(np.mean([len(w) for w in words]),3) if words else 0,
            "type_token_ratio": round(len(set(w.lower() for w in words))/len(words),3) if words else 0,
            "has_question": int("?" in txt),
            "compound":     round(sc["compound"],3),
            "pos":          round(sc["pos"],3),
            "neg":          round(sc["neg"],3),
            "neu":          round(sc["neu"],3),
        })
    return pd.DataFrame(rows)

def ttest_series(s21, s22, label):
    g21, g22 = s21.dropna().values, s22.dropna().values
    if len(g21)<2 or len(g22)<2:
        return {"label": label, "mean_fl21": float(np.mean(g21)), "mean_fl22": float(np.mean(g22)),
                "t": None, "p": None, "sig": "n/a", "d": None}
    t_s, p_v = stats.ttest_ind(g21, g22)
    sig  = "***" if p_v<0.001 else "**" if p_v<0.01 else "*" if p_v<0.05 else "ns"
    pool = np.sqrt(((len(g21)-1)*np.std(g21,ddof=1)**2+(len(g22)-1)*np.std(g22,ddof=1)**2)/(len(g21)+len(g22)-2))
    d    = (np.mean(g21)-np.mean(g22))/pool if pool>0 else 0
    return {"label": label,
            "mean_fl21": round(float(np.mean(g21)),3), "sd_fl21": round(float(np.std(g21,ddof=1)),3),
            "mean_fl22": round(float(np.mean(g22)),3), "sd_fl22": round(float(np.std(g22,ddof=1)),3),
            "t": round(float(t_s),4), "p": round(float(p_v),4), "sig": sig, "d": round(float(d),3)}

def run():
    df_clean, df_fl21, df_fl22 = load_clean_data()

    # ── Tours de conversation ─────────────────────────────
    df_clean["conv_ended"]   = df_clean.apply(detect_end, axis=1)
    df_clean["avg_msg_len"]  = df_clean.apply(lambda r: np.mean(
        [len(str(r.get(c,"")).split()) for c in MSG_COLS
         if pd.notna(r.get(c)) and str(r.get(c)).strip()]) or 0, axis=1)
    df_clean["engagement_trend"] = df_clean.apply(lambda r: (
        lambda ls: "décroissant" if len(ls)>=3 and ls[-1]<ls[0]*0.5
                   else "croissant" if len(ls)>=3 and ls[-1]>ls[0]*1.5
                   else "stable"
    )([len(str(r.get(c,"")).split()) for c in MSG_COLS
       if pd.notna(r.get(c)) and str(r.get(c)).strip()]), axis=1)

    g21_t = df_clean[df_clean["FL_13_DO"]=="FL_21"]["n_turns"].values
    g22_t = df_clean[df_clean["FL_13_DO"]=="FL_22"]["n_turns"].values
    t_turns = ttest_series(
        df_clean[df_clean["FL_13_DO"]=="FL_21"]["n_turns"],
        df_clean[df_clean["FL_13_DO"]=="FL_22"]["n_turns"],
        "Nombre de tours")

    ct = pd.crosstab(df_clean["FL_13_DO"], df_clean["conv_ended"])
    chi2_v, p_chi, _, _ = chi2_contingency(ct)

    turns_dist = {}
    for ver in ["FL_21","FL_22"]:
        grp = df_clean[df_clean["FL_13_DO"]==ver]
        turns_dist[ver] = grp["n_turns"].value_counts().sort_index().to_dict()

    # ── Métriques messages participants ───────────────────
    msgs_all  = collect_messages(df_clean, MSG_COLS)
    df_msg    = compute_msg_metrics(msgs_all)
    df_msg21  = df_msg[df_msg["version"]=="FL_21"]
    df_msg22  = df_msg[df_msg["version"]=="FL_22"]

    msg_metrics = [
        ("n_words",      "Nb mots / message"),
        ("n_chars",      "Nb caractères / message"),
        ("n_sentences",  "Nb phrases / message"),
        ("avg_word_len", "Longueur moy. des mots"),
        ("type_token_ratio","Richesse lexicale (TTR)"),
        ("has_question", "% messages avec '?'"),
        ("compound",     "Sentiment compound (VADER)"),
    ]
    msg_ttests = [ttest_series(df_msg21[col], df_msg22[col], lbl) for col, lbl in msg_metrics]

    # Agrégation par participant
    agg = df_msg.groupby(["respondent","version"]).agg(
        n_messages=("text","count"),
        total_words=("n_words","sum"),
        avg_words_per_msg=("n_words","mean"),
        avg_word_len=("avg_word_len","mean"),
        avg_ttr=("type_token_ratio","mean"),
        pct_questions=("has_question","mean"),
        avg_sentiment=("compound","mean"),
    ).reset_index()
    agg21 = agg[agg["version"]=="FL_21"]
    agg22 = agg[agg["version"]=="FL_22"]
    agg_ttests = [
        ttest_series(agg21[col], agg22[col], lbl) for col, lbl in [
            ("avg_words_per_msg","Moy mots/message (par part.)"),
            ("total_words","Total mots (par part.)"),
            ("avg_ttr","Richesse lexicale moy."),
            ("pct_questions","% messages avec '?' (par part.)"),
            ("avg_sentiment","Sentiment moy. par part."),
        ]
    ]

    # ── Métriques réponses chatbot ────────────────────────
    resps_all = collect_messages(df_clean, RESP_COLS)
    df_resp   = compute_msg_metrics(resps_all)
    df_resp21 = df_resp[df_resp["version"]=="FL_21"]
    df_resp22 = df_resp[df_resp["version"]=="FL_22"]
    resp_ttests = [ttest_series(df_resp21[col], df_resp22[col], lbl) for col, lbl in msg_metrics]

    # ── Sauvegarde ────────────────────────────────────────
    results = {
        "turns": {
            "ttest":         t_turns,
            "chi2_end_proper": {"chi2": round(chi2_v,4), "p": round(p_chi,4),
                                "sig": "*" if p_chi<0.05 else "ns"},
            "distribution":  {k: {str(kk):int(vv) for kk,vv in v.items()} for k,v in turns_dist.items()},
            "pct_ended_fl21": round(df_clean[df_clean["FL_13_DO"]=="FL_21"]["conv_ended"].mean()*100,1),
            "pct_ended_fl22": round(df_clean[df_clean["FL_13_DO"]=="FL_22"]["conv_ended"].mean()*100,1),
        },
        "msg_per_message":  msg_ttests,
        "msg_per_participant": agg_ttests,
        "bot_per_message":  resp_ttests,
    }

    out_path = OUTPUT_DIR / "metrics.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    # Sauvegarde des DataFrames pour les autres modules
    df_msg.to_json(OUTPUT_DIR / "df_msg.json",  orient="records", force_ascii=False)
    df_resp.to_json(OUTPUT_DIR / "df_resp.json", orient="records", force_ascii=False)
    agg.to_json(OUTPUT_DIR / "df_agg.json",      orient="records", force_ascii=False)
    df_clean[["ResponseId","FL_13_DO","n_turns","conv_ended","engagement_trend"]]\
        .to_json(OUTPUT_DIR / "df_turns.json",   orient="records", force_ascii=False)

    print(f"Métriques sauvegardées -> {out_path}")
    print(f"  Tours   FL21 moy={np.mean(g21_t):.2f}  FL22 moy={np.mean(g22_t):.2f}  p={t_turns['p']} {t_turns['sig']}")
    print(f"  Fin propre FL21={results['turns']['pct_ended_fl21']}%  FL22={results['turns']['pct_ended_fl22']}%  chi2 p={results['turns']['chi2_end_proper']['p']}")
    for r in msg_ttests:
        print(f"  MSG {r['label']:35s}  FL21={r['mean_fl21']:.3f}  FL22={r['mean_fl22']:.3f}  p={r.get('p','?')} {r.get('sig','')}")
    return results

if __name__ == "__main__":
    run()
PYEOF
Sortie

exit code 0
