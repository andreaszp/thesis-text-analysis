
"""
02_wordfreq.py — Fréquences de mots, TF-IDF, WordClouds.
Sauvegarde outputs/wordfreq.json et les images PNG.
"""
import sys, json
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import re
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from collections import Counter
from sklearn.feature_extraction.text import TfidfVectorizer

import nltk
for pkg in ["punkt","punkt_tab","stopwords"]:
    try: nltk.data.find(f"tokenizers/{pkg}" if "punkt" in pkg else f"corpora/{pkg}")
    except LookupError: nltk.download(pkg, quiet=True)
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from wordcloud import WordCloud

from config import MSG_COLS, RESP_COLS, OUTPUT_DIR, STOPWORDS_CUSTOM, PALETTE, COLOR_FL21, COLOR_FL22
from loader import load_clean_data, collect_messages

SW_NLTK = set(stopwords.words("french")) | set(stopwords.words("english"))
ALL_SW  = SW_NLTK | STOPWORDS_CUSTOM

def clean_tokens(text):
    text   = text.lower()
    text   = re.sub(r"[^a-zàâäéèêëîïôùûüœç\s]", " ", text)
    tokens = word_tokenize(text, preserve_line=True)
    return [t for t in tokens if len(t) > 2 and t not in ALL_SW]

def get_top_words(records, n=40):
    all_tok = []
    for r in records: all_tok.extend(clean_tokens(r["text"]))
    return Counter(all_tok).most_common(n)

def tfidf_distinctive(recs_a, recs_b, n=20):
    """Mots qui caractérisent chaque groupe vs l'autre (delta TF-IDF)."""
    docs_a = [" ".join(clean_tokens(r["text"])) for r in recs_a]
    docs_b = [" ".join(clean_tokens(r["text"])) for r in recs_b]
    all_docs = docs_a + docs_b
    labels   = ["A"]*len(docs_a) + ["B"]*len(docs_b)
    vec = TfidfVectorizer(max_features=500, min_df=2, ngram_range=(1,2))
    try:
        mat = vec.fit_transform(all_docs)
    except ValueError:
        return [], []
    feat = np.array(vec.get_feature_names_out())
    idx_a = [i for i,l in enumerate(labels) if l=="A"]
    idx_b = [i for i,l in enumerate(labels) if l=="B"]
    mean_a = np.asarray(mat[idx_a].mean(axis=0)).flatten()
    mean_b = np.asarray(mat[idx_b].mean(axis=0)).flatten()
    diff_a = mean_a - mean_b
    diff_b = mean_b - mean_a
    top_a = [(feat[i], round(float(mean_a[i]),4), round(float(diff_a[i]),4))
             for i in np.argsort(diff_a)[::-1][:n]]
    top_b = [(feat[i], round(float(mean_b[i]),4), round(float(diff_b[i]),4))
             for i in np.argsort(diff_b)[::-1][:n]]
    return top_a, top_b

def make_wordcloud_png(records, color_name, title, out_path):
    text = " ".join(" ".join(clean_tokens(r["text"])) for r in records)
    if not text.strip(): return
    wc = WordCloud(
        width=1000, height=500, background_color="white",
        colormap=("Greens" if "21" in title else "Oranges"),
        max_words=80, prefer_horizontal=0.85,
    ).generate(text)
    fig, ax = plt.subplots(figsize=(11, 5))
    ax.imshow(wc, interpolation="bilinear"); ax.axis("off")
    ax.set_title(title, fontsize=13, fontweight="bold",
                 color=COLOR_FL21 if "21" in title else COLOR_FL22)
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  WordCloud -> {out_path}")

def plot_top_words(top_a, top_b, label_a, label_b, title, out_path, top_n=20):
    fig, axes = plt.subplots(1, 2, figsize=(16, 8))
    fig.suptitle(title, fontsize=13, fontweight="bold")
    for ax, top, color, label in zip(axes, [top_a[:top_n], top_b[:top_n]], PALETTE, [label_a, label_b]):
        words  = [w for w,_ in top][::-1]
        counts = [c for _,c in top][::-1]
        bars   = ax.barh(words, counts, color=color, alpha=0.85, edgecolor="white")
        for bar, val in zip(bars, counts):
            ax.text(bar.get_width()+0.3, bar.get_y()+bar.get_height()/2,
                    str(val), va="center", ha="left", fontsize=8)
        ax.set_title(label, fontsize=11, color=color, fontweight="bold")
        ax.set_xlabel("Fréquence"); ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Top mots -> {out_path}")

def plot_tfidf(top_a, top_b, label_a, label_b, title, out_path, top_n=15):
    fig, axes = plt.subplots(1, 2, figsize=(15, 7))
    fig.suptitle(title, fontsize=12, fontweight="bold")
    for ax, top, color, label in zip(axes, [top_a[:top_n], top_b[:top_n]], PALETTE, [label_a, label_b]):
        words  = [w for w,_,_ in top][::-1]
        scores = [d for _,_,d in top][::-1]
        ax.barh(words, scores, color=color, alpha=0.85, edgecolor="white")
        ax.set_title(label, color=color, fontweight="bold")
        ax.set_xlabel("Δ TF-IDF (score distinctif)"); ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    fig.savefig(out_path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  TF-IDF -> {out_path}")

def run():
    df_clean, _, _ = load_clean_data()
    grp21 = df_clean[df_clean["FL_13_DO"]=="FL_21"]
    grp22 = df_clean[df_clean["FL_13_DO"]=="FL_22"]

    msgs_fl21  = collect_messages(grp21, MSG_COLS)
    msgs_fl22  = collect_messages(grp22, MSG_COLS)
    resps_fl21 = collect_messages(grp21, RESP_COLS)
    resps_fl22 = collect_messages(grp22, RESP_COLS)

    # ── Top mots ─────────────────────────────────────────
    top_m21 = get_top_words(msgs_fl21)
    top_m22 = get_top_words(msgs_fl22)
    top_r21 = get_top_words(resps_fl21)
    top_r22 = get_top_words(resps_fl22)

    # ── TF-IDF ───────────────────────────────────────────
    tfidf_ma, tfidf_mb = tfidf_distinctive(msgs_fl21,  msgs_fl22)
    tfidf_ra, tfidf_rb = tfidf_distinctive(resps_fl21, resps_fl22)

    # ── Graphiques ────────────────────────────────────────
    plot_top_words(top_m21, top_m22, "FL_21 Friendly","FL_22 Pro",
                   "Top mots — Participants (après stopwords)",
                   OUTPUT_DIR / "fig_topwords_participants.png")
    plot_top_words(top_r21, top_r22, "Chatbot FL_21","Chatbot FL_22",
                   "Top mots — Chatbot (après stopwords)",
                   OUTPUT_DIR / "fig_topwords_chatbot.png")
    plot_tfidf(tfidf_ma, tfidf_mb, "FL_21 Friendly","FL_22 Pro",
               "Mots distinctifs TF-IDF — Participants",
               OUTPUT_DIR / "fig_tfidf_participants.png")
    plot_tfidf(tfidf_ra, tfidf_rb, "Chatbot FL_21","Chatbot FL_22",
               "Mots distinctifs TF-IDF — Chatbot",
               OUTPUT_DIR / "fig_tfidf_chatbot.png")

    for key, recs, label in [
        ("wc_part_fl21", msgs_fl21,  "Participants FL_21 Friendly"),
        ("wc_part_fl22", msgs_fl22,  "Participants FL_22 Pro"),
        ("wc_bot_fl21",  resps_fl21, "Chatbot FL_21 Friendly"),
        ("wc_bot_fl22",  resps_fl22, "Chatbot FL_22 Pro"),
    ]:
        make_wordcloud_png(recs, key, label, OUTPUT_DIR / f"fig_{key}.png")

    # ── Sauvegarde JSON ───────────────────────────────────
    results = {
        "top_words_participants_fl21": [(w, int(c)) for w,c in top_m21],
        "top_words_participants_fl22": [(w, int(c)) for w,c in top_m22],
        "top_words_chatbot_fl21":      [(w, int(c)) for w,c in top_r21],
        "top_words_chatbot_fl22":      [(w, int(c)) for w,c in top_r22],
        "tfidf_distinctive_participants_fl21": [(w,s,d) for w,s,d in tfidf_ma],
        "tfidf_distinctive_participants_fl22": [(w,s,d) for w,s,d in tfidf_mb],
        "tfidf_distinctive_chatbot_fl21":      [(w,s,d) for w,s,d in tfidf_ra],
        "tfidf_distinctive_chatbot_fl22":      [(w,s,d) for w,s,d in tfidf_rb],
    }
    out_path = OUTPUT_DIR / "wordfreq.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print(f"Fréquences sauvegardées -> {out_path}")
    return results

if __name__ == "__main__":
    run()

exit code 0
