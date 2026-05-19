"""
04_wordfreq.py
Block 2 — Step 2: Word frequencies, TF-IDF, WordClouds.
Saves: outputs/wordfreq.json and PNG figures.
"""
import sys, json, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from collections import Counter
from sklearn.feature_extraction.text import TfidfVectorizer

import nltk
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from wordcloud import WordCloud

from config import MSG_COLS, RESP_COLS, OUTPUT_DIR, STOPWORDS_CUSTOM, COLOR_FL21, COLOR_FL22
from loader import load_clean_data, collect_messages

try:
    SW_FR = set(stopwords.words("french"))
    SW_EN = set(stopwords.words("english"))
except:
    nltk.download("stopwords", quiet=True)
    SW_FR = set(stopwords.words("french"))
    SW_EN = set(stopwords.words("english"))

ALL_SW = SW_FR | SW_EN | STOPWORDS_CUSTOM

def clean_tokens(text):
    text   = text.lower()
    text   = re.sub(r"[^a-zàâäéèêëîïôùûüœç\s]", " ", text)
    tokens = word_tokenize(text, preserve_line=True)
    return [t for t in tokens if len(t) > 2 and t not in ALL_SW]

def top_words(records, n=40):
    all_tok = []
    for r in records: all_tok.extend(clean_tokens(r["text"]))
    return Counter(all_tok).most_common(n)

def tfidf_distinctive(recs_a, recs_b, n=20):
    docs_a = [" ".join(clean_tokens(r["text"])) for r in recs_a]
    docs_b = [" ".join(clean_tokens(r["text"])) for r in recs_b]
    all_docs = docs_a + docs_b
    labels   = ["A"]*len(docs_a) + ["B"]*len(docs_b)
    try:
        vec = TfidfVectorizer(max_features=500, min_df=2, ngram_range=(1,2))
        mat = vec.fit_transform(all_docs)
    except ValueError:
        return [], []
    feat   = np.array(vec.get_feature_names_out())
    idx_a  = [i for i,l in enumerate(labels) if l=="A"]
    idx_b  = [i for i,l in enumerate(labels) if l=="B"]
    mean_a = np.asarray(mat[idx_a].mean(axis=0)).flatten()
    mean_b = np.asarray(mat[idx_b].mean(axis=0)).flatten()
    top_a  = [(feat[i], round(float(mean_a[i]),4), round(float(mean_a[i]-mean_b[i]),4))
              for i in np.argsort(mean_a-mean_b)[::-1][:n]]
    top_b  = [(feat[i], round(float(mean_b[i]),4), round(float(mean_b[i]-mean_a[i]),4))
              for i in np.argsort(mean_b-mean_a)[::-1][:n]]
    return top_a, top_b

def save_fig(fig, name):
    path = OUTPUT_DIR / name
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    print(f"  Saved {name}")

def plot_top_words(tw_a, tw_b, la, lb, title, fname, n=20):
    fig, axes = plt.subplots(1, 2, figsize=(16, 8))
    fig.suptitle(title, fontsize=13, fontweight="bold")
    for ax, tw, color, label in zip(axes, [tw_a[:n], tw_b[:n]],
                                    [COLOR_FL21, COLOR_FL22], [la, lb]):
        words  = [w for w,_ in tw][::-1]
        counts = [c for _,c in tw][::-1]
        bars   = ax.barh(words, counts, color=f"#{color}", alpha=0.85, edgecolor="white")
        for bar, val in zip(bars, counts):
            ax.text(bar.get_width()+.3, bar.get_y()+bar.get_height()/2,
                    str(val), va="center", ha="left", fontsize=8)
        ax.set_title(label, fontweight="bold", color=f"#{color}")
        ax.set_xlabel("Frequency"); ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    save_fig(fig, fname)

def plot_tfidf(ta, tb, la, lb, title, fname, n=15):
    fig, axes = plt.subplots(1, 2, figsize=(15, 7))
    fig.suptitle(title, fontsize=12, fontweight="bold")
    for ax, top, color, label in zip(axes, [ta[:n], tb[:n]],
                                      [COLOR_FL21, COLOR_FL22], [la, lb]):
        words  = [w for w,_,_ in top][::-1]
        scores = [d for _,_,d in top][::-1]
        ax.barh(words, scores, color=f"#{color}", alpha=0.85, edgecolor="white")
        ax.set_title(label, color=f"#{color}", fontweight="bold")
        ax.set_xlabel("Δ TF-IDF (distinctive score)"); ax.grid(axis="x", alpha=0.3)
    plt.tight_layout()
    save_fig(fig, fname)

def make_wordcloud(records, color, title, fname):
    text = " ".join(" ".join(clean_tokens(r["text"])) for r in records)
    if not text.strip(): return
    wc = WordCloud(width=900, height=450, background_color="white",
                   colormap="Greens" if "21" in fname else "Oranges",
                   max_words=80).generate(text)
    fig, ax = plt.subplots(figsize=(10, 5))
    ax.imshow(wc, interpolation="bilinear"); ax.axis("off")
    ax.set_title(title, fontsize=12, fontweight="bold", color=f"#{color}")
    plt.tight_layout()
    save_fig(fig, fname)

def run():
    df_clean, _, _ = load_clean_data()
    g21 = df_clean[df_clean["FL_13_DO"]=="FL_21"]
    g22 = df_clean[df_clean["FL_13_DO"]=="FL_22"]

    msgs_fl21  = collect_messages(g21, MSG_COLS)
    msgs_fl22  = collect_messages(g22, MSG_COLS)
    resps_fl21 = collect_messages(g21, RESP_COLS)
    resps_fl22 = collect_messages(g22, RESP_COLS)

    tw_m21 = top_words(msgs_fl21);  tw_m22 = top_words(msgs_fl22)
    tw_r21 = top_words(resps_fl21); tw_r22 = top_words(resps_fl22)
    tfi_ma, tfi_mb = tfidf_distinctive(msgs_fl21,  msgs_fl22)
    tfi_ra, tfi_rb = tfidf_distinctive(resps_fl21, resps_fl22)

    plot_top_words(tw_m21,tw_m22,"FL_21 Friendly","FL_22 Pro",
                   "Top words — Participants","fig_topwords_participants.png")
    plot_top_words(tw_r21,tw_r22,"Chatbot FL_21","Chatbot FL_22",
                   "Top words — Chatbot","fig_topwords_chatbot.png")
    plot_tfidf(tfi_ma,tfi_mb,"FL_21 Friendly","FL_22 Pro",
               "Distinctive words TF-IDF — Participants","fig_tfidf_participants.png")
    plot_tfidf(tfi_ra,tfi_rb,"Chatbot FL_21","Chatbot FL_22",
               "Distinctive words TF-IDF — Chatbot","fig_tfidf_chatbot.png")
    for key, recs, color, label in [
        ("fig_wc_part_fl21.png", msgs_fl21,  COLOR_FL21, "Participants FL_21 Friendly"),
        ("fig_wc_part_fl22.png", msgs_fl22,  COLOR_FL22, "Participants FL_22 Pro"),
        ("fig_wc_bot_fl21.png",  resps_fl21, COLOR_FL21, "Chatbot FL_21 Friendly"),
        ("fig_wc_bot_fl22.png",  resps_fl22, COLOR_FL22, "Chatbot FL_22 Pro"),
    ]:
        make_wordcloud(recs, color, label, key)

    results = {
        "top_words_participants_fl21": [(w,int(c)) for w,c in tw_m21],
        "top_words_participants_fl22": [(w,int(c)) for w,c in tw_m22],
        "top_words_chatbot_fl21":      [(w,int(c)) for w,c in tw_r21],
        "top_words_chatbot_fl22":      [(w,int(c)) for w,c in tw_r22],
        "tfidf_participants_fl21":     [(w,s,d) for w,s,d in tfi_ma],
        "tfidf_participants_fl22":     [(w,s,d) for w,s,d in tfi_mb],
        "tfidf_chatbot_fl21":          [(w,s,d) for w,s,d in tfi_ra],
        "tfidf_chatbot_fl22":          [(w,s,d) for w,s,d in tfi_rb],
    }
    with open(OUTPUT_DIR/"wordfreq.json","w",encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)
    print("Word frequencies complete")
    return results

if __name__ == "__main__":
    run()
