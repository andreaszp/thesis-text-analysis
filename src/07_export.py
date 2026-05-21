"""
07_export.py
Final Excel export — all sheets, all analyses, English labels.
Run this after all other scripts have completed.
Saves: outputs/Master_Thesis_Full_Analysis.xlsx
"""
import sys, json, io
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage

from config import (OUTPUT_FILE, OUTPUT_DIR, COLOR_FL21, COLOR_FL22,
                    COLOR_HDR, COLOR_ALT, SCORE_COLORS, ALL_Q_COLS, NUM_COLS,
                    EXCEL_LETTERS, Q_LABELS, CONSTRUCTS, COL_SCALE, LIKERT_7, CAPABLE_7)

# ================================================================
# STYLE HELPERS
# ================================================================

THIN = Border(
    left=Side(style="thin",color="BFBFBF"), right=Side(style="thin",color="BFBFBF"),
    top=Side(style="thin",color="BFBFBF"),  bottom=Side(style="thin",color="BFBFBF"),
)

def hdr(c, bg=COLOR_HDR, fc="FFFFFF", sz=10, bold=True):
    c.font      = Font(bold=bold, color=fc, name="Arial", size=sz)
    c.fill      = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = THIN

def dat(c, val, center=True, alt=False, bold=False, fc="000000", fmt=None):
    v = None if (isinstance(val,float) and np.isnan(val)) else val
    if isinstance(v, bool): v = "Yes" if v else "No"
    c.value     = v
    c.font      = Font(name="Arial", size=9, bold=bold, color=fc)
    c.alignment = Alignment(horizontal="center" if center else "left",
                            vertical="center", wrap_text=True)
    c.border    = THIN
    if alt: c.fill = PatternFill("solid", start_color=COLOR_ALT)
    if fmt: c.number_format = fmt

def pval(c, p):
    c.border    = THIN
    c.alignment = Alignment(horizontal="center", vertical="center")
    if p is None or (isinstance(p,float) and np.isnan(p)):
        c.value = "n/a"; c.font = Font(name="Arial",size=9); return
    if   p < 0.001: c.value=f"{p:.4f} ***"; c.fill=PatternFill("solid",start_color="1A6634"); c.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
    elif p < 0.01:  c.value=f"{p:.4f} **";  c.fill=PatternFill("solid",start_color="52BE80"); c.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
    elif p < 0.05:  c.value=f"{p:.4f} *";   c.fill=PatternFill("solid",start_color="F9E79F"); c.font=Font(name="Arial",size=9,bold=True,color="7D6608")
    else:           c.value=f"{p:.4f} ns";  c.fill=PatternFill("solid",start_color="FADBD8"); c.font=Font(name="Arial",size=9,color="922B21")

def delta_cell(c, d):
    c.border=THIN; c.alignment=Alignment(horizontal="center",vertical="center")
    if d is None or (isinstance(d,float) and np.isnan(d)): c.value="n/a"; return
    c.value=round(d,3); c.number_format="0.000"
    if d > 0.05:
        c.fill=PatternFill("solid",start_color="D5F5E3"); c.font=Font(name="Arial",size=9,bold=True,color="1A6634")
    elif d < -0.05:
        c.fill=PatternFill("solid",start_color="FADBD8"); c.font=Font(name="Arial",size=9,bold=True,color="922B21")
    else:
        c.fill=PatternFill("solid",start_color="F2F3F4"); c.font=Font(name="Arial",size=9,color="555555")

def title_row(ws, text, bg, ncols, row=1, sz=12):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    c = ws.cell(row=row,column=1,value=text)
    c.font      = Font(bold=True,color="FFFFFF",name="Arial",size=sz)
    c.fill      = PatternFill("solid",start_color=bg)
    c.alignment = Alignment(horizontal="center",vertical="center")
    ws.row_dimensions[row].height = 26

def write_df(ws, df_in, title, bg, start_row=1, freeze=True):
    title_row(ws, title, bg, len(df_in.columns), row=start_row)
    for j,col in enumerate(df_in.columns,1):
        hdr(ws.cell(row=start_row+1,column=j,value=col), bg="2C3E50", sz=9)
    ws.row_dimensions[start_row+1].height = 28
    for i,(_,row) in enumerate(df_in.iterrows()):
        alt=(i%2==0)
        for j,val in enumerate(row,1):
            v = None if (isinstance(val,float) and np.isnan(val)) else val
            dat(ws.cell(row=start_row+2+i,column=j), v, center=(j>1), alt=alt)
    for j,col in enumerate(df_in.columns,1):
        w = max(12, min(40, len(str(col))+3))
        ws.column_dimensions[get_column_letter(j)].width = w
    if freeze:
        ws.freeze_panes = f"A{start_row+2}"

def write_ttest_table(ws, rows, start_row, bg_hdr="2C3E50"):
    hdrs = ["Variable","Label","N FL_21","Mean FL_21","SD FL_21",
            "N FL_22","Mean FL_22","SD FL_22","Δ (21−22)","t-stat","p-value","Sig.","Cohen's d","Effect size"]
    for j,h in enumerate(hdrs,1):
        hdr(ws.cell(row=start_row,column=j,value=h), bg=bg_hdr, sz=9)
    ws.row_dimensions[start_row].height = 25
    for i,r in enumerate(rows):
        rr=start_row+1+i; alt=(i%2==0)
        dat(ws.cell(rr,1), r.get("variable",""), center=False, alt=alt, bold=True)
        dat(ws.cell(rr,2), r.get("label",""),   center=False, alt=alt)
        dat(ws.cell(rr,3), r.get("n_fl21"),     alt=alt)
        dat(ws.cell(rr,4), r.get("mean_fl21"),  alt=alt, fmt="0.000")
        dat(ws.cell(rr,5), r.get("sd_fl21"),    alt=alt, fmt="0.000")
        dat(ws.cell(rr,6), r.get("n_fl22"),     alt=alt)
        dat(ws.cell(rr,7), r.get("mean_fl22"),  alt=alt, fmt="0.000")
        dat(ws.cell(rr,8), r.get("sd_fl22"),    alt=alt, fmt="0.000")
        delta_cell(ws.cell(rr,9), r.get("delta"))
        dat(ws.cell(rr,10), r.get("t"),         alt=alt, fmt="0.0000")
        pval(ws.cell(rr,11), r.get("p"))
        dat(ws.cell(rr,12), r.get("sig",""),    alt=alt, bold=True)
        dat(ws.cell(rr,13), r.get("cohens_d"),  alt=alt, fmt="0.000")
        dat(ws.cell(rr,14), r.get("effect_size",""), alt=alt)
    col_widths=[20,38,8,11,9,8,11,9,12,10,14,8,11,12]
    for j,w in enumerate(col_widths,1):
        ws.column_dimensions[get_column_letter(j)].width=w
    return start_row+1+len(rows)

def write_chi2_table(ws, rows, start_row, bg_hdr="2C3E50"):
    hdrs = ["Variable","Label","% FL_21","% FL_22","% Total","Δ % (21−22)","Chi2","p-value","Sig."]
    for j,h in enumerate(hdrs,1):
        hdr(ws.cell(row=start_row,column=j,value=h), bg=bg_hdr, sz=9)
    for i,r in enumerate(rows):
        rr=start_row+1+i; alt=(i%2==0)
        dat(ws.cell(rr,1), r.get("variable",""),  center=False, alt=alt, bold=True)
        dat(ws.cell(rr,2), r.get("label",""),     center=False, alt=alt)
        dat(ws.cell(rr,3), r.get("pct_fl21"),     alt=alt, fmt="0.0")
        dat(ws.cell(rr,4), r.get("pct_fl22"),     alt=alt, fmt="0.0")
        dat(ws.cell(rr,5), r.get("pct_total"),    alt=alt, fmt="0.0")
        delta_cell(ws.cell(rr,6), r.get("delta_pct"))
        dat(ws.cell(rr,7), r.get("chi2"),         alt=alt, fmt="0.0000")
        pval(ws.cell(rr,8), r.get("p"))
        dat(ws.cell(rr,9), r.get("sig",""),       alt=alt, bold=True)
    for j,w in enumerate([20,40,10,10,10,14,10,14,8],1):
        ws.column_dimensions[get_column_letter(j)].width=w
    return start_row+1+len(rows)

def insert_png(ws, path, anchor, w_cm=18, h_cm=9):
    if not Path(path).exists(): return
    img=XLImage(str(path))
    img.width=int(w_cm*37.795); img.height=int(h_cm*37.795)
    ws.add_image(img,anchor)

# ================================================================
# PROGRESSION FIGURE
# ================================================================

def make_progression_fig(progression):
    if not progression: return None
    fig, ax = plt.subplots(figsize=(12,5))
    colors = {COLOR_FL21:"FL_21 Friendly", COLOR_FL22:"FL_22 Pro"}
    for ver, col, label in [("FL_21",COLOR_FL21,"FL_21 Friendly"),
                             ("FL_22",COLOR_FL22,"FL_22 Pro")]:
        data = progression.get(ver,{})
        if not data: continue
        turns = sorted(int(t) for t in data.keys())
        means = [data[str(t)]["mean"] for t in turns]
        sds   = [data[str(t)]["sd"] or 0 for t in turns]
        ax.plot(turns, means, marker="o", color=f"#{col}", linewidth=2, label=label)
        ax.fill_between(turns,
                        [m-s for m,s in zip(means,sds)],
                        [m+s for m,s in zip(means,sds)],
                        color=f"#{col}", alpha=0.15)
    ax.set_xlabel("Conversation turn"); ax.set_ylabel("Avg quality score (1-5)")
    ax.set_title("Quality score progression per turn — FL_21 vs FL_22", fontsize=12, fontweight="bold")
    ax.legend(); ax.grid(alpha=0.3); ax.set_ylim(1,5)
    plt.tight_layout()
    buf=io.BytesIO(); fig.savefig(buf,format="png",dpi=150,bbox_inches="tight")
    buf.seek(0); plt.close(fig)
    return buf

# ================================================================
# MAIN EXPORT
# ================================================================

def run():
    print("\nBuilding final Excel file...")

    # ── Load all data ─────────────────────────────────────
    df_clean    = pd.read_json(OUTPUT_DIR/"df_clean.json")   if (OUTPUT_DIR/"df_clean.json").exists()   else pd.DataFrame()
    df_fl21     = pd.read_json(OUTPUT_DIR/"df_fl21.json")    if (OUTPUT_DIR/"df_fl21.json").exists()    else pd.DataFrame()
    df_fl22     = pd.read_json(OUTPUT_DIR/"df_fl22.json")    if (OUTPUT_DIR/"df_fl22.json").exists()    else pd.DataFrame()
    df_coding   = pd.read_json(OUTPUT_DIR/"df_coding.json")  if (OUTPUT_DIR/"df_coding.json").exists()  else pd.DataFrame()
    df_desc     = pd.read_json(OUTPUT_DIR/"df_desc.json")    if (OUTPUT_DIR/"df_desc.json").exists()    else pd.DataFrame()
    df_agg      = pd.read_json(OUTPUT_DIR/"df_agg.json")     if (OUTPUT_DIR/"df_agg.json").exists()     else pd.DataFrame()
    df_merged   = pd.read_json(OUTPUT_DIR/"df_merged.json")  if (OUTPUT_DIR/"df_merged.json").exists()  else pd.DataFrame()
    df_ai       = pd.read_json(OUTPUT_DIR/"df_ai.json")      if (OUTPUT_DIR/"df_ai.json").exists()      else pd.DataFrame()
    df_prog     = pd.read_json(OUTPUT_DIR/"df_progression.json") if (OUTPUT_DIR/"df_progression.json").exists() else pd.DataFrame()

    q_ttests    = pd.read_json(OUTPUT_DIR/"q_ttests.json")    if (OUTPUT_DIR/"q_ttests.json").exists()    else pd.DataFrame()
    q_constructs= pd.read_json(OUTPUT_DIR/"q_constructs.json") if (OUTPUT_DIR/"q_constructs.json").exists() else pd.DataFrame()

    synthesis = {}
    if (OUTPUT_DIR/"synthesis.json").exists():
        with open(OUTPUT_DIR/"synthesis.json",encoding="utf-8") as f:
            synthesis = json.load(f)

    wordfreq = {}
    if (OUTPUT_DIR/"wordfreq.json").exists():
        with open(OUTPUT_DIR/"wordfreq.json",encoding="utf-8") as f:
            wordfreq = json.load(f)

    wb = Workbook()

    # ────────────────────────────────────────────────────────────
    # SHEET 1 — Raw data (cleaned)
    # ────────────────────────────────────────────────────────────
    ws1 = wb.active; ws1.title = "01_Raw_Data_Cleaned"
    if not df_clean.empty:
        cols_show = [c for c in df_clean.columns if not c.endswith("_num")][:50]
        write_df(ws1, df_clean[cols_show], f"Raw cleaned data — all respondents (n={len(df_clean)})", COLOR_HDR)
    print("  Sheet 01 — Raw data")

    # ────────────────────────────────────────────────────────────
    # SHEET 2 — FL_21 cleaned
    # ────────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("02_FL21_Friendly")
    if not df_fl21.empty:
        cols_show = [c for c in df_fl21.columns if not c.endswith("_num")][:50]
        write_df(ws2, df_fl21[cols_show], f"FL_21 — Friendly tone (n={len(df_fl21)})", COLOR_FL21)
    print("  Sheet 02 — FL_21")

    # ────────────────────────────────────────────────────────────
    # SHEET 3 — FL_22 cleaned
    # ────────────────────────────────────────────────────────────
    ws3 = wb.create_sheet("03_FL22_Professional")
    if not df_fl22.empty:
        cols_show = [c for c in df_fl22.columns if not c.endswith("_num")][:50]
        write_df(ws3, df_fl22[cols_show], f"FL_22 — Professional tone (n={len(df_fl22)})", COLOR_FL22)
    print("  Sheet 03 — FL_22")

    # ────────────────────────────────────────────────────────────
    # SHEET 4 — Likert scale coding
    # ────────────────────────────────────────────────────────────
    ws4 = wb.create_sheet("04_Scale_Coding")
    title_row(ws4, "Likert scale coding — questions T to AP", COLOR_HDR, 10)
    # Scale reference boxes
    scale_defs = [
        ("Likert Scale (7 pts)", "2E75B6",
         ["Strongly\ndisagree","Disagree","Somewhat\ndisagree","Neither agree\nnor disagree","Somewhat\nagree","Agree","Strongly\nagree"],
         "Columns: T U V W X Y  Z AA AB AC  AL AM AN AO AP"),
        ("Capable Scale (7 pts)", "1A6634",
         ["Not at all\ncapable","Very slightly\ncapable","Slightly\ncapable","Neither capable\nnor incapable","Somewhat\ncapable","Capable","Very\ncapable"],
         "Columns: AD AE"),
        ("Amount Scale (7 pts)", "7B3F00",
         ["Not at all","Very\nlittle","A little","A moderate\namount","Quite\na bit","A lot","A great\ndeal"],
         "Columns: AF AG AH AI  AJ AK"),
    ]
    row=3
    for sname, color, labels, col_info in scale_defs:
        ws4.merge_cells(start_row=row,start_column=1,end_row=row,end_column=10)
        c=ws4.cell(row=row,column=1,value=sname)
        c.font=Font(bold=True,color="FFFFFF",name="Arial",size=11)
        c.fill=PatternFill("solid",start_color=color)
        c.alignment=Alignment(horizontal="left",vertical="center")
        ws4.row_dimensions[row].height=22; row+=1
        ws4.merge_cells(start_row=row,start_column=1,end_row=row,end_column=10)
        c=ws4.cell(row=row,column=1,value=col_info)
        c.font=Font(italic=True,name="Arial",size=9,color="555555")
        ws4.row_dimensions[row].height=16; row+=1
        ws4.cell(row=row,column=1,value="Score →").font=Font(bold=True,name="Arial",size=9)
        for j,(score,label) in enumerate(zip(range(1,8),labels),2):
            bg,fc=SCORE_COLORS[score]
            c1=ws4.cell(row=row,column=j,value=score)
            c1.font=Font(bold=True,color=fc,name="Arial",size=11)
            c1.fill=PatternFill("solid",start_color=bg)
            c1.alignment=Alignment(horizontal="center",vertical="center"); c1.border=THIN
            c2=ws4.cell(row=row+1,column=j,value=label)
            c2.font=Font(name="Arial",size=8)
            c2.fill=PatternFill("solid",start_color="F8F9FA")
            c2.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c2.border=THIN
            ws4.column_dimensions[get_column_letter(j)].width=17
        ws4.row_dimensions[row].height=20; ws4.row_dimensions[row+1].height=34; row+=3
    # Coding table
    if not df_coding.empty:
        row+=1; write_df(ws4, df_coding, "Coding summary — all questions", "2C3E50", start_row=row, freeze=False)
    ws4.column_dimensions["A"].width=5
    print("  Sheet 04 — Scale coding")

    # ────────────────────────────────────────────────────────────
    # SHEET 5 — Questionnaire comparative analysis (p-values)
    # ────────────────────────────────────────────────────────────
    ws5 = wb.create_sheet("05_Questionnaire_PValues")
    title_row(ws5, "Comparative analysis — Questionnaire DVs (t-tests FL_21 vs FL_22)", COLOR_HDR, 14)
    if not q_ttests.empty:
        rows_t = q_ttests.to_dict("records")
        write_ttest_table(ws5, rows_t, start_row=2)
    ws5.freeze_panes="A3"
    print("  Sheet 05 — Questionnaire p-values")

    # ────────────────────────────────────────────────────────────
    # SHEET 6 — Construct summary
    # ────────────────────────────────────────────────────────────
    ws6 = wb.create_sheet("06_Construct_Summary")
    title_row(ws6, "Construct-level summary — FL_21 vs FL_22", COLOR_HDR, 12)
    if not q_constructs.empty:
        constr_colors = {
            "General Evaluation":"2E75B6","Perceived Manipulation":"C0392B",
            "AI Competence":"1A6634","Moral Responsibility":"7B3F00",
            "AI Independence":"6C3483","Chatbot Personality":"1A5276",
        }
        hdrs6=["Construct","N items","Mean FL_21","SD FL_21","N obs FL_21",
               "Mean FL_22","SD FL_22","N obs FL_22","Δ","t-stat","p-value","Cohen's d","Effect size"]
        for j,h in enumerate(hdrs6,1): hdr(ws6.cell(row=2,column=j,value=h),bg="2C3E50",sz=9)
        for i,(_,r) in enumerate(q_constructs.iterrows()):
            rr=i+3; alt=(i%2==0); constr=r.get("construct","")
            color=constr_colors.get(constr,"888888")
            c=ws6.cell(row=rr,column=1,value=constr)
            c.font=Font(name="Arial",size=10,bold=True,color="FFFFFF")
            c.fill=PatternFill("solid",start_color=color); c.border=THIN
            c.alignment=Alignment(horizontal="left",vertical="center")
            for j,col in enumerate(["n_items","mean_fl21","sd_fl21"],2):
                dat(ws6.cell(rr,j), r.get(col), alt=alt, fmt="0.000" if j>2 else None)
            dat(ws6.cell(rr,5), r.get("n_fl21"), alt=alt)
            for j,col in enumerate(["mean_fl22","sd_fl22"],6):
                dat(ws6.cell(rr,j), r.get(col), alt=alt, fmt="0.000")
            dat(ws6.cell(rr,8), r.get("n_fl22"), alt=alt)
            delta_cell(ws6.cell(rr,9), r.get("delta"))
            dat(ws6.cell(rr,10), r.get("t"), alt=alt, fmt="0.0000")
            pval(ws6.cell(rr,11), r.get("p"))
            dat(ws6.cell(rr,12), r.get("cohens_d"), alt=alt, fmt="0.000")
            dat(ws6.cell(rr,13), r.get("effect_size",""), alt=alt)
            ws6.row_dimensions[rr].height=22
    for j,w in enumerate([30,8,11,9,10,11,9,10,12,11,14,11,12],1):
        ws6.column_dimensions[get_column_letter(j)].width=w
    print("  Sheet 06 — Construct summary")

    # ────────────────────────────────────────────────────────────
    # SHEET 7 — Computed metrics (message length, TTR, sentiment)
    # ────────────────────────────────────────────────────────────
    ws7 = wb.create_sheet("07_Computed_Metrics")
    title_row(ws7, "Computed metrics — message length, lexical richness, sentiment (t-tests)", COLOR_HDR, 14)
    metrics = synthesis.get("ttests_continuous",[])
    metric_rows = [r for r in metrics if r.get("variable","") in
                   ["avg_words_per_msg","avg_word_len","avg_ttr","pct_questions","avg_sentiment","n_turns"]]
    if metric_rows:
        next_r = write_ttest_table(ws7, metric_rows, start_row=2)
    # Also add the raw metrics per participant
    if not df_agg.empty:
        df_agg_disp = df_agg.rename(columns={"respondent":"ResponseId"})
        row_off = (len(metric_rows)+5) if metric_rows else 5
        write_df(ws7, df_agg_disp,
                 "Aggregated metrics per participant", "2C3E50",
                 start_row=row_off, freeze=False)
    print("  Sheet 07 — Computed metrics")

    # ────────────────────────────────────────────────────────────
    # SHEET 8 — Word frequencies
    # ────────────────────────────────────────────────────────────
    ws8 = wb.create_sheet("08_Word_Frequencies")
    title_row(ws8, "Word frequencies & TF-IDF — Participants and Chatbot", COLOR_HDR, 8)
    if wordfreq:
        tw_m21 = wordfreq.get("top_words_participants_fl21",[])
        tw_m22 = wordfreq.get("top_words_participants_fl22",[])
        tw_r21 = wordfreq.get("top_words_chatbot_fl21",[])
        tw_r22 = wordfreq.get("top_words_chatbot_fl22",[])
        for j,h in enumerate(["Rank","Word FL_21 (Part.)","Freq","Word FL_22 (Part.)","Freq",
                               "Word Chatbot FL_21","Freq","Word Chatbot FL_22","Freq"],1):
            hdr(ws8.cell(row=2,column=j,value=h),
                bg=COLOR_FL21 if "21" in h else COLOR_FL22 if "22" in h else "2C3E50", sz=9)
        max_r = max(len(tw_m21),len(tw_m22),len(tw_r21),len(tw_r22),1)
        for i in range(max_r):
            r=3+i; alt=(i%2==0)
            dat(ws8.cell(r,1), i+1, alt=alt, bold=True)
            if i<len(tw_m21): dat(ws8.cell(r,2),tw_m21[i][0],center=False,alt=alt); dat(ws8.cell(r,3),tw_m21[i][1],alt=alt)
            if i<len(tw_m22): dat(ws8.cell(r,4),tw_m22[i][0],center=False,alt=alt); dat(ws8.cell(r,5),tw_m22[i][1],alt=alt)
            if i<len(tw_r21): dat(ws8.cell(r,6),tw_r21[i][0],center=False,alt=alt); dat(ws8.cell(r,7),tw_r21[i][1],alt=alt)
            if i<len(tw_r22): dat(ws8.cell(r,8),tw_r22[i][0],center=False,alt=alt); dat(ws8.cell(r,9),tw_r22[i][1],alt=alt)
        # TF-IDF section
        tfidf_start = max_r + 5
        title_row(ws8,"Distinctive words TF-IDF — FL_21 vs FL_22","C0392B",8,row=tfidf_start)
        for j,h in enumerate(["Rank","Distinctive FL_21 (Part.)","Δ TF-IDF",
                               "Distinctive FL_22 (Part.)","Δ TF-IDF",
                               "Distinctive Chatbot FL_21","Δ TF-IDF",
                               "Distinctive Chatbot FL_22","Δ TF-IDF"],1):
            hdr(ws8.cell(tfidf_start+1,j,value=h),bg="2C3E50",sz=9)
        for i in range(20):
            r=tfidf_start+2+i; alt=(i%2==0)
            dat(ws8.cell(r,1),i+1,alt=alt,bold=True)
            for col_idx, key in [(2,"tfidf_participants_fl21"),(4,"tfidf_participants_fl22"),
                                  (6,"tfidf_chatbot_fl21"),(8,"tfidf_chatbot_fl22")]:
                tw=wordfreq.get(key,[])
                if i<len(tw):
                    dat(ws8.cell(r,col_idx),tw[i][0],center=False,alt=alt)
                    dat(ws8.cell(r,col_idx+1),tw[i][2],alt=alt,fmt="0.0000")
    # Insert figures
    insert_png(ws8, OUTPUT_DIR/"fig_topwords_participants.png","K2",22,9)
    insert_png(ws8, OUTPUT_DIR/"fig_topwords_chatbot.png","K24",22,9)
    insert_png(ws8, OUTPUT_DIR/"fig_tfidf_participants.png","K46",22,9)
    insert_png(ws8, OUTPUT_DIR/"fig_tfidf_chatbot.png","K68",22,9)
    for j,w in enumerate([6,22,10,22,10,22,10,22,10],1):
        ws8.column_dimensions[get_column_letter(j)].width=w
    print("  Sheet 08 — Word frequencies")

    # ────────────────────────────────────────────────────────────
    # SHEET 9 — AI results per participant
    # ────────────────────────────────────────────────────────────
    ws9 = wb.create_sheet("09_AI_Results_Per_Participant")
    title_row(ws9, f"AI classification per participant (n={len(df_ai)})", "4A235A", 20)
    if not df_ai.empty:
        ai_cols = ["respondent_id","version","n_turns","quality_global","quality_precision",
                   "quality_examples","quality_relevance","quality_richness",
                   "action_global","action_concrete_pb","action_advice","action_use_case",
                   "profile_engagement","profile_elaboration","profile_coherence","profile_expertise",
                   "content_opinion","content_suggestion","content_emotion","content_competitor",
                   "breakpoint_exists","breakpoint_turn","completed_fully","end_type",
                   "bot_score_friendly","bot_score_professional","bot_compliance_score","bot_compliant",
                   "bot_tone_label","summary","key_verbatim"]
        ai_cols = [c for c in ai_cols if c in df_ai.columns]
        for j,col in enumerate(ai_cols,1):
            hdr(ws9.cell(row=2,column=j,value=col), bg="4A235A", sz=8)
        ws9.row_dimensions[2].height=28
        QCOLS={1:"FF4444",2:"FF8C00",3:"FFD700",4:"4CAF50",5:"1A6634"}
        QFCS ={1:"FFFFFF",2:"FFFFFF",3:"000000",4:"FFFFFF",5:"FFFFFF"}
        for i,(_,row) in enumerate(df_ai.sort_values(["version","quality_global"],ascending=[True,False]).iterrows()):
            r=i+3; alt=(i%2==0)
            for j,col in enumerate(ai_cols,1):
                v=row.get(col)
                c=ws9.cell(row=r,column=j)
                if col in ["quality_global","action_global","profile_engagement","profile_expertise",
                           "bot_score_friendly","bot_score_professional","bot_compliance_score"] and pd.notna(v):
                    vi=int(v); c.value=vi; c.border=THIN
                    c.alignment=Alignment(horizontal="center",vertical="center")
                    c.fill=PatternFill("solid",start_color=QCOLS.get(vi,"FFFFFF"))
                    c.font=Font(name="Arial",size=9,bold=True,color=QFCS.get(vi,"000000"))
                elif col=="version":
                    c.value=v; c.border=THIN
                    c.alignment=Alignment(horizontal="center",vertical="center")
                    c.fill=PatternFill("solid",start_color=COLOR_FL21 if v=="FL_21" else COLOR_FL22)
                    c.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
                else:
                    dat(c,v,center=(col not in ["respondent_id","summary","key_verbatim","end_type"]),alt=alt)
            ws9.row_dimensions[r].height=30
        for j,col in enumerate(ai_cols,1):
            ws9.column_dimensions[get_column_letter(j)].width=40 if col in ["summary","key_verbatim"] else 18
        ws9.freeze_panes="A3"
    print("  Sheet 09 — AI results per participant")

    # ────────────────────────────────────────────────────────────
    # SHEET 10 — p-values all DVs by tone
    # ────────────────────────────────────────────────────────────
    ws10 = wb.create_sheet("10_All_DVs_PValues")
    title_row(ws10, "All DVs — t-tests and chi2 by chatbot tone (FL_21 vs FL_22)", COLOR_HDR, 14)
    ttests = synthesis.get("ttests_continuous",[])
    chi2s  = synthesis.get("chi2_binary",[])
    next_r = write_ttest_table(ws10, ttests, start_row=2)
    if chi2s:
        ws10.cell(row=next_r+1,column=1,value="Binary DVs (Chi2 tests):").font=Font(bold=True,name="Arial",size=10)
        write_chi2_table(ws10, chi2s, start_row=next_r+2)
    # p-value legend
    r_leg = next_r + len(chi2s) + 5
    ws10.cell(row=r_leg,column=1,value="Significance legend:").font=Font(bold=True,name="Arial",size=9)
    for r2,(bg,fc,txt) in enumerate([("1A6634","FFFFFF","p < 0.001  ***  Highly significant"),
                                      ("52BE80","FFFFFF","p < 0.01   **   Significant"),
                                      ("F9E79F","7D6608","p < 0.05   *    Marginally significant"),
                                      ("FADBD8","922B21","p ≥ 0.05   ns   Not significant")],r_leg+1):
        ws10.merge_cells(start_row=r2,start_column=1,end_row=r2,end_column=5)
        c=ws10.cell(row=r2,column=1,value=txt)
        c.fill=PatternFill("solid",start_color=bg); c.font=Font(name="Arial",size=9,bold=True,color=fc); c.border=THIN
    ws10.freeze_panes="A3"
    print("  Sheet 10 — All DVs p-values")

    # ────────────────────────────────────────────────────────────
    # SHEET 11 — Correlations
    # ────────────────────────────────────────────────────────────
    ws11 = wb.create_sheet("11_Correlations")
    title_row(ws11, "Correlations between DVs — Pearson r (p < .05)", COLOR_HDR, 9)
    corrs = synthesis.get("correlations",[])
    if corrs:
        hdrs11=["Variable X","Label X","Variable Y","Label Y","r","p-value","Sig.","Strength","Direction"]
        for j,h in enumerate(hdrs11,1): hdr(ws11.cell(row=2,column=j,value=h),bg="2C3E50",sz=9)
        for i,r in enumerate(corrs):
            rr=i+3; alt=(i%2==0)
            dat(ws11.cell(rr,1), r.get("var_x",""),   center=False, alt=alt, bold=True)
            dat(ws11.cell(rr,2), r.get("label_x",""), center=False, alt=alt)
            dat(ws11.cell(rr,3), r.get("var_y",""),   center=False, alt=alt, bold=True)
            dat(ws11.cell(rr,4), r.get("label_y",""), center=False, alt=alt)
            rv = r.get("r")
            rc = ws11.cell(rr,5,value=rv)
            rc.border=THIN; rc.alignment=Alignment(horizontal="center",vertical="center")
            rc.number_format="0.000"
            if rv:
                if abs(rv)>=0.5:   rc.fill=PatternFill("solid",start_color="1A6634"); rc.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
                elif abs(rv)>=0.3: rc.fill=PatternFill("solid",start_color="A8D8A8"); rc.font=Font(name="Arial",size=9)
                else:              rc.font=Font(name="Arial",size=9)
                if alt and abs(rv) < 0.5: rc.fill=PatternFill("solid",start_color=COLOR_ALT)
            pval(ws11.cell(rr,6), r.get("p"))
            dat(ws11.cell(rr,7), r.get("sig",""),       alt=alt, bold=True)
            dat(ws11.cell(rr,8), r.get("strength",""),  alt=alt)
            dat(ws11.cell(rr,9), r.get("direction",""), alt=alt)
        for j,w in enumerate([20,32,20,32,8,14,8,12,12],1):
            ws11.column_dimensions[get_column_letter(j)].width=w
    ws11.freeze_panes="A3"
    print("  Sheet 11 — Correlations")

    # ────────────────────────────────────────────────────────────
    # SHEET 12 — Brief compliance
    # ────────────────────────────────────────────────────────────
    ws12 = wb.create_sheet("12_Brief_Compliance")
    title_row(ws12, "Chatbot brief compliance — Perceived tone vs requested tone", COLOR_HDR, 8)
    comp = synthesis.get("compliance",{})
    if comp:
        summary_data = [
            ["Metric","FL_21 (Friendly)","FL_22 (Professional)"],
            ["N respondents", comp.get("per_version",{}).get("FL_21",{}).get("n"), comp.get("per_version",{}).get("FL_22",{}).get("n")],
            ["Compliance score (mean)", comp.get("per_version",{}).get("FL_21",{}).get("compliance_mean"), comp.get("per_version",{}).get("FL_22",{}).get("compliance_mean")],
            ["Compliance score (SD)",   comp.get("per_version",{}).get("FL_21",{}).get("compliance_sd"),   comp.get("per_version",{}).get("FL_22",{}).get("compliance_sd")],
            ["% Compliant",            comp.get("per_version",{}).get("FL_21",{}).get("pct_compliant"),   comp.get("per_version",{}).get("FL_22",{}).get("pct_compliant")],
            ["% Emojis used",          comp.get("per_version",{}).get("FL_21",{}).get("pct_emojis"),      comp.get("per_version",{}).get("FL_22",{}).get("pct_emojis")],
            ["% Informal address",     comp.get("per_version",{}).get("FL_21",{}).get("pct_informal"),    comp.get("per_version",{}).get("FL_22",{}).get("pct_informal")],
            ["% Formal address",       comp.get("per_version",{}).get("FL_21",{}).get("pct_formal"),      comp.get("per_version",{}).get("FL_22",{}).get("pct_formal")],
            ["Avg encouragement phrases", comp.get("per_version",{}).get("FL_21",{}).get("avg_encouragements"), comp.get("per_version",{}).get("FL_22",{}).get("avg_encouragements")],
            ["Avg sober phrases",      comp.get("per_version",{}).get("FL_21",{}).get("avg_sober"),       comp.get("per_version",{}).get("FL_22",{}).get("avg_sober")],
            ["Avg friendly words",     comp.get("per_version",{}).get("FL_21",{}).get("avg_friendly_words"), comp.get("per_version",{}).get("FL_22",{}).get("avg_friendly_words")],
            ["Avg formal words",       comp.get("per_version",{}).get("FL_21",{}).get("avg_formal_words"),    comp.get("per_version",{}).get("FL_22",{}).get("avg_formal_words")],
        ]
        for j,h in enumerate(summary_data[0],1):
            hdr(ws12.cell(row=2,column=j,value=h),
                bg="2C3E50" if j==1 else COLOR_FL21 if j==2 else COLOR_FL22, sz=10)
        for i,row_d in enumerate(summary_data[1:],1):
            alt=(i%2==0)
            for j,val in enumerate(row_d,1):
                c=ws12.cell(row=2+i,column=j)
                dat(c,val,center=(j>1),alt=alt,bold=(j==1),fmt="0.00" if isinstance(val,float) else None)
        # T-test on compliance
        t_comp = comp.get("ttest_compliance",{})
        if t_comp:
            r_tc=len(summary_data)+4
            ws12.cell(row=r_tc,column=1,value="T-test (compliance score FL_21 vs FL_22):").font=Font(bold=True,name="Arial",size=10)
            write_ttest_table(ws12, [t_comp], start_row=r_tc+1)
        # Per-participant compliance table
        if not df_ai.empty and "bot_compliance_score" in df_ai.columns:
            comp_cols=["respondent_id","version","bot_score_friendly","bot_score_professional",
                       "bot_tone_label","bot_compliance_score","bot_compliant","bot_gap",
                       "bot_emojis","bot_informal_addr","bot_formal_addr",
                       "bot_n_friendly_words","bot_n_formal_words","bot_n_encouragements"]
            comp_cols=[c for c in comp_cols if c in df_ai.columns]
            row_off=len(summary_data)+12
            write_df(ws12, df_ai[comp_cols], "Compliance detail per participant", "4A235A", start_row=row_off, freeze=False)
    for j,w in enumerate([35,20,22],1):
        ws12.column_dimensions[get_column_letter(j)].width=w
    print("  Sheet 12 — Brief compliance")

    # ────────────────────────────────────────────────────────────
    # SHEET 13 — Dropout analysis
    # ────────────────────────────────────────────────────────────
    ws13 = wb.create_sheet("13_Dropout_Analysis")
    title_row(ws13, "Dropout analysis — Completers vs Dropouts", COLOR_HDR, 14)
    drop = synthesis.get("dropouts",{})
    if drop:
        summary13 = [
            ["","Global","FL_21","FL_22"],
            ["% Dropout", drop.get("pct_dropout_global"), drop.get("pct_dropout_fl21"), drop.get("pct_dropout_fl22")],
            ["N Dropouts", drop.get("n_dropouts"), None, None],
            ["N Completers", drop.get("n_completers"), None, None],
        ]
        for j,h in enumerate(summary13[0],1):
            hdr(ws13.cell(row=2,column=j,value=h), bg="2C3E50", sz=10)
        for i,row_d in enumerate(summary13[1:],1):
            for j,val in enumerate(row_d,1):
                dat(ws13.cell(2+i,j),val,center=(j>1),bold=(j==1),alt=(i%2==0))
        # End type distribution
        fin_dist=drop.get("end_type_distribution",{})
        if fin_dist:
            r_fd=7
            ws13.cell(row=r_fd,column=1,value="End type distribution (%):").font=Font(bold=True,name="Arial",size=10)
            all_types=set()
            for v in fin_dist.values(): all_types.update(v.keys())
            for j,h in enumerate(["End type","Global","FL_21","FL_22"],1):
                hdr(ws13.cell(r_fd+1,j,value=h),bg="2C3E50",sz=9)
            for i,et in enumerate(sorted(all_types)):
                r=r_fd+2+i; alt=(i%2==0)
                dat(ws13.cell(r,1),et,center=False,alt=alt,bold=True)
                for j,ver in enumerate(["global","FL_21","FL_22"],2):
                    dat(ws13.cell(r,j),fin_dist.get(ver,{}).get(et),alt=alt,fmt="0.0")
        # Profile comparison table
        prof=drop.get("profile_comparison",[])
        if prof:
            r_pc=7+len(all_types)+5 if fin_dist else 7
            ws13.cell(row=r_pc,column=1,value="Profile comparison — Dropouts vs Completers:").font=Font(bold=True,name="Arial",size=10)
            hdrs13p=["Variable","Label","Mean Dropouts","SD Dropouts","Mean Completers","SD Completers","t","p-value","Sig.","Cohen's d","Effect"]
            for j,h in enumerate(hdrs13p,1): hdr(ws13.cell(r_pc+1,j,value=h),bg="2C3E50",sz=9)
            for i,r in enumerate(prof):
                rr=r_pc+2+i; alt=(i%2==0)
                dat(ws13.cell(rr,1),r.get("variable",""),center=False,alt=alt,bold=True)
                dat(ws13.cell(rr,2),r.get("label",""),   center=False,alt=alt)
                dat(ws13.cell(rr,3),r.get("mean_dropouts"),alt=alt,fmt="0.000")
                dat(ws13.cell(rr,4),r.get("sd_dropouts"),  alt=alt,fmt="0.000")
                dat(ws13.cell(rr,5),r.get("mean_completers"),alt=alt,fmt="0.000")
                dat(ws13.cell(rr,6),r.get("sd_completers"),  alt=alt,fmt="0.000")
                dat(ws13.cell(rr,7),r.get("t"),alt=alt,fmt="0.0000")
                pval(ws13.cell(rr,8),r.get("p"))
                dat(ws13.cell(rr,9), r.get("sig",""),      alt=alt,bold=True)
                dat(ws13.cell(rr,10),r.get("cohens_d"),    alt=alt,fmt="0.000")
                dat(ws13.cell(rr,11),r.get("effect_size",""),alt=alt)
    for j,w in enumerate([22,35,14,12,16,12,10,14,8,11,12],1):
        ws13.column_dimensions[get_column_letter(j)].width=w
    print("  Sheet 13 — Dropout analysis")

    # ────────────────────────────────────────────────────────────
    # SHEET 14 — Turn-by-turn progression
    # ────────────────────────────────────────────────────────────
    ws14 = wb.create_sheet("14_Turn_Progression")
    title_row(ws14, "Quality score progression per conversation turn — FL_21 vs FL_22", COLOR_HDR, 8)
    prog = synthesis.get("progression",{})
    if prog:
        all_turns = sorted(set(int(t) for v in prog.values() for t in v.keys()))
        for j,h in enumerate(["Turn","Mean FL_21","SD FL_21","N FL_21","Mean FL_22","SD FL_22","N FL_22","Δ (21−22)"],1):
            hdr(ws14.cell(row=2,column=j,value=h),
                bg=COLOR_FL21 if "21" in h else COLOR_FL22 if "22" in h else "2C3E50", sz=9)
        for i,t in enumerate(all_turns):
            r=3+i; alt=(i%2==0); ts=str(t)
            d21=prog.get("FL_21",{}).get(ts,{}); d22=prog.get("FL_22",{}).get(ts,{})
            m21=d21.get("mean"); m22=d22.get("mean")
            dat(ws14.cell(r,1),t,alt=alt,bold=True)
            dat(ws14.cell(r,2),m21,alt=alt,fmt="0.000")
            dat(ws14.cell(r,3),d21.get("sd"),alt=alt,fmt="0.000")
            dat(ws14.cell(r,4),d21.get("n"),alt=alt)
            dat(ws14.cell(r,5),m22,alt=alt,fmt="0.000")
            dat(ws14.cell(r,6),d22.get("sd"),alt=alt,fmt="0.000")
            dat(ws14.cell(r,7),d22.get("n"),alt=alt)
            if m21 and m22:
                delta_cell(ws14.cell(r,8), m21-m22)
    # Insert progression figure
    fig_buf = make_progression_fig(prog)
    if fig_buf:
        fig_path = OUTPUT_DIR/"fig_progression.png"
        with open(fig_path,"wb") as f: f.write(fig_buf.read())
        insert_png(ws14, fig_path, "J2", 22, 9)
    # Per-participant turn data
    if not df_prog.empty:
        row_off=len(all_turns)+6 if prog else 5
        write_df(ws14, df_prog, "Raw progression data per participant per turn",
                 "2C3E50", start_row=row_off, freeze=False)
    for j,w in enumerate([8,12,10,8,12,10,8,12],1):
        ws14.column_dimensions[get_column_letter(j)].width=w
    print("  Sheet 14 — Turn progression")

# ────────────────────────────────────────────────────────────
    # SHEET 15 — Mediation analyses
    # ────────────────────────────────────────────────────────────
    ws15 = wb.create_sheet("15_Mediation_Analyses")
    title_row(ws15, "Simple mediation analyses — Bootstrap 5000 samples, 95% CI (Hayes/PROCESS style)", COLOR_HDR, 12)

    med_path = OUTPUT_DIR / "mediation.json"
    if med_path.exists():
        with open(med_path, encoding="utf-8") as f:
          content = f.read().strip()
        if not content:
          print("  ⚠️  mediation.json is empty — skipping Sheet 15")
          med_results = []
        else:
            med_results = json.loads(content)

        # Legend row
        ws15.merge_cells("A2:L2")
        leg = ws15.cell(row=2, column=1,
                        value="IV: Chatbot tone (FL_21=+1, FL_22=−1)  |  "
                              "Mediator: AI Competence_1  |  "
                              "Significant indirect effect = 95% CI excludes 0")
        leg.font      = Font(italic=True, name="Arial", size=9, color="555555")
        leg.alignment = Alignment(horizontal="left", vertical="center")

        current_row = 4
        PATH_COLORS = {
            "path_a":       "2E75B6",
            "path_b":       "1A6634",
            "path_c":       "7B3F00",
            "path_c_prime": "6C3483",
            "indirect_effect": "C0392B",
        }
        PATH_LABELS = {
            "path_a":        "Path a  — IV → Mediator",
            "path_b":        "Path b  — Mediator → DV (controlling IV)",
            "path_c":        "Path c  — IV → DV (total effect)",
            "path_c_prime":  "Path c' — IV → DV (direct effect)",
            "indirect_effect":"Indirect effect a×b (bootstrapped)",
        }

        for med in med_results:
            if "error" in med and "path_a" not in med:
                ws15.cell(row=current_row, column=1,
                          value=f"{med.get('name','')} — ERROR: {med.get('error','')}").font = Font(name="Arial", size=10, color="FF0000")
                current_row += 2
                continue

            # Mediation title block
            ws15.merge_cells(start_row=current_row, start_column=1,
                             end_row=current_row, end_column=12)
            c = ws15.cell(row=current_row, column=1,
                          value=f"{med.get('name','')}  —  "
                                f"Tone → Competence_1 → {med.get('dv_label','')}")
            c.font      = Font(bold=True, color="FFFFFF", name="Arial", size=11)
            c.fill      = PatternFill("solid", start_color="1F4E79")
            c.alignment = Alignment(horizontal="left", vertical="center")
            ws15.row_dimensions[current_row].height = 24
            current_row += 1

            # Info row
            ws15.cell(row=current_row, column=1,
                      value=f"n = {med.get('n','')}  |  "
                            f"Bootstrap samples = {med.get('n_bootstrap',5000)}  |  "
                            f"Confidence level = {int((1-med.get('alpha',0.05))*100)}%").font = Font(italic=True, name="Arial", size=9)
            current_row += 1

            # Column headers
            hdrs_med = ["Path", "Description", "b (coef)", "SE", "t", "p-value", "Sig.",
                        "95% CI lower", "95% CI upper", "Significant?", "Interpretation"]
            for j, h in enumerate(hdrs_med, 1):
                hdr(ws15.cell(row=current_row, column=j, value=h), bg="2C3E50", sz=9)
            ws15.row_dimensions[current_row].height = 22
            current_row += 1

            # Path rows
            for path_key, path_label in PATH_LABELS.items():
                pdata = med.get(path_key, {})
                if not pdata:
                    continue
                r = current_row
                color = PATH_COLORS[path_key]

                c = ws15.cell(row=r, column=1, value=path_key.replace("_", " ").upper())
                c.font  = Font(bold=True, color="FFFFFF", name="Arial", size=9)
                c.fill  = PatternFill("solid", start_color=color)
                c.border = THIN
                c.alignment = Alignment(horizontal="center", vertical="center")

                dat(ws15.cell(r, 2), path_label, center=False)
                dat(ws15.cell(r, 3), pdata.get("coef"),      fmt="0.0000")
                dat(ws15.cell(r, 4), pdata.get("se"),        fmt="0.0000")
                dat(ws15.cell(r, 5), pdata.get("t"),         fmt="0.0000")

                # p-value (not available for indirect)
                if path_key == "indirect_effect":
                    ws15.cell(row=r, column=6).value  = "—"
                    ws15.cell(row=r, column=6).border = THIN
                    ws15.cell(row=r, column=6).alignment = Alignment(horizontal="center")
                    ws15.cell(row=r, column=7).value  = "—"
                    ws15.cell(row=r, column=7).border = THIN
                else:
                    pval(ws15.cell(r, 6), pdata.get("p"))
                    dat(ws15.cell(r, 7), pdata.get("sig", ""), bold=True)

                # CI columns
                if path_key == "indirect_effect":
                    ci_low = pdata.get("ci_lower")
                    ci_up  = pdata.get("ci_upper")
                    sig_i  = pdata.get("significant", False)

                    c_low = ws15.cell(row=r, column=8, value=ci_low)
                    c_up  = ws15.cell(row=r, column=9, value=ci_up)
                    for cc in [c_low, c_up]:
                        cc.border = THIN
                        cc.alignment = Alignment(horizontal="center", vertical="center")
                        cc.number_format = "0.0000"
                        if ci_low is not None and ci_up is not None:
                            if sig_i:
                                cc.fill = PatternFill("solid", start_color="D5F5E3")
                                cc.font = Font(name="Arial", size=9, bold=True, color="1A6634")
                            else:
                                cc.fill = PatternFill("solid", start_color="FADBD8")
                                cc.font = Font(name="Arial", size=9, color="922B21")

                    sig_cell = ws15.cell(row=r, column=10,
                                         value="YES ✓" if sig_i else "NO ✗")
                    sig_cell.border = THIN
                    sig_cell.alignment = Alignment(horizontal="center", vertical="center")
                    if sig_i:
                        sig_cell.fill = PatternFill("solid", start_color="1A6634")
                        sig_cell.font = Font(name="Arial", size=9, bold=True, color="FFFFFF")
                    else:
                        sig_cell.fill = PatternFill("solid", start_color="FADBD8")
                        sig_cell.font = Font(name="Arial", size=9, color="922B21")

                    interp = ws15.cell(row=r, column=11,
                                        value=pdata.get("interpretation", ""))
                    interp.border = THIN
                    interp.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    interp.font = Font(name="Arial", size=9, italic=True)
                else:
                    for j in [8, 9, 10, 11]:
                        ws15.cell(row=r, column=j).value  = "—"
                        ws15.cell(row=r, column=j).border = THIN
                        ws15.cell(row=r, column=j).alignment = Alignment(horizontal="center")

                ws15.row_dimensions[r].height = 20
                current_row += 1

            # Mediation type summary
            med_type = med.get("mediation_type", "")
            ws15.merge_cells(start_row=current_row, start_column=1,
                             end_row=current_row, end_column=11)
            c = ws15.cell(row=current_row, column=1,
                          value=f"Conclusion: {med_type}")
            is_mediation = "No mediation" not in med_type and "Inconsistent" not in med_type
            c.font  = Font(bold=True, name="Arial", size=10,
                           color="FFFFFF" if is_mediation else "922B21")
            c.fill  = PatternFill("solid",
                                  start_color="1A6634" if is_mediation else "FADBD8")
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = THIN
            ws15.row_dimensions[current_row].height = 22
            current_row += 3  # space between analyses

        # Column widths
        col_widths_med = [16, 40, 10, 10, 10, 14, 8, 12, 12, 14, 35]
        for j, w in enumerate(col_widths_med, 1):
            ws15.column_dimensions[get_column_letter(j)].width = w

    else:
        ws15.cell(row=2, column=1,
                  value="mediation.json not found — run 08_mediation.py first").font = Font(
                  name="Arial", size=10, color="FF0000", italic=True)

    print("  Sheet 15 — Mediation analyses")

    # ── Save ─────────────────────────────────────────────────────
    wb.save(OUTPUT_FILE)
    print(f"\nExcel saved: {OUTPUT_FILE}")
    return str(OUTPUT_FILE)

if __name__ == "__main__":
    run()
