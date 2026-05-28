"""
07_export.py
Generates the 15-sheet Excel workbook.
Run AFTER all analysis scripts (01–06 + 08_new_analyses.py).
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
# STYLE CONSTANTS
# ================================================================

THIN = Border(
    left=Side(style="thin",color="BFBFBF"), right=Side(style="thin",color="BFBFBF"),
    top=Side(style="thin",color="BFBFBF"),  bottom=Side(style="thin",color="BFBFBF"),
)

# Block header colors
C_EVAL    = "2E75B6"   # blue    — evaluation questionnaire
C_QUALITY = "1A6634"   # green   — quality AI scores
C_METRICS = "7B3F00"   # brown   — text metrics
C_ACTION  = "6C3483"   # purple  — actionability
C_PSY     = "C0392B"   # red     — perception IA
C_CONV    = "1A5276"   # dark blue — conversation/dropout
C_COMPLY  = "17202A"   # near-black — compliance

# ================================================================
# SHARED FORMATTING HELPERS
# ================================================================

def hdr(c, bg=COLOR_HDR, fc="FFFFFF", sz=10, bold=True):
    c.font      = Font(bold=bold, color=fc, name="Arial", size=sz)
    c.fill      = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border    = THIN

def dat(c, val, center=True, alt=False, bold=False, fc="000000", fmt=None):
    v = None if (isinstance(val, float) and np.isnan(val)) else val
    if isinstance(v, bool): v = "Yes" if v else "No"
    c.value     = v
    c.font      = Font(name="Arial", size=9, bold=bold, color=fc)
    c.alignment = Alignment(horizontal="center" if center else "left",
                            vertical="center", wrap_text=True)
    c.border    = THIN
    if alt: c.fill = PatternFill("solid", start_color=COLOR_ALT)
    if fmt: c.number_format = fmt

def pval_cell(c, p):
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
    if   d > 0.05:  c.fill=PatternFill("solid",start_color="D5F5E3"); c.font=Font(name="Arial",size=9,bold=True,color="1A6634")
    elif d < -0.05: c.fill=PatternFill("solid",start_color="FADBD8"); c.font=Font(name="Arial",size=9,bold=True,color="922B21")
    else:           c.fill=PatternFill("solid",start_color="F2F3F4"); c.font=Font(name="Arial",size=9,color="555555")

def section_title(ws, text, bg, ncols, row, sz=11):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    c = ws.cell(row=row,column=1,value=text)
    c.font      = Font(bold=True,color="FFFFFF",name="Arial",size=sz)
    c.fill      = PatternFill("solid",start_color=bg)
    c.alignment = Alignment(horizontal="left",vertical="center")
    ws.row_dimensions[row].height = 22

def summary_box(ws, text, row, ncols, bg="EBF5FB"):
    ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=ncols)
    c = ws.cell(row=row,column=1,value=text)
    c.font      = Font(italic=True,name="Arial",size=9,color="1F4E79")
    c.fill      = PatternFill("solid",start_color=bg)
    c.alignment = Alignment(horizontal="left",vertical="center",wrap_text=True)
    c.border    = THIN
    ws.row_dimensions[row].height = 30

def write_ttest_block(ws, rows, start_row, bg_hdr, n_cols=14):
    hdrs = ["Variable","Label","N FL_21","Mean FL_21","SD FL_21",
            "N FL_22","Mean FL_22","SD FL_22","Δ (21−22)","t-stat","p-value","Sig.","Cohen's d","Effect size"]
    for j,h in enumerate(hdrs,1): hdr(ws.cell(start_row,j,value=h),bg=bg_hdr,sz=9)
    ws.row_dimensions[start_row].height = 24
    for i,r in enumerate(rows):
        rr=start_row+1+i; alt=(i%2==0)
        dat(ws.cell(rr,1), r.get("variable",""), center=False, alt=alt, bold=True)
        dat(ws.cell(rr,2), r.get("label",""),    center=False, alt=alt)
        dat(ws.cell(rr,3), r.get("n_fl21"),      alt=alt)
        dat(ws.cell(rr,4), r.get("mean_fl21"),   alt=alt, fmt="0.000")
        dat(ws.cell(rr,5), r.get("sd_fl21"),     alt=alt, fmt="0.000")
        dat(ws.cell(rr,6), r.get("n_fl22"),      alt=alt)
        dat(ws.cell(rr,7), r.get("mean_fl22"),   alt=alt, fmt="0.000")
        dat(ws.cell(rr,8), r.get("sd_fl22"),     alt=alt, fmt="0.000")
        delta_cell(ws.cell(rr,9), r.get("delta"))
        dat(ws.cell(rr,10), r.get("t"),          alt=alt, fmt="0.0000")
        pval_cell(ws.cell(rr,11), r.get("p"))
        dat(ws.cell(rr,12), r.get("sig",""),     alt=alt, bold=True)
        dat(ws.cell(rr,13), r.get("cohens_d"),   alt=alt, fmt="0.000")
        dat(ws.cell(rr,14), r.get("effect_size",""), alt=alt)
    widths=[20,38,8,11,9,8,11,9,12,10,14,8,11,12]
    for j,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(j)].width=w
    return start_row+1+len(rows)

def write_chi2_block(ws, rows, start_row, bg_hdr):
    hdrs=["Variable","Label","% FL_21","% FL_22","% Total","Δ% (21−22)","Chi²","p-value","Sig."]
    for j,h in enumerate(hdrs,1): hdr(ws.cell(start_row,j,value=h),bg=bg_hdr,sz=9)
    for i,r in enumerate(rows):
        rr=start_row+1+i; alt=(i%2==0)
        dat(ws.cell(rr,1), r.get("variable",""),  center=False, alt=alt, bold=True)
        dat(ws.cell(rr,2), r.get("label",""),     center=False, alt=alt)
        dat(ws.cell(rr,3), r.get("pct_fl21"),     alt=alt, fmt="0.0")
        dat(ws.cell(rr,4), r.get("pct_fl22"),     alt=alt, fmt="0.0")
        dat(ws.cell(rr,5), r.get("pct_total"),    alt=alt, fmt="0.0")
        delta_cell(ws.cell(rr,6), r.get("delta_pct"))
        dat(ws.cell(rr,7), r.get("chi2"),         alt=alt, fmt="0.0000")
        pval_cell(ws.cell(rr,8), r.get("p"))
        dat(ws.cell(rr,9), r.get("sig",""),       alt=alt, bold=True)
    return start_row+1+len(rows)

def write_mediation_block(ws, med, start_row, label):
    """Write one mediation model (5 paths) starting at start_row."""
    # Title
    ws.merge_cells(start_row=start_row,start_column=1,end_row=start_row,end_column=11)
    c=ws.cell(start_row,1,value=label)
    c.font=Font(bold=True,color="FFFFFF",name="Arial",size=10)
    c.fill=PatternFill("solid",start_color="1F4E79")
    c.alignment=Alignment(horizontal="left",vertical="center"); c.border=THIN
    ws.row_dimensions[start_row].height=22
    start_row+=1

    if "error" in med:
        ws.cell(start_row,1,value=f"ERROR: {med['error']}").font=Font(color="FF0000",name="Arial",size=9)
        return start_row+2

    # Info
    ws.cell(start_row,1,value=f"n={med.get('n','')} | Bootstrap=5000 | 95% CI").font=Font(italic=True,name="Arial",size=9)
    start_row+=1

    # Headers
    path_hdrs=["Path","Description","b","SE","t","p-value","Sig.","CI lower","CI upper","Sig. indirect","Conclusion"]
    for j,h in enumerate(path_hdrs,1): hdr(ws.cell(start_row,j,value=h),bg="2C3E50",sz=9)
    start_row+=1

    PATH_COLORS={"path_a":"2E75B6","path_b":"1A6634","path_c":"7B3F00","path_cp":"6C3483","indirect":"C0392B"}
    PATH_DESC={
        "path_a":  f"IV ({med.get('iv','')}) → Mediator ({med.get('mediator','')})",
        "path_b":  f"Mediator → DV ({med.get('dv','')}) | controlling IV",
        "path_c":  f"IV → DV (total effect)",
        "path_cp": f"IV → DV (direct effect, controlling M)",
        "indirect":f"Indirect effect (a × b) — bootstrapped",
    }
    path_keys=[("path_a","path_a"),("path_b","path_b"),("path_c","path_c"),("path_cp","path_cp"),("indirect","indirect")]

    for path_key, color_key in path_keys:
        pdata = med.get(path_key if path_key != "indirect" else "indirect",{})
        if path_key == "indirect": pdata = med.get("indirect",{})
        r=start_row
        c_path=ws.cell(r,1,value=path_key.upper())
        c_path.font=Font(bold=True,color="FFFFFF",name="Arial",size=9)
        c_path.fill=PatternFill("solid",start_color=PATH_COLORS[color_key])
        c_path.border=THIN; c_path.alignment=Alignment(horizontal="center",vertical="center")
        dat(ws.cell(r,2),PATH_DESC.get(path_key,""),center=False)
        dat(ws.cell(r,3),pdata.get("coef"),fmt="0.0000")
        dat(ws.cell(r,4),pdata.get("se"),fmt="0.0000")
        dat(ws.cell(r,5),pdata.get("t"),fmt="0.0000")
        if path_key=="indirect":
            ws.cell(r,6,value="—").border=THIN
            ws.cell(r,7,value="—").border=THIN
        else:
            pval_cell(ws.cell(r,6),pdata.get("p"))
            dat(ws.cell(r,7),pdata.get("sig",""),bold=True)
        if path_key=="indirect":
            sig_i=pdata.get("significant",False)
            for col_j,val in [(8,pdata.get("ci_low")),(9,pdata.get("ci_up"))]:
                cc=ws.cell(r,col_j,value=val)
                cc.border=THIN; cc.alignment=Alignment(horizontal="center",vertical="center")
                cc.number_format="0.0000"
                cc.fill=PatternFill("solid",start_color="D5F5E3" if sig_i else "FADBD8")
                cc.font=Font(name="Arial",size=9,bold=sig_i,color="1A6634" if sig_i else "922B21")
            sc=ws.cell(r,10,value="YES ✓" if sig_i else "NO ✗")
            sc.border=THIN; sc.alignment=Alignment(horizontal="center",vertical="center")
            sc.fill=PatternFill("solid",start_color="1A6634" if sig_i else "FADBD8")
            sc.font=Font(name="Arial",size=9,bold=True,color="FFFFFF" if sig_i else "922B21")
            dat(ws.cell(r,11),pdata.get("interpretation",""),center=False)
        else:
            for col_j in [8,9,10,11]:
                ws.cell(r,col_j,value="—").border=THIN
        ws.row_dimensions[r].height=18
        start_row+=1

    # Conclusion
    ws.merge_cells(start_row=start_row,start_column=1,end_row=start_row,end_column=11)
    med_type=med.get("mediation_type","")
    is_med="No mediation" not in med_type and "Inconsistent" not in med_type
    cc=ws.cell(start_row,1,value=f"Conclusion: {med_type}")
    cc.font=Font(bold=True,name="Arial",size=9,color="FFFFFF" if is_med else "922B21")
    cc.fill=PatternFill("solid",start_color="1A6634" if is_med else "FADBD8")
    cc.alignment=Alignment(horizontal="left",vertical="center"); cc.border=THIN
    ws.row_dimensions[start_row].height=18
    start_row+=2
    return start_row

def write_regression_block(ws, result, start_row, label):
    if not result: return start_row+2
    ws.merge_cells(start_row=start_row,start_column=1,end_row=start_row,end_column=8)
    c=ws.cell(start_row,1,value=label)
    c.font=Font(bold=True,color="FFFFFF",name="Arial",size=10)
    c.fill=PatternFill("solid",start_color="1F4E79")
    c.alignment=Alignment(horizontal="left",vertical="center"); c.border=THIN
    ws.row_dimensions[start_row].height=22; start_row+=1

    model=result.get("model",{}) if "model" in result else result
    if not model: return start_row+1

    n=model.get("n",""); r2=model.get("r2","")
    ws.cell(start_row,1,value=f"n={n} | R² = {r2}").font=Font(italic=True,name="Arial",size=9)
    start_row+=1
    for j,h in enumerate(["Predictor","Label","b","SE","t","p-value","Sig.",""],1):
        hdr(ws.cell(start_row,j,value=h),bg="2C3E50",sz=9)
    start_row+=1
    for i,pred in enumerate(model.get("predictors",[])):
        alt=(i%2==0)
        dat(ws.cell(start_row,1),pred.get("variable",""),center=False,alt=alt,bold=True)
        dat(ws.cell(start_row,2),pred.get("label",""),center=False,alt=alt)
        dat(ws.cell(start_row,3),pred.get("b"),alt=alt,fmt="0.0000")
        dat(ws.cell(start_row,4),pred.get("se"),alt=alt,fmt="0.0000")
        dat(ws.cell(start_row,5),pred.get("t"),alt=alt,fmt="0.0000")
        pval_cell(ws.cell(start_row,6),pred.get("p"))
        dat(ws.cell(start_row,7),pred.get("sig",""),alt=alt,bold=True)
        start_row+=1
    return start_row+1

def write_mediation_summary_table(ws, models_list, start_row, bloc_label):
    """One-line-per-model summary table for multi-mediation blocs."""
    section_title(ws,bloc_label,"C0392B",13,start_row); start_row+=1
    hdrs=["#","IV","Mediator","DV","n","a sig","b sig","Indirect","CI lower","CI upper","Sig.","Type"]
    for j,h in enumerate(hdrs,1): hdr(ws.cell(start_row,j,value=h),bg="2C3E50",sz=9)
    start_row+=1
    for i,item in enumerate(models_list):
        r=item.get("result",{})
        alt=(i%2==0)
        if "error" in r:
            dat(ws.cell(start_row,1),i+1,alt=alt)
            dat(ws.cell(start_row,2),item.get("iv_label",""),center=False,alt=alt)
            dat(ws.cell(start_row,3),item.get("mediator",""),center=False,alt=alt)
            dat(ws.cell(start_row,4),item.get("dv_label",""),center=False,alt=alt)
            dat(ws.cell(start_row,5),"ERROR",alt=alt)
            for j in range(6,13): ws.cell(start_row,j,value="—").border=THIN
        else:
            ind=r.get("indirect",{})
            dat(ws.cell(start_row,1),i+1,alt=alt)
            dat(ws.cell(start_row,2),item.get("iv_label",""),center=False,alt=alt)
            dat(ws.cell(start_row,3),item.get("mediator",""),center=False,alt=alt)
            dat(ws.cell(start_row,4),item.get("dv_label",""),center=False,alt=alt)
            dat(ws.cell(start_row,5),r.get("n"),alt=alt)
            dat(ws.cell(start_row,6),r.get("path_a",{}).get("sig",""),alt=alt,bold=True)
            dat(ws.cell(start_row,7),r.get("path_b",{}).get("sig",""),alt=alt,bold=True)
            dat(ws.cell(start_row,8),ind.get("coef"),alt=alt,fmt="0.0000")
            dat(ws.cell(start_row,9),ind.get("ci_low"),alt=alt,fmt="0.0000")
            dat(ws.cell(start_row,10),ind.get("ci_up"),alt=alt,fmt="0.0000")
            sig=ind.get("significant",False)
            sc=ws.cell(start_row,11,value="✓ SIG" if sig else "✗ ns")
            sc.border=THIN; sc.alignment=Alignment(horizontal="center",vertical="center")
            sc.fill=PatternFill("solid",start_color="D5F5E3" if sig else "FADBD8")
            sc.font=Font(name="Arial",size=9,bold=sig,color="1A6634" if sig else "922B21")
            dat(ws.cell(start_row,12),r.get("mediation_type",""),center=False,alt=alt)
        ws.row_dimensions[start_row].height=18; start_row+=1
    for j,w in enumerate([5,25,22,22,6,8,8,10,10,10,10,22],1):
        ws.column_dimensions[get_column_letter(j)].width=w
    return start_row+1

def insert_png(ws, path, anchor, w_cm=18, h_cm=9):
    if not Path(path).exists(): return
    img=XLImage(str(path))
    img.width=int(w_cm*37.795); img.height=int(h_cm*37.795)
    ws.add_image(img,anchor)

# ================================================================
# LOAD ALL DATA
# ================================================================

def load_all():
    def jload(name):
        p=OUTPUT_DIR/name
        if p.exists():
            with open(p,encoding="utf-8") as f: return json.load(f)
        return {}

    return {
        "df_clean":    pd.read_json(OUTPUT_DIR/"df_clean.json")   if (OUTPUT_DIR/"df_clean.json").exists()  else pd.DataFrame(),
        "df_fl21":     pd.read_json(OUTPUT_DIR/"df_fl21.json")    if (OUTPUT_DIR/"df_fl21.json").exists()   else pd.DataFrame(),
        "df_fl22":     pd.read_json(OUTPUT_DIR/"df_fl22.json")    if (OUTPUT_DIR/"df_fl22.json").exists()   else pd.DataFrame(),
        "df_coding":   pd.read_json(OUTPUT_DIR/"df_coding.json")  if (OUTPUT_DIR/"df_coding.json").exists() else pd.DataFrame(),
        "df_ai":       pd.read_json(OUTPUT_DIR/"df_ai.json")      if (OUTPUT_DIR/"df_ai.json").exists()     else pd.DataFrame(),
        "df_agg":      pd.read_json(OUTPUT_DIR/"df_agg.json")     if (OUTPUT_DIR/"df_agg.json").exists()    else pd.DataFrame(),
        "df_prog":     pd.read_json(OUTPUT_DIR/"df_progression.json") if (OUTPUT_DIR/"df_progression.json").exists() else pd.DataFrame(),
        "correlations":      jload("correlations_full.json"),
        "vif":               jload("vif_full.json"),
        "tone_comparisons":  jload("tone_comparisons.json"),
        "mediation_tone":    jload("mediation_tone.json"),
        "regression_noton":  jload("regression_noton.json"),
        "mediation_noton":   jload("mediation_noton.json"),
        "dropout":           jload("dropout_corrected.json"),
        "wordfreq":          jload("wordfreq.json"),
        "synthesis":         jload("synthesis.json"),
    }

# ================================================================
# SHEET BUILDERS
# ================================================================

def build_sheet1(wb, df_clean):
    ws = wb.active; ws.title="01_Raw_Data_Cleaned"
    if df_clean.empty: return
    cols=[c for c in df_clean.columns if not c.endswith("_num")]
    section_title(ws,f"Raw cleaned data — all respondents (n={len(df_clean)})",COLOR_HDR,len(cols),1)
    for j,col in enumerate(cols,1): hdr(ws.cell(2,j,value=col),bg="2C3E50",sz=9)
    for i,(_,row) in enumerate(df_clean[cols].iterrows()):
        for j,v in enumerate(row,1): dat(ws.cell(3+i,j),v,center=(j>2),alt=(i%2==0))
    ws.freeze_panes="A3"

def build_sheet2(wb, df_fl21):
    ws=wb.create_sheet("02_FL21_Friendly")
    if df_fl21.empty: return
    cols=[c for c in df_fl21.columns if not c.endswith("_num")]
    section_title(ws,f"FL_21 — Friendly tone (n={len(df_fl21)})",COLOR_FL21,len(cols),1)
    for j,col in enumerate(cols,1): hdr(ws.cell(2,j,value=col),bg=COLOR_FL21,sz=9)
    for i,(_,row) in enumerate(df_fl21[cols].iterrows()):
        for j,v in enumerate(row,1): dat(ws.cell(3+i,j),v,center=(j>2),alt=(i%2==0))
    ws.freeze_panes="A3"

def build_sheet3(wb, df_fl22):
    ws=wb.create_sheet("03_FL22_Professional")
    if df_fl22.empty: return
    cols=[c for c in df_fl22.columns if not c.endswith("_num")]
    section_title(ws,f"FL_22 — Professional tone (n={len(df_fl22)})",COLOR_FL22,len(cols),1)
    for j,col in enumerate(cols,1): hdr(ws.cell(2,j,value=col),bg=COLOR_FL22,sz=9)
    for i,(_,row) in enumerate(df_fl22[cols].iterrows()):
        for j,v in enumerate(row,1): dat(ws.cell(3+i,j),v,center=(j>2),alt=(i%2==0))
    ws.freeze_panes="A3"

def build_sheet4(wb, df_coding):
    ws=wb.create_sheet("04_Scale_Coding")
    section_title(ws,"Likert scale coding — questions T to AP",COLOR_HDR,10,1)
    scale_defs=[
        ("Likert Scale (7 pts)","2E75B6",
         ["Strongly\ndisagree","Disagree","Somewhat\ndisagree","Neither agree\nnor disagree","Somewhat\nagree","Agree","Strongly\nagree"],
         "Columns: evaluation_1-6, Perceived Manipulati_1-4, personnality-Manip_1-5"),
        ("Capable Scale (7 pts)","1A6634",
         ["Not at all\ncapable","Very slightly\ncapable","Slightly\ncapable","Neither capable\nnor incapable","Somewhat\ncapable","Capable","Very\ncapable"],
         "Columns: Competence_1-2"),
        ("Amount Scale (7 pts)","7B3F00",
         ["Not at all","Very\nlittle","A little","A moderate\namount","Quite\na bit","A lot","A great\ndeal"],
         "Columns: Moral Responsibility_1-4, Sense of Independenc_1-2"),
    ]
    row=3
    for sname,color,labels,col_info in scale_defs:
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=10)
        c=ws.cell(row,1,value=sname)
        c.font=Font(bold=True,color="FFFFFF",name="Arial",size=11)
        c.fill=PatternFill("solid",start_color=color)
        c.alignment=Alignment(horizontal="left",vertical="center")
        ws.row_dimensions[row].height=22; row+=1
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=10)
        c=ws.cell(row,1,value=col_info)
        c.font=Font(italic=True,name="Arial",size=9,color="555555"); row+=1
        ws.cell(row,1,value="Score →").font=Font(bold=True,name="Arial",size=9)
        for j,(score,label) in enumerate(zip(range(1,8),labels),2):
            bg,fc=SCORE_COLORS[score]
            c1=ws.cell(row,j,value=score)
            c1.font=Font(bold=True,color=fc,name="Arial",size=11)
            c1.fill=PatternFill("solid",start_color=bg)
            c1.alignment=Alignment(horizontal="center",vertical="center"); c1.border=THIN
            c2=ws.cell(row+1,j,value=label)
            c2.font=Font(name="Arial",size=8)
            c2.fill=PatternFill("solid",start_color="F8F9FA")
            c2.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c2.border=THIN
            ws.column_dimensions[get_column_letter(j)].width=17
        ws.row_dimensions[row].height=20; ws.row_dimensions[row+1].height=34; row+=3
    if not df_coding.empty:
        row+=1
        section_title(ws,"Coding summary — all questions","2C3E50",8,row); row+=1
        coding_hdrs=["Excel col","Variable","Scale","Construct","Label","N coded","N total","% coded"]
        for j,h in enumerate(coding_hdrs,1): hdr(ws.cell(row,j,value=h),bg="2C3E50",sz=9)
        row+=1
        for i,(_,r) in enumerate(df_coding.iterrows()):
            alt=(i%2==0)
            for j,col in enumerate(["excel_col","variable","scale","construct","label","n_coded","n_total","pct_coded"],1):
                dat(ws.cell(row,j),r.get(col),center=(j>2),alt=alt)
            row+=1

def build_sheet5(wb):
    """Variable Dictionary — definitions, scales, what each value means."""
    ws=wb.create_sheet("05_Variable_Dictionary")
    section_title(ws,"Variable Dictionary — all retained variables with definitions and scales",COLOR_HDR,7,1)

    variables = [
        # EVALUATION QUESTIONNAIRE
        ("EVALUATION QUESTIONNAIRE (Likert 1-7: Strongly disagree → Strongly agree)",C_EVAL,None,None,None,None),
        ("evaluation_1_num","E1 Required effort","Likert 1-7","Questionnaire","Continuous",
         "1=no effort at all, 7=a great deal of effort required. REVERSED: higher = more effort."),
        ("evaluation_2_num","E2 Engagement felt","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'I felt engaged while answering this survey.'"),
        ("evaluation_3_num","E3 Chatbot appreciation","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'I enjoyed interacting with the chatbot.'"),
        ("evaluation_4_num","E4 Conversation utility","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'I found the conversation useful for providing my feedback.'"),
        ("evaluation_5_num","E5 Reuse intention","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'I would consider using this chatbot again.'"),
        ("evaluation_6_num","E6 Chatbot preference","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'I would prefer to provide feedback to a chatbot rather than a human.'"),
        # PERCEIVED MANIPULATION
        ("PERCEIVED MANIPULATION (Likert 1-7)",C_PSY,None,None,None,None),
        ("Perceived Manipulati_1_num","PM1 Freedom threat","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'The chatbot threatened my freedom to choose.' (Based on psychological reactance theory)"),
        ("Perceived Manipulati_2_num","PM2 Decision override","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'The chatbot tried to make a decision for me.'"),
        ("Perceived Manipulati_3_num","PM3 Manipulation","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'The chatbot tried to manipulate me.'"),
        ("Perceived Manipulati_4_num","PM4 Pressure","Likert 1-7","Questionnaire","Continuous",
         "1=strongly disagree, 7=strongly agree. 'The chatbot tried to pressure me.'"),
        # COMPETENCE
        ("AI COMPETENCE (Capable scale 1-7)",C_PSY,None,None,None,None),
        ("Competence_1_num","Comp1 Skills judgment","Capable 1-7","Questionnaire","Continuous",
         "1=not at all capable, 7=very capable. 'How capable is an AI system of judging someone's abilities?' Source: Liang et al."),
        ("Competence_2_num","Comp2 Moral judgment","Capable 1-7","Questionnaire","Continuous",
         "1=not at all capable, 7=very capable. 'How capable is an AI system of judging how moral someone's behavior is?'"),
        # MORAL RESPONSIBILITY
        ("MORAL RESPONSIBILITY (Amount scale 1-7: Not at all → A great deal)",C_PSY,None,None,None,None),
        ("Moral Responsibility_1_num","MR1 AI harm wrong","Amount 1-7","Questionnaire","Continuous",
         "1=not at all, 7=a great deal. 'How morally wrong would it be for this entity (AI) to harm a person?' Source: Robots/Chatbots/Self-Driving Cars article"),
        ("Moral Responsibility_2_num","MR2 AI responsible","Amount 1-7","Questionnaire","Continuous",
         "1=not at all, 7=a great deal. 'To what extent would this entity deserve to be held morally responsible for causing a negative outcome?'"),
        ("Moral Responsibility_3_num","MR3 Human harm AI","Amount 1-7","Questionnaire","Continuous",
         "1=not at all, 7=a great deal. 'How morally wrong would it be for someone to harm this entity (AI)?'"),
        ("Moral Responsibility_4_num","MR4 AI moral concern","Amount 1-7","Questionnaire","Continuous",
         "1=not at all, 7=a great deal. 'To what extent does this entity deserve to be treated with moral concern?'"),
        # SENSE OF INDEPENDENCE
        ("AI SENSE OF INDEPENDENCE (Amount scale 1-7)",C_PSY,None,None,None,None),
        ("Sense of Independenc_1_num","Ind1 AI plans & goals","Amount 1-7","Questionnaire","Continuous",
         "1=not at all, 7=a great deal. 'To what extent can this entity make plans and work towards goals?' Source: Robots/Chatbots article"),
        ("Sense of Independenc_2_num","Ind2 AI self-control","Amount 1-7","Questionnaire","Continuous",
         "1=not at all, 7=a great deal. 'To what extent can this entity exercise self-control?'"),
        # AI QUALITY SCORES
        ("AI QUALITY SCORES (GPT-4o-mini rated, scale 1-5)",C_QUALITY,None,None,None,None),
        ("quality_global","Quality global","1-5","AI (GPT-4o-mini)","Continuous",
         "1=almost all 1-3 word answers, 2=few sentences poorly developed, 3=adequate but incomplete, 4=detailed with context and example, 5=rich nuanced with multiple examples"),
        ("quality_precision","Quality precision","1-5","AI (GPT-4o-mini)","Continuous",
         "1=very vague no exploitable info, 3=some precise elements insufficient, 5=very precise every statement grounded in fact"),
        ("quality_examples","Quality examples","1-5","AI (GPT-4o-mini)","Continuous",
         "1=no concrete example, 3=one correct example, 5=multiple detailed examples illustrating different aspects"),
        ("quality_relevance","Quality relevance","1-5","AI (GPT-4o-mini)","Continuous",
         "1=answers do not address questions, 3=answers correctly but generically, 5=answers perfectly anticipates follow-up"),
        ("quality_richness","Quality richness","1-5","AI (GPT-4o-mini)","Continuous",
         "1=very limited vocabulary, 3=correct vocabulary some nuances, 5=rich vocabulary precise register"),
        # ACTIONABILITY
        ("ACTIONABILITY (Binary, AI-rated)",C_ACTION,None,None,None,None),
        ("action_concrete_pb","Concrete problem","Binary 0/1","AI (GPT-4o-mini)","Binary",
         "1=participant describes a precise problem with context (when X happens, Y occurs because Z), 0=vague dissatisfaction only"),
        ("action_advice","Applicable advice","Binary 0/1","AI (GPT-4o-mini)","Binary",
         "1=participant formulates an improvement a product team could implement directly, 0=vague wish not directly implementable"),
        ("action_use_case","Precise use case","Binary 0/1","AI (GPT-4o-mini)","Binary",
         "1=participant describes precise usage context (when/where/how/why they use the app), 0=usage described generically"),
        # TEXT METRICS
        ("TEXT METRICS (computed with NLTK)",C_METRICS,None,None,None,None),
        ("avg_words_per_msg","Avg words per message","Count","NLTK tokenizer","Continuous",
         "Average number of alphabetic tokens per participant message, across all turns. Excludes punctuation."),
        ("avg_word_len","Avg word length","Characters","NLTK tokenizer","Continuous",
         "Average number of characters per word across all messages. Proxy for lexical complexity."),
        ("n_turns","Number of turns","Count","Auto-detected","Continuous",
         "Number of participant message turns in the conversation (1 turn = 1 participant message + 1 bot response)."),
        # CONVERSATION
        ("CONVERSATION VARIABLES",C_CONV,None,None,None,None),
        ("end_type","End type","Categorical","Auto+AI detection","Categorical",
         "proper_end=chatbot sent closing phrase; early_dropout=≤3 turns AND ≤5 words in last message; minimal_response=last message ≤3 words without proper close. Note: mineral_response and minimale_response are encoding errors merged into minimal_response."),
        ("breakpoint_exists","Breakpoint detected","Binary 0/1","AI (GPT-4o-mini)","Binary",
         "1=AI detected a moment where at least 2 consecutive turns show lasting degradation in response quality, 0=no breakpoint"),
        ("breakpoint_turn","Breakpoint turn","Turn number","AI (GPT-4o-mini)","Continuous",
         "The conversation turn at which quality started to consistently decline. Only meaningful when breakpoint_exists=1."),
        # CHATBOT COMPLIANCE
        ("CHATBOT COMPLIANCE SCORES (AI-rated)",C_COMPLY,None,None,None,None),
        ("bot_score_friendly","Bot friendly score","1-5","AI (GPT-4o-mini)","Continuous",
         "1=very cold/distant, 3=neutral, 5=very warm/enthusiastic with emojis and informal address"),
        ("bot_score_professional","Bot professional score","1-5","AI (GPT-4o-mini)","Continuous",
         "1=very casual, 3=neutral, 5=very formal with systematic formal address and no emojis"),
        ("bot_compliance_score","Bot compliance score","1-5","AI (GPT-4o-mini)","Continuous",
         "How well the chatbot followed its tone brief. 1=completely opposite, 3=partially compliant notable drift, 5=perfectly compliant"),
        ("bot_coherence_score","Bot tone coherence","1-5","AI (GPT-4o-mini)","Continuous",
         "Consistency of tone throughout the conversation. 1=major drifts, 5=perfectly consistent throughout"),
    ]

    hdrs=["Variable name","Short label","Scale","Source","Type","Definition & scoring"]
    for j,h in enumerate(hdrs,1): hdr(ws.cell(2,j,value=h),bg="2C3E50",sz=10)
    ws.row_dimensions[2].height=24

    row=3
    for item in variables:
        if len(item)==6 and item[2] is None:
            # Section header
            section_title(ws,item[0],item[1],6,row); row+=1
        else:
            varname,shortlbl,scale,source,vtype,defn = item
            alt=((row-3)%2==0)
            dat(ws.cell(row,1),varname,center=False,alt=alt,bold=True)
            dat(ws.cell(row,2),shortlbl,center=False,alt=alt)
            dat(ws.cell(row,3),scale,alt=alt)
            dat(ws.cell(row,4),source,alt=alt)
            dat(ws.cell(row,5),vtype,alt=alt)
            dat(ws.cell(row,6),defn,center=False,alt=alt)
            ws.row_dimensions[row].height=40; row+=1

    for j,w in enumerate([28,22,15,20,12,65],1): ws.column_dimensions[get_column_letter(j)].width=w

def build_sheet6(wb, corr_data):
    """Full correlation matrix + bloc zooms."""
    ws=wb.create_sheet("06_Correlations")
    if not corr_data: ws.cell(1,1,value="Run 08_new_analyses.py first"); return

    section_title(ws,"Correlation Matrix — all retained variables",COLOR_HDR,3,1)
    summary_box(ws,"Interpretation: r = Pearson correlation coefficient. |r|<0.3=weak, 0.3-0.5=moderate, >0.5=strong. Green=positive, Red=negative. * p<.05  ** p<.01  *** p<.001",2,3)

    fm=corr_data.get("full_matrix",{})
    labels=fm.get("labels",[]); cols=fm.get("columns",[]); data=fm.get("data",[])
    n=len(labels)

    # Headers row (column labels)
    ws.cell(4,1,value="↓ Variable / Variable →").font=Font(bold=True,name="Arial",size=8)
    for j,lbl in enumerate(labels,2):
        c=ws.cell(4,j,value=lbl)
        c.font=Font(bold=True,name="Arial",size=7); c.border=THIN
        c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True,text_rotation=90)
        c.fill=PatternFill("solid",start_color="EBF5FB")
        ws.column_dimensions[get_column_letter(j)].width=6
    ws.row_dimensions[4].height=90
    ws.column_dimensions["A"].width=28

    for i,row_data in enumerate(data):
        r=5+i
        ws.cell(r,1,value=row_data.get("label","")).font=Font(name="Arial",size=8,bold=True)
        ws.cell(r,1).border=THIN
        r_vals=row_data.get("r",[]); p_vals=row_data.get("p",[])
        for j,(rv,pv) in enumerate(zip(r_vals,p_vals),2):
            c=ws.cell(r,j)
            c.border=THIN; c.alignment=Alignment(horizontal="center",vertical="center")
            c.font=Font(name="Arial",size=8)
            if rv is None: c.value=""; continue
            if i==j-2: c.value="1"; c.fill=PatternFill("solid",start_color="F2F3F4"); continue
            c.value=f"{rv:.2f}"
            sig_stars="" if pv is None or pv>=0.05 else "***" if pv<0.001 else "**" if pv<0.01 else "*"
            if sig_stars: c.value=f"{rv:.2f}{sig_stars}"
            if rv >= 0.5:   c.fill=PatternFill("solid",start_color="1A6634"); c.font=Font(name="Arial",size=8,color="FFFFFF",bold=True)
            elif rv >= 0.3: c.fill=PatternFill("solid",start_color="A9DFBF"); c.font=Font(name="Arial",size=8,color="000000")
            elif rv <= -0.5:c.fill=PatternFill("solid",start_color="922B21"); c.font=Font(name="Arial",size=8,color="FFFFFF",bold=True)
            elif rv <= -0.3:c.fill=PatternFill("solid",start_color="F1948A"); c.font=Font(name="Arial",size=8)

    # Significant pairs table
    next_row = 5+n+3
    sig_pairs=corr_data.get("significant_pairs",[])
    section_title(ws,f"Significant pairs sorted by |r| (n={len(sig_pairs)} pairs with p<.05)",C_EVAL,9,next_row)
    next_row+=1
    for j,h in enumerate(["Variable X","Label X","Variable Y","Label Y","r","p","Sig.","Strength","Direction"],1):
        hdr(ws.cell(next_row,j,value=h),bg="2C3E50",sz=9)
    next_row+=1
    for i,pair in enumerate(sig_pairs[:50]):  # top 50
        alt=(i%2==0)
        dat(ws.cell(next_row,1),pair["var_x"],center=False,alt=alt,bold=True)
        dat(ws.cell(next_row,2),pair["label_x"],center=False,alt=alt)
        dat(ws.cell(next_row,3),pair["var_y"],center=False,alt=alt,bold=True)
        dat(ws.cell(next_row,4),pair["label_y"],center=False,alt=alt)
        rv=pair["r"]
        rc=ws.cell(next_row,5,value=rv); rc.number_format="0.000"; rc.border=THIN
        rc.alignment=Alignment(horizontal="center",vertical="center")
        if abs(rv)>=0.5: rc.fill=PatternFill("solid",start_color="1A6634"); rc.font=Font(name="Arial",size=9,bold=True,color="FFFFFF")
        elif abs(rv)>=0.3: rc.fill=PatternFill("solid",start_color="A9DFBF")
        pval_cell(ws.cell(next_row,6),pair["p"])
        dat(ws.cell(next_row,7),pair["sig"],alt=alt,bold=True)
        dat(ws.cell(next_row,8),pair["strength"],alt=alt)
        dat(ws.cell(next_row,9),pair["direction"],alt=alt)
        next_row+=1

    # Bloc zooms
    bloc_mats=corr_data.get("bloc_matrices",{})
    for bloc_name,bdata in bloc_mats.items():
        next_row+=2
        section_title(ws,f"Zoom: {bloc_name}","2C3E50",len(bdata.get("labels",[]))+1,next_row)
        next_row+=1
        blabels=bdata.get("labels",[]); bd=bdata.get("data",[])
        ws.cell(next_row,1,value="↓/→").font=Font(bold=True,name="Arial",size=8)
        for j,lbl in enumerate(blabels,2):
            c=ws.cell(next_row,j,value=lbl)
            c.font=Font(bold=True,name="Arial",size=8); c.border=THIN
            c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
            c.fill=PatternFill("solid",start_color="EBF5FB")
            ws.column_dimensions[get_column_letter(j)].width=max(6,ws.column_dimensions[get_column_letter(j)].width)
        next_row+=1
        for row_d in bd:
            ws.cell(next_row,1,value=row_d.get("label","")).font=Font(name="Arial",size=8,bold=True)
            ws.cell(next_row,1).border=THIN
            for j,cell in enumerate(row_d.get("cells",[]),2):
                c=ws.cell(next_row,j); c.border=THIN
                c.alignment=Alignment(horizontal="center",vertical="center")
                rv=cell.get("r"); pv=cell.get("p")
                if rv is None: c.value="—"; c.font=Font(name="Arial",size=8); continue
                if rv==1.0: c.value="1"; c.fill=PatternFill("solid",start_color="F2F3F4"); c.font=Font(name="Arial",size=8); continue
                sig_s="" if pv is None or pv>=0.05 else "***" if pv<0.001 else "**" if pv<0.01 else "*"
                c.value=f"{rv:.2f}{sig_s}"; c.font=Font(name="Arial",size=8)
                if   rv>=0.5:  c.fill=PatternFill("solid",start_color="1A6634"); c.font=Font(name="Arial",size=8,color="FFFFFF",bold=True)
                elif rv>=0.3:  c.fill=PatternFill("solid",start_color="A9DFBF")
                elif rv<=-0.5: c.fill=PatternFill("solid",start_color="922B21"); c.font=Font(name="Arial",size=8,color="FFFFFF",bold=True)
                elif rv<=-0.3: c.fill=PatternFill("solid",start_color="F1948A")
            next_row+=1

    ws.freeze_panes="B5"

def build_sheet7(wb, vif_data):
    ws=wb.create_sheet("07_VIF")
    section_title(ws,"VIF — Variance Inflation Factor (each variable regressed on all others)",COLOR_HDR,5,1)
    summary_box(ws,"VIF interpretation: <3=OK (green), 3-5=Moderate (orange), 5-10=High (red), >10=Critical (dark red). High VIF = variable is largely explained by other variables → avoid combining in same regression model.",2,5)
    for j,h in enumerate(["Variable","Label","VIF","R² (vs others)","Flag"],1):
        hdr(ws.cell(3,j,value=h),bg="2C3E50",sz=10)
    ws.row_dimensions[3].height=24

    FLAG_COLORS={"OK (<3)":"D5F5E3","MODERATE (3-5)":"FAD7A0","HIGH (5-10)":"F1948A","CRITICAL (>10)":"C0392B","insufficient data":"F2F3F4"}
    FLAG_FC={"OK (<3)":"1A6634","MODERATE (3-5)":"7D6608","HIGH (5-10)":"922B21","CRITICAL (>10)":"FFFFFF","insufficient data":"888888"}
    for i,item in enumerate(vif_data):
        r=4+i; alt=(i%2==0)
        dat(ws.cell(r,1),item.get("variable",""),center=False,alt=alt,bold=True)
        dat(ws.cell(r,2),item.get("label",""),center=False,alt=alt)
        vif=item.get("vif")
        vc=ws.cell(r,3,value=vif)
        vc.border=THIN; vc.alignment=Alignment(horizontal="center",vertical="center")
        vc.number_format="0.00"
        flag=item.get("flag","")
        if vif:
            vc.fill=PatternFill("solid",start_color=FLAG_COLORS.get(flag,"F2F3F4"))
            vc.font=Font(name="Arial",size=10,bold=True,color=FLAG_FC.get(flag,"000000"))
        dat(ws.cell(r,4),item.get("r2"),alt=alt,fmt="0.000")
        fc_=ws.cell(r,5,value=flag)
        fc_.border=THIN; fc_.alignment=Alignment(horizontal="left",vertical="center")
        fc_.fill=PatternFill("solid",start_color=FLAG_COLORS.get(flag,"F2F3F4"))
        fc_.font=Font(name="Arial",size=9,color=FLAG_FC.get(flag,"000000"))
    for j,w in enumerate([30,28,10,14,20],1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="A4"

def build_sheet9(wb, drop_data):
    ws=wb.create_sheet("09_Dropout_Analysis")
    section_title(ws,"Dropout Analysis — End type distribution + E1/E2 analysis",COLOR_HDR,10,1)

    if not drop_data: ws.cell(2,1,value="Run 08_new_analyses.py first"); return

    # Summary
    row=2
    summary_box(ws,f"Total: n={drop_data.get('n_total')} | Proper end: {drop_data.get('n_completers')} ({100-drop_data.get('pct_dropout',0):.1f}%) | Dropouts: {drop_data.get('n_dropouts')} ({drop_data.get('pct_dropout')}%) | FL_21: {drop_data.get('pct_dropout_fl21')}% dropout | FL_22: {drop_data.get('pct_dropout_fl22')}% dropout",row,10)
    row+=2

    # Distribution table
    section_title(ws,"End type distribution by tone (corrected — mineral/minimale merged into minimal_response)",C_CONV,6,row); row+=1
    for j,h in enumerate(["End type","N total","% total","N FL_21","% FL_21","N FL_22","% FL_22"],1):
        hdr(ws.cell(row,j,value=h),bg="2C3E50",sz=9)
    row+=1
    overall=drop_data.get("overall_distribution",{})
    by_tone=drop_data.get("by_tone",{})
    n_total=drop_data.get("n_total",1)
    n_fl21=by_tone.get("FL_21",{}).get("proper_end",0)+by_tone.get("FL_21",{}).get("early_dropout",0)+by_tone.get("FL_21",{}).get("minimal_response",0)
    n_fl22=by_tone.get("FL_22",{}).get("proper_end",0)+by_tone.get("FL_22",{}).get("early_dropout",0)+by_tone.get("FL_22",{}).get("minimal_response",0)
    # recalculate n_fl21/fl22 from totals
    n21_total = sum(by_tone.get("FL_21",{}).values()) or 1
    n22_total = sum(by_tone.get("FL_22",{}).values()) or 1
    for i,cat in enumerate(["proper_end","early_dropout","minimal_response"]):
        alt=(i%2==0)
        n_tot=overall.get(cat,0); n21=by_tone.get("FL_21",{}).get(cat,0); n22=by_tone.get("FL_22",{}).get(cat,0)
        dat(ws.cell(row,1),cat,center=False,alt=alt,bold=True)
        dat(ws.cell(row,2),n_tot,alt=alt)
        dat(ws.cell(row,3),round(n_tot/n_total*100,1),alt=alt,fmt="0.0")
        dat(ws.cell(row,4),n21,alt=alt)
        dat(ws.cell(row,5),round(n21/n21_total*100,1) if n21_total else None,alt=alt,fmt="0.0")
        dat(ws.cell(row,6),n22,alt=alt)
        dat(ws.cell(row,7),round(n22/n22_total*100,1) if n22_total else None,alt=alt,fmt="0.0")
        row+=1

    # Chi2 end_type by tone
    chi2=drop_data.get("chi2_by_tone",{})
    row+=1
    ws.cell(row,1,value=f"Chi² end_type × tone: χ²={chi2.get('chi2')}  p={chi2.get('p')}  {chi2.get('sig')}").font=Font(italic=True,name="Arial",size=9)
    row+=2

    # Profile comparison
    section_title(ws,"Profile comparison — Dropouts vs Completers",C_CONV,10,row); row+=1
    write_ttest_block(ws,drop_data.get("profile_comparison",[]),row,C_CONV,10)
    row+=len(drop_data.get("profile_comparison",[]))+3

    # E1/E2 by end_type
    section_title(ws,"E1 (Effort) and E2 (Engagement) by end_type — are they different for each group?",C_EVAL,10,row)
    row+=1
    summary_box(ws,"Reading: for each end_type category, we compare the mean E1/E2 of participants IN that category vs all OTHERS. Significant p = that group scored meaningfully higher/lower on E1 or E2.",row,10); row+=1

    e1e2=drop_data.get("e1_e2_by_endtype",[])
    if e1e2:
        hdrs=["End type","Variable","Label","N in","Mean in","SD in","N out","Mean out","SD out","t","p","Sig.","d","Interpretation"]
        for j,h in enumerate(hdrs,1): hdr(ws.cell(row,j,value=h),bg=C_EVAL,sz=9)
        row+=1
        for i,r in enumerate(e1e2):
            alt=(i%2==0)
            for j,col in enumerate(["end_type","variable","label","n_in","mean_in","sd_in","n_out","mean_out","sd_out"],1):
                dat(ws.cell(row,j),r.get(col),center=(j>3),alt=alt,fmt="0.000" if j in [5,6,8,9] else None)
            dat(ws.cell(row,10),r.get("t"),alt=alt,fmt="0.0000")
            pval_cell(ws.cell(row,11),r.get("p"))
            dat(ws.cell(row,12),r.get("sig",""),alt=alt,bold=True)
            dat(ws.cell(row,13),r.get("cohens_d"),alt=alt,fmt="0.000")
            dat(ws.cell(row,14),r.get("interpretation",""),center=False,alt=alt)
            row+=1
    for j,w in enumerate([18,22,20,6,10,8,6,10,8,10,14,8,10,40],1): ws.column_dimensions[get_column_letter(j)].width=w

def build_sheet10(wb, tone_data, df_ai):
    ws=wb.create_sheet("10_Tone_Comparisons")
    section_title(ws,"Tone Comparisons — All DVs by Chatbot Tone (FL_21 Friendly vs FL_22 Professional)",COLOR_HDR,14,1)

    if not tone_data: ws.cell(2,1,value="Run 08_new_analyses.py first"); return

    # Auto-summary
    all_tests = (tone_data.get("evaluation",[]) + tone_data.get("quality_ai",[]) +
                 tone_data.get("text_metrics",[]) + tone_data.get("perception_ai",[]) +
                 tone_data.get("conversation",[]))
    sig_tests=[r for r in all_tests if r and r.get("sig") not in ("ns","n/a",None)]
    action_sig=[r for r in tone_data.get("actionability",[]) if r and r.get("sig") not in ("ns","n/a",None)]
    row=2
    summary_box(ws,f"SUMMARY — Significant results: {len(sig_tests)+len(action_sig)} variables out of {len(all_tests)+len(tone_data.get('actionability',[]))} tested. See color-coded sections below. Green=FL_21(Friendly) higher, Red=FL_22(Professional) higher.",row,14)
    row+=2

    # EVALUATION
    section_title(ws,"EVALUATION QUESTIONNAIRE — How participants experienced the chatbot (Likert 1-7)",C_EVAL,14,row); row+=1
    row=write_ttest_block(ws,tone_data.get("evaluation",[]),row,C_EVAL)+1

    # QUALITY AI
    section_title(ws,"AI-RATED RESPONSE QUALITY — Quality of participant responses as judged by GPT-4o-mini (1-5)",C_QUALITY,14,row); row+=1
    row=write_ttest_block(ws,tone_data.get("quality_ai",[]),row,C_QUALITY)+1

    # TEXT METRICS
    section_title(ws,"TEXT METRICS — Objective NLP-computed metrics on participant messages",C_METRICS,14,row); row+=1
    row=write_ttest_block(ws,tone_data.get("text_metrics",[]),row,C_METRICS)+1

    # ACTIONABILITY (chi2)
    section_title(ws,"ACTIONABILITY — Binary indicators of response quality (Chi² tests)",C_ACTION,14,row); row+=1
    if tone_data.get("actionability"):
        write_chi2_block(ws,tone_data.get("actionability",[]),row,C_ACTION)
        row+=len(tone_data.get("actionability",[]))+2

    # PERCEPTION IA
    section_title(ws,"PERCEPTION OF AI — 12 psychological questionnaire items (Likert/Amount/Capable 1-7)",C_PSY,14,row); row+=1
    row=write_ttest_block(ws,tone_data.get("perception_ai",[]),row,C_PSY)+1

    # CONVERSATION + COMPLIANCE
    section_title(ws,"CONVERSATION & CHATBOT COMPLIANCE — Turn count, breakpoints, compliance scores",C_CONV,14,row); row+=1
    row=write_ttest_block(ws,tone_data.get("conversation",[]),row,C_CONV)+1

    # END TYPE + DROPOUT BY TONE
    section_title(ws,"END TYPE & DROPOUT — How conversations ended, by tone",C_CONV,14,row); row+=1
    et=tone_data.get("end_type",{})
    if et:
        ws.cell(row,1,value=f"Chi² end_type × tone: χ²={et.get('chi2')}  p={et.get('p')}  {et.get('sig')}").font=Font(italic=True,name="Arial",size=9)
        row+=1
        dist=et.get("distribution",{})
        for j,h in enumerate(["End type","N FL_21","% FL_21","N FL_22","% FL_22"],1):
            hdr(ws.cell(row,j,value=h),bg=C_CONV,sz=9)
        row+=1
        for ver in ["FL_21","FL_22"]: pass  # build from dist
        for cat in ["proper_end","early_dropout","minimal_response"]:
            n21=dist.get("FL_21",{}).get(cat,0); n22=dist.get("FL_22",{}).get(cat,0)
            dat(ws.cell(row,1),cat,center=False); dat(ws.cell(row,2),n21); dat(ws.cell(row,3),None)
            dat(ws.cell(row,4),n22); dat(ws.cell(row,5),None); row+=1

    # Dropout summary
    drop=tone_data.get("dropout_by_tone",{})
    if drop:
        row+=1
        ws.cell(row,1,value=f"Dropout rate: FL_21={drop.get('FL_21',{}).get('early_dropout_pct',0)+drop.get('FL_21',{}).get('minimal_response_pct',0):.1f}%  FL_22={drop.get('FL_22',{}).get('early_dropout_pct',0)+drop.get('FL_22',{}).get('minimal_response_pct',0):.1f}%").font=Font(italic=True,name="Arial",size=9,color="922B21")
        row+=1

    # CONSTRUCT COMPOSITES
    composites=tone_data.get("construct_composites",[])
    if composites:
        row+=2
        section_title(ws,"CONSTRUCT COMPOSITES — Mean score per theoretical construct (higher-level view)",C_PSY,14,row); row+=1
        summary_box(ws,"Each composite = mean of all items within the construct. Complements the individual item analysis above.",row,14); row+=1
        row=write_ttest_block(ws,composites,row,C_PSY)+1

    # P-value legend
    row+=2
    for bg,fc,txt in [("1A6634","FFFFFF","p < 0.001  *** Highly significant"),
                       ("52BE80","FFFFFF","p < 0.01   **  Significant"),
                       ("F9E79F","7D6608","p < 0.05   *   Marginally significant"),
                       ("FADBD8","922B21","p ≥ 0.05   ns  Not significant")]:
        ws.merge_cells(start_row=row,start_column=1,end_row=row,end_column=6)
        c=ws.cell(row,1,value=txt)
        c.fill=PatternFill("solid",start_color=bg); c.font=Font(name="Arial",size=9,bold=True,color=fc); c.border=THIN
        row+=1
    ws.freeze_panes="A3"

def build_sheet11(wb, med_tone, regression_noton):
    ws=wb.create_sheet("11_Tone_Effects")
    section_title(ws,"Tone Effects — Mediations & Regressions where Chatbot Tone is the Independent Variable",COLOR_HDR,11,1)
    summary_box(ws,"This sheet answers: HOW does tone affect outcomes? Via which pathways? Bootstrap=5000 samples, 95% CI. Significant indirect effect = CI excludes 0.",2,11)

    row=4
    # BLOC A — MEDIATIONS TONE AS IV
    section_title(ws,"BLOC A — MEDIATIONS: Tone → Mediator → Outcome (bootstrap 5000 samples)",C_EVAL,11,row); row+=2

    med_labels={
        # Q1.4 — Tone → Competence → Psych DVs (8 models)
        "tone_Competence_1_num_Sense of Independenc_1_num": "Q1.4 — Tone → Comp1(Skills) → Ind1 AI plans & goals",
        "tone_Competence_1_num_Sense of Independenc_2_num": "Q1.4 — Tone → Comp1(Skills) → Ind2 AI self-control",
        "tone_Competence_2_num_Sense of Independenc_1_num": "Q1.4 — Tone → Comp2(Morality) → Ind1 AI plans & goals",
        "tone_Competence_2_num_Sense of Independenc_2_num": "Q1.4 — Tone → Comp2(Morality) → Ind2 AI self-control",
        "tone_Competence_1_num_Perceived Manipulati_1_num": "Q1.4 — Tone → Comp1(Skills) → PM1 Freedom threat",
        "tone_Competence_1_num_Perceived Manipulati_2_num": "Q1.4 — Tone → Comp1(Skills) → PM2 Decision override",
        "tone_Competence_2_num_Perceived Manipulati_1_num": "Q1.4 — Tone → Comp2(Morality) → PM1 Freedom threat",
        "tone_Competence_2_num_Perceived Manipulati_2_num": "Q1.4 — Tone → Comp2(Morality) → PM2 Decision override",
        # Q2.4 / Q2.5 — Tone → engagement/effort → quality
        "Q2.4_tone_E2_quality": "Q2.4 — Tone → E2 (Engagement felt) → Quality global",
        "Q2.5_tone_E1_quality": "Q2.5 — Tone → E1 (Required effort) → Quality global",
        # Q3.4 — Tone → appreciation → preference
        "Q3.4_tone_E3_E6":      "Q3.4 — Tone → E3 (Chatbot appreciation) → E6 (Chatbot preference)",
        # Q4.5 — Tone → competence → engagement
        "Q4.5a_tone_comp1_E2":  "Q4.5a — Tone → Competence_1 (Skills) → E2 (Engagement felt)",
        "Q4.5b_tone_comp2_E2":  "Q4.5b — Tone → Competence_2 (Morality) → E2 (Engagement felt)",
        # Tone → emotion → quality (6 models)
        "tone_emotion_quality_global":    "Tone → Emotion expressed → Quality global",
        "tone_emotion_quality_precision": "Tone → Emotion expressed → Quality precision",
        "tone_emotion_quality_examples":  "Tone → Emotion expressed → Quality examples",
        "tone_emotion_quality_relevance": "Tone → Emotion expressed → Quality relevance",
        "tone_emotion_quality_richness":  "Tone → Emotion expressed → Quality richness",
        # Tone → emotion → psych/eval DVs
        "tone_emotion_Perceived Manipulati_1_num": "Tone → Emotion expressed → PM1 Freedom threat",
        "tone_emotion_Perceived Manipulati_2_num": "Tone → Emotion expressed → PM2 Decision override",
        "tone_emotion_Competence_1_num":            "Tone → Emotion expressed → Comp1 Skills judgment",
        "tone_emotion_Sense of Independenc_1_num":  "Tone → Emotion expressed → Ind1 AI plans & goals",
        "tone_emotion_evaluation_2_num":            "Tone → Emotion expressed → E2 Engagement felt",
        "tone_emotion_evaluation_3_num":            "Tone → Emotion expressed → E3 Chatbot appreciation",
    }
    # Summary table first
    summary_rows=[]
    for key,lbl in med_labels.items():
        item=med_tone.get(key,{})
        m=item.get("model",{})
        if m and "error" not in m:
            ind=m.get("indirect",{})
            summary_rows.append({
                "label":lbl,"iv":"tone","m":m.get("mediator",""),
                "dv":m.get("dv",""),"n":m.get("n"),
                "a_sig":m.get("path_a",{}).get("sig",""),"b_sig":m.get("path_b",{}).get("sig",""),
                "indirect":ind.get("coef"),"ci_low":ind.get("ci_low"),"ci_up":ind.get("ci_up"),
                "sig_ind":ind.get("significant",False),"type":m.get("mediation_type",""),
            })
    section_title(ws,"Quick Summary — all tone mediations",C_EVAL,11,row); row+=1
    hdrs=["Model","IV","Mediator","DV","n","a","b","Indirect","CI low","CI high","Sig.","Type"]
    for j,h in enumerate(hdrs,1): hdr(ws.cell(row,j,value=h),bg=C_EVAL,sz=9)
    row+=1
    for i,sr in enumerate(summary_rows):
        alt=(i%2==0)
        dat(ws.cell(row,1),sr["label"],center=False,alt=alt)
        dat(ws.cell(row,2),"tone",alt=alt)
        dat(ws.cell(row,3),sr["m"],center=False,alt=alt)
        dat(ws.cell(row,4),sr["dv"],center=False,alt=alt)
        dat(ws.cell(row,5),sr["n"],alt=alt)
        dat(ws.cell(row,6),sr["a_sig"],alt=alt,bold=True)
        dat(ws.cell(row,7),sr["b_sig"],alt=alt,bold=True)
        dat(ws.cell(row,8),sr["indirect"],alt=alt,fmt="0.0000")
        dat(ws.cell(row,9),sr["ci_low"],alt=alt,fmt="0.0000")
        dat(ws.cell(row,10),sr["ci_up"],alt=alt,fmt="0.0000")
        sig=sr["sig_ind"]
        sc=ws.cell(row,11,value="✓ SIG" if sig else "✗ ns")
        sc.border=THIN; sc.alignment=Alignment(horizontal="center",vertical="center")
        sc.fill=PatternFill("solid",start_color="D5F5E3" if sig else "FADBD8")
        sc.font=Font(name="Arial",size=9,bold=sig,color="1A6634" if sig else "922B21")
        dat(ws.cell(row,12),sr["type"],center=False,alt=alt); row+=1

    row+=2
    section_title(ws,"Detailed mediation results",C_EVAL,11,row); row+=2
    for key,lbl in med_labels.items():
        item=med_tone.get(key,{})
        m=item.get("model",{})
        row=write_mediation_block(ws,m,row,lbl)

    # BLOC B — REGRESSIONS WITH TONE (from regression_noton which includes some with tone context)
    row+=2
    section_title(ws,"BLOC B — Additional regressions providing tone context (from Variable Relationships)",C_METRICS,11,row); row+=1
    summary_box(ws,"These regressions show predictors of quality and chatbot preference in context — not tone mediations but complementary analyses.",row,11); row+=2
    # Show Q2.1, Q2.2 as contextual regressions
    for key in ["Q2.1_words_quality","Q2.2_E2_quality","Q3.2b_E1_E2_quality_E6"]:
        item=regression_noton.get(key,{})
        if item:
            row=write_regression_block(ws,item.get("model",{}),row,item.get("label",""))

    for j,w in enumerate([45,12,22,22,6,8,8,10,10,10,10,30],1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="A4"

def build_sheet12(wb, regression_noton):
    ws=wb.create_sheet("12_Variable_Relationships")
    section_title(ws,"Variable Relationships — Regressions and tests where Tone is NOT the independent variable",COLOR_HDR,8,1)
    summary_box(ws,"This sheet answers: what predicts quality? what predicts chatbot preference? what is the role of effort, engagement, appreciation? Tone is excluded as IV.",2,8)

    if not regression_noton: ws.cell(3,1,value="Run 08_new_analyses.py first"); return

    row=4
    blocks=[
        ("BLOC 1 — QUALITY PREDICTORS",C_QUALITY,[
            "Q2.1_words_quality","Q2.2_E2_quality","Q2.3_E4_quality",
            "emotion_quality_global","emotion_quality_precision",
            "emotion_quality_examples","emotion_quality_relevance","emotion_quality_richness"]),
        ("BLOC 2 — EMOTION → PSYCHOLOGICAL & EVALUATION VARIABLES",C_PSY,[
            "emotion_Perceived Manipulati_1_num","emotion_Perceived Manipulati_2_num",
            "emotion_Competence_1_num","emotion_Sense of Independenc_1_num",
            "emotion_evaluation_2_num","emotion_evaluation_3_num"]),
        ("BLOC 3 — REUSE INTENTION & CHATBOT PREFERENCE",C_EVAL,[
            "Q3.1_multi_E5","Q3.1b_full_multi_E5","Q3.2_quality_E6","Q3.2b_E1_E2_quality_E6"]),
        ("BLOC 4 — PERCEPTION AI → CHATBOT APPRECIATION",C_PSY,[
            "Q4.1_psy_E3"]),
        ("BLOC 5 — BREAKPOINT & CONVERSATION EFFECTS",C_CONV,[
            "Q5.2_bkpt_quality"]),
    ]
    for bloc_title,bloc_color,keys in blocks:
        section_title(ws,bloc_title,bloc_color,8,row); row+=2
        for key in keys:
            item=regression_noton.get(key,{})
            if item:
                row=write_regression_block(ws,item.get("model",{}) if "model" in item else item,row,item.get("label",key))
        row+=1

    # Q5.3 t-test
    q53=regression_noton.get("Q5.3_bkpt_exists_quality",{})
    if q53:
        section_title(ws,"BLOC 6 — BREAKPOINT EXISTS → QUALITY (t-test)",C_CONV,8,row); row+=2
        m=q53.get("model",{})
        ws.cell(row,1,value="Q5.3 — Do participants with no breakpoint have higher quality?").font=Font(bold=True,name="Arial",size=10)
        row+=1
        for j,h in enumerate(["Group","N","Mean","SD","t","p","Sig.","Cohen's d","Effect size"],1):
            hdr(ws.cell(row,j,value=h),bg=C_CONV,sz=9)
        row+=1
        dat(ws.cell(row,1),"No breakpoint (0)",center=False,bold=True)
        dat(ws.cell(row,2),m.get("n_no_breakpoint"))
        dat(ws.cell(row,3),m.get("mean_no_bkpt"),fmt="0.000")
        dat(ws.cell(row,4),m.get("sd_no_bkpt"),fmt="0.000")
        dat(ws.cell(row,5),m.get("t"),fmt="0.0000")
        pval_cell(ws.cell(row,6),m.get("p"))
        dat(ws.cell(row,7),m.get("sig",""),bold=True)
        dat(ws.cell(row,8),m.get("cohens_d"),fmt="0.000")
        dat(ws.cell(row,9),m.get("effect_size",""))
        row+=1
        dat(ws.cell(row,1),"Breakpoint detected (1)",center=False,bold=True,alt=True)
        dat(ws.cell(row,2),m.get("n_breakpoint"),alt=True)
        dat(ws.cell(row,3),m.get("mean_bkpt"),alt=True,fmt="0.000")
        dat(ws.cell(row,4),m.get("sd_bkpt"),alt=True,fmt="0.000")

    for j,w in enumerate([45,8,10,8,10,14,8,10,14],1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="A4"

def build_sheet13(wb, med_noton):
    ws=wb.create_sheet("13_Mediation_Moderation")
    section_title(ws,"Mediations — Variable interactions where Tone is NOT the independent variable",COLOR_HDR,13,1)
    summary_box(ws,"This sheet answers: does engagement mediate the link between AI perception and quality? Does appreciation mediate the link between utility and chatbot preference? Bootstrap=5000, 95% CI.",2,13)

    if not med_noton: ws.cell(3,1,value="Run 08_new_analyses.py first"); return

    row=4
    # Q3.3 single model
    section_title(ws,"Q3.3 — E4 (Utility) → E3 (Appreciation) → E6 (Chatbot preference)",C_EVAL,13,row); row+=2
    q33=med_noton.get("Q3.3_E4_E3_E6",{})
    row=write_mediation_block(ws,q33.get("model",{}),row,"Q3.3 — Utility → Appreciation → Chatbot preference")

    row+=2
    # Q2.6 — 12 mediations (summary table + detail for sig only)
    q26=med_noton.get("Q2.6_psy_E2_quality",{})
    models_q26=q26.get("models",[])
    row=write_mediation_summary_table(ws,models_q26,row,"Q2.6 — Perception AI → E2 (Engagement felt) → Quality global [12 models]")
    row+=1
    # Detail only for significant
    sig_q26=[m for m in models_q26 if m.get("result",{}).get("indirect",{}).get("significant",False)]
    if sig_q26:
        section_title(ws,f"Q2.6 Significant models — detailed paths ({len(sig_q26)} significant)","1A6634",13,row); row+=2
        for item in sig_q26:
            lbl=f"Q2.6: {item.get('iv_label','')} → E2(Engagement) → Quality"
            row=write_mediation_block(ws,item.get("result",{}),row,lbl)
    else:
        ws.cell(row,1,value="No significant mediation found in Q2.6 models.").font=Font(italic=True,name="Arial",size=9,color="922B21")
        row+=2

    row+=2
    # Q4.4 — 12 mediations (summary + sig detail)
    q44=med_noton.get("Q4.4_psy_E3_E6",{})
    models_q44=q44.get("models",[])
    row=write_mediation_summary_table(ws,models_q44,row,"Q4.4 — Perception AI → E3 (Appreciation) → E6 (Chatbot preference) [12 models]")
    row+=1
    sig_q44=[m for m in models_q44 if m.get("result",{}).get("indirect",{}).get("significant",False)]
    if sig_q44:
        section_title(ws,f"Q4.4 Significant models — detailed paths ({len(sig_q44)} significant)","1A6634",13,row); row+=2
        for item in sig_q44:
            lbl=f"Q4.4: {item.get('iv_label','')} → E3(Appreciation) → E6(Preference)"
            row=write_mediation_block(ws,item.get("result",{}),row,lbl)
    else:
        ws.cell(row,1,value="No significant mediation found in Q4.4 models.").font=Font(italic=True,name="Arial",size=9,color="922B21")
        row+=1

    # EMOTION MEDIATIONS
    row+=3
    section_title(ws,"EMOTION MEDIATIONS — Content emotion as IV or mediator",C_ACTION,13,row); row+=2

    em_E2_q=med_noton.get("emotion_E2_quality",{})
    row=write_mediation_block(ws,em_E2_q.get("model",{}),row,
                              "Emotion expressed → E2 (Engagement felt) → Quality global")

    em_E3_E6=med_noton.get("emotion_E3_E6",{})
    row=write_mediation_block(ws,em_E3_E6.get("model",{}),row,
                              "Emotion expressed → E3 (Chatbot appreciation) → E6 (Chatbot preference)")

    for j,w in enumerate([45,22,22,22,6,8,8,10,10,10,10,30],1): ws.column_dimensions[get_column_letter(j)].width=w
    ws.freeze_panes="A4"

def build_sheet14(wb, df_prog, synthesis):
    ws=wb.create_sheet("14_Quality_Progression")
    section_title(ws,"Quality Score Progression per Conversation Turn — FL_21 vs FL_22",COLOR_HDR,9,1)
    summary_box(ws,"Quality score (1-5) rated per turn by GPT-4o-mini. Curves show when engagement declines. Only turns with N≥10 per condition are plotted.",2,9)

    prog=synthesis.get("progression",{}) if synthesis else {}
    if not prog and not df_prog.empty:
        # Rebuild from df_prog
        for _,r in df_prog.iterrows():
            ver=r.get("version",""); t=str(int(r.get("turn",0)))
            if ver not in prog: prog[ver]={}
            if t not in prog[ver]: prog[ver][t]={"scores":[]}
            prog[ver][t].setdefault("scores",[]).append(r.get("quality_score"))
        for ver in prog:
            for t in prog[ver]:
                scores=[s for s in prog[ver][t].get("scores",[]) if s is not None]
                prog[ver][t]={"mean":round(float(np.mean(scores)),3) if scores else None,
                              "sd":round(float(np.std(scores,ddof=1)),3) if len(scores)>1 else None,
                              "n":len(scores)}

    all_turns=sorted(set(int(t) for v in prog.values() for t in v.keys())) if prog else []
    row=4
    for j,h in enumerate(["Turn","Mean FL_21","SD FL_21","N FL_21","Mean FL_22","SD FL_22","N FL_22","Δ(21-22)"],1):
        hdr(ws.cell(row,j,value=h),bg=COLOR_FL21 if "21" in h else COLOR_FL22 if "22" in h else "2C3E50",sz=9)
    row+=1
    for t in all_turns:
        ts=str(t); alt=(t%2==0)
        d21=prog.get("FL_21",{}).get(ts,{}); d22=prog.get("FL_22",{}).get(ts,{})
        m21=d21.get("mean"); m22=d22.get("mean")
        dat(ws.cell(row,1),t,alt=alt,bold=True)
        dat(ws.cell(row,2),m21,alt=alt,fmt="0.000")
        dat(ws.cell(row,3),d21.get("sd"),alt=alt,fmt="0.000")
        dat(ws.cell(row,4),d21.get("n"),alt=alt)
        dat(ws.cell(row,5),m22,alt=alt,fmt="0.000")
        dat(ws.cell(row,6),d22.get("sd"),alt=alt,fmt="0.000")
        dat(ws.cell(row,7),d22.get("n"),alt=alt)
        if m21 and m22: delta_cell(ws.cell(row,8),m21-m22)
        else: ws.cell(row,8,value="n/a").border=THIN
        row+=1

    # Progression figure
    if prog and all_turns:
        fig,ax=plt.subplots(figsize=(12,5))
        for ver,col,lbl in [("FL_21",COLOR_FL21,"FL_21 Friendly"),("FL_22",COLOR_FL22,"FL_22 Professional")]:
            data_v=prog.get(ver,{})
            turns_v=sorted(int(t) for t in data_v.keys() if data_v[str(t) if str(t) in data_v else t].get("n",0)>=8)
            means_v=[data_v.get(str(t),{}).get("mean") for t in turns_v]
            sds_v=[data_v.get(str(t),{}).get("sd") or 0 for t in turns_v]
            valid=[(t,m,s) for t,m,s in zip(turns_v,means_v,sds_v) if m is not None]
            if valid:
                tv,mv,sv=zip(*valid)
                ax.plot(tv,mv,marker="o",color=f"#{col}",linewidth=2,label=lbl)
                ax.fill_between(tv,[m-s for m,s in zip(mv,sv)],[m+s for m,s in zip(mv,sv)],
                                color=f"#{col}",alpha=0.15)
        ax.set_xlabel("Conversation turn"); ax.set_ylabel("Avg quality score (1-5)")
        ax.set_title("Quality score per turn — FL_21 vs FL_22",fontsize=12,fontweight="bold")
        ax.legend(); ax.grid(alpha=0.3); ax.set_ylim(1,5)
        plt.tight_layout()
        fig_path=OUTPUT_DIR/"fig_progression_new.png"
        fig.savefig(fig_path,dpi=150,bbox_inches="tight"); plt.close(fig)
        if fig_path.exists():
            insert_png(ws,fig_path,f"J4",22,9)

    for j,w in enumerate([8,12,10,8,12,10,8,12],1): ws.column_dimensions[get_column_letter(j)].width=w

def build_sheet15(wb, wordfreq):
    ws=wb.create_sheet("15_Word_Frequencies")
    section_title(ws,"Word Frequencies & TF-IDF — Participants and Chatbot responses",COLOR_HDR,8,1)
    if not wordfreq: ws.cell(2,1,value="Run 04_wordfreq.py first"); return

    tw_m21=wordfreq.get("top_words_participants_fl21",[])
    tw_m22=wordfreq.get("top_words_participants_fl22",[])
    tw_r21=wordfreq.get("top_words_chatbot_fl21",[])
    tw_r22=wordfreq.get("top_words_chatbot_fl22",[])
    row=3
    for j,h in enumerate(["Rank","Word FL_21 (Part.)","Freq","Word FL_22 (Part.)","Freq",
                           "Word Chatbot FL_21","Freq","Word Chatbot FL_22","Freq"],1):
        hdr(ws.cell(row,j,value=h),bg=COLOR_FL21 if "21" in h else COLOR_FL22 if "22" in h else "2C3E50",sz=9)
    row+=1
    for i in range(min(40,max(len(tw_m21),len(tw_m22),len(tw_r21),len(tw_r22),1))):
        alt=(i%2==0)
        dat(ws.cell(row,1),i+1,alt=alt,bold=True)
        if i<len(tw_m21): dat(ws.cell(row,2),tw_m21[i][0],center=False,alt=alt); dat(ws.cell(row,3),tw_m21[i][1],alt=alt)
        if i<len(tw_m22): dat(ws.cell(row,4),tw_m22[i][0],center=False,alt=alt); dat(ws.cell(row,5),tw_m22[i][1],alt=alt)
        if i<len(tw_r21): dat(ws.cell(row,6),tw_r21[i][0],center=False,alt=alt); dat(ws.cell(row,7),tw_r21[i][1],alt=alt)
        if i<len(tw_r22): dat(ws.cell(row,8),tw_r22[i][0],center=False,alt=alt); dat(ws.cell(row,9),tw_r22[i][1],alt=alt)
        row+=1

    # TF-IDF
    row+=2
    section_title(ws,"Distinctive words TF-IDF — FL_21 vs FL_22","C0392B",8,row); row+=1
    for j,h in enumerate(["Rank","Distinctive FL_21 (Part.)","Δ TF-IDF","Distinctive FL_22 (Part.)","Δ TF-IDF",
                           "Distinctive Chatbot FL_21","Δ TF-IDF","Distinctive Chatbot FL_22","Δ TF-IDF"],1):
        hdr(ws.cell(row,j,value=h),bg="2C3E50",sz=9)
    row+=1
    for i in range(20):
        alt=(i%2==0)
        dat(ws.cell(row,1),i+1,alt=alt,bold=True)
        for col_i,key in [(2,"tfidf_participants_fl21"),(4,"tfidf_participants_fl22"),
                           (6,"tfidf_chatbot_fl21"),(8,"tfidf_chatbot_fl22")]:
            tw=wordfreq.get(key,[])
            if i<len(tw):
                dat(ws.cell(row,col_i),tw[i][0],center=False,alt=alt)
                dat(ws.cell(row,col_i+1),tw[i][2],alt=alt,fmt="0.0000")
        row+=1

    # Word clouds
    for fname,anchor in [("fig_wc_part_fl21.png","K3"),("fig_wc_part_fl22.png","K18"),
                          ("fig_wc_bot_fl21.png","K33"),("fig_wc_bot_fl22.png","K48")]:
        insert_png(ws,OUTPUT_DIR/fname,anchor,20,8)
    for j,w in enumerate([6,22,10,22,10,22,10,22,10],1): ws.column_dimensions[get_column_letter(j)].width=w

# ================================================================
# MAIN
# ================================================================

def run():
    print("\nBuilding final Excel file (15 sheets)...")
    d=load_all()

    wb=Workbook()

    print("  Sheet 01 — Raw data")
    build_sheet1(wb, d["df_clean"])
    print("  Sheet 02 — FL_21 Friendly")
    build_sheet2(wb, d["df_fl21"])
    print("  Sheet 03 — FL_22 Professional")
    build_sheet3(wb, d["df_fl22"])
    print("  Sheet 04 — Scale coding")
    build_sheet4(wb, d["df_coding"])
    print("  Sheet 05 — Variable dictionary")
    build_sheet5(wb)
    print("  Sheet 06 — Correlations")
    build_sheet6(wb, d["correlations"])
    print("  Sheet 07 — VIF")
    build_sheet7(wb, d["vif"] if isinstance(d["vif"],list) else [])
    print("  Sheet 08 — Chatbot compliance (from synthesis)")
    # Sheet 8 reuses compliance from existing synthesis
    ws8=wb.create_sheet("08_Chatbot_Compliance")
    comp=d["synthesis"].get("compliance",{}) if d["synthesis"] else {}
    if comp:
        section_title(ws8,"Chatbot Compliance — Perceived tone vs requested brief",COLOR_HDR,8,1)
        per_ver=comp.get("per_version",{})
        hdrs8=["Metric","FL_21 (Friendly)","FL_22 (Professional)"]
        for j,h in enumerate(hdrs8,1): hdr(ws8.cell(2,j,value=h),bg="2C3E50" if j==1 else COLOR_FL21 if j==2 else COLOR_FL22,sz=10)
        metrics8=[("N respondents","n"),("Compliance score (mean)","compliance_mean"),("Compliance SD","compliance_sd"),
                  ("% Compliant","pct_compliant"),("% Emojis","pct_emojis"),("% Informal address","pct_informal"),
                  ("% Formal address","pct_formal"),("Avg encouragements","avg_encouragements"),
                  ("Avg sober phrases","avg_sober"),("Avg friendly words","avg_friendly_words"),("Avg formal words","avg_formal_words")]
        for i,(lbl,key) in enumerate(metrics8):
            r=3+i; alt=(i%2==0)
            dat(ws8.cell(r,1),lbl,center=False,alt=alt,bold=True)
            dat(ws8.cell(r,2),per_ver.get("FL_21",{}).get(key),alt=alt,fmt="0.00" if isinstance(per_ver.get("FL_21",{}).get(key),float) else None)
            dat(ws8.cell(r,3),per_ver.get("FL_22",{}).get(key),alt=alt,fmt="0.00" if isinstance(per_ver.get("FL_22",{}).get(key),float) else None)
        tc=comp.get("ttest_compliance",{})
        if tc:
            r=3+len(metrics8)+2
            ws8.cell(r,1,value="T-test compliance score:").font=Font(bold=True,name="Arial",size=10)
            write_ttest_block(ws8,[tc],r+1,"2C3E50")
        for j,w in enumerate([30,20,22],1): ws8.column_dimensions[get_column_letter(j)].width=w
    else:
        ws8.cell(1,1,value="Run 06_synthesis.py first")

    print("  Sheet 09 — Dropout analysis")
    build_sheet9(wb, d["dropout"])
    print("  Sheet 10 — Tone comparisons")
    build_sheet10(wb, d["tone_comparisons"], d["df_ai"])
    print("  Sheet 11 — Tone effects (mediations)")
    build_sheet11(wb, d["mediation_tone"], d["regression_noton"])
    print("  Sheet 12 — Variable relationships (regressions)")
    build_sheet12(wb, d["regression_noton"])
    print("  Sheet 13 — Mediation/Moderation (no tone)")
    build_sheet13(wb, d["mediation_noton"])
    print("  Sheet 14 — Quality progression")
    build_sheet14(wb, d["df_prog"], d["synthesis"])
    print("  Sheet 15 — Word frequencies")
    build_sheet15(wb, d["wordfreq"])

    wb.save(OUTPUT_FILE)
    print(f"\nExcel saved: {OUTPUT_FILE}")
    return str(OUTPUT_FILE)

if __name__ == "__main__":
    run()
