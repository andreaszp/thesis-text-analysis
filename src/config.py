"""
config.py
Shared configuration: paths, constants, stopwords.
"""
import os
from pathlib import Path
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
if not OPENAI_API_KEY:
    print("WARNING: OPENAI_API_KEY not found. Create a .env file (see .env.example).")

DATA_DIR   = ROOT / "data"
OUTPUT_DIR = ROOT / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)

EXCEL_FILE  = DATA_DIR / "Master_Thesis_May+12,+2026_02.02.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "Master_Thesis_Full_Analysis.xlsx"
SHEET_NAME  = "Sheet0"
MAX_TURNS   = 20
MSG_COLS    = [f"msg_{i}"      for i in range(1, MAX_TURNS + 1)]
RESP_COLS   = [f"response_{i}" for i in range(1, MAX_TURNS + 1)]

# ── Likert scale mappings ─────────────────────────────────
LIKERT_7 = {
    "Strongly disagree": 1, "Disagree": 2, "Somewhat disagree": 3,
    "Neither agree nor disagree": 4, "Somewhat agree": 5,
    "Agree": 6, "Strongly agree": 7,
}
CAPABLE_7 = {
    "Not at all capable": 1, "Very slightly capable": 2, "Slightly capable": 3,
    "Neither capable nor incapable": 4, "Somewhat capable": 5,
    "Capable": 6, "Very capable": 7,
}
AMOUNT_7 = {
    "Not at all": 1, "Very little": 2, "A little": 3,
    "A moderate amount": 4, "Quite a bit": 5, "A lot": 6, "A great deal": 7,
}

COL_SCALE = {
    "evaluation_1": LIKERT_7, "evaluation_2": LIKERT_7, "evaluation_3": LIKERT_7,
    "evaluation_4": LIKERT_7, "evaluation_5": LIKERT_7, "evaluation_6": LIKERT_7,
    "Perceived Manipulati_1": LIKERT_7, "Perceived Manipulati_2": LIKERT_7,
    "Perceived Manipulati_3": LIKERT_7, "Perceived Manipulati_4": LIKERT_7,
    "Competence_1": CAPABLE_7, "Competence_2": CAPABLE_7,
    "Moral Responsibility_1": AMOUNT_7, "Moral Responsibility_2": AMOUNT_7,
    "Moral Responsibility_3": AMOUNT_7, "Moral Responsibility_4": AMOUNT_7,
    "Sense of Independenc_1": AMOUNT_7, "Sense of Independenc_2": AMOUNT_7,
    "personnality - Manip_1": LIKERT_7, "personnality - Manip_2": LIKERT_7,
    "personnality - Manip_3": LIKERT_7, "personnality - Manip_4": LIKERT_7,
    "personnality - Manip_5": LIKERT_7,
}
ALL_Q_COLS   = list(COL_SCALE.keys())
NUM_COLS     = [c + "_num" for c in ALL_Q_COLS]
EXCEL_LETTERS = [
    "T","U","V","W","X","Y","Z","AA","AB","AC",
    "AD","AE","AF","AG","AH","AI","AJ","AK",
    "AL","AM","AN","AO","AP"
]
Q_LABELS = [
    "Required effort","Engagement","Chatbot appreciation","Conversation utility",
    "Reuse intention","Chatbot preference",
    "Freedom threat","Decision override","Manipulation","Pressure",
    "Judgment capability (skills)","Judgment capability (morality)",
    "Moral harm (AI→human)","AI moral responsibility",
    "Moral harm (human→AI)","AI moral concern",
    "AI plans & goals","AI self-control",
    "Friendly","Professional","Accessible","Warm","Formal",
]
CONSTRUCTS = {
    "General Evaluation":   ["evaluation_1","evaluation_2","evaluation_3","evaluation_4","evaluation_5","evaluation_6"],
    "Perceived Manipulation":["Perceived Manipulati_1","Perceived Manipulati_2","Perceived Manipulati_3","Perceived Manipulati_4"],
    "AI Competence":        ["Competence_1","Competence_2"],
    "Moral Responsibility": ["Moral Responsibility_1","Moral Responsibility_2","Moral Responsibility_3","Moral Responsibility_4"],
    "AI Independence":      ["Sense of Independenc_1","Sense of Independenc_2"],
    "Chatbot Personality":  ["personnality - Manip_1","personnality - Manip_2","personnality - Manip_3","personnality - Manip_4","personnality - Manip_5"],
}

# ── Custom stopwords (bilingual) ──────────────────────────
STOPWORDS_CUSTOM = {
    "merci","thank","thanks","thankyou","sharing","shared","partager","partage",
    "commentaires","suggestions","apprécier","appréciez","apprécies","apprécie",
    "savoir","voir","faire","donner","pouvoir","permet","utilise","utilisez",
    "utilises","ajouter","connais","sais","peut","think","know","use","using",
    "used","make","makes","making","want","wanted","get","got","give","say",
    "said","come","go","going","would","could","like","est-ce","qu'est-ce",
    "qu'il","a-t-il","avez-vous","pouvez-vous","aimerais","aimeriez","lorsque",
    "est","ce","que","qui","quand","comme","avec","aussi","très","bien",
    "avoir","être","fait","tout","même","encore","toujours","jamais","déjà",
    "vraiment","mais","plus","pas","pour","dans","sur","par","cela","ça","ca",
    "chose","quelque","autre","aspects","exemple","etc","parfois","souvent",
    "tous","jours","temps","time","rien","sans","something","someone","anything",
    "nothing","everything","much","many","every","really","thing","things",
    "just","maybe","sure","okay","ok","yeah","well","actually","basically",
    "kind","though","bit","lot","one","also","musique","musiques","song","songs",
    "playlist","playlists","écouter","listen","listening","artiste","artistes",
    "artist","artists","chanson","chansons","podcast","podcasts","service",
    "services","plateforme","platform","soundflow","application","app","streaming",
    "nan","yes","no","oui","non","nope","je","tu","il","elle","nous","vous",
    "ils","elles","le","la","les","un","une","des","du","de","en","au","aux",
    "et","ou","ni","car","donc","or","the","this","that","these","those","and",
    "but","for","with","from","have","has","are","was","were","been","being",
    "its","it","they","their","them","our","your","his","her","you",
}

# ── Excel style colours ───────────────────────────────────
COLOR_FL21  = "1A6634"
COLOR_FL22  = "7B3F00"
COLOR_HDR   = "1F4E79"
COLOR_ALT   = "EBF5FB"
SCORE_COLORS = {
    1:("FF4444","FFFFFF"), 2:("FF8C00","FFFFFF"), 3:("FFD700","000000"),
    4:("F2F3F4","000000"), 5:("A8D8A8","000000"), 6:("4CAF50","FFFFFF"),
    7:("1A6634","FFFFFF"),
}
