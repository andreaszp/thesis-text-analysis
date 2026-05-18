"""
config.py — Configuration partagée entre tous les scripts.
Chargement securise de la cle API + stopwords + constantes.
"""
import os
import re
from pathlib import Path
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent.parent
load_dotenv(ROOT / ".env")

OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
if not OPENAI_API_KEY:
    print("WARNING: Aucune cle API. Creez .env a la racine (voir .env.example).")

DATA_DIR    = ROOT / "data"
OUTPUT_DIR  = ROOT / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)
EXCEL_FILE  = DATA_DIR / "Master_Thesis_April_7__2026_04_28.xlsx"
OUTPUT_FILE = OUTPUT_DIR / "Master_Thesis_Text_Analysis.xlsx"
SHEET_NAME  = "Sheet0"
MAX_TURNS   = 20
MSG_COLS    = [f"msg_{i}"      for i in range(1, MAX_TURNS + 1)]
RESP_COLS   = [f"response_{i}" for i in range(1, MAX_TURNS + 1)]

STOPWORDS_CUSTOM = {
    # Formules de politesse
    "merci","thank","thanks","thankyou","sharing","shared",
    "partager","partage","commentaires","suggestions",
    "apprécier","appréciez","apprécies","apprécie",
    # Verbes generiques
    "savoir","voir","faire","donner","pouvoir","permet",
    "utilise","utilisez","utilises","utilisé","utilise",
    "ajouter","connais","sais","peut","think","know",
    "use","using","used","make","makes","making",
    "want","wanted","get","got","give","say","said",
    "come","go","going","would","could","like",
    # Mots interrogatifs chatbot
    "est-ce","qu'est-ce","qu'il","a-t-il","avez-vous",
    "pouvez-vous","aimerais","aimeriez","lorsque",
    "est","ce","que","qui","quand","comme","avec",
    "aussi","très","bien","avoir","être","fait",
    "tout","même","encore","toujours","jamais","déjà",
    "vraiment","mais","plus","pas","pour","dans","sur",
    "par","cela","ça","ca",
    # Mots vagues
    "chose","quelque","autre","aspects","exemple","etc",
    "parfois","souvent","tous","jours","temps","time",
    "rien","sans","something","someone","anything","nothing",
    "everything","much","many","every","really","thing",
    "things","just","maybe","sure","okay","ok","yeah",
    "well","actually","basically","kind","though","bit",
    "lot","one","also",
    # Contexte musical omniprésent
    "musique","musiques","song","songs","playlist","playlists",
    "écouter","listen","listening","artiste","artistes",
    "artist","artists","chanson","chansons","podcast","podcasts",
    "service","services","plateforme","platform","soundflow",
    "application","app","streaming","spotify",
    # Mots de remplissage
    "nan","yes","no","oui","non","nope",
    "je","tu","il","elle","nous","vous","ils","elles",
    "le","la","les","un","une","des","du","de","en",
    "au","aux","et","ou","ni","car","donc","or",
    "the","this","that","these","those","and","but",
    "for","with","from","have","has","are","was",
    "were","been","being","its","it","they","their",
    "them","our","your","his","her","you",
}

COLOR_FL21 = "#1A6634"
COLOR_FL22 = "#7B3F00"
COLOR_HDR  = "#1F4E79"
COLOR_ALT  = "#EBF5FB"
PALETTE    = [COLOR_FL21, COLOR_FL22]

