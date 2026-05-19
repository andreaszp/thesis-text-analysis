"""
03_api_analysis.py — Classification IA de chaque conversation via OpenAI API.
Toutes les variables sont définies avec des grilles précises dans le prompt.
Sauvegarde outputs/ai_results.json (checkpoint : reprend si interrompu).
"""
import sys, json, time, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
from openai import OpenAI

from config import OPENAI_API_KEY, OUTPUT_DIR, MSG_COLS, RESP_COLS
from loader import load_clean_data, reconstruct_conversation

# ================================================================
# BRIEFS EXACTS DES DEUX VERSIONS DU CHATBOT
# ================================================================

BRIEF_FL21 = """Version FL_21 - Ton FRIENDLY/AMICAL :
Prompt exact donne au chatbot :
'You should maintain a friendly, warm, and engaging tone throughout the conversation.'
Caracteristiques ATTENDUES pour ce ton :
- Chaleureux, engageant, enthousiaste
- Tutoiement possible et naturel
- Emojis autorises et encourages
- Formules d'encouragement (super !, c'est interessant !, genial)
- Relances personnalisees et chaleureuses
- Registre decontracte mais professionnel"""

BRIEF_FL22 = """Version FL_22 - Ton PROFESSIONNEL/FORMEL :
Prompt exact donne au chatbot :
'You should maintain a professional and respectful tone throughout the conversation.'
Caracteristiques ATTENDUES pour ce ton :
- Neutre, sobre, respectueux
- Vouvoiement systematique
- Aucun emoji
- Formules sobres et directes
- Pas d'exclamations enthousiastes
- Registre formel et structure"""

# ================================================================
# DETECTION FIN DE CONVERSATION
# ================================================================

END_PHRASES = [
    "tu peux cliquer sur la fleche",
    "you can click the arrow",
    "cliquer sur la fleche en bas",
    "click the arrow at the bottom",
    "<end_of_interview>",
    "avez-vous quelque chose a ajouter",
    "as-tu quelque chose a ajouter",
    "have you anything else to add",
    "is there anything else",
    "y a-t-il autre chose",
    "quelque chose a ajouter",
    "anything else to add",
    "anything else you would like to share",
    "quelque chose d'autre a partager",
    "do you have anything else",
    "avez-vous autre chose",
]

def detect_proper_end(row):
    n = row["n_turns"]
    for i in range(max(1, n - 1), n + 2):
        resp = row.get(f"response_{i}")
        if pd.notna(resp):
            txt = str(resp).lower()
            for phrase in END_PHRASES:
                if phrase.lower() in txt:
                    return True
    return False

def classify_last_msg(row):
    n        = row["n_turns"]
    last_msg = str(row.get(f"msg_{n}", "") or "").strip()
    n_words  = len(last_msg.split())
    if row.get("conv_ended_properly", False):
        return "fin_propre"
    elif n <= 3 and n_words <= 5:
        return "abandon_precoce"
    elif n_words <= 3:
        return "reponse_minimale"
    else:
        return "fin_sans_cloture"

# ================================================================
# PROMPT PRINCIPAL
# ================================================================

MAIN_PROMPT = """Tu es un expert senior en recherche UX qualitative et en analyse d'interviews utilisateurs.

Analyse cette conversation entre un chatbot de recherche UX et un participant.
Contexte : etude sur SoundFlow, une application de streaming musical.

BRIEF DU CHATBOT (ce qui lui etait demande) :
{brief}

CONVERSATION COMPLETE :
{conv_text}

---

GRILLES DE SCORING - lis-les attentivement avant d'analyser.

QUALITE DE REPONSE (score 1-5) :

score_global :
  1 = Reponses quasi toutes en 1-3 mots, sans aucun developpement
  2 = Quelques phrases mais peu developpees, peu de contexte
  3 = Developpement correct mais incomplet, peu d'exemples
  4 = Reponses detaillees avec contexte et au moins un exemple
  5 = Reponses riches, precises, nuancees, avec plusieurs exemples concrets

score_precision :
  1 = Tres vague, aucune information exploitable
  2 = Imprecis, generalites sans ancrage
  3 = Quelques elements precis mais insuffisants
  4 = Precis sur la plupart des points
  5 = Tres precis, chaque affirmation ancree dans un fait ou exemple concret

score_exemples :
  1 = Aucun exemple concret dans toute la conversation
  2 = Un exemple vague ou mal developpe
  3 = Un exemple correct
  4 = Plusieurs exemples pertinents
  5 = Exemples multiples, detailles, illustrant differents aspects

score_pertinence :
  1 = Les reponses ne repondent pas aux questions posees
  2 = Repond partiellement, souvent hors sujet
  3 = Repond correctement mais de maniere generique
  4 = Repond directement et precisement aux questions
  5 = Repond parfaitement, anticipe les questions, apporte du contexte non demande

score_richesse :
  1 = Vocabulaire tres limite, phrases simples repetitives
  2 = Peu varie, registre pauvre
  3 = Vocabulaire correct, quelques nuances
  4 = Vocabulaire varie, formulations nuancees
  5 = Vocabulaire riche, registre precis, nuances fines

ACTIONABILITE (score 1-5) :
score_global :
  1 = Rien d'actionnable, que des generalites
  2 = Une vague piste mais non exploitable directement
  3 = Un element concret mais insuffisamment developpe
  4 = Plusieurs elements directement exploitables par une equipe produit
  5 = Feedback tres riche : problemes precis + contexte + suggestions applicables

contient_probleme_concret :
  true = Le participant decrit un probleme precis avec contexte
  false = Mecontentement vague sans description du probleme

contient_conseil_applicable :
  true = Le participant formule une amelioration implementable directement (ex: pouvoir telecharger les playlists hors ligne)
  false = Souhait vague non implementable (ex: ce serait bien que ce soit mieux)

contient_use_case_precis :
  true = Le participant decrit un contexte d'usage precis (quand, ou, comment, pourquoi)
  false = Usage decrit de maniere generique sans contexte

TYPE DE CONTENU :
opinion_personnelle :
  true = Le participant exprime clairement son point de vue personnel
  false = Se limite a des faits ou descriptions sans position personnelle

probleme_concret_vs_vague :
  true = Au moins un probleme decrit avec contexte precis
  false = Insatisfactions exprimees de maniere floue uniquement

suggestion_feature_request :
  true = Le participant propose explicitement une fonctionnalite ou amelioration
  false = Aucune suggestion concrete

emotion_frustration_satisfaction :
  true = Le participant exprime une emotion claire avec contexte
  false = Ton purement factuel, aucune emotion exprimee

comparaison_concurrents :
  true = Le participant mentionne explicitement une autre app ou service (Spotify, Apple Music, Deezer, etc.)
  false = Aucune comparaison externe

PROFIL DU PARTICIPANT :
engagement_global (1-5) :
  1 = Repond en 1-3 mots systematiquement, ne developpe jamais
  2 = Phrases courtes, peu de contexte, developpement rare
  3 = Developpe parfois, reste souvent superficiel
  4 = Developpe souvent, donne du contexte, semble implique
  5 = Developpe systematiquement, donne des exemples spontanement, montre un vrai interet

elaboration :
  court = messages en moyenne moins de 10 mots
  moyen = messages en moyenne 10 a 30 mots
  detaille = messages en moyenne plus de 30 mots

coherence :
  faible = reponses contradictoires ou hors sujet frequents
  moyenne = quelques incoherences ou digressions
  elevee = discours coherent, logique, facile a suivre

expertise_percue (maitrise du sujet dans son feedback) :
  1 = Feedback vague, generalites sans ancrage
  2 = Quelques details mais sans vocabulaire precis ni contexte clair
  3 = Decrit son usage avec contexte, quelques exemples, reste en surface
  4 = Feedback structure, vocabulaire precis, sait identifier ce qui ne va pas
  5 = Feedback expert : identifie des problemes systemiques, propose des solutions precises

PROGRESSION PAR TOUR :
IMPORTANT :
- Un tour = une reponse du participant
- Un oui isole ne suffit PAS a degrader le score si la reponse suivante est riche
- Le point de rupture = moment ou AU MOINS 2 tours CONSECUTIFS montrent une degradation durable
- Si pas de rupture claire : point_de_rupture.existe = false et tour = null

CHATBOT - TON PERCU :
score_amical (1-5) :
  1 = Tres froid, distant, aucune chaleur
  2 = Formel et sobre, peu de chaleur
  3 = Neutre, ni froid ni chaleureux
  4 = Chaleureux, engageant, quelques marques d'enthousiasme
  5 = Tres chaleureux, enthousiaste, formules d'encouragement frequentes, emojis

score_professionnel (1-5) :
  1 = Tres decontracte, presque familier
  2 = Peu formel, registre detendu
  3 = Neutre, ni formel ni decontracte
  4 = Formel, sobre, respectueux
  5 = Tres formel, vouvoiement systematique, aucun emoji, registre soutenu

CONFORMITE AU BRIEF (score 1-5) :
  1 = Completement a l'oppose du ton demande
  2 = Quelques elements conformes mais majorite non conforme
  3 = Partiellement conforme, derive notable
  4 = Globalement conforme, quelques ecarts mineurs
  5 = Parfaitement conforme au ton demande

---

Reponds UNIQUEMENT en JSON valide, sans markdown, sans texte avant ou apres.

{
  "participant": {

    "qualite_globale": {
      "score_global": <1-5>,
      "score_precision": <1-5>,
      "score_exemples": <1-5>,
      "score_pertinence": <1-5>,
      "score_richesse": <1-5>,
      "justification": "<1 phrase factuelle>"
    },

    "actionabilite": {
      "score_global": <1-5>,
      "contient_probleme_concret": <true|false>,
      "contient_conseil_applicable": <true|false>,
      "contient_use_case_precis": <true|false>,
      "justification": "<1 phrase>"
    },

    "type_contenu": {
      "opinion_personnelle": <true|false>,
      "probleme_concret_vs_vague": <true|false>,
      "suggestion_feature_request": <true|false>,
      "emotion_frustration_satisfaction": <true|false>,
      "comparaison_concurrents": <true|false>
    },

    "profil": {
      "engagement_global": <1-5>,
      "elaboration": "<court|moyen|detaille>",
      "coherence": "<faible|moyenne|elevee>",
      "expertise_percue": <1-5>
    },

    "progression_par_tour": [
      {
        "tour": <numero>,
        "score_qualite": <1-5>,
        "nb_mots_approx": <entier>,
        "note": "<1 phrase courte si notable, sinon null>"
      }
    ],

    "point_de_rupture": {
      "existe": <true|false>,
      "tour": <numero ou null>,
      "explication": "<description ou null>"
    },

    "fin_conversation": {
      "a_repondu_jusqu_au_bout": <true|false>,
      "type_fin": "<fin_propre|reponse_minimale|abandon_precoce|fin_sans_cloture>",
      "dernier_message_participant": "<texte exact>"
    },

    "resume": "<2-3 phrases resumant la contribution>",
    "verbatim_cle": "<citation exacte la plus representative, max 120 caracteres>",
    "langue_principale": "<fr|en|mixte>"
  },

  "chatbot": {

    "ton_percu": {
      "score_amical": <1-5>,
      "score_professionnel": <1-5>,
      "label_dominant": "<tres amical|amical|neutre|professionnel|tres professionnel>",
      "justification": "<1 phrase avec exemple concret>"
    },

    "coherence_ton": {
      "score": <1-5>,
      "derive_detectee": <true|false>,
      "description_derive": "<description si derive sinon null>"
    },

    "marqueurs_ton": {
      "mots_amicaux_detectes": ["<liste>"],
      "mots_formels_detectes": ["<liste>"],
      "emojis_utilises": <true|false>,
      "liste_emojis": ["<emojis detectes ou liste vide>"],
      "tutoiement": <true|false>,
      "vouvoiement": <true|false>,
      "formules_encouragement": ["<ex: super!, genial>"],
      "formules_sobres": ["<ex: je comprends, en effet>"]
    },

    "conformite_brief": {
      "score": <1-5>,
      "conforme": <true|false>,
      "ecart_detecte": "<description precise ou null>",
      "justification": "<1 phrase avec exemple>"
    },

    "style_questions": {
      "questions_ouvertes_pct": <0-100>,
      "relances_personnalisees": <true|false>,
      "pression_ressentie": "<aucune|legere|moderee|forte>"
    }
  }
}"""


# ================================================================
# APPEL API
# ================================================================

def call_openai(prompt, client, model="gpt-4o-mini", max_retries=3):
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model,
                max_tokens=2500,
                response_format={"type": "json_object"},
                messages=[
                    {
                        "role": "system",
                        "content": (
                            "Tu es un expert en analyse qualitative UX. "
                            "Reponds uniquement en JSON valide en respectant exactement la structure demandee."
                        )
                    },
                    {"role": "user", "content": prompt}
                ]
            )
            raw = response.choices[0].message.content.strip()
            return json.loads(raw)

        except json.JSONDecodeError as e:
            print(f"    JSON invalide (tentative {attempt + 1}) : {e}")
            if attempt == max_retries - 1:
                return {"error": f"json_decode: {e}", "ok": False}
            time.sleep(2)

        except Exception as e:
            err = str(e)
            if "rate_limit" in err.lower() or "429" in err:
                wait = 30 * (attempt + 1)
                print(f"    Rate limit - attente {wait}s...")
                time.sleep(wait)
            elif attempt == max_retries - 1:
                return {"error": err, "ok": False}
            else:
                time.sleep(5)

    return {"error": "max_retries", "ok": False}


def analyze_conversation(row, client):
    version   = row["FL_13_DO"]
    brief     = BRIEF_FL21 if version == "FL_21" else BRIEF_FL22
    conv_text = reconstruct_conversation(row)[:4000]
    prompt    = MAIN_PROMPT.format(brief=brief, conv_text=conv_text)
    result    = call_openai(prompt, client)

    result["respondent_id"]       = str(row["ResponseId"])
    result["version"]             = version
    result["n_turns"]             = int(row["n_turns"])
    result["conv_ended_properly"] = bool(row.get("conv_ended_properly", False))
    result["last_msg_type"]       = row.get("last_msg_type", "")
    result["ok"]                  = "error" not in result
    return result


def flatten_result(r):
    p    = r.get("participant", {})
    c    = r.get("chatbot", {})
    qg   = p.get("qualite_globale", {})
    act  = p.get("actionabilite", {})
    cnt  = p.get("type_contenu", {})
    prf  = p.get("profil", {})
    rup  = p.get("point_de_rupture", {})
    fin  = p.get("fin_conversation", {})
    ton  = c.get("ton_percu", {})
    coh  = c.get("coherence_ton", {})
    mrk  = c.get("marqueurs_ton", {})
    conf = c.get("conformite_brief", {})
    stq  = c.get("style_questions", {})

    return {
        "respondent_id":          r.get("respondent_id", ""),
        "version":                r.get("version", ""),
        "n_turns":                r.get("n_turns"),
        "conv_ended_properly":    r.get("conv_ended_properly"),
        "last_msg_type":          r.get("last_msg_type", ""),
        "ok":                     r.get("ok", False),
        "qualite_global":         qg.get("score_global"),
        "qualite_precision":      qg.get("score_precision"),
        "qualite_exemples":       qg.get("score_exemples"),
        "qualite_pertinence":     qg.get("score_pertinence"),
        "qualite_richesse":       qg.get("score_richesse"),
        "qualite_justification":  qg.get("justification", ""),
        "action_global":          act.get("score_global"),
        "action_pb_concret":      int(act.get("contient_probleme_concret", False)),
        "action_conseil":         int(act.get("contient_conseil_applicable", False)),
        "action_usecase":         int(act.get("contient_use_case_precis", False)),
        "action_justification":   act.get("justification", ""),
        "contenu_opinion":        int(cnt.get("opinion_personnelle", False)),
        "contenu_pb_concret":     int(cnt.get("probleme_concret_vs_vague", False)),
        "contenu_suggestion":     int(cnt.get("suggestion_feature_request", False)),
        "contenu_emotion":        int(cnt.get("emotion_frustration_satisfaction", False)),
        "contenu_concurrent":     int(cnt.get("comparaison_concurrents", False)),
        "profil_engagement":      prf.get("engagement_global"),
        "profil_elaboration":     prf.get("elaboration", ""),
        "profil_coherence":       prf.get("coherence", ""),
        "profil_expertise":       prf.get("expertise_percue"),
        "rupture_existe":         int(rup.get("existe", False)),
        "rupture_tour":           rup.get("tour"),
        "rupture_explication":    rup.get("explication", ""),
        "fin_jusqu_au_bout":      int(fin.get("a_repondu_jusqu_au_bout", False)),
        "fin_type":               fin.get("type_fin", ""),
        "fin_dernier_msg":        fin.get("dernier_message_participant", ""),
        "resume":                 p.get("resume", ""),
        "verbatim_cle":           p.get("verbatim_cle", ""),
        "langue":                 p.get("langue_principale", ""),
        "bot_score_amical":       ton.get("score_amical"),
        "bot_score_pro":          ton.get("score_professionnel"),
        "bot_ton_label":          ton.get("label_dominant", ""),
        "bot_ton_just":           ton.get("justification", ""),
        "bot_coherence_score":    coh.get("score"),
        "bot_derive":             int(coh.get("derive_detectee", False)),
        "bot_derive_desc":        coh.get("description_derive", ""),
        "bot_mots_amicaux":       " | ".join(mrk.get("mots_amicaux_detectes", [])),
        "bot_mots_formels":       " | ".join(mrk.get("mots_formels_detectes", [])),
        "bot_nb_mots_amicaux":    len(mrk.get("mots_amicaux_detectes", [])),
        "bot_nb_mots_formels":    len(mrk.get("mots_formels_detectes", [])),
        "bot_emojis":             int(mrk.get("emojis_utilises", False)),
        "bot_liste_emojis":       " ".join(mrk.get("liste_emojis", [])),
        "bot_tutoiement":         int(mrk.get("tutoiement", False)),
        "bot_vouvoiement":        int(mrk.get("vouvoiement", False)),
        "bot_nb_encouragements":  len(mrk.get("formules_encouragement", [])),
        "bot_nb_formules_sobres": len(mrk.get("formules_sobres", [])),
        "bot_conformite_score":   conf.get("score"),
        "bot_conforme":           int(conf.get("conforme", False)),
        "bot_ecart":              conf.get("ecart_detecte", ""),
        "bot_conformite_just":    conf.get("justification", ""),
        "bot_pct_ouvertes":       stq.get("questions_ouvertes_pct"),
        "bot_relances":           int(stq.get("relances_personnalisees", False)),
        "bot_pression":           stq.get("pression_ressentie", ""),
    }


# ================================================================
# RUN
# ================================================================

def run():
    if not OPENAI_API_KEY:
        print("ERREUR : OPENAI_API_KEY non trouvee dans .env")
        return {}

    checkpoint = OUTPUT_DIR / "ai_results.json"
    if checkpoint.exists():
        with open(checkpoint, encoding="utf-8") as f:
            done = json.load(f)
        print(f"Checkpoint trouve : {len(done)} conversations deja analysees.")
    else:
        done = {}

    df_clean, _, _ = load_clean_data()
    df_clean["conv_ended_properly"] = df_clean.apply(detect_proper_end, axis=1)
    df_clean["last_msg_type"]       = df_clean.apply(classify_last_msg, axis=1)

    client = OpenAI(api_key=OPENAI_API_KEY)
    total  = len(df_clean)

    print(f"\nAnalyse IA - {total} conversations")
    print(f"Modele      : gpt-4o-mini")
    print(f"Cout estime : ~$0.15-0.25")
    print(f"Temps estime: ~6-10 minutes\n")

    for i, (_, row) in enumerate(df_clean.iterrows()):
        rid = str(row["ResponseId"])
        if rid in done:
            continue
        print(f"  [{i+1}/{total}] {rid[:20]}... FL={row['FL_13_DO']} {row['n_turns']} tours", end=" ")
        result = analyze_conversation(row, client)
        print("OK" if result.get("ok") else f"ERREUR {result.get('error','')[:40]}")
        done[rid] = result
        if (i + 1) % 5 == 0 or i == total - 1:
            with open(checkpoint, "w", encoding="utf-8") as f:
                json.dump(done, f, ensure_ascii=False, indent=2)
        time.sleep(0.8)

    n_ok = sum(1 for v in done.values() if v.get("ok"))
    print(f"\nTermine - {n_ok}/{len(done)} succes")

    rows_flat = [flatten_result(v) for v in done.values() if v.get("ok")]
    df_ai = pd.DataFrame(rows_flat)
    df_ai.to_json(OUTPUT_DIR / "df_ai.json", orient="records", force_ascii=False)

    prog_rows = []
    for rid, v in done.items():
        if not v.get("ok"):
            continue
        for t in v.get("participant", {}).get("progression_par_tour", []):
            prog_rows.append({
                "respondent_id": rid,
                "version":       v.get("version", ""),
                "tour":          t.get("tour"),
                "score_qualite": t.get("score_qualite"),
                "nb_mots":       t.get("nb_mots_approx"),
                "note":          t.get("note", ""),
            })
    pd.DataFrame(prog_rows).to_json(
        OUTPUT_DIR / "df_progression.json", orient="records", force_ascii=False
    )

    print(f"Sauvegardes : df_ai.json ({len(df_ai)} lignes), df_progression.json ({len(prog_rows)} lignes)")
    return done


if __name__ == "__main__":
    run()
