"""
03_api_analysis.py — Classification IA de chaque conversation via OpenAI API.
La cle API est lue depuis .env — jamais en dur dans le code.
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
# PROMPTS
# ================================================================

# Brief exact des deux versions du chatbot
BRIEF_FL21 = """Version FL_21 — Ton FRIENDLY/AMICAL :
'You should maintain a friendly, warm, and engaging tone throughout the conversation.
Goal: gather rich, detailed, and experience-based insights while keeping the interaction
concise and pleasant (3-4 minutes max).'
Caractéristiques attendues : chaleureux, engageant, émojis possibles, vouvoiement ou
tutoiement décontracté, encouragements, reformulations enthousiastes."""

BRIEF_FL22 = """Version FL_22 — Ton PROFESSIONNEL/FORMEL :
'You should maintain a professional and respectful tone throughout the conversation.
Goal: gather rich, detailed, and experience-based insights while keeping the interaction
concise and pleasant (3-4 minutes max).'
Caractéristiques attendues : neutre, formel, respectueux, vouvoiement soutenu,
pas d'émojis, questions directes et sobres."""

# Phrases de fin de conversation détectables
END_PHRASES = [
    # Phrases exactes du système
    "tu peux cliquer sur la flèche",
    "you can click the arrow",
    "cliquer sur la flèche en bas",
    "click the arrow at the bottom",
    "<END_OF_INTERVIEW>",
    # Formules de clôture génériques
    "avez-vous quelque chose à ajouter",
    "have you anything else to add",
    "is there anything else",
    "y a-t-il autre chose",
    "quelque chose à ajouter",
    "anything else to add",
    "anything else you would like to share",
    "quelque chose d'autre à partager",
]

def detect_proper_end(row):
    """
    Retourne True si la conversation s'est terminée proprement.
    Vérifie les 2 dernières réponses du chatbot.
    """
    n = row["n_turns"]
    for i in range(max(1, n - 1), n + 2):
        resp = row.get(f"response_{i}")
        if pd.notna(resp):
            txt = str(resp).lower()
            for phrase in END_PHRASES:
                if phrase.lower() in txt:
                    return True
    return False

def classify_last_participant_msg(row):
    """
    Catégorise comment le participant a terminé.
    - fin_propre        : la conversation a atteint la clôture naturelle
    - reponse_minimale  : dernier message très court (1-3 mots) MAIS conversation déjà longue
    - abandon_precoce   : arrêt avant la fin avec peu de tours
    """
    n = row["n_turns"]
    last_msg = row.get(f"msg_{n}", "")
    if pd.isna(last_msg):
        last_msg = ""
    last_msg = str(last_msg).strip()
    n_words = len(last_msg.split())

    if row.get("conv_ended_properly", False):
        return "fin_propre"
    elif n <= 3 and n_words <= 5:
        return "abandon_precoce"
    elif n_words <= 3:
        return "reponse_minimale"
    else:
        return "fin_sans_cloture"

# ================================================================
# PROMPT PRINCIPAL — analyse participant + chatbot en 1 seul appel
# ================================================================

MAIN_PROMPT = """Tu es un expert senior en recherche UX qualitative et analyse d'interviews.

Analyse cette conversation entre un chatbot de recherche UX et un participant.
Contexte : étude sur SoundFlow, une application de streaming musical.

BRIEF DU CHATBOT (ce qui lui était demandé) :
{brief}

CONVERSATION COMPLÈTE :
{conv_text}

---

Tu dois produire une analyse structurée en JSON.
Réponds UNIQUEMENT en JSON valide, sans markdown, sans commentaire, sans texte avant ou après.

Critères de scoring pour la QUALITÉ des réponses participant (échelle 1-5) :
  1 = Réponse d'un seul mot / "oui" / "non" / vide de sens
  2 = Phrase courte sans détail ni contexte
  3 = Réponse correcte mais peu développée
  4 = Réponse détaillée avec contexte ou exemple
  5 = Réponse riche, précise, avec exemple concret et nuance

IMPORTANT pour la progression :
- Ne classe pas un "oui" isolé comme dégradation si la réponse suivante est riche
- Le point de rupture = moment où AU MOINS 2 tours consécutifs montrent une dégradation soutenue
- Si pas de dégradation claire, point_de_rupture = null

Structure JSON exacte à retourner :

{{
  "participant": {{

    "qualite_globale": {{
      "score_global": <1-5>,
      "score_precision": <1-5>,
      "score_exemples": <1-5>,
      "score_pertinence": <1-5>,
      "score_richesse": <1-5>,
      "justification": "<1 phrase factuelle>"
    }},

    "actionabilite": {{
      "score_global": <1-5>,
      "contient_probleme_concret": <true|false>,
      "contient_conseil_applicable": <true|false>,
      "contient_use_case_precis": <true|false>,
      "justification": "<1 phrase>"
    }},

    "type_contenu": {{
      "opinion_personnelle": <true|false>,
      "probleme_concret_vs_vague": <true|false>,
      "suggestion_feature_request": <true|false>,
      "emotion_frustration_satisfaction": <true|false>,
      "comparaison_concurrents": <true|false>
    }},

    "profil": {{
      "engagement_global": <1-5>,
      "elaboration": "<court|moyen|détaillé>",
      "coherence": "<faible|moyenne|élevée>",
      "expertise_percue": "<novice|intermédiaire|expert>"
    }},

    "progression_par_tour": [
      {{
        "tour": <numéro du tour>,
        "score_qualite": <1-5>,
        "nb_mots_approx": <nombre de mots estimé>,
        "note": "<très court : 1 phrase max si notable, sinon laisser vide>"
      }}
    ],

    "point_de_rupture": {{
      "existe": <true|false>,
      "tour": <numéro du tour où ça se dégrade durablement, ou null si pas de rupture>,
      "explication": "<1 phrase ou null>"
    }},

    "fin_conversation": {{
      "a_repondu_jusqu_au_bout": <true|false>,
      "type_fin": "<fin_propre|reponse_minimale|abandon_precoce|fin_sans_cloture>",
      "dernier_message_participant": "<texte exact du dernier message>"
    }},

    "resume": "<résumé de la contribution du participant en 1-2 phrases>",
    "verbatim_cle": "<citation exacte la plus représentative, max 120 caractères>",
    "langue_principale": "<fr|en|mixte>"

  }},

  "chatbot": {{

    "ton_percu": {{
      "score_amical": <1-5, où 1=très formel/froid et 5=très amical/chaleureux>,
      "score_professionnel": <1-5, où 1=très décontracté et 5=très formel>,
      "label_dominant": "<très amical|amical|neutre|professionnel|très professionnel>",
      "justification": "<1 phrase>"
    }},

    "coherence_ton": {{
      "score": <1-5>,
      "derive_detectee": <true|false>,
      "description_derive": "<si dérive : à quel moment et vers quoi, sinon null>"
    }},

    "marqueurs_ton": {{
      "mots_amicaux_detectes": ["<mot ou expression>"],
      "mots_formels_detectes": ["<mot ou expression>"],
      "emojis_utilises": <true|false>,
      "tutoiement": <true|false>,
      "vouvoiement": <true|false>
    }},

    "conformite_brief": {{
      "score": <1-5, où 5=parfaitement conforme au brief>,
      "conforme": <true|false>,
      "ecart_detecte": "<description de l'écart si non conforme, sinon null>",
      "justification": "<1 phrase>"
    }},

    "style_questions": {{
      "questions_ouvertes_pct": <0-100>,
      "relances_personnalisees": <true|false>,
      "pression_ressentie": "<aucune|légère|modérée|forte>"
    }}

  }}
}}"""


# ================================================================
# FONCTIONS
# ================================================================

def call_openai(prompt, client, model="gpt-4o-mini", max_retries=3):
    """Appel API OpenAI avec retry automatique."""
    for attempt in range(max_retries):
        try:
            response = client.chat.completions.create(
                model=model,
                max_tokens=2000,
                response_format={"type": "json_object"},
                messages=[
                    {
                        "role": "system",
                        "content": "Tu es un expert en analyse qualitative UX. Réponds uniquement en JSON valide."
                    },
                    {
                        "role": "user",
                        "content": prompt
                    }
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
                print(f"    Rate limit — attente {wait}s...")
                time.sleep(wait)
            elif attempt == max_retries - 1:
                return {"error": err, "ok": False}
            else:
                time.sleep(5)

    return {"error": "max_retries", "ok": False}


def analyze_conversation(row, client):
    """
    Analyse complète d'une conversation :
    participant (qualité, actionabilité, type contenu, profil, progression)
    + chatbot (ton perçu, cohérence, conformité brief)
    """
    version = row["FL_13_DO"]
    brief   = BRIEF_FL21 if version == "FL_21" else BRIEF_FL22
    conv    = reconstruct_conversation(row)

    # Limite à 4000 caractères pour rester dans les tokens raisonnables
    conv_short = conv[:4000]

    prompt = MAIN_PROMPT.format(brief=brief, conv_text=conv_short)
    result = call_openai(prompt, client)

    # Ajoute les métadonnées
    result["respondent_id"]       = str(row["ResponseId"])
    result["version"]             = version
    result["n_turns"]             = int(row["n_turns"])
    result["conv_ended_properly"] = bool(row.get("conv_ended_properly", False))
    result["last_msg_type"]       = row.get("last_msg_type", "")
    result["ok"]                  = "error" not in result

    return result


def flatten_result(r):
    """
    Aplatit le JSON imbriqué en une ligne de DataFrame.
    Une ligne = un participant.
    """
    p   = r.get("participant", {})
    c   = r.get("chatbot", {})
    qg  = p.get("qualite_globale", {})
    act = p.get("actionabilite", {})
    cnt = p.get("type_contenu", {})
    prf = p.get("profil", {})
    rup = p.get("point_de_rupture", {})
    fin = p.get("fin_conversation", {})
    ton = c.get("ton_percu", {})
    coh = c.get("coherence_ton", {})
    mrk = c.get("marqueurs_ton", {})
    conf= c.get("conformite_brief", {})
    stq = c.get("style_questions", {})

    return {
        # Identification
        "respondent_id":         r.get("respondent_id", ""),
        "version":               r.get("version", ""),
        "n_turns":               r.get("n_turns"),
        "conv_ended_properly":   r.get("conv_ended_properly"),
        "last_msg_type":         r.get("last_msg_type", ""),
        "ok":                    r.get("ok", False),

        # Qualité globale
        "qualite_score_global":   qg.get("score_global"),
        "qualite_score_precision":qg.get("score_precision"),
        "qualite_score_exemples": qg.get("score_exemples"),
        "qualite_score_pertinence":qg.get("score_pertinence"),
        "qualite_score_richesse": qg.get("score_richesse"),
        "qualite_justification":  qg.get("justification", ""),

        # Actionabilité
        "action_score_global":    act.get("score_global"),
        "action_pb_concret":      act.get("contient_probleme_concret"),
        "action_conseil":         act.get("contient_conseil_applicable"),
        "action_usecase":         act.get("contient_use_case_precis"),
        "action_justification":   act.get("justification", ""),

        # Type de contenu (binaire)
        "contenu_opinion":        cnt.get("opinion_personnelle"),
        "contenu_probleme_concret":cnt.get("probleme_concret_vs_vague"),
        "contenu_suggestion":     cnt.get("suggestion_feature_request"),
        "contenu_emotion":        cnt.get("emotion_frustration_satisfaction"),
        "contenu_concurrent":     cnt.get("comparaison_concurrents"),

        # Profil participant
        "profil_engagement":      prf.get("engagement_global"),
        "profil_elaboration":     prf.get("elaboration", ""),
        "profil_coherence":       prf.get("coherence", ""),
        "profil_expertise":       prf.get("expertise_percue", ""),

        # Point de rupture
        "rupture_existe":         rup.get("existe"),
        "rupture_tour":           rup.get("tour"),
        "rupture_explication":    rup.get("explication", ""),

        # Fin de conversation
        "fin_jusqu_au_bout":      fin.get("a_repondu_jusqu_au_bout"),
        "fin_type":               fin.get("type_fin", ""),
        "fin_dernier_msg":        fin.get("dernier_message_participant", ""),

        # Synthèse participant
        "resume":                 p.get("resume", ""),
        "verbatim_cle":           p.get("verbatim_cle", ""),
        "langue":                 p.get("langue_principale", ""),

        # Chatbot — ton perçu
        "bot_score_amical":       ton.get("score_amical"),
        "bot_score_pro":          ton.get("score_professionnel"),
        "bot_ton_label":          ton.get("label_dominant", ""),
        "bot_ton_justification":  ton.get("justification", ""),

        # Chatbot — cohérence du ton
        "bot_coherence_score":    coh.get("score"),
        "bot_derive":             coh.get("derive_detectee"),
        "bot_derive_description": coh.get("description_derive", ""),

        # Chatbot — marqueurs
        "bot_mots_amicaux":       " | ".join(mrk.get("mots_amicaux_detectes", [])),
        "bot_mots_formels":       " | ".join(mrk.get("mots_formels_detectes", [])),
        "bot_emojis":             mrk.get("emojis_utilises"),
        "bot_tutoiement":         mrk.get("tutoiement"),
        "bot_vouvoiement":        mrk.get("vouvoiement"),

        # Chatbot — conformité au brief
        "bot_conformite_score":   conf.get("score"),
        "bot_conforme":           conf.get("conforme"),
        "bot_ecart":              conf.get("ecart_detecte", ""),
        "bot_conformite_just":    conf.get("justification", ""),

        # Chatbot — style questions
        "bot_pct_ouvertes":       stq.get("questions_ouvertes_pct"),
        "bot_relances":           stq.get("relances_personnalisees"),
        "bot_pression":           stq.get("pression_ressentie", ""),
    }


# ================================================================
# RUN PRINCIPAL
# ================================================================

def run():
    if not OPENAI_API_KEY:
        print("ERREUR : OPENAI_API_KEY non trouvée dans .env")
        print("Créez un fichier .env avec : OPENAI_API_KEY=sk-proj-...")
        return {}

    # Checkpoint : reprend si interrompu
    checkpoint = OUTPUT_DIR / "ai_results.json"
    if checkpoint.exists():
        with open(checkpoint, encoding="utf-8") as f:
            done = json.load(f)
        print(f"Checkpoint trouvé : {len(done)} conversations déjà analysées.")
    else:
        done = {}

    # Chargement données
    df_clean, _, _ = load_clean_data()

    # Détection fin de conversation
    df_clean["conv_ended_properly"] = df_clean.apply(detect_proper_end, axis=1)
    df_clean["last_msg_type"]       = df_clean.apply(classify_last_participant_msg, axis=1)

    client = OpenAI(api_key=OPENAI_API_KEY)
    total  = len(df_clean)

    print(f"\nAnalyse IA — {total} conversations")
    print(f"Modèle        : gpt-4o-mini")
    print(f"Coût estimé   : ~$0.10-0.20 pour {total} conversations")
    print(f"Temps estimé  : ~5-8 minutes\n")

    for i, (_, row) in enumerate(df_clean.iterrows()):
        rid = str(row["ResponseId"])

        if rid in done:
            continue  # déjà traité, on passe

        print(f"  [{i+1}/{total}] {rid[:25]}... (FL={row['FL_13_DO']}, {row['n_turns']} tours)", end=" ")

        result = analyze_conversation(row, client)

        if result.get("ok"):
            print("✓")
        else:
            print(f"✗ {result.get('error','?')[:50]}")

        done[rid] = result

        # Checkpoint toutes les 5 conversations
        if (i + 1) % 5 == 0 or i == total - 1:
            with open(checkpoint, "w", encoding="utf-8") as f:
                json.dump(done, f, ensure_ascii=False, indent=2)

        time.sleep(0.8)  # respecte le rate limit

    # Rapport final
    n_ok  = sum(1 for v in done.values() if v.get("ok"))
    n_err = len(done) - n_ok
    print(f"\nClassification terminée")
    print(f"  Succès  : {n_ok}/{len(done)}")
    print(f"  Erreurs : {n_err}")

    # Sauvegarde DataFrame aplati pour 04_export.py
    rows = [flatten_result(v) for v in done.values() if v.get("ok")]
    df_ai = pd.DataFrame(rows)
    df_ai.to_json(OUTPUT_DIR / "df_ai.json", orient="records", force_ascii=False)

    # Sauvegarde progression par tour séparément
    progression_rows = []
    for rid, v in done.items():
        if not v.get("ok"):
            continue
        version = v.get("version", "")
        prog    = v.get("participant", {}).get("progression_par_tour", [])
        for t in prog:
            progression_rows.append({
                "respondent_id": rid,
                "version":       version,
                "tour":          t.get("tour"),
                "score_qualite": t.get("score_qualite"),
                "nb_mots":       t.get("nb_mots_approx"),
                "note":          t.get("note", ""),
            })
    df_prog = pd.DataFrame(progression_rows)
    df_prog.to_json(OUTPUT_DIR / "df_progression.json", orient="records", force_ascii=False)

    print(f"\nFichiers sauvegardés :")
    print(f"  {checkpoint}")
    print(f"  {OUTPUT_DIR / 'df_ai.json'}")
    print(f"  {OUTPUT_DIR / 'df_progression.json'}")

    return done


if __name__ == "__main__":
    run()
