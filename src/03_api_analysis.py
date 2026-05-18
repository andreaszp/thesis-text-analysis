
"""
03_api_analysis.py — Classification IA de chaque conversation via Claude API.
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

# ── Prompt système (analyse des participants) ─────────────
PROMPT_PARTICIPANT = """Tu es un expert senior en recherche UX qualitative et en analyse d'interviews utilisateurs.

Analyse cette conversation entre un chatbot de recherche et un participant.
Le chatbot recueille des retours sur SoundFlow, une application de streaming musical.

CONVERSATION :
{conv_text}

---
Objectif : évaluer la QUALITÉ et la RICHESSE des réponses du PARTICIPANT uniquement
(ignore le style du chatbot).

Réponds UNIQUEMENT en JSON valide, sans markdown, sans commentaire.
Structure exacte :

{{
  "qualite_reponse": {{
    "score": <1 à 5>,
    "label": "<très faible|faible|moyenne|bonne|excellente>",
    "justification": "<1 phrase factuelle>"
  }},
  "actionabilite": {{
    "score": <1 à 5>,
    "contient_probleme_concret": <true|false>,
    "contient_conseil_applicable": <true|false>,
    "contient_use_case_précis": <true|false>,
    "justification": "<1 phrase>"
  }},
  "profil_participant": {{
    "engagement": "<passif|neutre|actif|très actif>",
    "elaboration": "<court|moyen|détaillé>",
    "coherence": "<faible|moyenne|élevée>",
    "expertise_perçue": "<novice|intermédiaire|expert>"
  }},
  "contenu": {{
    "opinion_personnelle":     <true|false>,
    "frustration_exprimee":   <true|false>,
    "suggestion_concrete":    <true|false>,
    "experience_vecue":       <true|false>,
    "comparaison_concurrents":<true|false>,
    "feature_request":        <true|false>
  }},
  "themes": ["<theme1>", "<theme2>"],
  "sentiment_dominant": "<négatif|neutre|positif|mitigé>",
  "langue": "<fr|en|mixte>",
  "abandon_premature": <true|false>,
  "resume": "<1 phrase max, en français>",
  "verbatim_cle": "<citation exacte du participant, max 100 caractères>"
}}"""

# ── Prompt pour analyse du chatbot ───────────────────────
PROMPT_CHATBOT = """Tu es un expert en analyse de style de communication conversationnelle.

Analyse les RÉPONSES DU CHATBOT dans cette conversation.
Note : le chatbot existe en deux versions — Friendly (FL_21, ton chaleureux) et Pro (FL_22, ton professionnel).
Version de ce chatbot : {version}

CONVERSATION :
{conv_text}

---
Objectif : caractériser le STYLE DE COMMUNICATION du chatbot.

Réponds UNIQUEMENT en JSON valide, sans markdown.

{{
  "ton_perçu": {{
    "score_chaleur": <1 à 5>,
    "score_formalisme": <1 à 5>,
    "score_empathie": <1 à 5>,
    "label_dominant": "<très formel|formel|neutre|chaleureux|très chaleureux>"
  }},
  "style_questions": {{
    "ouvertes_pct": <0 à 100>,
    "relances_personnalisees": <true|false>,
    "formules_encouragement": <true|false>,
    "pression_ressentie": "<aucune|légère|modérée|forte>"
  }},
  "marqueurs_ton": {{
    "friendly_markers": ["<mot ou expression>"],
    "pro_markers": ["<mot ou expression>"]
  }},
  "coherence_ton": "<faible|moyenne|élevée>",
  "resume_style": "<1 phrase décrivant le style du chatbot, en français>"
}}"""

def classify_one(conv_text, prompt_template, client, **kwargs):
    prompt = prompt_template.format(conv_text=conv_text[:3500], **kwargs)
    for attempt in range(3):
        try:
            resp = client.chat.completions.create(
                model="gpt-4o",       # ou "gpt-4o-mini" pour moins cher
                max_tokens=1000,
                response_format={"type": "json_object"},  # force le JSON
                messages=[{"role": "user", "content": prompt}]
            )
            raw = resp.choices[0].message.content.strip()
            return json.loads(raw)
        except json.JSONDecodeError:
            if attempt == 2: return {"error": "json_decode", "classification_ok": False}
            time.sleep(2)
        except Exception as e:
            err = str(e)
            if "rate_limit" in err.lower():
                print("    Rate limit — attente 30s...")
                time.sleep(30)
            elif attempt == 2:
                return {"error": err, "classification_ok": False}
            else:
                time.sleep(5)
    return {"error": "max_retries", "classification_ok": False}

def run():
    if not OPENAI_API_KEY:
      print("ERREUR : OPENAI_API_KEY non trouvée dans .env — abandon.")
      return {}

    checkpoint_path = OUTPUT_DIR / "ai_results.json"
    # Reprise si interrompu
    if checkpoint_path.exists():
        with open(checkpoint_path) as f:
            done = json.load(f)
        print(f"Checkpoint trouvé : {len(done)} conversations déjà traitées.")
    else:
        done = {}

    df_clean, _, _ = load_clean_data()
    client = OpenAI(api_key=OPENAI_API_KEY)

    total = len(df_clean)
    print(f"\nClassification IA — {total} conversations")
    print(f"Modèle : claude-sonnet-4-20250514")
    print(f"Coût estimé : ~$0.50-0.80 (114 conversations × ~1800 tokens)\n")

    for i, (_, row) in enumerate(df_clean.iterrows()):
        rid     = str(row["ResponseId"])
        version = row["FL_13_DO"]

        if rid in done:
            continue  # déjà traité

        conv_text = reconstruct_conversation(row)

        # Analyse participant
        res_part = classify_one(conv_text, PROMPT_PARTICIPANT, client)
        res_part["classification_ok"] = "error" not in res_part

        # Analyse chatbot
        res_bot  = classify_one(conv_text, PROMPT_CHATBOT, client, version=version)
        res_bot["classification_ok"]  = "error" not in res_bot

        done[rid] = {
            "respondent_id": rid,
            "version":       version,
            "n_turns":       int(row["n_turns"]),
            "participant":   res_part,
            "chatbot":       res_bot,
        }

        # Checkpoint toutes les 5 conversations
        if (i + 1) % 5 == 0 or i == total - 1:
            with open(checkpoint_path, "w", encoding="utf-8") as f:
                json.dump(done, f, ensure_ascii=False, indent=2)
            n_ok = sum(1 for v in done.values() if v["participant"].get("classification_ok"))
            print(f"  [{i+1}/{total}] Checkpoint — {n_ok} succès")

        time.sleep(0.8)  # respecte le rate limit

    n_ok = sum(1 for v in done.values() if v["participant"].get("classification_ok"))
    print(f"\nClassification terminée : {n_ok}/{len(done)} succès")
    print(f"Résultats -> {checkpoint_path}")
    return done

if __name__ == "__main__":
    run()
