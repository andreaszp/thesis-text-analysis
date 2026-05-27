"""
05_api_analysis.py
Block 2 — Step 3: AI classification of each conversation via OpenAI API.
API key loaded from .env — never hardcoded.
Checkpoint every 5 conversations — safe to restart.
Saves: outputs/ai_results.json, df_ai.json, df_progression.json

VARIABLES REMOVED vs previous version:
- profile_engagement  (r=0.937 with quality_global — redundant)
- profile_expertise   (r=0.827 with quality_global — redundant)
- action_global       (r=0.794 with quality_global — redundant)
- content_competitor  (not relevant to thesis scope)
- content_suggestion  (r=0.966 with action_advice — duplicate)
- content_concrete_pb (duplicate of action_concrete_pb)
- content_opinion     (95% base rate — no discriminant power)
- content_emotion     (removed per variable reduction)
"""
import sys, json, time, re
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))

import pandas as pd
from openai import OpenAI
from config import OPENAI_API_KEY, OUTPUT_DIR
from loader import load_clean_data, reconstruct_conversation

BRIEF_FL21 = """Version FL_21 — FRIENDLY/WARM tone:
Exact prompt given to the chatbot:
'You should maintain a friendly, warm, and engaging tone throughout the conversation.'
EXPECTED characteristics:
- Warm, engaging, enthusiastic
- Informal address (tutoiement) possible
- Emojis encouraged
- Encouragement phrases (super!, great!, interesting!)
- Personalised follow-up questions
- Relaxed but professional register"""

BRIEF_FL22 = """Version FL_22 — PROFESSIONAL/FORMAL tone:
Exact prompt given to the chatbot:
'You should maintain a professional and respectful tone throughout the conversation.'
EXPECTED characteristics:
- Neutral, sober, respectful
- Formal address (vouvoiement) throughout
- No emojis
- Sober and direct phrasing
- No enthusiastic exclamations
- Formal, structured register"""

END_PHRASES = [
    "tu peux cliquer sur la fleche","you can click the arrow",
    "click the arrow at the bottom","<end_of_interview>",
    "avez-vous quelque chose a ajouter","as-tu quelque chose a ajouter",
    "is there anything else","y a-t-il autre chose",
    "anything else to add","anything else you would like to share",
    "do you have anything else","avez-vous autre chose",
]

def detect_proper_end(row):
    n = row["n_turns"]
    for i in range(max(1, n-1), n+2):
        resp = row.get(f"response_{i}")
        if pd.notna(resp) and any(p in str(resp).lower() for p in END_PHRASES):
            return True
    return False

def classify_last_msg(row):
    n = row["n_turns"]
    last = str(row.get(f"msg_{n}","") or "").strip()
    nw   = len(last.split())
    if row.get("conv_ended_properly", False): return "proper_end"
    if n <= 3 and nw <= 5:                   return "early_dropout"
    if nw <= 3:                              return "minimal_response"
    return "end_without_closing"

PROMPT = """You are a senior expert in qualitative UX research and user interview analysis.

Analyse this conversation between a UX research chatbot and a participant.
Context: study on SoundFlow, a music streaming application.

CHATBOT BRIEF (what it was instructed to do):
{brief}

FULL CONVERSATION:
{conv_text}

---

SCORING GRIDS — read carefully before analysing.

RESPONSE QUALITY (1-5):
score_global:
  1 = Almost all answers are 1-3 words, no development whatsoever
  2 = A few sentences but poorly developed, little context
  3 = Adequate development but incomplete, few examples
  4 = Detailed answers with context and at least one example
  5 = Rich, precise, nuanced answers with multiple concrete examples

score_precision:
  1 = Very vague, no exploitable information
  2 = Imprecise, generalities without grounding
  3 = Some precise elements but insufficient
  4 = Precise on most points
  5 = Very precise, every statement grounded in a fact or concrete example

score_examples:
  1 = No concrete example in the entire conversation
  2 = One vague or poorly developed example
  3 = One correct example
  4 = Several relevant examples
  5 = Multiple detailed examples illustrating different aspects

score_relevance:
  1 = Answers do not address the questions asked
  2 = Partially answers, often off-topic
  3 = Answers correctly but generically
  4 = Answers directly and precisely
  5 = Answers perfectly, anticipates follow-up, brings unsolicited context

score_richness:
  1 = Very limited vocabulary, repetitive simple sentences
  2 = Little variety, poor register
  3 = Correct vocabulary, some nuances
  4 = Varied vocabulary, nuanced phrasing
  5 = Rich vocabulary, precise register, fine nuances

ACTIONABILITY (1-5):
contains_concrete_problem:
  true = Participant describes a precise problem with context (when X happens, Y occurs, because Z)
  false = Vague dissatisfaction without problem description

contains_applicable_advice:
  true = Participant formulates an improvement a product team could implement directly
  false = Vague wish not directly implementable

CONTENT TYPE:
suggestion_feature_request:
  true = Participant explicitly proposes a feature or improvement
  false = No concrete suggestion

PARTICIPANT PROFILE:
elaboration:
  "short" = messages average < 10 words
  "medium" = messages average 10-30 words
  "detailed" = messages average > 30 words

coherence:
  "low" = contradictory or off-topic answers frequent
  "medium" = some inconsistencies or digressions
  "high" = coherent, logical discourse, easy to follow

TURN-BY-TURN PROGRESSION:
IMPORTANT:
- One turn = one participant response [P1], [P2], etc.
- An isolated "yes" does NOT downgrade the score if the next answer is rich
- Breakpoint = moment where AT LEAST 2 CONSECUTIVE turns show lasting degradation
- If no clear breakpoint: breakpoint.exists = false and turn = null

CHATBOT — PERCEIVED TONE:
score_friendly (1-5):
  1 = Very cold, distant, no warmth
  2 = Formal and sober, little warmth
  3 = Neutral, neither cold nor warm
  4 = Warm, engaging, some enthusiasm markers
  5 = Very warm, enthusiastic, frequent encouragement, emojis, informal address

score_professional (1-5):
  1 = Very casual, almost familiar
  2 = Not very formal, relaxed register
  3 = Neutral, neither formal nor casual
  4 = Formal, sober, respectful
  5 = Very formal, systematic formal address, no emojis, elevated register

BRIEF COMPLIANCE (1-5):
  1 = Completely opposite to the requested tone
  2 = Some compliant elements but majority non-compliant
  3 = Partially compliant, notable drift
  4 = Broadly compliant, minor deviations
  5 = Perfectly compliant with requested tone

---

Reply ONLY with valid JSON, no markdown, no text before or after.

{{
  "participant": {{
    "quality": {{
      "score_global": <1-5>,
      "score_precision": <1-5>,
      "score_examples": <1-5>,
      "score_relevance": <1-5>,
      "score_richness": <1-5>,
      "justification": "<1 factual sentence>"
    }},
    "actionability": {{
      "contains_concrete_problem": <true|false>,
      "contains_applicable_advice": <true|false>,
      "justification": "<1 sentence with example from conversation>"
    }},
    "content_type": {{
      "suggestion_feature_request": <true|false>
    }},
    "profile": {{
      "elaboration": "<short|medium|detailed>",
      "coherence": "<low|medium|high>"
    }},
    "turn_progression": [
      {{
        "turn": <number>,
        "quality_score": <1-5>,
        "approx_words": <integer>,
        "note": "<1 short sentence if notable, else null>"
      }}
    ],
    "breakpoint": {{
      "exists": <true|false>,
      "turn": <number or null>,
      "explanation": "<description or null>"
    }},
    "conversation_end": {{
      "completed_fully": <true|false>,
      "end_type": "<proper_end|minimal_response|early_dropout|end_without_closing>",
      "last_participant_message": "<exact text>"
    }},
    "summary": "<2-3 sentences summarising participant contribution>",
    "key_verbatim": "<exact most representative quote, max 120 chars>",
    "main_language": "<fr|en|mixed>"
  }},
  "chatbot": {{
    "perceived_tone": {{
      "score_friendly": <1-5>,
      "score_professional": <1-5>,
      "dominant_label": "<very friendly|friendly|neutral|professional|very professional>",
      "justification": "<1 sentence with concrete example>"
    }},
    "tone_coherence": {{
      "score": <1-5>,
      "drift_detected": <true|false>,
      "drift_description": "<description if drift, else null>"
    }},
    "tone_markers": {{
      "friendly_words_detected": ["<list>"],
      "formal_words_detected": ["<list>"],
      "emojis_used": <true|false>,
      "emoji_list": ["<detected emojis or empty>"],
      "informal_address": <true|false>,
      "formal_address": <true|false>,
      "encouragement_phrases": ["<e.g. great!, super!>"],
      "sober_phrases": ["<e.g. I understand, indeed>"]
    }},
    "brief_compliance": {{
      "score": <1-5>,
      "compliant": <true|false>,
      "detected_gap": "<precise description if non-compliant, else null>",
      "justification": "<1 sentence with example>"
    }},
    "question_style": {{
      "open_questions_pct": <0-100>,
      "personalised_follow_ups": <true|false>,
      "perceived_pressure": "<none|slight|moderate|strong>"
    }}
  }}
}}"""

def call_api(prompt, client, model="gpt-4o-mini", max_retries=3):
    for attempt in range(max_retries):
        try:
            resp = client.chat.completions.create(
                model=model, max_tokens=2500,
                response_format={"type":"json_object"},
                messages=[
                    {"role":"system","content":"You are a UX qualitative research expert. Reply only with valid JSON."},
                    {"role":"user","content":prompt}
                ]
            )
            return json.loads(resp.choices[0].message.content.strip())
        except json.JSONDecodeError as e:
            if attempt==max_retries-1: return {"error":str(e),"ok":False}
            time.sleep(2)
        except Exception as e:
            err=str(e)
            if "rate_limit" in err.lower() or "429" in err:
                time.sleep(30*(attempt+1))
            elif attempt==max_retries-1:
                return {"error":err,"ok":False}
            else:
                time.sleep(5)
    return {"error":"max_retries","ok":False}

def analyze(row, client):
    brief = BRIEF_FL21 if row["FL_13_DO"]=="FL_21" else BRIEF_FL22
    conv  = reconstruct_conversation(row)[:4000]
    res   = call_api(PROMPT.format(brief=brief, conv_text=conv), client)
    res["respondent_id"]       = str(row["ResponseId"])
    res["version"]             = row["FL_13_DO"]
    res["n_turns"]             = int(row["n_turns"])
    res["conv_ended_properly"] = bool(row.get("conv_ended_properly",False))
    res["last_msg_type"]       = row.get("last_msg_type","")
    res["ok"]                  = "error" not in res
    return res

def flatten(r):
    p   = r.get("participant",{})
    c   = r.get("chatbot",{})
    q   = p.get("quality",{})
    act = p.get("actionability",{})
    cnt = p.get("content_type",{})
    prf = p.get("profile",{})
    brk = p.get("breakpoint",{})
    fin = p.get("conversation_end",{})
    ton = c.get("perceived_tone",{})
    coh = c.get("tone_coherence",{})
    mrk = c.get("tone_markers",{})
    com = c.get("brief_compliance",{})
    stq = c.get("question_style",{})
    return {
        "respondent_id":        r.get("respondent_id",""),
        "version":              r.get("version",""),
        "n_turns":              r.get("n_turns"),
        "conv_ended_properly":  r.get("conv_ended_properly"),
        "last_msg_type":        r.get("last_msg_type",""),
        "ok":                   r.get("ok",False),
        # Quality (all 5 subscores + global kept)
        "quality_global":       q.get("score_global"),
        "quality_precision":    q.get("score_precision"),
        "quality_examples":     q.get("score_examples"),
        "quality_relevance":    q.get("score_relevance"),
        "quality_richness":     q.get("score_richness"),
        "quality_justification":q.get("justification",""),
        # Actionability (action_global REMOVED; concrete_pb and advice kept)
        "action_concrete_pb":   int(act.get("contains_concrete_problem",False)),
        "action_advice":        int(act.get("contains_applicable_advice",False)),
        "action_justification": act.get("justification",""),
        # Content type (only suggestion kept; opinion/emotion/competitor/concrete_pb REMOVED)
        "content_suggestion":   int(cnt.get("suggestion_feature_request",False)),
        # Profile (engagement and expertise REMOVED; elaboration and coherence kept)
        "profile_elaboration":  prf.get("elaboration",""),
        "profile_coherence":    prf.get("coherence",""),
        # Breakpoint
        "breakpoint_exists":    int(brk.get("exists",False)),
        "breakpoint_turn":      brk.get("turn"),
        "breakpoint_expl":      brk.get("explanation",""),
        # Conversation end
        "completed_fully":      int(fin.get("completed_fully",False)),
        "end_type":             fin.get("end_type",""),
        "last_msg":             fin.get("last_participant_message",""),
        # Summary
        "summary":              p.get("summary",""),
        "key_verbatim":         p.get("key_verbatim",""),
        "language":             p.get("main_language",""),
        # Chatbot tone
        "bot_score_friendly":   ton.get("score_friendly"),
        "bot_score_professional":ton.get("score_professional"),
        "bot_tone_label":       ton.get("dominant_label",""),
        "bot_tone_just":        ton.get("justification",""),
        # Chatbot coherence
        "bot_coherence_score":  coh.get("score"),
        "bot_drift":            int(coh.get("drift_detected",False)),
        "bot_drift_desc":       coh.get("drift_description",""),
        # Chatbot markers
        "bot_friendly_words":   " | ".join(mrk.get("friendly_words_detected",[])),
        "bot_formal_words":     " | ".join(mrk.get("formal_words_detected",[])),
        "bot_n_friendly_words": len(mrk.get("friendly_words_detected",[])),
        "bot_n_formal_words":   len(mrk.get("formal_words_detected",[])),
        "bot_emojis":           int(mrk.get("emojis_used",False)),
        "bot_emoji_list":       " ".join(mrk.get("emoji_list",[])),
        "bot_informal_addr":    int(mrk.get("informal_address",False)),
        "bot_formal_addr":      int(mrk.get("formal_address",False)),
        "bot_n_encouragements": len(mrk.get("encouragement_phrases",[])),
        "bot_n_sober":          len(mrk.get("sober_phrases",[])),
        # Brief compliance
        "bot_compliance_score": com.get("score"),
        "bot_compliant":        int(com.get("compliant",False)),
        "bot_gap":              com.get("detected_gap",""),
        "bot_compliance_just":  com.get("justification",""),
        # Question style
        "bot_open_q_pct":       stq.get("open_questions_pct"),
        "bot_follow_ups":       int(stq.get("personalised_follow_ups",False)),
        "bot_pressure":         stq.get("perceived_pressure",""),
    }

def run():
    if not OPENAI_API_KEY:
        print("ERROR: OPENAI_API_KEY not found in .env"); return {}

    checkpoint = OUTPUT_DIR/"ai_results.json"
    done = {}
    if checkpoint.exists():
        with open(checkpoint,encoding="utf-8") as f: done=json.load(f)
        print(f"Checkpoint found: {len(done)} conversations already done.")

    df_clean, _, _ = load_clean_data()
    df_clean["conv_ended_properly"] = df_clean.apply(detect_proper_end, axis=1)
    df_clean["last_msg_type"]       = df_clean.apply(classify_last_msg,  axis=1)

    import os
    api_key = os.environ.get("OPENAI_API_KEY", "").strip() or OPENAI_API_KEY.strip()
    client = OpenAI(api_key=api_key)
    total  = len(df_clean)
    print(f"\nAI Analysis — {total} conversations | model: gpt-4o-mini | est. cost: ~$0.15-0.25\n")

    for i, (_, row) in enumerate(df_clean.iterrows()):
        rid = str(row["ResponseId"])
        if rid in done: continue
        print(f"  [{i+1}/{total}] {rid[:20]}... FL={row['FL_13_DO']} {row['n_turns']} turns", end=" ")
        res = analyze(row, client)
        print("OK" if res.get("ok") else f"ERROR {res.get('error','')[:40]}")
        done[rid] = res
        if (i+1)%5==0 or i==total-1:
            with open(checkpoint,"w",encoding="utf-8") as f:
                json.dump(done, f, ensure_ascii=False, indent=2)
        time.sleep(0.8)

    n_ok = sum(1 for v in done.values() if v.get("ok"))
    print(f"\nDone — {n_ok}/{len(done)} successful")

    rows_flat = [flatten(v) for v in done.values() if v.get("ok")]
    df_ai = pd.DataFrame(rows_flat)
    df_ai.to_json(OUTPUT_DIR/"df_ai.json", orient="records", force_ascii=False)

    prog = []
    for rid, v in done.items():
        if not v.get("ok"): continue
        for t in v.get("participant",{}).get("turn_progression",[]):
            prog.append({"respondent_id":rid,"version":v.get("version",""),
                         "turn":t.get("turn"),"quality_score":t.get("quality_score"),
                         "nb_words":t.get("approx_words"),"note":t.get("note","")})
    pd.DataFrame(prog).to_json(OUTPUT_DIR/"df_progression.json",orient="records",force_ascii=False)
    print(f"Saved: df_ai.json ({len(df_ai)} rows), df_progression.json ({len(prog)} rows)")
    return done

if __name__ == "__main__":
    run()
