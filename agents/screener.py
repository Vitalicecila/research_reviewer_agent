import os
from dotenv import load_dotenv
from openai import OpenAI
import json

load_dotenv("key.env")
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

CRITERIA_PROMPT = """You are a systematic literature review screener.

Your task is to decide whether to INCLUDE or EXCLUDE a paper based on strict criteria.

=== INCLUSION CRITERIA (ALL must apply to include) ===
1. The paper presents a DESIGNED and IMPLEMENTED system, interface, or prototype
   (e.g., AR/VR reading systems, interactive books, AI reading companions)
2. The system DIRECTLY mediates or shapes the ACT OF READING — reading is the central activity, not secondary
3. The paper includes an EVALUATION with actual users
   (e.g., user study, experiment, usability test involving real participants)
4. The contribution focuses on the READING EXPERIENCE itself
   (e.g., engagement, immersion, enjoyment, interaction with text)
5. The reading context is GENERAL or PLEASURE-ORIENTED — NOT restricted to learning outcomes or educational performance

=== EXCLUSION CRITERIA (ANY one is enough to exclude) ===
- No system/interface is designed (conceptual papers, frameworks, guidelines, theoretical work)
- No evaluation with users (proposals or prototypes without a user study)
- Only existing/commercial systems are used without meaningful modification
- Reading is not the central activity
- Primary focus is on LEARNING OUTCOMES (comprehension improvement, literacy training, educational performance)
- Purely technical contribution (NLP, summarization, translation) with no user-facing reading interaction

=== ONE-LINE RULE ===
Include ONLY papers that DESIGN and EVALUATE a system that directly shapes the READING EXPERIENCE (not learning outcomes).

=== YOUR RESPONSE FORMAT (strict JSON, no extra text) ===
{{
  "decision": "include" | "exclude",
  "confidence": "high" | "medium" | "low",
  "reasoning": "one or two sentences explaining your decision"
}}

=== PAPER TO SCREEN ===
Title: {title}

Abstract: {abstract}
"""


def screen_paper_once(title: str, abstract: str, temperature: float = 0.2) -> dict:
    prompt = CRITERIA_PROMPT.format(title=title, abstract=abstract or "(no abstract available)")
    response = client.chat.completions.create(
        model="gpt-4o",
        messages=[{"role": "user", "content": prompt}],
        temperature=temperature,
        response_format={"type": "json_object"},
    )
    raw = response.choices[0].message.content
    try:
        result = json.loads(raw)
        decision = result.get("decision", "").strip().lower()
        if decision not in ("include", "exclude"):
            decision = "exclude"
        return {
            "decision": decision,
            "confidence": result.get("confidence", "low"),
            "reasoning": result.get("reasoning", ""),
        }
    except Exception:
        return {"decision": "exclude", "confidence": "low", "reasoning": f"Parse error: {raw[:200]}"}


def run_screener(title: str, abstract: str) -> dict:
    """Run 3 independent screening passes and return consensus + consistency info."""
    # Slightly vary temperature across runs to encourage independent reasoning
    temperatures = [0.0, 0.3, 0.5]
    runs = [screen_paper_once(title, abstract, t) for t in temperatures]

    decisions = [r["decision"] for r in runs]
    include_count = decisions.count("include")
    exclude_count = decisions.count("exclude")

    consensus = "include" if include_count >= 2 else "exclude"
    consistent = include_count == 3 or exclude_count == 3

    return {
        "run_1_decision": runs[0]["decision"],
        "run_1_confidence": runs[0]["confidence"],
        "run_1_reasoning": runs[0]["reasoning"],
        "run_2_decision": runs[1]["decision"],
        "run_2_confidence": runs[1]["confidence"],
        "run_2_reasoning": runs[1]["reasoning"],
        "run_3_decision": runs[2]["decision"],
        "run_3_confidence": runs[2]["confidence"],
        "run_3_reasoning": runs[2]["reasoning"],
        "llm_consensus": consensus,
        "llm_consistent": consistent,
        "include_votes": include_count,
    }
