"""
utils/sentiment.py
──────────────────
Hybrid sentiment analyser for Feedback Automation Tool.

Layer 1 — Rule-based (instant, zero cost)
  Handles: empty / NA / "no comments from the learner" / "none" / "nil"
  These rows get sentiment inferred from their numeric rating alone.

Layer 2 — LLM batch (claude-haiku-4-5, ~₹0.02 per report)
  Handles: every row with real, substantive text.
  All substantive rows sent in ONE API call per report — not per row.

Public API:
    from utils.sentiment import analyse_from_excel_rows
    sentiments = analyse_from_excel_rows(raw_rows)
    # raw_rows: list of (sr_no, best_part, rating, improvement) tuples
"""

import os
import json
import logging
import re
from typing import Any

logger = logging.getLogger(__name__)

MODEL = "claude-haiku-4-5-20251001"

_NO_ANSWER = {
    "no comments from the learner", "no comments", "na", "n/a",
    "none", "nil", "-", "", ".",
}

_POS_WORDS = {
    "good", "great", "excellent", "thank", "thanks", "well", "helpful",
    "clear", "amazing", "wonderful", "perfect", "best", "love", "loved",
    "enjoyed", "fantastic", "superb", "brilliant", "awesome", "nice",
    "all good", "very good", "informative", "insightful", "engaging",
    "useful", "effective", "outstanding",
}
_NEG_WORDS = {
    "bad", "poor", "worst", "disappoint", "disappointed", "boring",
    "confus", "confused", "slow", "unclear", "waste", "not good",
    "lacking", "improve", "improvement needed", "difficult", "hard to follow",
    "too fast", "too slow", "not helpful", "didn't understand",
}


def _is_empty(text: str) -> bool:
    return text.strip().lower() in _NO_ANSWER


def _rule_sentiment(best_part: str, improvement: str, rating: int) -> dict:
    imp = improvement.strip()
    imp_lower = imp.lower()

    if _is_empty(imp):
        if rating == 5:
            return {"sentiment": "Positive", "confidence": "high",
                    "source": "rule", "reason": "Rating 5, no improvement comment"}
        elif rating == 4:
            return {"sentiment": "Positive", "confidence": "medium",
                    "source": "rule", "reason": "Rating 4, no improvement comment"}
        elif rating == 3:
            return {"sentiment": "Neutral", "confidence": "medium",
                    "source": "rule", "reason": "Rating 3, no improvement comment"}
        else:
            return {"sentiment": "Negative", "confidence": "medium",
                    "source": "rule", "reason": f"Rating {rating}, no improvement comment"}

    words = set(re.sub(r"[^\w\s]", " ", imp_lower).split())
    neg_hit = words & _NEG_WORDS
    pos_hit = words & _POS_WORDS

    if neg_hit and not pos_hit:
        return {"sentiment": "Negative", "confidence": "medium",
                "source": "rule", "reason": f"Negative keywords: {neg_hit}"}
    if pos_hit and not neg_hit:
        return {"sentiment": "Positive", "confidence": "medium",
                "source": "rule", "reason": f"Positive keywords: {pos_hit}"}
    if pos_hit and neg_hit:
        sent = "Positive" if rating >= 4 else "Neutral"
        return {"sentiment": sent, "confidence": "low", "source": "rule",
                "reason": f"Mixed keywords (pos={pos_hit}, neg={neg_hit}), rating={rating}"}

    if rating >= 4:
        return {"sentiment": "Positive", "confidence": "low",
                "source": "rule", "reason": f"No keyword match, rating={rating} (fallback)"}
    elif rating == 3:
        return {"sentiment": "Neutral", "confidence": "low",
                "source": "rule", "reason": f"No keyword match, rating={rating} (fallback)"}
    else:
        return {"sentiment": "Negative", "confidence": "low",
                "source": "rule", "reason": f"No keyword match, rating={rating} (fallback)"}


_SYSTEM_PROMPT = """You are a sentiment classifier for learner feedback on educational sessions.

You will receive a JSON array of feedback rows. Each row has:
  - row: row number (integer)
  - best_part: what the learner said was best about the session
  - improvement: what the learner said could be improved (may be empty/NA)
  - rating: numeric rating 1-5

Classify each row's OVERALL sentiment as exactly one of:
  "Positive"  — learner is satisfied, no real concerns
  "Neutral"   — mixed, ambiguous, or constructively critical
  "Negative"  — dissatisfied, disappointed, or clearly critical

Rules:
  1. Base sentiment primarily on IMPROVEMENT text when it has real content.
  2. If improvement is empty / "na" / "none" / "no comments", use best_part + rating.
  3. Detect sarcasm → Negative.
  4. Handle Hinglish / Hindi text correctly.
  5. "Good content but too fast" → Neutral (mixed).
  6. confidence: "high" if clear signal, "medium" if some ambiguity, "low" if guessing.

Respond ONLY with a valid JSON array — no markdown, no explanation, no preamble:
[
  {"row": 1, "sentiment": "Positive", "confidence": "high", "reason": "..."},
  ...
]"""


def _llm_batch(substantive_rows: list[dict], api_key: str) -> list[dict] | None:
    try:
        import anthropic
    except ImportError:
        logger.warning("anthropic package not installed — pip install anthropic")
        return None

    client = anthropic.Anthropic(api_key=api_key)
    payload = json.dumps(substantive_rows, ensure_ascii=False)

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1024,
            system=_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": payload}],
        )
        raw = response.content[0].text.strip()
        raw = re.sub(r"^```(?:json)?", "", raw).strip()
        raw = re.sub(r"```$", "", raw).strip()
        results = json.loads(raw)
        for item in results:
            assert "row" in item and "sentiment" in item
            assert item["sentiment"] in ("Positive", "Neutral", "Negative")
        return results
    except Exception as exc:
        logger.warning(f"LLM sentiment call failed ({exc}) — falling back to rules")
        return None


def analyse_batch(
    rows: list[dict[str, Any]],
    api_key: str | None = None,
) -> list[dict[str, Any]]:
    resolved_key = api_key or os.environ.get("ANTHROPIC_API_KEY")
    results_by_row: dict[int, dict] = {}
    substantive: list[dict] = []

    for r in rows:
        row_n  = r["row"]
        bp     = r.get("best_part", "")
        imp    = r.get("improvement", "")
        rating = int(r.get("rating", 3))

        if _is_empty(imp):
            result = _rule_sentiment(bp, imp, rating)
            result["row"] = row_n
            results_by_row[row_n] = result
        else:
            substantive.append({"row": row_n, "best_part": bp,
                                 "improvement": imp, "rating": rating})

    if substantive:
        if resolved_key:
            llm_results = _llm_batch(substantive, resolved_key)
        else:
            logger.warning("No ANTHROPIC_API_KEY found — rules-only mode.")
            llm_results = None

        if llm_results:
            llm_by_row = {item["row"]: item for item in llm_results}
            for sub_row in substantive:
                row_n = sub_row["row"]
                if row_n in llm_by_row:
                    item = llm_by_row[row_n]
                    results_by_row[row_n] = {
                        "row": row_n, "sentiment": item["sentiment"],
                        "confidence": item.get("confidence", "medium"),
                        "source": "llm", "reason": item.get("reason", "LLM classification"),
                    }
                else:
                    result = _rule_sentiment(
                        sub_row["best_part"], sub_row["improvement"], sub_row["rating"])
                    result["row"] = row_n
                    results_by_row[row_n] = result
        else:
            for sub_row in substantive:
                row_n = sub_row["row"]
                result = _rule_sentiment(
                    sub_row["best_part"], sub_row["improvement"], sub_row["rating"])
                result["row"] = row_n
                results_by_row[row_n] = result

    return [results_by_row[r["row"]] for r in rows]


def analyse_from_excel_rows(
    raw_rows: list[tuple],
    api_key: str | None = None,
) -> list[dict]:
    """
    Accepts your tool's existing tuple format: (sr_no, best_part, rating, improvement)
    Returns: [{"row":1, "sentiment":"Positive", "confidence":"high", "source":"rule"}, ...]
    """
    normalised = [
        {"row": sr, "best_part": bp, "improvement": imp, "rating": rt}
        for sr, bp, rt, imp in raw_rows
    ]
    return analyse_batch(normalised, api_key=api_key)
