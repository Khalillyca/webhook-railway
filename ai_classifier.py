"""
ai_classifier.py — AI classification using OpenAI gpt-4o

Analyzes full thread text and returns structured classification:
  summary, department, priority, status, reason

Includes forwarded/escalated detection.
Robust fallback to status='Open' on any failure.
"""

import os
import json
import logging
from typing import Optional

log = logging.getLogger(__name__)

MODEL = "gpt-4o"

SYSTEM_PROMPT = """You are an intelligent email analyst for corporate operations. Your job is to deeply understand email conversations and classify them accurately based on context, tone, and intent — not just keywords.

Analyze the full email thread and return a JSON classification.

UNDERSTANDING THE THREAD:
- Read every message in the thread carefully
- Understand who is waiting for what, and from whom
- Consider the tone: urgent, casual, formal, frustrated, satisfied
- Look at the last message — who sent it and what did they need
- Consider the overall arc: what started, what happened, where it stands now

YOUR CLASSIFICATIONS:

status: What is the current state of this thread?
- "Pending Reply" → The ball is in someone's court and they haven't responded yet
- "Open" → Active conversation, things are moving, no one is stuck waiting
- "Resolved" → The matter is closed — confirmed, completed, thanked, or acknowledged as done
- "Forwarded" → Responsibility has been handed to someone else — delegated, escalated, reassigned, looped in

priority: How urgent is this, really?
- "Critical" → Legal, financial, security, or executive risk. Needs immediate attention.
- "High" → Time-sensitive. Client or business impact within 24 hours.
- "Medium" → Normal operational matter. Should be addressed soon but not urgent.
- "Low" → Informational, routine, or no action required.

department: Which team owns this? Choose the most fitting:
IT, HR, Finance, Legal, Operations, Sales, Support, Marketing, Infrastructure, Management, General

summary: One sentence. What is this thread about and what is the current situation?

reason: One sentence. Why did you choose this status specifically?

Think carefully. Use your judgment. Context matters more than exact phrases.

Return ONLY this JSON, nothing else:
{
  "summary": "...",
  "department": "...",
  "priority": "Low|Medium|High|Critical",
  "status": "Pending Reply|Open|Resolved|Forwarded",
  "reason": "..."
}"""


def _get_client():
    """Create OpenAI client lazily so missing key doesn't crash at import time."""
    from openai import OpenAI
    api_key = os.getenv("OPENAI_API_KEY", "")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY environment variable is not set.")
    return OpenAI(api_key=api_key)


def classify_thread(full_thread_text: str, subject: str = "") -> dict:
    fallback = {
        "summary": f"Thread: {subject}" if subject else "Unable to classify",
        "department": "General",
        "priority": "Medium",
        "status": "Open",
        "reason": "AI classification failed — defaulting to Open",
    }

    if not full_thread_text or not full_thread_text.strip():
        log.warning("[AI] Empty thread text — returning fallback.")
        return fallback

    user_prompt = f"Subject: {subject}\n\n--- FULL THREAD ---\n{full_thread_text[:12000]}"

    try:
        log.info(f"[AI] Classifying thread: {subject[:80]}...")

        client = _get_client()
        response = client.chat.completions.create(
            model=MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.1,
            max_tokens=500,
            response_format={"type": "json_object"},
        )

        raw_text = response.choices[0].message.content.strip()
        log.info(f"[AI] Raw response: {raw_text[:300]}")

        result = _parse_response(raw_text)
        if result:
            _validate_result(result)
            log.info(
                f"[AI] Classification: status={result['status']}, "
                f"priority={result['priority']}, dept={result['department']}"
            )
            return result

        log.warning("[AI] Failed to parse response — using fallback.")
        return fallback

    except Exception as e:
        log.error(f"[AI] Classification error: {e}", exc_info=True)
        return fallback


def _parse_response(raw_text: str) -> Optional[dict]:
    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        pass

    import re
    json_match = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", raw_text, re.DOTALL)
    if json_match:
        try:
            return json.loads(json_match.group(1))
        except json.JSONDecodeError:
            pass

    brace_match = re.search(r"\{[^{}]*\}", raw_text, re.DOTALL)
    if brace_match:
        try:
            return json.loads(brace_match.group(0))
        except json.JSONDecodeError:
            pass

    return None


def _validate_result(result: dict):
    valid_statuses = {"Pending Reply", "Open", "Resolved", "Forwarded"}
    valid_priorities = {"Low", "Medium", "High", "Critical"}

    if "summary" not in result or not result["summary"]:
        result["summary"] = "No summary available"
    if "department" not in result or not result["department"]:
        result["department"] = "General"
    if "priority" not in result or result["priority"] not in valid_priorities:
        result["priority"] = "Medium"
    if "status" not in result or result["status"] not in valid_statuses:
        result["status"] = "Open"
    if "reason" not in result or not result["reason"]:
        result["reason"] = "Classification completed"