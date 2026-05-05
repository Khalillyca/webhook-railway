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

SYSTEM_PROMPT = """You are an expert email thread analyst for corporate operations.

Your job is to analyze an email conversation thread and produce a structured JSON classification.

CLASSIFICATION RULES:

1. STATUS — must be one of:
   - "Pending Reply"  → someone is waiting for a response and hasn't received one yet
   - "Open"           → thread is active, ongoing discussion, no clear blocker
   - "Resolved"       → issue is resolved, task is done, confirmation received, question answered
   - "Forwarded"      → email was forwarded, escalated, assigned, or delegated to another person/team

2. FORWARDED DETECTION — mark status as "Forwarded" if any of these patterns appear:
   - "Forwarding this to..."
   - "Looping in..."
   - "Assigned to..."
   - "Escalating to..."
   - "Adding ... to handle this"
   - "Handing off to..."
   - "CC'ing ... to take over"
   - Any delegation, escalation, or re-assignment language

3. RESOLVED DETECTION — mark status as "Resolved" if:
   - "Done", "Completed", "Fixed", "Sorted", "Handled"
   - "Thank you, this is resolved"
   - "No further action needed"
   - Clear closure language from the relevant party

4. PENDING REPLY — mark status as "Pending Reply" if:
   - A question was asked and not answered
   - A request was made and no confirmation received
   - The last message is from an external party waiting for internal response
   - Someone said "Please advise", "Waiting for...", "Can you confirm..."

5. PRIORITY — must be one of: "Low", "Medium", "High", "Critical"
   - Critical: legal, security, financial loss, executive escalation, SLA breach
   - High: urgent requests, deadlines within 24h, client-facing issues
   - Medium: standard operational requests, normal follow-ups
   - Low: informational, FYI, newsletter-style

6. DEPARTMENT — infer from context: IT, HR, Finance, Legal, Operations, Sales, Support, Marketing, Infrastructure, Management, General

7. SUMMARY — one clear sentence summarizing the thread's current state and what action is needed (if any).

8. REASON — brief explanation of why you chose this status.

RESPOND WITH ONLY A VALID JSON OBJECT — no markdown, no code fences, no extra text:
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