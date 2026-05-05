"""
bootstrap.py — First-run baseline scan + old pending thread analysis

On first run (or manual trigger):
  1. Fetch all messages since the configured scan date
  2. Group by conversationId (threadId)
  3. Store baseline msgCount in SQLite (no live trigger for baseline)
  4. AI-classify each thread
  5. Send overdue alerts for old threads that are 'Pending Reply'
  6. Skip 'Resolved' and 'Forwarded' threads
"""

import logging
from datetime import datetime, timezone
from collections import defaultdict
from typing import Optional

from auth import get_token
from graph import fetch_thread, is_from_target
from ai_classifier import classify_thread
from trigger import send_trigger_mail
from db import (
    get_thread_state,
    upsert_thread_state,
    set_thread_triggered,
    get_config_targets,
    get_config_thread_filter,
    get_config_overdue_hours,
    get_config_scan_date,
)

import requests

log = logging.getLogger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"


def _headers() -> dict:
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type": "application/json",
    }


def _fetch_messages_since(scan_date: str) -> list:
    """
    Fetch ALL messages since scan_date using pagination.
    Searches across Inbox + Sent + all folders.
    Returns list of raw message dicts.
    """
    all_messages = []
    url = f"{GRAPH}/me/messages"
    params = {
        "$filter": f"receivedDateTime ge {scan_date}T00:00:00Z",
        "$select": "id,conversationId,subject,from,toRecipients,ccRecipients,"
                   "receivedDateTime,bodyPreview",
        "$orderby": "receivedDateTime desc",
        "$top": 100,
    }

    page = 1
    while url:
        log.info(f"[BOOTSTRAP] Fetching message page {page}...")
        r = requests.get(url, headers=_headers(), params=params if page == 1 else None)
        r.raise_for_status()
        data = r.json()

        batch = data.get("value", [])
        all_messages.extend(batch)
        log.info(f"[BOOTSTRAP] Page {page}: {len(batch)} messages (total: {len(all_messages)})")

        url = data.get("@odata.nextLink")
        page += 1

    return all_messages


def _group_by_thread(messages: list) -> dict:
    """Group messages by conversationId → list of messages."""
    threads = defaultdict(list)
    for msg in messages:
        conv_id = msg.get("conversationId", "")
        if conv_id:
            threads[conv_id].append(msg)
    return dict(threads)


def _is_thread_relevant(messages: list, targets: list) -> bool:
    """Check if any message in the thread involves a target email/domain."""
    if not targets:
        return True  # No filter → all threads relevant

    for msg in messages:
        all_addresses = []
        from_addr = msg.get("from", {}).get("emailAddress", {}).get("address", "")
        if from_addr:
            all_addresses.append(from_addr.lower())
        for recip in msg.get("toRecipients", []):
            addr = recip.get("emailAddress", {}).get("address", "")
            if addr:
                all_addresses.append(addr.lower())
        for recip in msg.get("ccRecipients", []):
            addr = recip.get("emailAddress", {}).get("address", "")
            if addr:
                all_addresses.append(addr.lower())

        for addr in all_addresses:
            for target in targets:
                target = target.lower().strip()
                if "@" in target:
                    if addr == target:
                        return True
                else:
                    if addr.endswith(f"@{target}"):
                        return True
    return False


def run_bootstrap():
    """
    Execute the full bootstrap scan:
      1. Fetch messages since scan_date
      2. Group by thread
      3. Filter by targets + threadId filter
      4. Store baseline msgCount
      5. AI-classify
      6. Send overdue alerts for Pending Reply threads
    """
    scan_date = get_config_scan_date()
    if not scan_date:
        log.warning("[BOOTSTRAP] No scan_date configured — skipping bootstrap.")
        return

    targets = get_config_targets()
    thread_filter = get_config_thread_filter()
    overdue_hours = get_config_overdue_hours()

    log.info("=" * 60)
    log.info("[BOOTSTRAP] Starting baseline scan...")
    log.info(f"  Scan from  : {scan_date}")
    log.info(f"  Targets    : {targets}")
    log.info(f"  Thread filter: {thread_filter or '(all threads)'}")
    log.info(f"  Overdue hrs: {overdue_hours}")
    log.info("=" * 60)

    # Step 1: Fetch all messages
    raw_messages = _fetch_messages_since(scan_date)
    log.info(f"[BOOTSTRAP] Total messages fetched: {len(raw_messages)}")

    if not raw_messages:
        log.info("[BOOTSTRAP] No messages found — nothing to bootstrap.")
        return

    # Step 2: Group by thread
    thread_groups = _group_by_thread(raw_messages)
    log.info(f"[BOOTSTRAP] Total threads found: {len(thread_groups)}")

    # Step 3: Process each thread
    processed = 0
    pending_count = 0
    overdue_count = 0

    for conv_id, msgs in thread_groups.items():
        # Apply threadId filter
        if thread_filter and conv_id != thread_filter:
            continue

        # Apply target filter
        if targets and not _is_thread_relevant(msgs, targets):
            continue

        # Check if already in DB (skip if already baselined)
        existing = get_thread_state(conv_id)
        if existing:
            log.debug(f"[BOOTSTRAP] Thread {conv_id[:30]}... already in DB — skipping.")
            continue

        subject = msgs[0].get("subject", "(no subject)") if msgs else ""
        msg_count = len(msgs)

        try:
            # Fetch full thread with body content for AI
            thread = fetch_thread(conv_id)
            if not thread:
                log.warning(f"[BOOTSTRAP] Could not fetch thread {conv_id[:30]}...")
                continue

            msg_count = thread.get("msg_count", msg_count)
            full_text = thread.get("full_thread_text", "")
            last_activity = thread.get("last_activity_at", "")

            # AI classify
            ai_result = classify_thread(full_text, subject)
            status = ai_result.get("status", "Open")
            log.info(
                f"[BOOTSTRAP] Thread: {subject[:60]} | "
                f"msgs={msg_count} | status={status} | "
                f"priority={ai_result.get('priority')}"
            )

            # Store baseline — NO live trigger
            upsert_thread_state(
                thread_id=conv_id,
                msg_count=msg_count,
                subject=subject,
                last_status=status,
                last_summary=ai_result.get("summary", ""),
                last_department=ai_result.get("department", ""),
                last_priority=ai_result.get("priority", "Medium"),
                last_activity=last_activity,
            )
            processed += 1

            # Push data to your Django API
            try:
                api_payload = {
                    "thread_id": conv_id,
                    "subject": subject,
                    "msg_count": msg_count,
                    "status": status,
                    "summary": ai_result.get("summary", ""),
                    "department": ai_result.get("department", ""),
                    "priority": ai_result.get("priority", "Medium"),
                    "last_activity": last_activity,
                    "participants": [p['email'] for p in thread.get('all_participants', [])] if thread else []
                }
                log.info(f"[BOOTSTRAP] Pushing thread data to Django API for {conv_id[:30]}...")
                api_resp = requests.post(
                    "https://d1vg1nar1ura2r.cloudfront.net/api/email-tracker/",
                    json=api_payload,
                    headers={"Content-Type": "application/json"},
                    timeout=10
                )
                log.info(f"[BOOTSTRAP] Django API Response: {api_resp.status_code}")
            except Exception as api_err:
                log.error(f"[BOOTSTRAP] Failed to push data to Django API: {api_err}")

            # Send overdue alert for Pending Reply threads
            if status == "Pending Reply":
                pending_count += 1
                now = datetime.now(timezone.utc)
                if last_activity:
                    try:
                        act_dt = datetime.fromisoformat(
                            last_activity.replace("Z", "+00:00")
                        )
                        if act_dt.tzinfo is None:
                            act_dt = act_dt.replace(tzinfo=timezone.utc)
                        hours_elapsed = (now - act_dt).total_seconds() / 3600

                        if hours_elapsed >= overdue_hours:
                            log.info(
                                f"[BOOTSTRAP] Overdue alert: {subject[:60]} "
                                f"({hours_elapsed:.1f}h elapsed)"
                            )
                            send_trigger_mail(
                                thread,
                                event_type="OVERDUE",
                                ai_result=ai_result,
                            )
                            set_thread_triggered(conv_id)
                            overdue_count += 1
                    except (ValueError, TypeError) as e:
                        log.warning(f"[BOOTSTRAP] Date parse error: {e}")

        except Exception as e:
            log.error(f"[BOOTSTRAP] Error processing thread {conv_id[:30]}: {e}", exc_info=True)
            # Store minimal state even on error
            upsert_thread_state(
                thread_id=conv_id,
                msg_count=msg_count,
                subject=subject,
                last_status="Open",
                last_activity="",
            )
            processed += 1

    log.info("=" * 60)
    log.info(f"[BOOTSTRAP] Scan complete!")
    log.info(f"  Threads processed : {processed}")
    log.info(f"  Pending Reply     : {pending_count}")
    log.info(f"  Overdue alerts sent: {overdue_count}")
    log.info("=" * 60)
