"""
webhook_server.py — Flask server deployed on Railway
Microsoft Graph POSTs to this when emails arrive.

Configuration via environment variables (set in Railway dashboard):
  TARGETS          — comma-separated emails/domains to track
  OVERDUE_HOURS    — hours before overdue alert (default: 6)
  TOKEN_CACHE_B64  — base64-encoded token cache from local machine
  RAILWAY_API_URL  — Django API URL to push thread data to
  CLIENT_STATE_SECRET — webhook secret (default: my-outlook-tracker-secret-2026)
"""

import os
import json
import logging
import threading
import time
import requests
from datetime import datetime, timezone
from flask import Flask, request, jsonify

from auth import get_token, get_my_email
from graph import fetch_message, fetch_thread, is_from_target
from trigger import send_trigger_mail
from ai_classifier import classify_thread
from db import (
    init_db,
    is_message_processed,
    mark_message_processed,
    get_thread_state,
    upsert_thread_state,
    set_thread_triggered,
    get_pending_threads,
    get_config_targets,
    get_config_thread_filter,
    get_config_overdue_hours,
    get_config_scan_date,
    set_config,
    cleanup_old_processed,
)
from bootstrap import run_bootstrap

# ------------------------------------------------------------------ #
#  Config from environment variables                                   #
# ------------------------------------------------------------------ #

CLIENT_STATE_SECRET = os.getenv("CLIENT_STATE_SECRET", "my-outlook-tracker-secret-2026")
RAILWAY_API_URL = os.getenv("RAILWAY_API_URL", "https://web-production-5af2f.up.railway.app/api/email-tracker/")

# ------------------------------------------------------------------ #
#  Initialize DB and config at import time (required for gunicorn)     #
# ------------------------------------------------------------------ #

init_db()

def load_config_from_env():
    targets = os.getenv("TARGETS", "")
    if targets:
        set_config("targets", targets)
    overdue_hours = os.getenv("OVERDUE_HOURS", "6")
    set_config("overdue_hours", overdue_hours)
    thread_filter = os.getenv("THREAD_FILTER", "")
    if thread_filter:
        set_config("thread_filter", thread_filter)
    scan_date = os.getenv("SCAN_DATE", "")
    if scan_date:
        set_config("scan_date", scan_date)

load_config_from_env()

# ------------------------------------------------------------------ #
#  Logging                                                             #
# ------------------------------------------------------------------ #

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
    handlers=[logging.StreamHandler()]
)
log = logging.getLogger(__name__)

# ------------------------------------------------------------------ #
#  Flask app                                                           #
# ------------------------------------------------------------------ #

app = Flask(__name__)


@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    # ── CASE 1: Validation handshake ──────────────────────────────── #
    validation_token = request.args.get("validationToken")
    if validation_token:
        log.info("[WEBHOOK] Validation handshake received.")
        return validation_token, 200, {"Content-Type": "text/plain"}

    # ── CASE 2: Real notification ─────────────────────────────────── #
    try:
        data = request.get_json(force=True, silent=True)
        if not data:
            log.warning("[WEBHOOK] Empty or non-JSON body received.")
            return "", 202

        notifications = data.get("value", [])
        log.info(f"[WEBHOOK] Received {len(notifications)} notification(s).")

        for notif in notifications:
            if notif.get("clientState") != CLIENT_STATE_SECRET:
                log.warning("[WEBHOOK] clientState mismatch — ignoring.")
                continue

            change_type = notif.get("changeType", "")
            resource    = notif.get("resource", "")
            resource_id = notif.get("resourceData", {}).get("id", "")
            message_id  = resource_id or resource.split("/")[-1]

            if not message_id:
                continue

            if change_type not in ("created", "updated"):
                continue

            if is_message_processed(message_id):
                continue

            mark_message_processed(message_id)

            threading.Thread(
                target=process_notification,
                args=(message_id,),
                daemon=True,
            ).start()

    except Exception as e:
        log.error(f"[WEBHOOK] Error: {e}", exc_info=True)

    return "", 202


def process_notification(message_id: str):
    try:
        targets = get_config_targets()
        thread_filter = get_config_thread_filter()

        message = fetch_message(message_id)

        log.info(
            f"[MESSAGE] Subject: {message['subject']} | "
            f"From: {message['from']['email']}"
        )

        if targets and not is_from_target(message, targets):
            log.info("[PROCESSING] Not from target — skipping.")
            return

        thread_id = message["thread_id"]
        if thread_filter and thread_id != thread_filter:
            log.info("[PROCESSING] Thread does not match filter — skipping.")
            return

        thread = fetch_thread(thread_id)
        if not thread:
            return

        new_msg_count = thread.get("msg_count", 0)
        subject       = thread.get("subject", "")
        last_activity = thread.get("last_activity_at", "")

        existing  = get_thread_state(thread_id)
        old_count = existing["msg_count"] if existing else 0

        if new_msg_count <= old_count:
            upsert_thread_state(
                thread_id=thread_id,
                msg_count=new_msg_count,
                subject=subject,
                last_status=existing["last_status"] if existing else "Open",
                last_activity=last_activity,
            )
            return

        log.info(f"[AI] Classifying: {subject[:60]}...")
        ai_result = classify_thread(thread.get("full_thread_text", ""), subject)
        status    = ai_result.get("status", "Open")

        log.info(f"[AI RESULT] Status={status} | Priority={ai_result.get('priority')}")

        upsert_thread_state(
            thread_id=thread_id,
            msg_count=new_msg_count,
            subject=subject,
            last_status=status,
            last_summary=ai_result.get("summary", ""),
            last_department=ai_result.get("department", ""),
            last_priority=ai_result.get("priority", "Medium"),
            last_activity=last_activity,
        )

        # Push to Django API
        try:
            api_payload = {
                "thread_id":    thread_id,
                "subject":      subject,
                "msg_count":    new_msg_count,
                "status":       status,
                "summary":      ai_result.get("summary", ""),
                "department":   ai_result.get("department", ""),
                "priority":     ai_result.get("priority", "Medium"),
                "last_activity": last_activity,
                "participants": [p['email'] for p in thread.get('all_participants', [])]
            }
            api_resp = requests.post(
                RAILWAY_API_URL,
                json=api_payload,
                headers={"Content-Type": "application/json"},
                timeout=10
            )
            log.info(f"[API] Django API Response: {api_resp.status_code}")
        except Exception as api_err:
            log.error(f"[API] Failed to push data: {api_err}")

        if status in ("Resolved", "Forwarded"):
            return

        send_trigger_mail(thread, event_type="NEW_REPLY", ai_result=ai_result)
        set_thread_triggered(thread_id)
        log.info(f"[DONE] Trigger mail sent for: {subject}")

    except Exception as e:
        log.error(f"[ERROR] {e}", exc_info=True)


# ------------------------------------------------------------------ #
#  Overdue check loop                                                  #
# ------------------------------------------------------------------ #

def overdue_check_loop():
    log.info("[OVERDUE] Background loop started.")
    while True:
        try:
            time.sleep(30 * 60)

            overdue_hours = get_config_overdue_hours()
            pending       = get_pending_threads(overdue_hours)

            if not pending:
                continue

            log.info(f"[OVERDUE] Found {len(pending)} overdue thread(s).")

            for ts in pending:
                thread_id    = ts["thread_id"]
                last_triggered = ts.get("last_triggered", "")
                if last_triggered:
                    try:
                        trig_dt = datetime.fromisoformat(last_triggered.replace("Z", "+00:00"))
                        if trig_dt.tzinfo is None:
                            trig_dt = trig_dt.replace(tzinfo=timezone.utc)
                        hours_since = (datetime.now(timezone.utc) - trig_dt).total_seconds() / 3600
                        if hours_since < overdue_hours:
                            continue
                    except (ValueError, TypeError):
                        pass

                try:
                    thread = fetch_thread(thread_id)
                    if not thread:
                        continue

                    ai_result = classify_thread(
                        thread.get("full_thread_text", ""),
                        thread.get("subject", ""),
                    )

                    if ai_result.get("status") == "Pending Reply":
                        send_trigger_mail(thread, event_type="OVERDUE", ai_result=ai_result)
                        set_thread_triggered(thread_id)
                    else:
                        upsert_thread_state(
                            thread_id=thread_id,
                            msg_count=thread.get("msg_count", 0),
                            subject=thread.get("subject", ""),
                            last_status=ai_result.get("status", "Open"),
                            last_summary=ai_result.get("summary", ""),
                            last_department=ai_result.get("department", ""),
                            last_priority=ai_result.get("priority", "Medium"),
                            last_activity=thread.get("last_activity_at", ""),
                        )
                except Exception as e:
                    log.error(f"[OVERDUE] Error: {e}", exc_info=True)

            cleanup_old_processed(keep_last_n=5000)

        except Exception as e:
            log.error(f"[OVERDUE] Loop error: {e}", exc_info=True)


# ------------------------------------------------------------------ #
#  Health check                                                        #
# ------------------------------------------------------------------ #

@app.route("/health", methods=["GET"])
def health():
    return jsonify({
        "status": "running",
        "targets": get_config_targets(),
        "thread_filter": get_config_thread_filter() or "(all)",
        "overdue_hours": get_config_overdue_hours(),
    }), 200


# ------------------------------------------------------------------ #
#  Startup                                                             #
# ------------------------------------------------------------------ #

if __name__ == "__main__":
    port = int(os.getenv("PORT", 5000))
    log.info(f"[STARTUP] Server running on port {port}")
    app.run(host="0.0.0.0", port=port, debug=False)