"""
trigger.py — Send trigger mail as reply-all in the SAME thread

Uses Graph createReplyAll so it stays inside the same conversationId.
Builds a clean HTML notification with all extracted data.

Enhanced with:
  - AI classification data in trigger mails
  - OVERDUE event type with red alert template
  - Detailed logging
"""

import requests
import logging
from auth import get_token
from typing import Optional

log = logging.getLogger(__name__)
GRAPH = "https://graph.microsoft.com/v1.0"


def _headers() -> dict:
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type": "application/json",
    }


def send_trigger_mail(
    thread: dict,
    event_type: str = "NEW_REPLY",
    ai_result: Optional[dict] = None,
):
    """
    Sends a reply-all notification inside the same thread.

    Args:
        thread:     Full thread object from graph.fetch_thread()
        event_type: "NEW_REPLY" or "OVERDUE"
        ai_result:  Optional AI classification dict with summary, department, priority, status, reason
    """
    # Use first message ID for reply to ensure we stay in the original thread
    message_id = thread.get("first_message_id") or thread.get("latest_message_id")
    if not message_id:
        log.error("[TRIGGER] No message_id in thread — cannot send trigger mail.")
        return

    html = _build_html(thread, event_type, ai_result)

    try:
        # Step 1: Create reply-all draft
        log.info(f"[TRIGGER] Creating reply-all draft for: {thread.get('subject', '')[:60]}...")
        r = requests.post(
            f"{GRAPH}/me/messages/{message_id}/createReplyAll",
            headers=_headers(),
            json={},
        )

        # If first message fails, try latest
        if r.status_code != 200:
            latest_id = thread.get("latest_message_id")
            if latest_id and latest_id != message_id:
                log.warning(f"[TRIGGER] First message failed ({r.status_code}), trying latest...")
                r = requests.post(
                    f"{GRAPH}/me/messages/{latest_id}/createReplyAll",
                    headers=_headers(),
                    json={},
                )

        r.raise_for_status()
        draft_id = r.json()["id"]
        log.info(f"[TRIGGER] Created reply-all draft: {draft_id[:40]}...")

        # Step 2: Patch the draft body with our HTML
        r = requests.patch(
            f"{GRAPH}/me/messages/{draft_id}",
            headers=_headers(),
            json={
                "body": {
                    "contentType": "HTML",
                    "content": html,
                }
            },
        )
        if r.status_code != 200:
            log.error(f"[TRIGGER] Failed to patch draft body: {r.status_code} — {r.text}")
            r.raise_for_status()
        log.info("[TRIGGER] Patched draft body successfully.")

        # Step 3: Send the draft
        r = requests.post(
            f"{GRAPH}/me/messages/{draft_id}/send",
            headers=_headers(),
            json={},
        )
        if r.status_code not in (200, 202):
            log.error(f"[TRIGGER] Failed to send draft: {r.status_code} — {r.text}")
            r.raise_for_status()
        log.info(
            f"[TRIGGER] {event_type} mail sent for thread: "
            f"{thread.get('subject', '')[:60]}"
        )

    except requests.exceptions.HTTPError as e:
        log.error(f"[TRIGGER] HTTP error sending trigger: {e}", exc_info=True)
        raise
    except Exception as e:
        log.error(f"[TRIGGER] Unexpected error: {e}", exc_info=True)
        raise


def _build_html(
    thread: dict,
    event_type: str,
    ai_result: Optional[dict] = None,
) -> str:
    """Build HTML body for trigger mail."""

    # ── Color scheme based on event type ──────────────────────────── #
    if event_type == "OVERDUE":
        title_color = "#d93025"
        title_text  = "⏰ OVERDUE ALERT: Pending Too Long"
        subtitle    = "This thread requires immediate attention — no reply received within threshold"
        accent_bg   = "#fef0ef"
    else:
        title_color = "#1a73e8"
        title_text  = "⚡ LIVE TRIGGER: New Reply Detected"
        subtitle    = "A new message was added to this tracked thread"
        accent_bg   = "#f0f7ff"

    # ── AI classification section ─────────────────────────────────── #
    ai_html = ""
    if ai_result:
        priority = ai_result.get("priority", "Medium")
        status   = ai_result.get("status", "Open")

        # Priority color
        priority_colors = {
            "Critical": "#d93025",
            "High":     "#ea8600",
            "Medium":   "#1a73e8",
            "Low":      "#34a853",
        }
        p_color = priority_colors.get(priority, "#1a73e8")

        # Status color
        status_colors = {
            "Pending Reply": "#ea8600",
            "Open":          "#1a73e8",
            "Resolved":      "#34a853",
            "Forwarded":     "#9334e6",
        }
        s_color = status_colors.get(status, "#1a73e8")

        ai_html = f"""
        <div style="padding:14px 20px;background:#fafafa;border-bottom:1px solid #eee;">
          <div style="font-size:12px;color:#666;margin-bottom:10px;">
            <b>🤖 AI ANALYSIS</b>
          </div>
          <table style="width:100%;border-collapse:collapse;">
            <tr>
              <td style="padding:4px 0;width:120px;color:#666;font-size:12px;"><b>Summary</b></td>
              <td style="padding:4px 0;font-size:13px;color:#222;">{ai_result.get('summary', 'N/A')}</td>
            </tr>
            <tr>
              <td style="padding:4px 0;color:#666;font-size:12px;"><b>Status</b></td>
              <td style="padding:4px 0;">
                <span style="background:{s_color};color:white;padding:2px 10px;
                             border-radius:10px;font-size:11px;font-weight:bold;">
                  {status}
                </span>
              </td>
            </tr>
            <tr>
              <td style="padding:4px 0;color:#666;font-size:12px;"><b>Priority</b></td>
              <td style="padding:4px 0;">
                <span style="background:{p_color};color:white;padding:2px 10px;
                             border-radius:10px;font-size:11px;font-weight:bold;">
                  {priority}
                </span>
              </td>
            </tr>
            <tr>
              <td style="padding:4px 0;color:#666;font-size:12px;"><b>Department</b></td>
              <td style="padding:4px 0;font-size:12px;">{ai_result.get('department', 'General')}</td>
            </tr>
            <tr>
              <td style="padding:4px 0;color:#666;font-size:12px;"><b>Reason</b></td>
              <td style="padding:4px 0;font-size:12px;color:#555;font-style:italic;">
                {ai_result.get('reason', 'N/A')}
              </td>
            </tr>
          </table>
        </div>
        """

    # ── Format participants ───────────────────────────────────────── #
    participants_html = ", ".join(
        f"{p['name']} &lt;{p['email']}&gt;"
        for p in thread.get("all_participants", [])
    )

    # ── Format reply chain summary ────────────────────────────────── #
    chain = thread.get("reply_chain", [])
    chain_rows = ""
    for i, msg in enumerate(chain[-10:]):  # Last 10 messages max
        chain_rows += f"""
        <tr style="background:{'#f8f9fa' if i % 2 == 0 else 'white'}">
          <td style="padding:6px 10px;font-size:12px;color:#555;">#{i+1}</td>
          <td style="padding:6px 10px;font-size:12px;">
            <b>{msg['from']['name']}</b> &lt;{msg['from']['email']}&gt;
          </td>
          <td style="padding:6px 10px;font-size:12px;color:#777;">{msg['received_at'][:19].replace('T', ' ')}</td>
          <td style="padding:6px 10px;font-size:12px;color:#444;">{msg['body_preview'][:100]}...</td>
        </tr>
        """

    # ── Latest reply info ─────────────────────────────────────────── #
    last_replier = thread.get("last_replier", {})

    return f"""
    <div style="font-family:Arial,sans-serif;max-width:700px;border:1px solid #ddd;
                border-radius:8px;overflow:hidden;margin:10px 0;">

      <!-- Header -->
      <div style="background:{title_color};color:white;padding:16px 20px;">
        <div style="font-size:17px;font-weight:bold;">{title_text}</div>
        <div style="font-size:12px;margin-top:4px;opacity:0.9;">{subtitle}</div>
      </div>

      <!-- Thread Summary -->
      <div style="padding:16px 20px;border-bottom:1px solid #eee;">
        <table style="width:100%;border-collapse:collapse;">
          <tr>
            <td style="padding:5px 0;width:160px;color:#666;font-size:13px;"><b>Subject</b></td>
            <td style="padding:5px 0;font-size:13px;">{thread.get('subject', '')}</td>
          </tr>
          <tr>
            <td style="padding:5px 0;color:#666;font-size:13px;"><b>Thread ID</b></td>
            <td style="padding:5px 0;font-size:12px;color:#888;font-family:monospace;">
              {thread.get('thread_id', '')[:40]}...
            </td>
          </tr>
          <tr>
            <td style="padding:5px 0;color:#666;font-size:13px;"><b>Original From</b></td>
            <td style="padding:5px 0;font-size:13px;">
              {thread.get('from', {}).get('name', '')}
              &lt;{thread.get('from', {}).get('email', '')}&gt;
            </td>
          </tr>
          <tr>
            <td style="padding:5px 0;color:#666;font-size:13px;"><b>Total Messages</b></td>
            <td style="padding:5px 0;font-size:13px;font-weight:bold;color:{title_color};">
              {thread.get('msg_count', 0)}
            </td>
          </tr>
          <tr>
            <td style="padding:5px 0;color:#666;font-size:13px;"><b>Latest Reply By</b></td>
            <td style="padding:5px 0;font-size:13px;">
              <b>{last_replier.get('name', '')}</b>
              &lt;{last_replier.get('email', '')}&gt;
            </td>
          </tr>
          <tr>
            <td style="padding:5px 0;color:#666;font-size:13px;"><b>Last Activity</b></td>
            <td style="padding:5px 0;font-size:13px;">
              {str(thread.get('last_activity_at', ''))[:19].replace('T', ' ')} UTC
            </td>
          </tr>
          <tr>
            <td style="padding:5px 0;color:#666;font-size:13px;"><b>All Participants</b></td>
            <td style="padding:5px 0;font-size:12px;color:#555;">{participants_html}</td>
          </tr>
        </table>
      </div>

      {ai_html}

      <!-- Latest Reply Snippet -->
      <div style="padding:14px 20px;background:{accent_bg};border-bottom:1px solid #eee;">
        <div style="font-size:12px;color:#666;margin-bottom:6px;"><b>LATEST REPLY SNIPPET</b></div>
        <div style="font-size:13px;color:#333;font-style:italic;line-height:1.5;">
          "{thread.get('last_reply_preview', '')[:300]}..."
        </div>
      </div>

      <!-- Reply Chain Table -->
      <div style="padding:14px 20px;">
        <div style="font-size:12px;color:#666;margin-bottom:8px;"><b>REPLY CHAIN</b></div>
        <table style="width:100%;border-collapse:collapse;border:1px solid #eee;border-radius:4px;">
          <tr style="background:#f1f3f4;">
            <th style="padding:6px 10px;font-size:11px;text-align:left;color:#666;">#</th>
            <th style="padding:6px 10px;font-size:11px;text-align:left;color:#666;">FROM</th>
            <th style="padding:6px 10px;font-size:11px;text-align:left;color:#666;">TIME</th>
            <th style="padding:6px 10px;font-size:11px;text-align:left;color:#666;">PREVIEW</th>
          </tr>
          {chain_rows}
        </table>
      </div>

      <!-- Footer -->
      <div style="padding:10px 20px;background:#f8f9fa;font-size:11px;color:#999;
                  border-top:1px solid #eee;text-align:center;">
        Sent by Outlook Thread Tracker — automated notification
      </div>
    </div>
    """