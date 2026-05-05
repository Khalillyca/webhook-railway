"""
graph.py — Microsoft Graph API calls

Extracts EVERYTHING from a message:
  from, to, cc, bcc, subject, body, threadId,
  receivedDateTime, attachments, participants

Also builds complete thread object when a reply is detected.

Enhanced with:
  - Full mailbox search (Inbox + Sent + all folders)
  - Pagination for complete thread fetch
  - Detailed logging
"""

import requests
import logging
from typing import Optional
from auth import get_token

log = logging.getLogger(__name__)

GRAPH = "https://graph.microsoft.com/v1.0"


def _headers() -> dict:
    return {
        "Authorization": f"Bearer {get_token()}",
        "Content-Type": "application/json",
    }


# ------------------------------------------------------------------ #
#  FETCH SINGLE MESSAGE — full extraction                              #
# ------------------------------------------------------------------ #

def fetch_message(message_id: str) -> dict:
    """
    Fetch a single message with ALL fields.
    Returns a clean dict with everything extracted.
    """
    log.info(f"[GRAPH] Fetching message: {message_id[:40]}...")
    r = requests.get(
        f"{GRAPH}/me/messages/{message_id}",
        headers=_headers(),
        params={
            "$select": (
                "id,conversationId,subject,"
                "from,toRecipients,ccRecipients,bccRecipients,"
                "receivedDateTime,sentDateTime,"
                "body,bodyPreview,"
                "hasAttachments,importance,isRead,"
                "internetMessageId,parentFolderId"
            )
        },
    )
    r.raise_for_status()
    raw = r.json()
    msg = _extract_message(raw)
    log.info(f"[GRAPH] Message fetched: {msg['subject'][:60]}")
    return msg


def _extract_message(raw: dict) -> dict:
    """
    Parse raw Graph message into clean structured dict.
    Extracts from, to, cc, bcc, subject, body, thread ID — everything.
    """
    def _addr(obj: dict) -> dict:
        ea = obj.get("emailAddress", {})
        return {
            "name":  ea.get("name", ""),
            "email": ea.get("address", "").lower(),
        }

    def _addr_list(lst: list) -> list:
        return [_addr(item) for item in (lst or [])]

    # Body — prefer HTML, fallback to text
    body_obj = raw.get("body", {})
    body_type    = body_obj.get("contentType", "text")   # "html" or "text"
    body_content = body_obj.get("content", "")

    return {
        # IDs
        "message_id":         raw.get("id", ""),
        "thread_id":          raw.get("conversationId", ""),   # conversationId = threadId
        "internet_message_id": raw.get("internetMessageId", ""),

        # Addresses
        "from":               _addr(raw.get("from", {})),
        "to":                 _addr_list(raw.get("toRecipients", [])),
        "cc":                 _addr_list(raw.get("ccRecipients", [])),
        "bcc":                _addr_list(raw.get("bccRecipients", [])),

        # Content
        "subject":            raw.get("subject", "(no subject)"),
        "body_preview":       raw.get("bodyPreview", ""),
        "body_type":          body_type,
        "body":               body_content,

        # Meta
        "received_at":        raw.get("receivedDateTime", ""),
        "sent_at":            raw.get("sentDateTime", ""),
        "has_attachments":    raw.get("hasAttachments", False),
        "importance":         raw.get("importance", "normal"),
        "is_read":            raw.get("isRead", False),
        "folder_id":          raw.get("parentFolderId", ""),
    }


# ------------------------------------------------------------------ #
#  FETCH FULL THREAD — all messages in the conversation               #
#  Searches across ALL folders (Inbox + Sent + Drafts + etc.)         #
# ------------------------------------------------------------------ #

def fetch_thread(conversation_id: str) -> dict:
    """
    Fetch ALL messages in a thread (conversation) across the entire mailbox.
    Uses pagination to get complete conversation history.
    Returns complete thread object with reply chain.
    """
    log.info(f"[GRAPH] Fetching full thread: {conversation_id[:40]}...")

    all_messages = []
    url = f"{GRAPH}/me/messages"
    params = {
        "$filter":  f"conversationId eq '{conversation_id}'",
        "$select": (
            "id,conversationId,subject,"
            "from,toRecipients,ccRecipients,"
            "receivedDateTime,bodyPreview,body"
        ),
        "$orderby": "receivedDateTime asc",
        "$top":     100,
    }

    page = 1
    while url:
        r = requests.get(
            url,
            headers=_headers(),
            params=params if page == 1 else None,
        )

        # Personal Outlook accounts don't support conversationId $filter
        # Fall back to search if we get a 400
        if r.status_code == 400:
            log.warning(
                f"[GRAPH] conversationId filter not supported — "
                f"falling back to inbox search for thread {conversation_id[:30]}"
            )
            return _fetch_thread_fallback(conversation_id)

        r.raise_for_status()
        data = r.json()

        batch = data.get("value", [])
        all_messages.extend(batch)
        log.debug(f"[GRAPH] Thread page {page}: {len(batch)} messages")

        url = data.get("@odata.nextLink")
        page += 1

    if not all_messages:
        log.warning(f"[GRAPH] No messages found for thread {conversation_id[:40]}")
        return {}

    # Parse each message
    parsed = [_extract_message(m) for m in all_messages]

    # Collect all participants across the thread
    all_participants = {}
    for msg in parsed:
        for addr in [msg["from"]] + msg["to"] + msg["cc"]:
            if addr["email"]:
                all_participants[addr["email"]] = addr["name"]

    first = parsed[0]
    last  = parsed[-1]

    # Build full thread text (reply chain readable format)
    chain_text_parts = []
    for i, msg in enumerate(parsed):
        chain_text_parts.append(
            f"--- Message {i+1} | From: {msg['from']['name']} <{msg['from']['email']}> "
            f"| {msg['received_at']} ---\n"
            f"Subject: {msg['subject']}\n"
            f"{msg['body_preview']}"
        )
    full_thread_text = "\n\n".join(chain_text_parts)

    log.info(
        f"[GRAPH] Thread fetched: {first['subject'][:60]} | "
        f"{len(parsed)} messages across all folders"
    )

    return {
        # Thread identity
        "thread_id":          conversation_id,
        "subject":            first["subject"],

        # Original sender
        "from":               first["from"],

        # Stats
        "msg_count":          len(parsed),
        "has_reply":          len(parsed) > 1,

        # Latest activity
        "last_replier":       last["from"],
        "last_reply_preview": last["body_preview"],
        "last_activity_at":   last["received_at"],

        # First message
        "created_at":         first["received_at"],
        "first_message_id":   first["message_id"],
        "latest_message_id":  last["message_id"],

        # All participants
        "all_participants":   [
            {"name": name, "email": email}
            for email, name in all_participants.items()
        ],

        # Full reply chain
        "reply_chain":        parsed,
        "full_thread_text":   full_thread_text,
    }


# ------------------------------------------------------------------ #
#  FALLBACK: fetch thread via inbox search (personal Outlook)         #
# ------------------------------------------------------------------ #

def _fetch_thread_fallback(conversation_id: str) -> dict:
    """
    Fallback for personal Outlook accounts where conversationId $filter
    returns 400. Searches inbox for messages with matching conversationId
    using $search or iterates recent messages.
    """
    log.info(f"[GRAPH] Fallback: searching inbox for conversationId {conversation_id[:30]}...")

    all_messages = []
    url = f"{GRAPH}/me/mailFolders/inbox/messages"
    params = {
        "$select": (
            "id,conversationId,subject,"
            "from,toRecipients,ccRecipients,"
            "receivedDateTime,bodyPreview,body"
        ),
        "$orderby": "receivedDateTime desc",
        "$top": 50,
    }

    r = requests.get(url, headers=_headers(), params=params)
    if r.status_code != 200:
        log.error(f"[GRAPH] Fallback inbox fetch failed: {r.status_code}")
        return {}

    data = r.json()
    for msg in data.get("value", []):
        if msg.get("conversationId") == conversation_id:
            all_messages.append(msg)

    # Also check sent items
    sent_url = f"{GRAPH}/me/mailFolders/sentitems/messages"
    r2 = requests.get(sent_url, headers=_headers(), params=params)
    if r2.status_code == 200:
        for msg in r2.json().get("value", []):
            if msg.get("conversationId") == conversation_id:
                all_messages.append(msg)

    if not all_messages:
        log.warning(f"[GRAPH] Fallback: no messages found for {conversation_id[:30]}")
        return {}

    # Sort by received time
    all_messages.sort(key=lambda m: m.get("receivedDateTime", ""))

    parsed = [_extract_message(m) for m in all_messages]

    all_participants = {}
    for msg in parsed:
        for addr in [msg["from"]] + msg["to"] + msg["cc"]:
            if addr["email"]:
                all_participants[addr["email"]] = addr["name"]

    first = parsed[0]
    last  = parsed[-1]

    chain_text_parts = []
    for i, msg in enumerate(parsed):
        chain_text_parts.append(
            f"--- Message {i+1} | From: {msg['from']['name']} <{msg['from']['email']}> "
            f"| {msg['received_at']} ---\n"
            f"Subject: {msg['subject']}\n"
            f"{msg['body_preview']}"
        )
    full_thread_text = "\n\n".join(chain_text_parts)

    log.info(f"[GRAPH] Fallback thread fetched: {first['subject'][:60]} | {len(parsed)} messages")

    return {
        "thread_id":          conversation_id,
        "subject":            first["subject"],
        "from":               first["from"],
        "msg_count":          len(parsed),
        "has_reply":          len(parsed) > 1,
        "last_replier":       last["from"],
        "last_reply_preview": last["body_preview"],
        "last_activity_at":   last["received_at"],
        "created_at":         first["received_at"],
        "first_message_id":   first["message_id"],
        "latest_message_id":  last["message_id"],
        "all_participants":   [
            {"name": name, "email": email}
            for email, name in all_participants.items()
        ],
        "reply_chain":        parsed,
        "full_thread_text":   full_thread_text,
    }


# ------------------------------------------------------------------ #
#  CHECK IF MESSAGE IS FROM TARGET                                    #
# ------------------------------------------------------------------ #

def is_from_target(message: dict, target_emails: list) -> bool:
    """
    Check if a message involves any of the target emails/domains.
    Checks sender AND all recipients so replies are not skipped.
    Targets can be full emails or just domains.
    """
    all_addresses = (
        [message["from"]["email"]]
        + [a["email"] for a in message.get("to", [])]
        + [a["email"] for a in message.get("cc", [])]
    )

    for addr in all_addresses:
        addr = addr.lower()
        for target in target_emails:
            target = target.lower()
            if "@" in target:
                if addr == target:
                    return True
            else:
                if addr.endswith(f"@{target}"):
                    return True
    return False