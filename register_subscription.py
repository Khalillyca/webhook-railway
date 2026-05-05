"""
register_subscription.py

Usage:
    python register_subscription.py register <webhook_url>
    python register_subscription.py list
    python register_subscription.py delete <subscription_id>
"""

import sys
import json
import logging
from datetime import datetime, timedelta, timezone

import requests

from auth import get_token

# ------------------------------------------------------------------
# Logging
# ------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s — %(message)s",
)

log = logging.getLogger(__name__)

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

# Must match webhook_server.py / your validation logic
CLIENT_STATE_SECRET = "my-outlook-tracker-secret-2026"


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------

def headers(token: str):
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def register_subscription(webhook_url: str):
    """
    Register Microsoft Graph webhook subscription.
    """

    token = get_token()

    # IMPORTANT:
    # Keep URL exactly as passed.
    # Do NOT strip trailing slash.
    webhook_url = webhook_url

    expires = (
        datetime.now(timezone.utc) + timedelta(days=3)
    ).replace(microsecond=0).isoformat().replace("+00:00", "Z")

    payload = {
        "changeType": "created,updated",
        "notificationUrl": webhook_url,
        "resource": "/me/mailFolders/Inbox/messages",
        "expirationDateTime": expires,
        "clientState": CLIENT_STATE_SECRET,
    }

    log.info("[SUBSCRIPTION] Registering subscription...")
    log.info(f"  Webhook URL : {webhook_url}")
    log.info(f"  Resource    : {payload['resource']}")
    log.info(f"  Expires     : {expires}")
    log.info(f"  Change types: {payload['changeType']}")

    response = requests.post(
        f"{GRAPH_BASE}/subscriptions",
        headers=headers(token),
        json=payload,
        timeout=30,
    )

    if response.status_code in (200, 201):
        data = response.json()

        log.info("[SUBSCRIPTION] Registration successful")
        log.info(f"  ID      : {data.get('id')}")
        log.info(f"  Expires : {data.get('expirationDateTime')}")

        print("\nSUCCESS")
        print(json.dumps(data, indent=2))

    else:
        log.error(
            f"[SUBSCRIPTION] Registration failed: {response.status_code}"
        )
        log.error(response.text)


def list_subscriptions():
    """
    List existing subscriptions.
    """

    token = get_token()

    response = requests.get(
        f"{GRAPH_BASE}/subscriptions",
        headers=headers(token),
        timeout=30,
    )

    if response.status_code == 200:
        data = response.json()

        print("\nACTIVE SUBSCRIPTIONS")
        print(json.dumps(data, indent=2))

    else:
        log.error(
            f"[SUBSCRIPTION] Failed to list subscriptions: {response.status_code}"
        )
        log.error(response.text)


def delete_subscription(subscription_id: str):
    """
    Delete a subscription.
    """

    token = get_token()

    response = requests.delete(
        f"{GRAPH_BASE}/subscriptions/{subscription_id}",
        headers=headers(token),
        timeout=30,
    )

    if response.status_code == 204:
        print(f"\nDeleted subscription: {subscription_id}")
    else:
        log.error(
            f"[SUBSCRIPTION] Failed to delete: {response.status_code}"
        )
        log.error(response.text)


# ------------------------------------------------------------------
# CLI
# ------------------------------------------------------------------

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        return

    cmd = sys.argv[1].lower()

    if cmd == "register":
        if len(sys.argv) < 3:
            print("Usage: python register_subscription.py register <webhook_url>")
            return

        webhook_url = sys.argv[2]
        register_subscription(webhook_url)

    elif cmd == "list":
        list_subscriptions()

    elif cmd == "delete":
        if len(sys.argv) < 3:
            print("Usage: python register_subscription.py delete <subscription_id>")
            return

        delete_subscription(sys.argv[2])

    else:
        print(__doc__)


if __name__ == "__main__":
    main()