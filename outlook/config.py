"""Generic, environment-backed settings for Outlook Tools.

Set values locally; never commit a file containing real account details or tokens.
"""
from __future__ import annotations

import os
from pathlib import Path

ACCOUNT = os.environ.get("OUTLOOK_PERSONAL_USER", "").strip().casefold()
GRAPH_CLIENT_ID = os.environ.get("OUTLOOK_PERSONAL_GRAPH_CLIENT_ID", "").strip()
TENANT = os.environ.get("OUTLOOK_PERSONAL_GRAPH_TENANT", "consumers").strip()
SCOPES = os.environ.get(
    "OUTLOOK_PERSONAL_GRAPH_SCOPES",
    "offline_access User.Read Mail.ReadWrite Mail.Send MailboxSettings.ReadWrite "
    "Calendars.Read Calendars.ReadWrite Contacts.Read",
).strip()
TIMEZONE = os.environ.get("OUTLOOK_TIMEZONE", "UTC").strip()
HISTORY_CUTOFF = os.environ.get("OUTLOOK_HISTORY_CUTOFF", "").strip()
GROUPS_CDP_PORT = int(os.environ.get("OUTLOOK_GROUPS_CDP_PORT", "9222"))
STATE_DIR = Path(os.environ.get(
    "OUTLOOK_TOOLS_STATE_DIR",
    str(Path.home() / ".local" / "state" / "outlook-tools"),
)).expanduser()
GROUPS_STATE_DIR = Path(os.environ.get(
    "OUTLOOK_GROUPS_STATE_DIR", str(STATE_DIR / "groups-leaver")
)).expanduser()


def require_account() -> str:
    if not ACCOUNT:
        raise RuntimeError("OUTLOOK_PERSONAL_USER_REQUIRED")
    return ACCOUNT


def require_graph_client_id() -> str:
    if not GRAPH_CLIENT_ID:
        raise RuntimeError("OUTLOOK_PERSONAL_GRAPH_CLIENT_ID_REQUIRED")
    return GRAPH_CLIENT_ID
