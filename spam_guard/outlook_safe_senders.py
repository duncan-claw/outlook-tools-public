#!/usr/bin/env python3
"""Sync Outlook.com Safe Senders into the sender filter, via the browser.

Graph exposes no Safe Senders API for personal accounts. Outlook on the web
loads the list itself with `service.svc?action=GetMailboxJunkEmailConfiguration`;
this opens the junk-email options page in a NEW tab of the dedicated,
authenticated Chrome through its configured loopback CDP endpoint, captures that response, verifies the
signed-in account, and closes the tab. It never touches other tabs.

Stored as SHA-256 hashes only (addresses, and '@domain' entries) in the sender
filter's private state. Free-mail domain entries (e.g. gmail.com) are not
honoured, matching the local allowlist rule. Prints counts/fixed codes only.
"""
from __future__ import annotations

import asyncio
from datetime import datetime, timezone
import json
import sys
import time
import urllib.request

import websockets

from outlook import config
from spam_guard import leave_groups_via_cdp as cdp
from spam_guard import sender_filter as sf
from spam_guard.sender_identity import FREEMAIL, split_address

OPTIONS_URL = "https://outlook.live.com/mail/0/options/mail/junkEmail"
ACTION = "action=GetMailboxJunkEmailConfiguration"
SNAPSHOT = "outlook_safe_senders.json"
FRESH_FOR_S = 3600


class SyncError(Exception):
    pass


def find_lists(payload) -> dict:
    """Locate the junk configuration lists anywhere in the OWA response."""
    found = {}

    def walk(node):
        if isinstance(node, dict):
            for key, value in node.items():
                if key in ("TrustedSendersAndDomains", "ContactsTrusted", "TrustedListsOnly") and key not in found:
                    found[key] = value
                walk(value)
        elif isinstance(node, list):
            for value in node:
                walk(value)
    walk(payload)
    return found


def parse(payload, now=None) -> dict:
    lists = find_lists(payload)
    entries = lists.get("TrustedSendersAndDomains")
    if not isinstance(entries, list):
        raise SyncError("SAFE_SENDERS_NOT_FOUND")
    addresses, domains, ignored = set(), set(), 0
    for entry in entries:
        if not isinstance(entry, str):
            ignored += 1
            continue
        entry = entry.strip().lower()
        parts = split_address(entry)
        domain = entry.lstrip("@")
        if parts and not entry.startswith("@"):
            addresses.add(sf.digest("@".join(parts)))
        elif "@" not in domain and "." in domain and domain not in FREEMAIL:
            domains.add(sf.digest("@" + domain))
        else:
            ignored += 1
    return {
        "synced_at": (now or datetime.now(timezone.utc)).isoformat(),
        "contacts_trusted": lists.get("ContactsTrusted") is True,
        "trusted_lists_only": lists.get("TrustedListsOnly") is True,
        "addresses": sorted(addresses), "domains": sorted(domains),
        "entries": len(entries), "ignored": ignored,
    }


def is_fresh(now=None) -> bool:
    try:
        data = json.loads((sf.STATE / SNAPSHOT).read_text())
        age = (now or datetime.now(timezone.utc)) - datetime.fromisoformat(data["synced_at"])
        return age.total_seconds() < FRESH_FOR_S
    except (OSError, ValueError, KeyError, TypeError):
        return False


async def capture(port: int, timeout_s: float) -> dict:
    opener = cdp.bypass_opener()
    request = urllib.request.Request(f"http://127.0.0.1:{port}/json/new?about:blank", method="PUT")
    try:
        with opener.open(request, timeout=10) as response:
            tab = json.loads(response.read())
    except Exception:
        raise SyncError("CDP_UNREACHABLE") from None
    try:
        async with websockets.connect(tab["webSocketDebuggerUrl"], proxy=None, max_size=None) as ws:
            await cdp.cdp_send(ws, "Network.enable")
            await ws.send(json.dumps({"id": 900001, "method": "Page.navigate", "params": {"url": OPTIONS_URL}}))
            wanted, body_request, body = set(), None, None
            deadline = time.time() + timeout_s
            while time.time() < deadline and body is None:
                try:
                    message = json.loads(await asyncio.wait_for(ws.recv(), timeout=2))
                except asyncio.TimeoutError:
                    continue
                method = message.get("method")
                if method == "Network.responseReceived":
                    response = message["params"]["response"]
                    if ACTION in response.get("url", "") and response.get("status") == 200:
                        wanted.add(message["params"]["requestId"])
                elif method == "Network.loadingFinished" and message["params"]["requestId"] in wanted and body_request is None:
                    body_request = 900002
                    await ws.send(json.dumps({"id": body_request, "method": "Network.getResponseBody",
                                              "params": {"requestId": message["params"]["requestId"]}}))
                elif body_request is not None and message.get("id") == body_request:
                    body = (message.get("result") or {}).get("body")
            if body is None:
                text = await cdp.page_text(ws)
                raise SyncError("LOGIN_REQUIRED" if cdp.is_login_wall(text) else "SAFE_SENDERS_NOT_CAPTURED")
            identity = await cdp.verify_identity(ws)
            if identity is not None:
                raise SyncError(identity)
            try:
                return json.loads(body)
            except ValueError:
                raise SyncError("SAFE_SENDERS_INVALID") from None
    finally:
        try:
            opener.open(f"http://127.0.0.1:{port}/json/close/{tab['id']}", timeout=5)
        except Exception:
            pass


def sync(port=cdp.DEFAULT_PORT, timeout_s=60.0, force=False) -> dict:
    if not force and is_fresh():
        return {"result": "FRESH"}
    snapshot = parse(asyncio.run(capture(port, timeout_s)))
    sf.write_private(sf.STATE / SNAPSHOT, json.dumps(snapshot))
    return {"result": "SYNCED", "entries": snapshot["entries"], "addresses": len(snapshot["addresses"]),
            "domains": len(snapshot["domains"]), "ignored": snapshot["ignored"],
            "contacts_trusted": snapshot["contacts_trusted"], "trusted_lists_only": snapshot["trusted_lists_only"]}


def main(argv=None) -> int:
    args = argv if argv is not None else sys.argv[1:]
    force = "--force" in args
    lock = None
    if "--caller-holds-lock" not in args:
        # Share the Groups leaver's lock so two drivers never use Chrome at once.
        import fcntl
        lock = (config.GROUPS_STATE_DIR / "lock").open("w")
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            print('{"error":"BROWSER_BUSY"}')
            return 1
    try:
        print(json.dumps(sync(force=force)))
        return 0
    except SyncError as error:
        print(json.dumps({"error": str(error)}))
    except Exception:
        print('{"error":"SAFE_SENDERS_SYNC_FAILED"}')
    return 1


if __name__ == "__main__":
    sys.exit(main())
