#!/usr/bin/env python3
"""Sender-based Inbox filter: recoverable moves to Junk Email. Default dry-run.

Reads only message id, internetMessageId, from (address + display name),
receivedDateTime and parentFolderId. Never reads subjects, bodies, links or
attachments, and never sends sender data to an agent or LLM. Stdout carries
counts and fixed codes only; sender details go to a private human-viewable
report outside Git.

Safe list = own address + Outlook contacts + Outlook Safe Senders + private
local allowlist + people the account owner has sent mail to. Graph exposes no Outlook.com
Safe Senders API (verified 2026-10-18); outlook_safe_senders.py syncs a hashed
snapshot through the dedicated browser instead.

Modes (private config.json): off | dry-run | scored | strict.
  scored: safe-listed senders protected; others moved when suspicious or when
          an unknown sender uses a free personal mail domain.
  strict: every sender outside the safe list moved.

Learning: the account owner's own moves teach the filter. Moving a message Junk -> Inbox
marks its sender safe; moving one Inbox -> Junk marks its sender blocked. The
latest action per sender wins, and overrides contacts/allowlist (not self).
"""
from __future__ import annotations

import argparse
from collections import Counter
from datetime import datetime, timedelta, timezone
import fcntl
import hashlib
import json
import os
from pathlib import Path
import sys
import urllib.error
import urllib.parse
import urllib.request

from outlook import config
from outlook import outlook_personal_graph as auth
from spam_guard.group_calendar_guard import BASE, Client as CalendarClient, GuardError
from spam_guard.sender_identity import FREEMAIL, THRESHOLD, score_sender, split_address

ACCOUNT = config.ACCOUNT
STATE = config.STATE_DIR / "sender-filter"
MODES = ("off", "dry-run", "scored", "strict")
SOURCES = ("contacts", "local", "correspondents", "outlook_safe")
OUTLOOK_SAFE_STALE = timedelta(days=7)
SELECT = "id,from,internetMessageId,receivedDateTime,parentFolderId"
CORRESPONDENT_REFRESH = timedelta(hours=1)
CORRESPONDENT_MAX_AGE = timedelta(hours=72)
CORRESPONDENT_DAYS = 730
DEFAULT_MAX_MOVES = 25


class Client(CalendarClient):
    """Calendar guard client plus JSON bodies and immutable message IDs.

    Immutable IDs survive folder moves, so a recorded ID still finds the
    message in Junk for restore.
    """

    def request(self, method, path, body=None):
        url = path if path.startswith("https://") else BASE + path
        parsed = urllib.parse.urlsplit(url)
        if parsed.scheme != "https" or parsed.netloc != "graph.microsoft.com" or not parsed.path.startswith("/v1.0/me/"):
            raise GuardError("INVALID_API_URL")
        headers = {"Authorization": "Bearer " + self.token, "Accept": "application/json",
                   "Prefer": 'IdType="ImmutableId"'}
        data = None
        if body is not None:
            data = json.dumps(body).encode()
            headers["Content-Type"] = "application/json"
        req = urllib.request.Request(url, data=data, method=method, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=30, context=auth.http_context()) as response:
                raw = response.read()
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as error:
            raise GuardError("HTTP_" + str(error.code)) from None
        except Exception:
            raise GuardError("REQUEST_FAILED") from None


def digest(value: str) -> str:
    return hashlib.sha256(value.encode()).hexdigest()


def now_utc() -> datetime:
    return datetime.now(timezone.utc)


# ---------------------------------------------------------------- safe list

class SafeList:
    def __init__(self, addresses=(), domains=(), hashed=()):
        self.addresses = {a.strip().lower() for a in addresses}
        self.domains = {d.strip().lower().lstrip("@") for d in domains}
        self.hashed = set(hashed)
        self.outlook_addresses: set[str] = set()
        self.outlook_domains: set[str] = set()
        self.sources: dict[str, str] = {}

    def add(self, address, source):
        address = address.strip().lower()
        self.addresses.add(address)
        self.sources.setdefault(address, source)

    def check(self, address: str) -> str | None:
        if address == ACCOUNT:
            return "SAFE_SELF"
        if address in self.addresses:
            return self.sources.get(address, "SAFE_LOCAL_ALLOWLIST")
        domain = address.rsplit("@", 1)[-1]
        if any(domain == d or domain.endswith("." + d) for d in self.domains):
            return "SAFE_LOCAL_DOMAIN"
        if digest(address) in self.outlook_addresses:
            return "SAFE_OUTLOOK_SAFE_SENDER"
        labels = domain.split(".")
        if any(digest("@" + ".".join(labels[i:])) in self.outlook_domains for i in range(len(labels) - 1)):
            return "SAFE_OUTLOOK_SAFE_DOMAIN"
        if digest(address) in self.hashed:
            return "SAFE_CORRESPONDENT"
        return None


def parse_allowlist(text: str) -> tuple[set[str], set[str], int]:
    """Lines: 'name@example.com' or '@example.com' (domain and subdomains).

    Free-mail domains are ignored as domain entries: allowing @gmail.com would
    protect every attacker on Gmail.
    """
    addresses, domains, ignored = set(), set(), 0
    for line in text.splitlines():
        entry = line.split("#", 1)[0].strip().lower()
        if not entry:
            continue
        if entry.startswith("@"):
            if entry[1:] in FREEMAIL or "." not in entry[1:]:
                ignored += 1
            else:
                domains.add(entry[1:])
        elif split_address(entry):
            addresses.add(entry)
        else:
            ignored += 1
    return addresses, domains, ignored


def load_contacts(client) -> set[str]:
    found = set()
    roots = ["/me/contacts?$select=emailAddresses&$top=100"]
    for folder in client.pages("/me/contactFolders?$select=id&$top=100"):
        roots.append("/me/contactFolders/" + urllib.parse.quote(folder["id"], safe="")
                     + "/contacts?$select=emailAddresses&$top=100")
    for root in roots:
        for contact in client.pages(root):
            for entry in contact.get("emailAddresses") or []:
                parts = split_address(entry.get("address"))
                if parts:
                    found.add("@".join(parts))
    return found


def refresh_correspondents(client, now=None) -> int:
    """Cache SHA-256 hashes of addresses the account owner has emailed."""
    now = now or now_utc()
    since = (now - timedelta(days=CORRESPONDENT_DAYS)).strftime("%Y-%m-%dT%H:%M:%SZ")
    query = urllib.parse.urlencode({"$select": "toRecipients,ccRecipients,bccRecipients",
                                    "$top": 250, "$filter": "sentDateTime ge " + since}, safe="$,:")
    hashes = set()
    for message in client.pages("/me/mailFolders/sentitems/messages?" + query):
        for field in ("toRecipients", "ccRecipients", "bccRecipients"):
            for recipient in message.get(field) or []:
                parts = split_address((recipient.get("emailAddress") or {}).get("address"))
                if parts:
                    hashes.add(digest("@".join(parts)))
    write_private(STATE / "correspondents.json",
                  json.dumps({"refreshed_at": now.isoformat(), "hashes": sorted(hashes)}))
    return len(hashes)


def load_correspondents(now=None) -> set[str]:
    now = now or now_utc()
    try:
        data = json.loads((STATE / "correspondents.json").read_text())
        age = now - datetime.fromisoformat(data["refreshed_at"])
    except (OSError, ValueError, KeyError, TypeError):
        raise GuardError("CORRESPONDENTS_UNAVAILABLE") from None
    if age > CORRESPONDENT_MAX_AGE:
        raise GuardError("CORRESPONDENTS_STALE")
    return set(data.get("hashes") or [])


def correspondents_due(now=None) -> bool:
    try:
        data = json.loads((STATE / "correspondents.json").read_text())
        return (now or now_utc()) - datetime.fromisoformat(data["refreshed_at"]) > CORRESPONDENT_REFRESH
    except (OSError, ValueError, KeyError, TypeError):
        return True


def build_safe_list(client, sources, summary) -> SafeList:
    """Any source failure raises, so no message is moved on a partial list."""
    safe = SafeList()
    if "local" in sources:
        try:
            text = (STATE / "allowlist.txt").read_text()
        except FileNotFoundError:
            text = ""
        addresses, domains, ignored = parse_allowlist(text)
        for address in addresses:
            safe.add(address, "SAFE_LOCAL_ALLOWLIST")
        safe.domains |= domains
        summary["allowlist_ignored_entries"] = ignored
    if "contacts" in sources:
        for address in load_contacts(client):
            safe.add(address, "SAFE_CONTACT")
    if "correspondents" in sources:
        safe.hashed = load_correspondents()
    if "outlook_safe" in sources:
        try:
            data = json.loads((STATE / "outlook_safe_senders.json").read_text())
            age = now_utc() - datetime.fromisoformat(data["synced_at"])
        except (OSError, ValueError, KeyError, TypeError):
            raise GuardError("OUTLOOK_SAFE_SENDERS_UNAVAILABLE") from None
        # A stale snapshot still protects everything it lists; it only misses
        # recent additions, which a drag-back also teaches.
        summary["outlook_safe_stale"] = int(age > OUTLOOK_SAFE_STALE)
        safe.outlook_addresses = set(data.get("addresses") or [])
        safe.outlook_domains = set(data.get("domains") or [])
        summary["outlook_safe_entries"] = len(safe.outlook_addresses) + len(safe.outlook_domains)
    summary["safe_addresses"] = len(safe.addresses)
    summary["safe_domains"] = len(safe.domains)
    summary["safe_correspondents"] = len(safe.hashed)
    return safe


# ---------------------------------------------------------------- decisions

def sender_of(message) -> tuple[str | None, str | None]:
    email = (message.get("from") or {}).get("emailAddress") or {}
    return email.get("address"), email.get("name")


def message_key(message) -> str:
    return digest(message.get("internetMessageId") or message.get("id") or "")


def sender_hash(message) -> str | None:
    parts = split_address(sender_of(message)[0])
    return digest("@".join(parts)) if parts else None


def decide(message, safe: SafeList, mode: str, learned=None) -> dict:
    """Pure decision for one message. action is 'junk' or 'keep'."""
    address, name = sender_of(message)
    parts = split_address(address)
    if parts is None:
        return {"action": "keep", "codes": ["SKIP_MALFORMED_ADDRESS"], "score": 0}
    address = "@".join(parts)
    if address == ACCOUNT:
        return {"action": "keep", "codes": ["SAFE_SELF"], "score": 0}
    verdict = ((learned or {}).get(digest(address)) or {}).get("verdict")
    if verdict == "block":
        return {"action": "junk", "codes": ["USER_BLOCKED_SENDER"], "score": 0}
    if verdict == "safe":
        return {"action": "keep", "codes": ["SAFE_USER_RESCUED"], "score": 0}
    protected = safe.check(address)
    if protected:
        return {"action": "keep", "codes": [protected], "score": 0}
    result = score_sender(address, name)
    codes = result["signals"] + result["notes"]
    if mode == "strict":
        return {"action": "junk", "codes": ["STRICT_NOT_SAFE_LISTED"] + codes, "score": result["score"]}
    if result["suspicious"]:
        return {"action": "junk", "codes": ["SCORED_SUSPICIOUS"] + codes, "score": result["score"]}
    if parts[1] in FREEMAIL:
        # Legitimate correspondents are almost always contacts or
        # correspondents; strangers on personal free-mail are junk by default.
        return {"action": "junk", "codes": ["SCORED_FREEMAIL_UNKNOWN"] + codes, "score": result["score"]}
    return {"action": "keep", "codes": ["KEEP_BELOW_THRESHOLD"] + codes, "score": result["score"]}


# ---------------------------------------------------------------- state

def write_private(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(path.suffix + ".tmp")
    temp.write_text(text)
    os.chmod(temp, 0o600)
    os.replace(temp, path)


def append_private(path: Path, record: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a") as stream:
        os.chmod(path, 0o600)
        stream.write(json.dumps(record) + "\n")


LEDGER_ACTIONS = ("moved", "restored", "rescued", "user_junked")


def load_ledger() -> dict[str, str]:
    """Latest action per message key. Any entry means: never auto-moved again."""
    ledger = {}
    try:
        for line in (STATE / "ledger.jsonl").read_text().splitlines():
            record = json.loads(line)
            if record.get("action") in LEDGER_ACTIONS:
                ledger[record["key"]] = record["action"]
    except FileNotFoundError:
        pass
    return ledger


def load_learned() -> dict:
    try:
        return json.loads((STATE / "senders.json").read_text())
    except FileNotFoundError:
        return {}
    except ValueError:
        raise GuardError("INVALID_LEARNED_STATE") from None


def learn(inbox_messages, junk_messages, ledger, learned, summary, client=None, folders=None) -> None:
    """Learn from the account owner's own moves since the previous run.

    seen.json holds a snapshot {message key: folder, sender hash}. A message
    that was in Junk and is now in Inbox (or that we moved and is back in
    Inbox) was rescued: sender becomes safe. A message that was in Inbox and is
    now in Junk, not moved by us, was moved by the account owner: block sender.
    Mail Outlook delivers straight to Junk is never seen in Inbox, so it never
    teaches a block. A message that left Junk for any other folder (e.g. dragged
    straight into Personal) also counts as rescued; one that went to Deleted
    Items, Recoverable Items (Outlook emptying Junk) or was purged teaches nothing. Each transition is recorded once.
    """
    try:
        previous = json.loads((STATE / "seen.json").read_text())
    except FileNotFoundError:
        previous = {}
    except ValueError:
        raise GuardError("INVALID_SEEN_STATE") from None
    stamp = now_utc().isoformat()
    snapshot = {}

    def teach(key, sender, message, verdict, action, code):
        if ledger.get(key) == action:
            return
        learned[sender] = {"verdict": verdict, "at": stamp, "code": code}
        ledger[key] = action
        summary["learned_" + verdict] += 1
        append_private(STATE / "ledger.jsonl", {"at": stamp, "key": key, "action": action})
        address, name = sender_of(message)
        append_private(STATE / "learned.jsonl", {"at": stamp, "verdict": verdict, "code": code,
                                                 "from_address": address, "from_name": name})

    for folder, messages in (("inbox", inbox_messages), ("junk", junk_messages)):
        for message in messages:
            key, sender = message_key(message), sender_hash(message)
            if sender is None:
                continue
            snapshot[key] = {"f": folder, "s": sender, "id": message.get("id")}
            before = (previous.get(key) or {}).get("f")
            if folder == "inbox" and (before == "junk" or ledger.get(key) in ("moved", "user_junked")):
                teach(key, sender, message, "safe", "rescued", "USER_RESCUED")
            elif folder == "junk" and before == "inbox" and ledger.get(key) != "moved":
                teach(key, sender, message, "block", "user_junked", "USER_JUNKED")
    if client is not None and folders:
        # Left Junk but not for Inbox: find where it went (immutable IDs).
        for key, entry in previous.items():
            if entry.get("f") != "junk" or key in snapshot or not entry.get("id"):
                continue
            try:
                current = client.request("GET", "/me/messages/" + urllib.parse.quote(entry["id"], safe="")
                                         + "?$select=id,from,parentFolderId")
            except GuardError:
                continue  # purged or permanently deleted
            if current.get("parentFolderId") not in folders["not_rescue"]:
                teach(key, entry["s"], current, "safe", "rescued", "USER_RESCUED")
    write_private(STATE / "senders.json", json.dumps(learned))
    write_private(STATE / "seen.json", json.dumps(snapshot))


def load_config() -> dict:
    try:
        config = json.loads((STATE / "config.json").read_text())
    except FileNotFoundError:
        return {"mode": "off"}
    except ValueError:
        raise GuardError("INVALID_CONFIG") from None
    if config.get("mode") not in MODES or not set(config.get("sources", SOURCES)) <= set(SOURCES):
        raise GuardError("INVALID_CONFIG")
    return config


# ---------------------------------------------------------------- scan / move

def folder_id(client, name) -> str:
    folder = client.request("GET", "/me/mailFolders/" + name + "?$select=id")
    if not folder.get("id"):
        raise GuardError("FOLDER_UNAVAILABLE")
    return folder["id"]


def not_rescue_folders(client, junk) -> set:
    """Folders a message can leave Junk for without the account owner rescuing it.

    Outlook's own Junk clean-up moves mail into the hidden Recoverable Items
    folders, which are reachable by ID; that must never teach a safe sender.
    """
    ids = {junk, folder_id(client, "deleteditems"), folder_id(client, "recoverableitemsdeletions")}
    for name in ("recoverableitemspurges", "recoverableitemsversions"):
        try:
            ids.add(folder_id(client, name))
        except GuardError:
            pass
    return ids


def scan(client, mode="dry-run", apply=False, since=None, sources=SOURCES,
         max_moves=DEFAULT_MAX_MOVES, report=None, learn_moves=None) -> dict:
    if not ACCOUNT:
        raise GuardError("ACCOUNT_NOT_CONFIGURED")
    if apply and mode not in ("scored", "strict"):
        raise GuardError("APPLY_REQUIRES_POLICY")
    summary = Counter(checked=0, safe=0, malformed=0, already_processed=0,
                      would_junk_scored=0, would_junk_strict=0, matched=0, moved=0, errors=0,
                      learned_safe=0, learned_block=0)
    identity = client.request("GET", "/me/?$select=mail,userPrincipalName")
    if not any(str(identity.get(k, "")).casefold() == ACCOUNT for k in ("mail", "userPrincipalName")):
        raise GuardError("ACCOUNT_MISMATCH")
    safe = build_safe_list(client, sources, summary)
    inbox, junk = folder_id(client, "inbox"), folder_id(client, "junkemail")
    ledger = load_ledger()
    learned = load_learned()
    query = "?" + urllib.parse.urlencode({"$select": SELECT, "$top": 100}, safe="$,:")
    # Complete discovery before moving so paging cannot shift under us. The
    # whole Inbox is listed (for learning); only messages in the window act.
    inbox_messages = list(client.pages("/me/mailFolders/inbox/messages" + query))
    if learn_moves if learn_moves is not None else apply:
        junk_messages = list(client.pages("/me/mailFolders/junkemail/messages" + query))
        folders = {"not_rescue": not_rescue_folders(client, junk)}
        learn(inbox_messages, junk_messages, ledger, learned, summary, client, folders)
    reasons = Counter()
    candidates = []
    for message in inbox_messages:
        if since and str(message.get("receivedDateTime") or "") < since:
            continue
        summary["checked"] += 1
        if message_key(message) in ledger:
            summary["already_processed"] += 1
            continue
        scored = decide(message, safe, "scored", learned)
        strict = decide(message, safe, "strict", learned)
        first = scored["codes"][0]
        if first.startswith("SAFE_"):
            summary["safe"] += 1
        if first == "SKIP_MALFORMED_ADDRESS":
            summary["malformed"] += 1
        summary["would_junk_scored"] += scored["action"] == "junk"
        summary["would_junk_strict"] += strict["action"] == "junk"
        chosen = strict if mode == "strict" else scored
        reasons[chosen["codes"][0]] += 1
        for code in chosen["codes"][1:]:
            reasons[code] += 1
        if chosen["action"] == "junk":
            candidates.append(message["id"])
        if report is not None and chosen["action"] == "junk":
            address, name = sender_of(message)
            report.append({"received": message.get("receivedDateTime"), "from_address": address,
                           "from_name": name, "mode": mode, "score": chosen["score"],
                           "codes": chosen["codes"], "key": message_key(message)})
    summary["matched"] = len(candidates)
    if apply:
        for message_id in candidates:
            if summary["moved"] >= max_moves:
                summary["move_cap_reached"] = 1
                break
            try:
                if move_one(client, message_id, safe, mode, inbox, junk, learned):
                    summary["moved"] += 1
            except GuardError as error:
                summary["errors"] += 1
                append_private(STATE / "ledger.jsonl", {"at": now_utc().isoformat(), "action": "error", "code": str(error)})
    result = dict(summary)
    result.update(mode=mode, applied=apply, reasons=dict(sorted(reasons.items())))
    return result


def move_one(client, message_id, safe, mode, inbox, junk, learned=None) -> bool:
    """Re-read, re-decide, move to Junk Email, verify. Returns True if moved."""
    path = "/me/messages/" + urllib.parse.quote(message_id, safe="")
    current = client.request("GET", path + "?$select=" + SELECT)
    if current.get("parentFolderId") != inbox:
        return False
    decision = decide(current, safe, mode, learned)
    if decision["action"] != "junk":
        return False
    key = message_key(current)
    if key in load_ledger():
        return False
    append_private(STATE / "ledger.jsonl", {"at": now_utc().isoformat(), "key": key, "action": "move_intent"})
    moved = client.request("POST", path + "/move", {"destinationId": junk})
    new_id = moved.get("id")
    if not new_id:
        raise GuardError("MOVE_NOT_VERIFIED")
    check = client.request("GET", "/me/messages/" + urllib.parse.quote(new_id, safe="") + "?$select=id,parentFolderId")
    if check.get("parentFolderId") != junk:
        raise GuardError("MOVE_NOT_VERIFIED")
    stamp = now_utc().isoformat()
    append_private(STATE / "ledger.jsonl", {"at": stamp, "key": key, "action": "moved", "codes": decision["codes"]})
    address, name = sender_of(current)
    append_private(STATE / "moves.jsonl", {
        "at": stamp, "key": key, "message_id": new_id, "received": current.get("receivedDateTime"),
        "from_address": address, "from_name": name, "mode": mode,
        "score": decision["score"], "codes": decision["codes"]})
    return True


def restore(client, keys=None, since=None) -> dict:
    """Move recorded messages from Junk Email back to Inbox; never re-filtered."""
    junk, inbox = folder_id(client, "junkemail"), folder_id(client, "inbox")
    ledger = load_ledger()
    summary = Counter(selected=0, restored=0, not_in_junk=0, errors=0)
    try:
        records = [json.loads(line) for line in (STATE / "moves.jsonl").read_text().splitlines()]
    except FileNotFoundError:
        records = []
    for record in records:
        full_key = record["key"]
        if ledger.get(full_key) != "moved":
            continue
        if keys and not any(full_key.startswith(k.lower()) for k in keys if len(k) >= 8):
            continue
        if since and record["at"] < since:
            continue
        summary["selected"] += 1
        path = "/me/messages/" + urllib.parse.quote(record["message_id"], safe="")
        try:
            current = client.request("GET", path + "?$select=id,parentFolderId")
            if current.get("parentFolderId") != junk:
                summary["not_in_junk"] += 1
                continue
            moved = client.request("POST", path + "/move", {"destinationId": inbox})
            if not moved.get("id"):
                raise GuardError("RESTORE_NOT_VERIFIED")
            append_private(STATE / "ledger.jsonl", {"at": now_utc().isoformat(), "key": full_key, "action": "restored"})
            parts = split_address(record.get("from_address"))
            if parts:
                learned = load_learned()
                learned[digest("@".join(parts))] = {"verdict": "safe", "at": now_utc().isoformat(), "code": "USER_RESTORED"}
                write_private(STATE / "senders.json", json.dumps(learned))
            summary["restored"] += 1
        except GuardError:
            summary["errors"] += 1
    return dict(summary)


# ---------------------------------------------------------------- CLI

def add_allow(entry: str) -> str:
    entry = entry.strip().lower()
    addresses, domains, ignored = parse_allowlist(entry)
    if ignored or not (addresses or domains):
        raise GuardError("INVALID_ALLOWLIST_ENTRY")
    append = STATE / "allowlist.txt"
    append.parent.mkdir(parents=True, exist_ok=True)
    with append.open("a") as stream:
        os.chmod(append, 0o600)
        stream.write(entry + "\n")
    return "ALLOWLIST_ENTRY_ADDED"


def scheduled(client) -> dict:
    config = load_config()
    mode = config["mode"]
    if mode == "off":
        return {"mode": "off"}
    sources = tuple(config.get("sources", SOURCES))
    if "correspondents" in sources and correspondents_due():
        refresh_correspondents(client)
    since = config.get("since")
    if mode in ("scored", "strict") and not since:
        raise GuardError("LIVE_MODE_REQUIRES_SINCE")
    return scan(client, mode=mode, apply=mode in ("scored", "strict"), since=since,
                sources=sources, max_moves=int(config.get("max_moves", DEFAULT_MAX_MOVES)),
                learn_moves=True)


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    p_scan = sub.add_parser("scan", help="Classify Inbox senders; dry-run unless --apply")
    p_scan.add_argument("--mode", choices=("scored", "strict"), default="scored")
    p_scan.add_argument("--apply", action="store_true")
    p_scan.add_argument("--since", help="Only messages received at/after this UTC ISO time")
    p_scan.add_argument("--max-moves", type=int, default=DEFAULT_MAX_MOVES)
    p_scan.add_argument("--report", action="store_true", help="Write private preview.jsonl (outside Git)")
    sub.add_parser("scheduled", help="Run per private config.json (used by the launchd runner)")
    sub.add_parser("refresh-correspondents")
    p_restore = sub.add_parser("restore", help="Move recorded messages back to Inbox")
    group = p_restore.add_mutually_exclusive_group(required=True)
    group.add_argument("--key", action="append", help="key (or 8+ char prefix) from moves.jsonl")
    group.add_argument("--since", help="Restore moves recorded at/after this ISO time")
    p_explain = sub.add_parser("explain", help="Show how one sender would be treated, with reasons; moves nothing")
    p_explain.add_argument("address")
    p_explain.add_argument("name", nargs="?")
    p_allow = sub.add_parser("allow", help="Add address or @domain to the private allowlist")
    p_allow.add_argument("entry")
    p_enable = sub.add_parser("enable", help="Set scheduled mode; live modes start from now")
    p_enable.add_argument("mode", choices=MODES)
    p_enable.add_argument("--max-moves", type=int, default=DEFAULT_MAX_MOVES)
    p_enable.add_argument("--sources", default=",".join(SOURCES))
    args = parser.parse_args(argv)

    os.umask(0o077)
    STATE.mkdir(parents=True, exist_ok=True)
    with (STATE / "lock").open("w") as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            print('{"error":"ALREADY_RUNNING"}')
            return 1
        try:
            if args.command == "explain":
                client = Client()
                safe = build_safe_list(client, tuple(load_config().get("sources", SOURCES)), Counter())
                message = {"from": {"emailAddress": {"address": args.address, "name": args.name}}}
                decision = decide(message, safe, "scored", load_learned())
                parts = split_address(args.address)
                result = {"action": decision["action"], "codes": decision["codes"],
                          "score": decision["score"], "threshold": THRESHOLD}
                if parts:
                    result["identity_check"] = score_sender("@".join(parts), args.name)
            elif args.command == "allow":
                result = {"result": add_allow(args.entry)}
            elif args.command == "enable":
                sources = [s for s in args.sources.split(",") if s]
                if not set(sources) <= set(SOURCES):
                    raise GuardError("INVALID_SOURCES")
                config = {"mode": args.mode, "sources": sources, "max_moves": args.max_moves,
                          "since": now_utc().strftime("%Y-%m-%dT%H:%M:%SZ"), "enabled_at": now_utc().isoformat()}
                write_private(STATE / "config.json", json.dumps(config, indent=2))
                result = {"result": "CONFIG_WRITTEN", "mode": args.mode}
            elif args.command == "refresh-correspondents":
                result = {"correspondents": refresh_correspondents(Client())}
            elif args.command == "restore":
                result = restore(Client(), keys=args.key, since=args.since)
            elif args.command == "scheduled":
                result = scheduled(Client())
            else:
                report = [] if args.report else None
                result = scan(Client(), mode=args.mode, apply=args.apply, since=args.since,
                              max_moves=args.max_moves, report=report)
                if report is not None:
                    write_private(STATE / "preview.jsonl", "".join(json.dumps(r) + "\n" for r in report))
            print(json.dumps(result))
            return int(bool(result.get("errors")))
        except GuardError as error:
            print(json.dumps({"error": str(error)}))
            return 1
        except Exception:
            print('{"error":"INTERNAL_ERROR"}')
            return 1


if __name__ == "__main__":
    sys.exit(main())
