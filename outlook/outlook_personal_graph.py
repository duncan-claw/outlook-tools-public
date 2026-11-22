#!/usr/bin/env python3
"""OAuth-backed Microsoft Graph helper for a personal Outlook mailbox.

Credentials and tokens stay outside the repository in the configured state directory.
"""
from __future__ import annotations

import argparse
import fcntl
import tempfile
import json
import os
import ssl
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path

from outlook import config

CLIENT_ID = config.GRAPH_CLIENT_ID
TENANT = config.TENANT
SCOPES = config.SCOPES
STATE_DIR = config.STATE_DIR / "oauth-graph"
TOKEN_PATH = STATE_DIR / "token.json"
DEVICE_URL = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/devicecode"
TOKEN_URL = f"https://login.microsoftonline.com/{TENANT}/oauth2/v2.0/token"
GRAPH_BASE = "https://graph.microsoft.com/v1.0"


def http_context() -> ssl.SSLContext:
    """Use the configured trust bundle when HTTPS inspection is enabled."""
    bundle = os.environ.get("SSL_CERT_FILE") or os.environ.get("REQUESTS_CA_BUNDLE")
    if bundle and os.path.exists(bundle):
        ctx = ssl.create_default_context(cafile=bundle)
    else:
        ctx = ssl.create_default_context()
    ctx.verify_flags = ssl.VERIFY_X509_PARTIAL_CHAIN
    return ctx


def post_form(url: str, data: dict[str, str]) -> dict:
    req = urllib.request.Request(url, data=urllib.parse.urlencode(data).encode(), headers={"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req, timeout=30, context=http_context()) as response:
            return json.loads(response.read().decode())
    except urllib.error.HTTPError as error:
        try:
            return json.loads(error.read().decode())
        except Exception:
            raise RuntimeError(f"HTTP {error.code}: {error.reason}") from error


def load_token() -> dict | None:
    try:
        return json.loads(TOKEN_PATH.read_text())
    except FileNotFoundError:
        return None


def save_token(token: dict) -> None:
    STATE_DIR.mkdir(mode=0o700, parents=True, exist_ok=True)
    os.chmod(STATE_DIR, 0o700)
    token = dict(token)
    if "expires_in" in token:
        token["expires_at"] = int(time.time()) + int(token.get("expires_in", 0))
    fd, name = tempfile.mkstemp(prefix=".token-", dir=STATE_DIR)
    try:
        with os.fdopen(fd, "w") as stream:
            os.fchmod(stream.fileno(), 0o600)
            json.dump(token, stream, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(name, TOKEN_PATH)
    finally:
        if os.path.exists(name):
            os.unlink(name)


def token_is_fresh(token: dict) -> bool:
    return bool(token.get("access_token")) and int(token.get("expires_at", 0)) > int(time.time()) + 120


def _refresh_token_locked(token: dict) -> dict | None:
    client_id = config.require_graph_client_id()
    refresh = token.get("refresh_token")
    if not refresh:
        return None
    refreshed = post_form(TOKEN_URL, {
        "client_id": client_id,
        "grant_type": "refresh_token",
        "refresh_token": refresh,
        "scope": " ".join(sorted(set(SCOPES.split()) | set(token.get("scope", "").split()))),
    })
    if "access_token" not in refreshed:
        return None
    if "refresh_token" not in refreshed:
        refreshed["refresh_token"] = refresh
    save_token(refreshed)
    return refreshed


def refresh_token(token: dict) -> dict | None:
    STATE_DIR.mkdir(mode=0o700, parents=True, exist_ok=True)
    os.chmod(STATE_DIR, 0o700)
    with (STATE_DIR / "refresh.lock").open("a") as lock:
        os.chmod(STATE_DIR / "refresh.lock", 0o600)
        fcntl.flock(lock, fcntl.LOCK_EX)
        current = load_token() or token
        if token_is_fresh(current):
            return current
        return _refresh_token_locked(current)


def device_login() -> dict:
    client_id = config.require_graph_client_id()
    device = post_form(DEVICE_URL, {"client_id": client_id, "scope": SCOPES})
    if "device_code" not in device:
        raise RuntimeError(json.dumps(device, indent=2))
    print(device.get("message") or f"Open {device.get('verification_uri')} and enter {device.get('user_code')}", flush=True)
    interval = int(device.get("interval", 5))
    deadline = time.time() + int(device.get("expires_in", 900))
    while time.time() < deadline:
        time.sleep(interval)
        token = post_form(TOKEN_URL, {
            "client_id": client_id,
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "device_code": device["device_code"],
        })
        error = token.get("error")
        if "access_token" in token:
            save_token(token)
            return token
        if error == "authorization_pending":
            print("waiting_for_sign_in", flush=True)
            continue
        if error == "slow_down":
            interval += 5
            continue
        raise RuntimeError(json.dumps(token, indent=2))
    raise TimeoutError("Device code expired")


def get_token(interactive: bool = True) -> dict:
    token = load_token()
    if token_is_fresh(token or {}):
        return token  # type: ignore[return-value]
    if token:
        refreshed = refresh_token(token)
        if refreshed:
            return refreshed
    if not interactive:
        raise RuntimeError("No valid cached token. Run: outlook_personal_graph.py login")
    return device_login()


def graph_request(method: str, path: str, body: dict | None = None, query: dict | None = None) -> dict:
    token = get_token(interactive=False)
    url = path if path.startswith("https://") else GRAPH_BASE + path
    if query:
        url += ("&" if "?" in url else "?") + urllib.parse.urlencode(query, doseq=True, safe="(),:$'")
    data = None if body is None else json.dumps(body).encode()
    headers = {"Authorization": "Bearer " + token["access_token"], "Accept": "application/json"}
    if body is not None:
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30, context=http_context()) as response:
            # Graph returns 202/204 with an empty body for accepted async sends
            # (e.g. /me/sendMail), so never assume JSON.
            raw = response.read().decode(errors="replace")
            if not raw.strip():
                return {"ok": True, "status": response.status}
            try:
                return json.loads(raw)
            except json.JSONDecodeError:
                return {"ok": True, "status": response.status, "raw": raw[:500]}
    except urllib.error.HTTPError as error:
        text = error.read().decode(errors="replace")
        try:
            detail = json.loads(text)
        except Exception:
            detail = text
        raise RuntimeError(json.dumps({"status": error.code, "error": detail}, indent=2)) from error


def status() -> dict:
    token = load_token()
    if not token:
        return {"tokenPath": str(TOKEN_PATH), "authenticated": False, "reason": "no token", "scopes": SCOPES}
    if token_is_fresh(token):
        return {"tokenPath": str(TOKEN_PATH), "authenticated": True, "mode": "cached", "expiresAt": token.get("expires_at"), "scopes": SCOPES}
    refreshed = refresh_token(token)
    if refreshed:
        return {"tokenPath": str(TOKEN_PATH), "authenticated": True, "mode": "refreshed", "expiresAt": refreshed.get("expires_at"), "scopes": SCOPES}
    return {"tokenPath": str(TOKEN_PATH), "authenticated": False, "reason": "refresh failed", "scopes": SCOPES}


def simplify_message(m: dict) -> dict:
    return {
        "id": m.get("id"),
        "receivedDateTime": m.get("receivedDateTime"),
        "from": ((m.get("from") or {}).get("emailAddress") or {}),
        "subject": m.get("subject"),
        "isRead": m.get("isRead"),
        "categories": m.get("categories") or [],
        "bodyPreview": m.get("bodyPreview"),
    }


def list_messages(limit: int, category: str | None = None) -> dict:
    query = {
        "$top": str(limit),
        "$orderby": "receivedDateTime desc",
        "$select": "id,receivedDateTime,from,subject,isRead,categories,bodyPreview",
    }
    if category:
        escaped = category.replace("'", "''")
        query["$filter"] = f"categories/any(c:c eq '{escaped}')"
    data = graph_request("GET", "/me/messages", query=query)
    return {"messages": [simplify_message(m) for m in data.get("value", [])], "nextLink": data.get("@odata.nextLink")}


def read_message(message_id: str, body_format: str) -> dict:
    query = {"$select": "id,receivedDateTime,from,toRecipients,ccRecipients,subject,isRead,categories,body,bodyPreview,internetMessageId"}
    data = graph_request("GET", f"/me/messages/{urllib.parse.quote(message_id, safe='')}", query=query)
    body = data.get("body") or {}
    result = simplify_message(data)
    result["toRecipients"] = data.get("toRecipients") or []
    result["ccRecipients"] = data.get("ccRecipients") or []
    result["internetMessageId"] = data.get("internetMessageId")
    if body_format in {"text", "full"}:
        result["bodyText"] = body.get("content") if body.get("contentType") == "text" else data.get("bodyPreview")
    if body_format in {"html", "full"}:
        result["body"] = body
    return {"message": result}


def list_categories() -> dict:
    data = graph_request("GET", "/me/outlook/masterCategories")
    return {"categories": data.get("value", [])}


def create_category(name: str, color: str) -> dict:
    return graph_request("POST", "/me/outlook/masterCategories", body={"displayName": name, "color": color})


def set_message_categories(message_id: str, categories: list[str], mode: str) -> dict:
    if mode == "add":
        current = graph_request("GET", f"/me/messages/{urllib.parse.quote(message_id, safe='')}", query={"$select": "id,categories"})
        merged = list(dict.fromkeys((current.get("categories") or []) + categories))
        categories = merged
    return graph_request("PATCH", f"/me/messages/{urllib.parse.quote(message_id, safe='')}", body={"categories": categories})


def list_calendars() -> dict:
    data = graph_request("GET", "/me/calendars")
    return {"calendars": data.get("value", []) if data else []}

def list_calendar_groups() -> dict:
    """Fetch calendar groups (My Calendars, People's Calendars, etc.) and the calendars within each."""
    groups_data = graph_request("GET", "/me/calendargroups", query={"$select": "id,name"})
    groups = groups_data.get("value", []) if groups_data else []
    result = {"groups": []}
    for g in groups:
        group_id = g.get("id")
        cals_data = graph_request("GET", f"/me/calendargroups/{group_id}/calendars", query={"$select": "id,name"})
        cals = cals_data.get("value", []) if cals_data else []
        result["groups"].append({
            "id": group_id,
            "name": g.get("name"),
            "calendars": cals,
        })
    # Also fetch orphaned calendars (not in a group)
    orphaned_data = graph_request("GET", "/me/calendars", query={"$select": "id,name"})
    result["orphaned_calendars"] = orphaned_data.get("value", []) if orphaned_data else []
    return result


def send_message(subject: str, to: str, body: str, body_format: str = "html") -> dict:
    payload = {
        "subject": subject,
        "body": {"contentType": "html" if body_format == "html" else "text", "content": body},
        "toRecipients": [{"emailAddress": {"address": to}}],
    }
    graph_request("POST", "/me/sendMail", body={"message": payload, "saveToSentItems": True})
    return {"ok": True, "to": to, "subject": subject}


def list_calendar_events(limit: int, calendar_id: str | None = None, days_ahead: int = 30) -> dict:
    now = datetime.now(timezone.utc)
    end = now + timedelta(days=days_ahead)
    # Graph API calendarview endpoint requires startDateTime and endDateTime as query params (no $ prefix)
    if calendar_id:
        cal_path = f"/me/calendars/{urllib.parse.quote(calendar_id, safe='')}/calendarview"
    else:
        cal_path = "/me/calendarview"
    query = {
        "startDateTime": now.isoformat(),
        "endDateTime": end.isoformat(),
        "$top": str(limit),
        "$orderby": "start/dateTime asc",
        "$select": "id,subject,start,end,location,isAllDay,isCancelled,organizer",
    }
    data = graph_request("GET", cal_path, query=query)
    return {"events": data.get("value", []), "nextLink": data.get("@odata.nextLink")}


def create_calendar_event(
    subject: str,
    start: str,
    end: str | None = None,
    calendar_id: str | None = None,
    location: str | None = None,
    body: str | None = None,
    is_all_day: bool = False,
    private: bool = False,
    tz: str = "Australia/Melbourne",
    categories: list[str] | None = None,
) -> dict:
    """Create a calendar event. `start`/`end` accept ISO 8601 (with or without offset).

    If `end` is omitted, defaults to start + 30 minutes (or +1 day for all-day).
    `private=True` sets the event sensitivity to private so it is not shared with
    delegates and does not expose its detail to others. Naive datetimes are
    interpreted in `tz`.
    """
    def shape(value: str, duration_fallback: timedelta) -> dict:
        raw = value.strip()
        try:
            parsed = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        except ValueError as error:
            raise RuntimeError(f"invalid ISO 8601 datetime: {value!r}") from error
        if parsed.tzinfo is None:
            return {"dateTime": parsed.isoformat(), "timeZone": tz}
        return {"dateTime": parsed.isoformat(), "timeZone": "UTC"}

    start_shape = shape(start, timedelta(minutes=30))
    if end:
        end_shape = shape(end, timedelta(minutes=30))
    else:
        raw_start = datetime.fromisoformat(start.strip().replace("Z", "+00:00"))
        delta = timedelta(days=1) if is_all_day else timedelta(minutes=30)
        raw_end = raw_start + delta
        extra = "" if raw_start.tzinfo is None else "+00:00"
        end_shape = {"dateTime": raw_end.isoformat() + extra, "timeZone": start_shape["timeZone"]}

    payload: dict = {"subject": subject, "start": start_shape, "end": end_shape, "isAllDay": is_all_day}
    if location:
        payload["location"] = {"displayName": location}
    if body:
        payload["body"] = {"contentType": "text", "content": body}
    if categories:
        payload["categories"] = [c for c in categories if c]
    if private:
        payload["sensitivity"] = "private"

    path = f"/me/calendars/{urllib.parse.quote(calendar_id, safe='')}/events" if calendar_id else "/me/events"
    created = graph_request("POST", path, body=payload)
    return {
        "ok": True,
        "id": created.get("id"),
        "subject": created.get("subject"),
        "start": (created.get("start") or {}).get("dateTime"),
        "end": (created.get("end") or {}).get("dateTime"),
        "categories": created.get("categories") or [],
        "sensitivity": created.get("sensitivity"),
        "isAllDay": created.get("isAllDay"),
        "webLink": created.get("webLink"),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Outlook.com personal Graph helper for categories")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("login")
    sub.add_parser("status")

    mail = sub.add_parser("mail")
    mail_sub = mail.add_subparsers(dest="mail_command", required=True)
    mail_list = mail_sub.add_parser("list")
    mail_list.add_argument("--limit", "-n", type=int, default=10)
    mail_list.add_argument("--category")
    mail_read = mail_sub.add_parser("read")
    mail_read.add_argument("id")
    mail_read.add_argument("--format", choices=["text", "html", "full"], default="text")
    mail_cat = mail_sub.add_parser("categorize")
    mail_cat.add_argument("id")
    mail_cat.add_argument("categories", nargs="+")
    mail_cat.add_argument("--mode", choices=["add", "replace"], default="add")

    cal = sub.add_parser("calendar")
    cal_sub = cal.add_subparsers(dest="calendar_command", required=True)
    cal_list = cal_sub.add_parser("list")
    cal_list.add_argument("--limit", "-n", type=int, default=20)
    cal_list.add_argument("--calendar")
    cal_list.add_argument("--days", "-d", type=int, default=30, help="Look ahead this many days from now")
    cal_sub.add_parser("calendars")
    cal_sub.add_parser("groups")
    cal_add = cal_sub.add_parser("add", help="Create an event")
    cal_add.add_argument("--subject", "-s", required=True)
    cal_add.add_argument("--start", required=True, help="ISO 8601, e.g. 2026-10-19T09:00:00 or with +11:00")
    cal_add.add_argument("--end", help="ISO 8601; defaults to start + 30m")
    cal_add.add_argument("--calendar", help="Target calendar id (default: primary)")
    cal_add.add_argument("--location", "-l")
    cal_add.add_argument("--body", "-b", help="Event notes/body")
    cal_add.add_argument("--all-day", action="store_true")
    cal_add.add_argument("--tz", default="Australia/Melbourne", help="Timezone for naive datetimes")
    cal_add.add_argument("--category", "-c", action="append", default=None,
                         help="Category name to apply (repeatable), e.g. --category reminder")

    mail = sub.choices.get("mail") if hasattr(sub, "choices") else None
    mail_send = mail_sub.add_parser("send", help="Send mail from this account")
    mail_send.add_argument("--to", required=True)
    mail_send.add_argument("--subject", "-s", required=True)
    mail_send.add_argument("--body", "-b", required=True)
    mail_send.add_argument("--format", choices=["text", "html"], default="html")

    cats = sub.add_parser("categories")
    cats_sub = cats.add_subparsers(dest="categories_command", required=True)
    cats_sub.add_parser("list")
    cat_create = cats_sub.add_parser("create")
    cat_create.add_argument("name")
    cat_create.add_argument("--color", default="preset0", help="Graph category color, e.g. preset0..preset24 or none")

    args = parser.parse_args()
    if args.command == "login":
        token = device_login()
        print(json.dumps({"ok": True, "tokenPath": str(TOKEN_PATH), "expiresAt": token.get("expires_at")}, indent=2))
    elif args.command == "status":
        print(json.dumps(status(), indent=2))
    elif args.command == "mail" and args.mail_command == "list":
        print(json.dumps(list_messages(args.limit, args.category), indent=2))
    elif args.command == "mail" and args.mail_command == "read":
        print(json.dumps(read_message(args.id, args.format), indent=2))
    elif args.command == "mail" and args.mail_command == "categorize":
        print(json.dumps(set_message_categories(args.id, args.categories, args.mode), indent=2))
    elif args.command == "mail" and args.mail_command == "send":
        print(json.dumps(send_message(args.subject, args.to, args.body, args.format), indent=2))
    elif args.command == "categories" and args.categories_command == "list":
        print(json.dumps(list_categories(), indent=2))
    elif args.command == "categories" and args.categories_command == "create":
        print(json.dumps(create_category(args.name, args.color), indent=2))
    elif args.command == "calendar" and args.calendar_command == "list":
        print(json.dumps(list_calendar_events(args.limit, args.calendar, args.days), indent=2))
    elif args.command == "calendar" and args.calendar_command == "calendars":
        print(json.dumps(list_calendars(), indent=2))
    elif args.command == "calendar" and args.calendar_command == "groups":
        print(json.dumps(list_calendar_groups(), indent=2))
    elif args.command == "calendar" and args.calendar_command == "add":
        print(json.dumps(create_calendar_event(
            subject=args.subject,
            start=args.start,
            end=args.end,
            calendar_id=args.calendar,
            location=args.location,
            body=args.body,
            is_all_day=args.all_day,
            tz=args.tz,
            categories=args.category,
        ), indent=2))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"ERROR: {error}", file=sys.stderr)
        raise SystemExit(1)
