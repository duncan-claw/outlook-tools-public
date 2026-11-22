#!/usr/bin/env python3
"""Leave Outlook.com Groups via Chrome DevTools Protocol (CDP).

Bounded, fail-closed automation for a *dedicated* Chrome instance that is
already authenticated to outlook.live.com. It is used by the scheduled
`bin/run_groups_leaver.py` job and can also be run by hand.

Privacy contract (enforced here, not just by convention):
  - No group name, group id, URL or screenshot bytes are ever printed to
    stdout/stderr, written to the status/audit files, or logged.
  - Output is fixed status codes and counts only.
  - Group names are used in-memory for exact group-specific confirmation text
    and are never persisted.
  - Screenshots are written only to the private state directory (mode 0600)
    and their paths are not printed.

Safety contract:
  - Default is dry-run. Live leave requires `--confirm`.
  - Login/consent prompts abort the run (fail closed) without leaving groups.
  - Departure must be *verified* by the group disappearing from the
    navigation tree; unverified departures are counted as errors.
  - One group's error never aborts the rest of the run.

Prerequisites:
  - Chrome running with `--remote-debugging-port` and a dedicated
    `--user-data-dir`, already signed in to outlook.live.com.
  - Loopback proxy bypass; never route CDP through the egress proxy.
"""
from __future__ import annotations

import argparse
import asyncio
import base64
import json
import os
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path

from outlook import config

import websockets

# --- Configuration -----------------------------------------------------------

GROUP_HOME_URL = "https://outlook.live.com/groups/home"
MELBOURNE = "Australia/Melbourne"

# The account is supplied in local configuration and is never logged.

# Identity is read from the account UI itself. The live outlook.live.com shell
# exposes the avatar button but no address in static markup, so the probe opens
# the avatar menu, reads address-shaped text, then closes it again. Only
# addresses are extracted; nothing else from the menu is read or stored.
IDENTITY_JS = r"""
(async () => {
  const ADDRESS = /[^\s@<>()\[\],;]+@[^\s@<>()\[\],;]+\.[A-Za-z]{2,}/g;
  const collect = () => {
    const found = [];
    for (const el of document.querySelectorAll('[aria-label],[title]')) {
      for (const attr of ['aria-label', 'title']) {
        const v = el.getAttribute(attr) || '';
        const m = v.match(ADDRESS);
        if (m) found.push(...m);
      }
    }
    for (const el of document.querySelectorAll('*')) {
      if (el.children.length) continue;
      const m = (el.textContent || '').trim().match(ADDRESS);
      if (m) found.push(...m);
    }
    return found;
  };

  let found = collect();
  if (found.length) return found.join('\n');

  // Open the account menu via the avatar affordance, then re-probe.
  const avatar = document.querySelector('[id^="avatar-r"]');
  const opener = avatar && (avatar.closest('[role="button"],button,a') || avatar);
  if (!opener) return '';
  opener.click();
  await new Promise(r => setTimeout(r, 1500));
  found = collect();
  document.body.click();
  await new Promise(r => setTimeout(r, 200));
  return found.join('\n');
})()
"""

# Group home / sign-in markers. Presence of any of these in the page text means
# we are not authenticated and must stop without touching groups.
LOGIN_MARKERS = (
    "sign in",
    "signin",
    "enter password",
    "use another account",
    "pick an account",
    "stay signed in?",
    "we couldn't find your account",
    "your account or password is incorrect",
    "account.microsoft.com",
)

STATE_DIR = Path(
    os.environ.get(
        "OUTLOOK_GROUPS_STATE_DIR",
        str(config.GROUPS_STATE_DIR),
    )
)
SCREENSHOT_DIR = STATE_DIR / "screenshots"

DEFAULT_PORT = config.GROUPS_CDP_PORT
DEFAULT_TIMEOUT = 240.0
PER_GROUP_SLEEP = 0.5

# CDP command ids are per-connection; a module counter is fine for one process.
_cdp_msg_id = 1


class Blocked(Exception):
    """Fail-closed stop: nothing was left and nothing should be retried blindly."""


# --- CDP plumbing ------------------------------------------------------------


def bypass_opener() -> urllib.request.OpenerDirector:
    """HTTP opener that never uses the injected egress proxy."""
    return urllib.request.build_opener(urllib.request.ProxyHandler({}))


def http_get_json(url: str, retries: int = 10, delay: float = 1.0) -> dict:
    opener = bypass_opener()
    last: Exception | None = None
    for attempt in range(retries):
        try:
            with opener.open(url, timeout=10) as response:
                return json.loads(response.read().decode())
        except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError) as error:
            last = error
            if attempt < retries - 1:
                time.sleep(delay)
    raise Blocked(f"CDP_UNREACHABLE:{type(last).__name__ if last else 'UNKNOWN'}")


async def cdp_send(ws, method: str, params: dict | None = None, timeout_s: float = 15.0) -> dict:
    global _cdp_msg_id
    msg_id = _cdp_msg_id
    _cdp_msg_id += 1

    payload: dict = {"id": msg_id, "method": method}
    if params is not None:
        payload["params"] = params
    await ws.send(json.dumps(payload))

    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            raw = await asyncio.wait_for(ws.recv(), timeout=3.0)
        except asyncio.TimeoutError:
            continue
        try:
            data = json.loads(raw)
        except json.JSONDecodeError:
            continue
        if data.get("id") == msg_id:
            return data
    return {"error": {"message": f"TIMEOUT:{method}"}}


async def evaluate(ws, expression: str, timeout_s: float = 15.0):
    resp = await cdp_send(
        ws,
        "Runtime.evaluate",
        {"expression": expression, "returnByValue": True, "awaitPromise": True},
        timeout_s=timeout_s,
    )
    if "error" in resp:
        return None
    result = resp.get("result", {}).get("result", {})
    if result.get("subtype") == "error":
        return None
    return result.get("value")


async def page_text(ws) -> str:
    """Visible page text, used only for login detection. Never persisted."""
    value = await evaluate(ws, "(() => (document.body && document.body.innerText) || '')()")
    return value if isinstance(value, str) else ""


async def take_screenshot(ws, name: str) -> str | None:
    """Private evidence screenshot. Path is never printed or persisted."""
    try:
        SCREENSHOT_DIR.mkdir(parents=True, exist_ok=True)
        os.chmod(SCREENSHOT_DIR, 0o700)
    except OSError:
        return None
    resp = await cdp_send(ws, "Page.captureScreenshot", {"format": "png"})
    data = (resp.get("result") or {}).get("data")
    if not data:
        return None
    path = SCREENSHOT_DIR / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{name}.png"
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "wb") as handle:
            handle.write(base64.b64decode(data))
    except OSError:
        return None
    return str(path)


async def extract_groups(ws) -> list[dict]:
    """Group entries from the navigation tree as {gid, name} in memory only."""
    js = """
    (() => Array.from(document.querySelectorAll('[role="treeitem"][data-folder-name][aria-label][aria-level="2"]'))
      .map(e => ({gid: e.getAttribute('data-folder-name') || '', name: e.getAttribute('aria-label') || ''}))
      .filter(g => /^[A-Za-z0-9_-]+$/.test(g.gid)))()
    """
    value = await evaluate(ws, js, timeout_s=20.0)
    if not isinstance(value, list):
        return []
    groups = []
    for item in value:
        if not isinstance(item, dict):
            continue
        gid = str(item.get("gid") or "")
        name = str(item.get("name") or "")
        if gid:
            groups.append({"gid": gid, "name": name})
    return groups


def identity_is_expected(identity_text: str) -> bool | None:
    """True/False for a definite answer, None when the UI showed no identity."""
    candidates = []
    for line in identity_text.splitlines():
        line = line.strip()
        if "@" in line and " " not in line:
            candidates.append(line)
        else:
            for token in line.split():
                if "@" in token:
                    candidates.append(token.strip("<>()[],;"))
    candidates = [c for c in candidates if c]
    if not candidates:
        return None
    expected = config.ACCOUNT
    if not expected:
        return None
    return any(c.casefold() == expected for c in candidates)


async def verify_identity(ws) -> str | None:
    """Mechanically confirm the signed-in account. Returns a fixed code or None.

    Fail-closed by design: an unreadable identity is reported as
    ACCOUNT_UNVERIFIED rather than assumed correct, so a signed-out or
    wrong-account browser can never drive live group leaving.
    """
    if not config.ACCOUNT:
        return "ACCOUNT_NOT_CONFIGURED"
    value = await evaluate(ws, IDENTITY_JS, timeout_s=20.0)
    text = value if isinstance(value, str) else ""
    verdict = identity_is_expected(text)
    if verdict is True:
        return None
    if verdict is False:
        return "WRONG_ACCOUNT"
    return "ACCOUNT_UNVERIFIED"


def is_login_wall(text: str) -> bool:
    lowered = text.lower()
    if "are you sure you want to leave" in lowered:
        return False
    return any(marker in lowered for marker in LOGIN_MARKERS)


def summarize(groups: list[dict]) -> dict:
    """Counts only. Must never include names or ids.

    Every discovered group is a target: there are no name/id denylists. A group
    that cannot be left (for example no Leave button is offered) is reported as
    `leave_unavailable` and does not stop the rest of the pass.
    """
    return {"discovered": len(groups), "targets": len(groups)}


# --- Leave flow --------------------------------------------------------------


async def leave_one(ws, group: dict, timeout_s: float) -> str:
    """Attempt one group. Returns a fixed outcome code.

    "left"               departure verified from the navigation tree
    "leave_unavailable"  no enabled Leave control was offered for this group
    "leave_failed"       a control existed but departure was not verified

    Raises Blocked when the page looks unauthenticated so the caller stops
    rather than treating a login wall as a per-group error. A missing button is
    never treated as a successful leave, and never stops the remaining groups.
    """
    gid = group["gid"]
    name = group["name"]
    started = time.monotonic()

    # The group page is a client-side route. Navigating straight to the members
    # URL is rewritten by the SPA to a non-group path where no group controls
    # render, so instead: load the groups home, then open this group from the
    # navigation tree, which lands on the group page. The Members tab must then
    # be selected before the Leave control exists.
    await cdp_send(ws, "Page.navigate", {"url": GROUP_HOME_URL})
    await asyncio.sleep(3)
    await _wait_for_ready(ws, timeout_s, cap=15.0)

    if is_login_wall(await page_text(ws)):
        raise Blocked("LOGIN_REQUIRED")

    opened = await evaluate(
        ws,
        """((gid) => {
            const items = Array.from(document.querySelectorAll('[role="treeitem"][data-folder-name][aria-level="2"]'));
            const el = items.find(e => e.getAttribute('data-folder-name') === gid);
            if (!el) return false;
            const target = el.closest('[role="button"],[role="treeitem"]') || el;
            target.click();
            return true;
        })(%s)"""
        % json.dumps(gid),
    )
    if opened is not True:
        # The group is no longer in the tree: it is gone, which is success.
        remaining = await extract_groups(ws)
        if remaining and not any(item["gid"] == gid for item in remaining):
            return "left"
        if not remaining:
            return "leave_failed"
        return "leave_unavailable"

    await asyncio.sleep(4)
    await _wait_for_ready(ws, timeout_s, cap=15.0)

    if is_login_wall(await page_text(ws)):
        raise Blocked("LOGIN_REQUIRED")

    # Select the Members tab; the Leave control only renders in that panel.
    await evaluate(
        ws,
        """(() => {
            const tabs = Array.from(document.querySelectorAll('[role="tab"]'));
            const members = tabs.find(t => (t.textContent || '').trim().toLowerCase() === 'members');
            if (!members) return false;
            members.click();
            return true;
        })()""",
    )
    await asyncio.sleep(3)

    # Exact group-specific confirmation requires the real name in-memory.
    clicked = await evaluate(
        ws,
        """(() => {
            const b = document.querySelector('button[aria-label="Leave group"]');
            if (!b || b.disabled) return false;
            b.click(); return true;
        })()""",
    )
    if clicked is not True:
        # No enabled Leave affordance for this group. Record it and continue.
        return "leave_unavailable"
    await asyncio.sleep(1.5)

    confirmed = await evaluate(
        ws,
        """((name) => {
            const dialogs = Array.from(document.querySelectorAll('[role="alertdialog"],[role="dialog"]'));
            const d = dialogs.find(x => (x.innerText || '').includes('Are you sure you want to leave ' + name + '?'));
            if (!d) return false;
            const b = Array.from(d.querySelectorAll('button')).find(b => (b.textContent || '').trim() === 'Leave');
            if (!b || b.disabled) return false;
            b.click(); return true;
        })(%s)"""
        % json.dumps(name),
    )
    if confirmed is not True:
        # Dialog state is ambiguous; do not retry blindly in the same run.
        return "leave_failed"

    # Verify departure: the group must disappear from the navigation tree.
    deadline = min(time.monotonic() + 25.0, started + timeout_s)
    while time.monotonic() < deadline:
        await asyncio.sleep(1.2)
        remaining = await _discover(ws, attempts=2, delay=0.5)
        if remaining is not None and not any(item["gid"] == gid for item in remaining):
            # Empty is a valid success when the final membership disappears.
            # Recheck after a home reload rather than trusting a transient tree.
            await cdp_send(ws, "Page.navigate", {"url": GROUP_HOME_URL})
            await asyncio.sleep(3)
            confirmed = await _discover(ws)
            if confirmed is not None and not any(item["gid"] == gid for item in confirmed):
                return "left"
    return "leave_failed"


async def _discover(ws, attempts: int = 3, delay: float = 2.0) -> list[dict] | None:
    """Return group entries, [] for a confirmed-empty tree, None if unreadable.

    A transient render failure must not be mistaken for a clean account, and a
    genuinely empty tree must not be reported as a discovery failure.
    """
    marker = await evaluate(
        ws,
        """(() => {
            const tree = document.querySelector('[role="tree"]');
            if (tree) return 'tree';
            const home = document.querySelector('[role="main"], main');
            if (home) return 'main';
            return '';
        })()""",
    )
    for attempt in range(attempts):
        groups = await extract_groups(ws)
        if groups:
            return groups
        if marker:
            # Container is present and readable -> an empty result is genuine.
            return []
        await asyncio.sleep(delay)
        marker = await evaluate(
            ws,
            """(() => (document.querySelector('[role="tree"],[role="main"],main') ? 'ok' : ''))()""",
        )
    return None


async def _wait_for_ready(ws, timeout_s: float, cap: float = 20.0) -> None:
    deadline = time.monotonic() + cap
    while time.monotonic() < deadline:
        ready = await evaluate(ws, "document.readyState")
        if ready == "complete":
            return
        await asyncio.sleep(0.5)


def _pick_target(targets: list[dict]) -> dict | None:
    pages = [t for t in targets if t.get("type") == "page"]
    if not pages:
        return None
    preferred = [
        t
        for t in pages
        if urllib.parse.urlparse(t.get("url", "")).netloc.endswith("outlook.live.com")
        and "groups" in urllib.parse.urlparse(t.get("url", "")).path
    ]
    return (preferred or pages)[0]


def _open_page_target(cdp_http: str) -> dict | None:
    request = urllib.request.Request(f"{cdp_http}/json/new?about:blank", method="PUT")
    try:
        with bypass_opener().open(request, timeout=10) as response:
            target = json.loads(response.read().decode())
    except (urllib.error.URLError, TimeoutError, json.JSONDecodeError):
        return None
    return target if target.get("webSocketDebuggerUrl") else None


async def run(
    port: int = DEFAULT_PORT,
    confirm: bool = False,
    timeout_s: float = DEFAULT_TIMEOUT,
    dry_run: bool = False,
    only: str | None = None,
) -> dict:
    """One bounded pass. Returns a fixed-code/count summary; never raw content."""
    started = time.monotonic()
    result = {
        "mode": "dry-run" if (dry_run or not confirm) else "live",
        "discovered": 0,
        "targets": 0,
        "left": 0,
        "leave_unavailable": 0,
        "errors": 0,
        "remaining": 0,
        "blocked": None,
        "timed_out": False,
        "identity": None,
        "state": None,
    }

    cdp_http = f"http://127.0.0.1:{port}"
    try:
        http_get_json(f"{cdp_http}/json/version", retries=3, delay=0.5)
    except Blocked as error:
        result["blocked"] = str(error)
        return result

    targets = http_get_json(f"{cdp_http}/json/list", retries=3, delay=0.5)
    target = _pick_target(targets)
    if target is None:
        # A Chrome restored with --no-startup-window has no tabs; open one.
        target = _open_page_target(cdp_http)
    if target is None:
        result["blocked"] = "NO_PAGE_TARGET"
        return result

    ws = await websockets.connect(target["webSocketDebuggerUrl"], proxy=None, max_size=None)
    try:
        await cdp_send(ws, "Page.enable")
        await cdp_send(ws, "Runtime.enable")

        await cdp_send(ws, "Page.navigate", {"url": GROUP_HOME_URL})
        await asyncio.sleep(4)
        await _wait_for_ready(ws, timeout_s, cap=20.0)

        if is_login_wall(await page_text(ws)):
            result["blocked"] = "LOGIN_REQUIRED"
            return result

        # Confirm the signed-in account mechanically before touching anything.
        identity = await verify_identity(ws)
        result["identity"] = "ok" if identity is None else identity
        if identity is not None:
            result["blocked"] = identity
            return result

        groups = await _discover(ws)
        if groups is None:
            # Discovery failed: not proof of an empty account. Fail closed.
            result["blocked"] = "DISCOVERY_FAILED"
            return result
        if only is not None:
            groups = [group for group in groups if group["gid"] == only]
        result.update(summarize(groups))
        if not groups:
            # Authenticated and the tree was read successfully with zero groups.
            # That is a legitimate fully-cleaned state, not an error.
            result["state"] = "empty_clean"
            return result

        if not confirm:
            # Dry-run: prove discovery and authentication only. Leave nothing.
            await take_screenshot(ws, "groups_home")
            await cdp_send(ws, "Page.navigate", {"url": GROUP_HOME_URL})
            return result

        for group in groups:
            if time.monotonic() - started > timeout_s:
                result["timed_out"] = True
                break
            try:
                outcome = await leave_one(ws, group, timeout_s)
                if outcome == "left":
                    result["left"] += 1
                elif outcome == "leave_unavailable":
                    result["leave_unavailable"] += 1
                else:
                    result["errors"] += 1
            except Blocked as error:
                result["blocked"] = str(error)
                break
            except (websockets.WebSocketException, OSError, asyncio.TimeoutError):
                result["errors"] += 1
            await asyncio.sleep(PER_GROUP_SLEEP)

        # Final discovery for a bounded remaining count.
        try:
            await cdp_send(ws, "Page.navigate", {"url": GROUP_HOME_URL})
            await asyncio.sleep(3)
            await _wait_for_ready(ws, timeout_s, cap=15.0)
            final = await _discover(ws, attempts=2, delay=2.0)
            if final is not None:
                result["remaining"] = len(final)
                result["state"] = "empty_clean" if not final else "remaining"
        except (websockets.WebSocketException, OSError, asyncio.TimeoutError):
            pass
    finally:
        try:
            await ws.close()
        except Exception:
            pass

    return result


# --- Status / notification ---------------------------------------------------


def write_status(result: dict, exit_code: int) -> None:
    """Atomic 0600 status write. Counts and fixed codes only."""
    try:
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        os.chmod(STATE_DIR, 0o700)
    except OSError:
        return
    payload = {
        "timestamp": datetime.now().astimezone().isoformat(),
        "exit_code": exit_code,
        "result": result,
    }
    temp = STATE_DIR / "status.tmp"
    try:
        fd = os.open(temp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w") as handle:
            handle.write(json.dumps(payload))
        os.replace(temp, STATE_DIR / "status.json")
    except OSError:
        pass

    try:
        with (STATE_DIR / "audit.jsonl").open("a") as stream:
            os.chmod(STATE_DIR / "audit.jsonl", 0o600)
            stream.write(json.dumps({"timestamp": payload["timestamp"], "exit_code": exit_code, **result}) + "\n")
    except OSError:
        pass


def notify_failure(code: str) -> None:
    """Fixed local alert; no dynamic string ever enters the AppleScript."""
    if not sys.platform == "darwin":
        return
    import subprocess

    subprocess.run(
        [
            "/usr/bin/osascript",
            "-e",
            f'display notification "Fixed code {code}. Check status.json." with title "Outlook groups leaver failed"',
        ],
        capture_output=True,
        timeout=5,
    )


def failure_code(result: dict) -> str:
    if result.get("blocked"):
        return str(result["blocked"]).split(":", 1)[0]
    if result.get("timed_out"):
        return "JOB_TIMEOUT"
    if result.get("errors"):
        return "GROUP_ERRORS"
    # `leave_unavailable` is retried on the next pass and is not fatal.
    # `state == "empty_clean"` and a normal run are both success.
    return ""


def load_config() -> dict:
    """Private settings in STATE_DIR/config.json; missing file means defaults.

    notify_group_errors (default true): raise the desktop alert when a pass
    only had per-group errors (GROUP_ERRORS). A group Microsoft has broken
    server-side fails every pass, so this alerts every 5 minutes. Blocked and
    timeout failures always alert; status.json and the exit code are unchanged.
    """
    defaults = {"notify_group_errors": True}
    try:
        config = json.loads((STATE_DIR / "config.json").read_text())
    except FileNotFoundError:
        return defaults
    except (OSError, ValueError):
        return defaults  # unreadable config must not silence alerts
    return {**defaults, **(config if isinstance(config, dict) else {})}


def should_notify(code: str, config: dict) -> bool:
    if not code:
        return False
    return code != "GROUP_ERRORS" or config.get("notify_group_errors") is not False


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Leave Outlook.com Groups via CDP (bounded)")
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument("--confirm", action="store_true", help="Actually leave groups")
    parser.add_argument("--dry-run", action="store_true", help="Discover only (default)")
    parser.add_argument("--timeout", type=float, default=DEFAULT_TIMEOUT, help="Wall-clock budget seconds")
    parser.add_argument("--group", type=str, help="Restrict to one numeric-free group id (debug only)")
    parser.add_argument("--status", action="store_true", help="Write status.json/audit.jsonl")
    args = parser.parse_args(argv)

    if args.dry_run and args.confirm:
        parser.error("--dry-run and --confirm are mutually exclusive")

    os.umask(0o077)
    confirm = args.confirm and not args.dry_run
    result = asyncio.run(
        run_one(port=args.port, confirm=confirm, timeout_s=args.timeout, only=args.group)
    )
    print(json.dumps(result, sort_keys=True))

    code = failure_code(result)
    if args.status:
        write_status(result, 1 if code else 0)
    if code:
        if should_notify(code, load_config()):
            notify_failure(code)
        return 1
    return 0


async def run_one(port: int, confirm: bool, timeout_s: float, only: str | None = None) -> dict:
    return await run(port=port, confirm=confirm, timeout_s=timeout_s, dry_run=not confirm, only=only)


if __name__ == "__main__":
    sys.exit(main())
