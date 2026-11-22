#!/usr/bin/env python3
"""Scheduled browser Groups leaver. Fixed status/counts only.

Runs one bounded live pass of `spam_guard.leave_groups_via_cdp --confirm`,
records a 0600 status file, and raises a fixed-code local notification only on
a genuine failure. No group name, id, URL or account text is ever printed,
logged, or written to status.
"""
import datetime
import fcntl
import json
import os
from pathlib import Path
import subprocess
import sys

from outlook import config

root = Path(__file__).resolve().parents[1]
state = config.GROUPS_STATE_DIR
port = config.GROUPS_CDP_PORT
os.umask(0o077)
state.mkdir(mode=0o700, parents=True, exist_ok=True)
os.chmod(state, 0o700)

LOCK_TIMEOUT_S = 270
RUN_TIMEOUT_S = 240


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def write_status(payload: dict) -> None:
    temp = state / "status.tmp"
    try:
        fd = os.open(temp, os.O_WRONLY | os.O_CREAT | os.O_TRUNC, 0o600)
        with os.fdopen(fd, "w") as handle:
            handle.write(json.dumps(payload))
        os.replace(temp, state / "status.json")
    except OSError:
        pass


status = {"timestamp": now(), "exit_code": 1, "result": {"error": "NOT_STARTED"}}

# Overlap lock: launchd serialises a job, but a long pass must not stack with
# the next tick if the schedule is ever tightened.
lock_path = state / "lock"
lock = open(lock_path, "w")
try:
    fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
except BlockingIOError:
    status = {
        "timestamp": now(),
        "exit_code": 0,
        "result": {"error": "ALREADY_RUNNING", "skipped": True},
    }
    write_status(status)
    print(json.dumps(status))
    sys.exit(0)

# Reuse only the dedicated isolated Chrome profile. Restart it if CDP is down.
import urllib.request
import time
opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
try:
    with opener.open(f'http://127.0.0.1:{port}/json/version', timeout=3):
        pass
except Exception:
    chrome = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome'
    profile = state / 'chrome-profile'
    subprocess.Popen([chrome, '--user-data-dir='+str(profile), '--remote-debugging-address=127.0.0.1', f'--remote-debugging-port={port}', '--no-first-run', 'https://outlook.live.com/groups/home'], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, start_new_session=True)
    # Bounded readiness wait, not a scheduling loop.
    for _ in range(20):
        try:
            with opener.open(f'http://127.0.0.1:{port}/json/version', timeout=2):
                break
        except Exception:
            time.sleep(0.5)

try:
    run = subprocess.run(
        [
            sys.executable,
            "-m",
            "spam_guard.leave_groups_via_cdp",
            "--confirm",
            f"--port={port}",
            f"--timeout={RUN_TIMEOUT_S}",
        ],
        cwd=root,
        capture_output=True,
        text=True,
        timeout=LOCK_TIMEOUT_S,
    )
    try:
        result = json.loads(run.stdout)
    except Exception:
        result = {"error": "INVALID_JOB_OUTPUT"}
    status = {"timestamp": now(), "exit_code": run.returncode, "result": result}
except subprocess.TimeoutExpired:
    status = {"timestamp": now(), "exit_code": 1, "result": {"error": "JOB_TIMEOUT"}}

# Outlook Safe Senders snapshot for the sender filter (hourly; same browser,
# same lock). Counts/fixed codes only.
try:
    sync = subprocess.run([sys.executable, "-m", "spam_guard.outlook_safe_senders", "--caller-holds-lock"],
                          cwd=root, capture_output=True, text=True, timeout=90)
    status["safe_senders"] = json.loads(sync.stdout)
    if sync.returncode:
        status["exit_code"] = 1
except Exception:
    status["safe_senders"] = {"error": "SAFE_SENDERS_JOB_FAILED"}
    status["exit_code"] = 1

write_status(status)
if status["exit_code"]:
    code = ((status.get("result") or {}).get("blocked") or (status.get("result") or {}).get("error")
            or (status.get("safe_senders") or {}).get("error") or "JOB_FAILED")
    # Fixed local alert only; no attacker-controlled string enters AppleScript.
    subprocess.run(
        [
            "/usr/bin/osascript",
            "-e",
            f'display notification "Fixed code {str(code).split(":", 1)[0]}. Check status.json." with title "Outlook groups leaver failed"',
        ],
        capture_output=True,
        timeout=5,
    )

print(json.dumps(status))
sys.exit(status["exit_code"])
