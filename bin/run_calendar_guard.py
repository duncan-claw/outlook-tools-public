#!/usr/bin/env python3
"""Bounded guard job; fixed status only. Never logs invitation content."""
import datetime
import json
import os
from pathlib import Path
import subprocess
import sys

from outlook import config

root = Path(__file__).resolve().parents[1]

state = config.STATE_DIR / 'calendar-guard'
os.umask(0o077)
state.mkdir(mode=0o700, parents=True, exist_ok=True)
os.chmod(state, 0o700)
try:
    run = subprocess.run([sys.executable, '-m', 'spam_guard.group_calendar_guard', '--apply'], cwd=root, capture_output=True, text=True, timeout=55)
    try:
        result=json.loads(run.stdout)
    except Exception:
        result={'error':'INVALID_JOB_OUTPUT'}
    status={'timestamp':datetime.datetime.now(datetime.timezone.utc).isoformat(), 'exit_code':run.returncode, 'result':result}
except subprocess.TimeoutExpired:
    status={'timestamp':datetime.datetime.now(datetime.timezone.utc).isoformat(), 'exit_code':1, 'result':{'error':'JOB_TIMEOUT'}}
target=state/'status.json'
temp=state/'status.tmp'
temp.write_text(json.dumps(status))
os.replace(temp,target)
# Calendar policy remains fail-closed independently of email cleanup.
if status['exit_code'] == 0:
    try:
        mail = subprocess.run([sys.executable, '-m', 'spam_guard.invitation_mail', '--apply'], cwd=root, capture_output=True, text=True, timeout=45)
        status['mail'] = json.loads(mail.stdout)
        if mail.returncode:
            status['exit_code'] = 1
    except Exception:
        status['mail'] = {'error':'MAIL_JOB_FAILED'}
        status['exit_code'] = 1
    # Sender filter is a separate Junk Email policy; its mode lives in private
    # config (absent = off). Counts/fixed codes only.
    try:
        sender = subprocess.run([sys.executable, '-m', 'spam_guard.sender_filter', 'scheduled'], cwd=root, capture_output=True, text=True, timeout=50)
        status['sender'] = json.loads(sender.stdout)
        if sender.returncode:
            status['exit_code'] = 1
    except Exception:
        status['sender'] = {'error':'SENDER_JOB_FAILED'}
        status['exit_code'] = 1
    temp.write_text(json.dumps(status))
    os.replace(temp,target)
if status['exit_code']:
    # Fixed local alert only; no attacker-controlled strings enter AppleScript.
    subprocess.run(['/usr/bin/osascript', '-e', 'display notification "Check Outlook Tools status.json for fixed error code." with title "Outlook guard failed"'], capture_output=True, timeout=5)
print(json.dumps(status))
sys.exit(status['exit_code'])
