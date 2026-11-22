#!/usr/bin/env python3
"""Metadata-only Outlook calendar cleanup. Default dry-run; no LLM integration."""
import argparse
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
import fcntl
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from outlook import config
from outlook import outlook_personal_graph as auth

BASE = 'https://graph.microsoft.com/v1.0'
STATE = config.STATE_DIR / 'calendar-guard'
SELECT = 'id,organizer,isOrganizer,start,type,recurrence'
if config.HISTORY_CUTOFF:
    CUTOFF = datetime.fromisoformat(config.HISTORY_CUTOFF)
    if CUTOFF.tzinfo is None:
        CUTOFF = CUTOFF.replace(tzinfo=ZoneInfo(config.TIMEZONE))
else:
    # No historical deletions until an operator explicitly configures a cutoff.
    CUTOFF = datetime.max.replace(tzinfo=ZoneInfo(config.TIMEZONE))

def in_scope(event):
    # Delete entire series only if both its first start and recurrence range
    # begin at/after cutoff. Older/ambiguous masters are expanded, never deleted.
    if event.get('type') == 'seriesMaster':
        first = ((event.get('recurrence') or {}).get('range') or {}).get('startDate')
        try:
            if datetime.fromisoformat(first).date() < CUTOFF.date():
                return False
        except (TypeError, ValueError):
            return False
    start = event.get('start') or {}
    try:
        value = datetime.fromisoformat(start['dateTime'].replace('Z', '+00:00'))
        if value.tzinfo is None:
            zone = start.get('timeZone', 'UTC')
            if zone == 'AUS Eastern Standard Time':
                zone = 'Australia/Melbourne'
            value = value.replace(tzinfo=ZoneInfo(zone))
        return value >= CUTOFF
    except (KeyError, ValueError, TypeError):
        return False

class GuardError(Exception):
    pass

def matches(event):
    address = ((event.get('organizer') or {}).get('emailAddress') or {}).get('address')
    return (isinstance(address, str)
            and re.fullmatch(r'[^\s@<>]+@groups\.outlook\.com', address, re.I) is not None
            and event.get('isOrganizer') is False)

class Client:
    def __init__(self):
        self.token = auth.get_token(interactive=False)['access_token']

    def request(self, method, path):
        url = path if path.startswith('https://') else BASE + path
        parsed = urllib.parse.urlsplit(url)
        if parsed.scheme != 'https' or parsed.netloc != 'graph.microsoft.com' or not parsed.path.startswith('/v1.0/me/'):
            raise GuardError('INVALID_API_URL')
        req = urllib.request.Request(url, method=method, headers={
            'Authorization': 'Bearer ' + self.token, 'Accept': 'application/json'})
        try:
            with urllib.request.urlopen(req, timeout=30, context=auth.http_context()) as response:
                raw = response.read()
                return json.loads(raw) if raw else {}
        except urllib.error.HTTPError as error:
            # Never expose server error bodies, URLs or attacker-controlled strings.
            raise GuardError('HTTP_' + str(error.code)) from None
        except Exception:
            raise GuardError('REQUEST_FAILED') from None

    def pages(self, path):
        seen = set()
        while path:
            if path in seen:
                raise GuardError('PAGINATION_LOOP')
            seen.add(path)
            result = self.request('GET', path)
            yield from result.get('value', [])
            path = result.get('@odata.nextLink')

def unknown_organizer(event, allowed):
    address = ((event.get('organizer') or {}).get('emailAddress') or {}).get('address')
    return (event.get('isOrganizer') is False and isinstance(address, str)
            and re.fullmatch(r'[^\s@<>]+@[^\s@<>]+', address) is not None
            and address.lower() not in allowed)


def discover_events(client, root, allowed, contacts_enabled, summary):
    """Expand historical suspicious series into occurrences, preserving history.

    Unbounded series cannot be enumerated to infinity. Historical masters use a
    rolling two-year horizon, refreshed every minute. New series can be removed
    as a whole, regardless of how far in the future they extend.
    """
    seen = set()
    for event in client.pages(root + '?$select=' + SELECT + '&$top=100'):
        if event.get('id') not in seen:
            seen.add(event.get('id'))
            yield event
        if (event.get('type') != 'seriesMaster' or in_scope(event)
                or not (matches(event) or (contacts_enabled and unknown_organizer(event, allowed)))):
            continue
        summary['series_expanded'] += 1
        end = datetime.now(ZoneInfo('Australia/Melbourne')) + timedelta(days=730)
        query = urllib.parse.urlencode({'startDateTime':CUTOFF.isoformat(),
            'endDateTime':end.isoformat(), '$select':SELECT, '$top':100})
        path = root + '/' + urllib.parse.quote(event['id'], safe='') + '/instances?' + query
        for instance in client.pages(path):
            # Never accept an unexpected master from an instances endpoint.
            if instance.get('type') not in ('occurrence', 'exception'):
                raise GuardError('INVALID_INSTANCE_TYPE')
            if instance.get('id') not in seen:
                seen.add(instance.get('id'))
                summary['instances_scanned'] += 1
                yield instance


def scan(client, apply=False, contacts_preview=False, contacts_apply=False):
    summary = dict(calendars=0, events=0, matched=0, deleted=0, skipped_calendars=0, errors=0, series_expanded=0, instances_scanned=0)
    expected_account = config.ACCOUNT
    if not expected_account:
        raise GuardError('ACCOUNT_NOT_CONFIGURED')
    identity = client.request('GET', '/me/?$select=mail,userPrincipalName')
    if not any(str(identity.get(k, '')).casefold() == expected_account for k in ('mail', 'userPrincipalName')):
        raise GuardError('ACCOUNT_MISMATCH')
    allowed = {expected_account}
    if contacts_preview or contacts_apply:
        # Fully load contacts before any calendar processing; never expose addresses.
        for contact in client.pages('/me/contacts?$select=emailAddresses&$top=100'):
            for entry in contact.get('emailAddresses') or []:
                address = entry.get('address')
                if isinstance(address, str):
                    allowed.add(address.strip().lower())
        summary['allowlist_addresses'] = len(allowed)
        summary['unknown_organizers'] = 0
    for cal in client.pages('/me/calendars?$select=id,owner,canEdit&$top=100'):
        # Never delete from a shared calendar or the organiser's own calendar.
        if str((cal.get('owner') or {}).get('address') or '').casefold() != expected_account or cal.get('canEdit') is not True:
            summary['skipped_calendars'] += 1
            continue
        summary['calendars'] += 1
        root = '/me/calendars/' + urllib.parse.quote(cal['id'], safe='') + '/events'
        # Complete discovery before mutations so deletion cannot shift paging.
        candidates = []
        for event in discover_events(client, root, allowed, contacts_preview or contacts_apply, summary):
            summary['events'] += 1
            unknown = False
            if (contacts_preview or contacts_apply) and in_scope(event) and event.get('isOrganizer') is False:
                address = ((event.get('organizer') or {}).get('emailAddress') or {}).get('address')
                if isinstance(address, str) and re.fullmatch(r'[^\s@<>]+@[^\s@<>]+', address) and address.lower() not in allowed:
                    summary['unknown_organizers'] += 1
                    unknown = True
            if in_scope(event) and (matches(event) or (contacts_apply and unknown)):
                candidates.append(event['id'])
        summary['matched'] += len(candidates)
        for event_id in candidates:
            if not apply:
                continue
            path = root + '/' + urllib.parse.quote(event_id, safe='')
            try:
                current = client.request('GET', path + '?$select=' + SELECT)
                address = ((current.get('organizer') or {}).get('emailAddress') or {}).get('address')
                unknown = (contacts_apply and current.get('isOrganizer') is False and isinstance(address, str)
                           and re.fullmatch(r'[^\s@<>]+@[^\s@<>]+', address) is not None and address.lower() not in allowed)
                if not in_scope(current) or not (matches(current) or unknown):
                    continue
                audit({'id_hash': hashlib.sha256(event_id.encode()).hexdigest(), 'action': 'delete_intent'})
                report = {'timestamp': datetime.now(ZoneInfo('Australia/Melbourne')).isoformat(), 'event_start': current.get('start'), 'event_type': current.get('type', 'singleInstance'), 'organizer': address, 'reason': 'groups_domain' if matches(current) else 'unknown_organizer', 'id_hash': hashlib.sha256(event_id.encode()).hexdigest()}
                client.request('DELETE', path)
                try:
                    client.request('GET', path + '?$select=id')
                    raise GuardError('DELETE_NOT_VERIFIED')
                except GuardError as error:
                    if str(error) != 'HTTP_404':
                        raise
                summary['deleted'] += 1
                with (STATE / 'deletions.jsonl').open('a') as stream:
                    stream.write(json.dumps(report) + '\n')
                audit({'id_hash': hashlib.sha256(event_id.encode()).hexdigest(), 'action': 'delete_verified'})
            except GuardError as error:
                summary['errors'] += 1
                audit({'action': 'error', 'code': str(error)})
    return summary

def audit(record):
    with (STATE / 'audit.jsonl').open('a') as stream:
        stream.write(json.dumps(record) + '\n')

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--apply', action='store_true')
    parser.add_argument('--contacts-preview', action='store_true', help='Count unknown organisers only; never delete in this mode')
    parser.add_argument('--contacts-apply', action='store_true', help='Apply the configured contacts-only policy; requires --apply')
    args = parser.parse_args()
    if args.contacts_apply and not args.apply:
        parser.error('--contacts-apply requires --apply')
    if args.contacts_preview and args.apply:
        parser.error('--contacts-preview cannot be combined with --apply')
    os.umask(0o077)
    STATE.mkdir(parents=True, exist_ok=True)
    with (STATE / 'lock').open('w') as lock:
        try:
            fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
        except BlockingIOError:
            print('{"error":"ALREADY_RUNNING"}')
            return 1
        try:
            result = scan(Client(), args.apply, args.contacts_preview, args.contacts_apply)
            print(json.dumps(result))
            return int(result['errors'] > 0)
        except GuardError as error:
            print(json.dumps({'error': str(error)}))
            return 1
        except Exception:
            print('{"error":"INTERNAL_ERROR"}')
            return 1

if __name__ == '__main__':
    sys.exit(main())
