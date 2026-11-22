# Outlook Tools

A small Python toolkit for reducing Outlook spam without reading message or invitation content. It combines metadata-only calendar invitation cleanup, narrow Groups-invitation mail handling, and an optional deterministic sender filter.

## Why this exists

This project started with a familiar frustration: Outlook’s built-in filtering was missing two different kinds of unwanted mail. Outlook Groups invitations were appearing as calendar events even after their senders were blocked, while inbox spam was arriving from plausible display names paired with suspicious-looking addresses. Ordinary Outlook rules could block known senders, but could not express the inverse rule that mattered here: keep mail from people and senders already trusted, and carefully handle suspicious or unknown senders.

The design was shaped as much by privacy as by the spam itself. Because an attacker controls the content of spam, handing message text to an AI model would create an unnecessary prompt-injection path into a mailbox and calendar. These tools instead use deterministic rules over the smallest useful metadata, avoid reading subjects and bodies, and favour reversible actions with fail-closed safeguards. AI helped develop and test the software; it is not in the runtime decision path.

The original motivation, implementation trade-offs, and lessons learned are in [A sender-based spam filter for Outlook.com that never reads the email](https://www.gr8ful.dev/2026/10/18/outlook-sender-spam-filter.html).

This repository is a prototype, not a one-click installer. Review the policies and configuration before enabling scheduled jobs. The calendar guard is dry-run by default; sender filtering is off by default. Group removal is a separate browser automation that requires explicit installation and sign-in.

## Components

- `spam_guard/group_calendar_guard.py` — identifies eligible calendar events from organiser and event metadata. It preserves self-organised events and requires an explicit configured history cutoff before any deletion can match.
- `spam_guard/invitation_mail.py` — recognizes only Outlook Groups event-message types from the exact `groups.outlook.com` domain; it does not inspect message bodies or subjects.
- `spam_guard/sender_filter.py` — optionally moves selected Inbox messages to Junk Email using sender metadata, safe lists, and deterministic scoring. It never reads subjects, bodies, links, or attachments.
- `spam_guard/leave_groups_via_cdp.py` — optionally leaves Outlook Groups using a dedicated, already-authenticated browser profile. It verifies the account before acting and fails closed if identity cannot be confirmed.
- `spam_guard/outlook_safe_senders.py` — refreshes a hashed snapshot of Outlook.com Safe Senders using the browser session. This depends on Outlook web internals and may need maintenance if Microsoft changes them.

The SMS-to-calendar prototype and IMAP helper are intentionally not included in this distribution.

## Requirements

- Python 3.11 or newer and macOS for the supplied LaunchAgent examples.
- A Microsoft Entra public-client app configured for the account type and delegated Graph scopes required by the selected features. The app's client ID is configuration, not a client secret.
- For Groups removal and Safe Senders sync: Google Chrome, a separate profile signed in to Outlook, and human completion of any login or MFA challenge.

The tools do not require OpenClaw. Calendar and sender-filter calls use Microsoft Graph. The Groups browser features use loopback Chrome DevTools Protocol and do not send group content to an LLM.

## Local configuration

Supply configuration through the process environment. Do not commit real values. For example, in a private shell session:

```sh
export OUTLOOK_PERSONAL_USER='owner@example.com'
export OUTLOOK_PERSONAL_GRAPH_CLIENT_ID='your-own-app-client-id'
export OUTLOOK_PERSONAL_GRAPH_TENANT='consumers'
export OUTLOOK_TOOLS_STATE_DIR="$HOME/.local/state/outlook-tools"
export OUTLOOK_TIMEZONE='Etc/UTC'
export OUTLOOK_GROUPS_CDP_PORT='9222'
# Set OUTLOOK_HISTORY_CUTOFF only after choosing a cutoff; unset means fail-closed.
```

Replace the example values locally. `OUTLOOK_PERSONAL_USER` and the Graph client ID are required for account-bound operations. If no history cutoff is configured, calendar deletion matching fails closed. Store OAuth tokens, allowlists, sender ledgers, browser profiles, and logs outside the repository with user-only permissions.

LaunchAgents do not inherit interactive shell variables. Configure the required environment in a local, untracked launchd plist; never put a real account or client configuration in a committed template.

## Run safely

```sh
python3 -m venv .venv
.venv/bin/pip install -e .
.venv/bin/python -m spam_guard.group_calendar_guard
.venv/bin/python -m spam_guard.invitation_mail
.venv/bin/python -m spam_guard.sender_filter explain sender@example.com 'Example Sender'
```

These commands are read-only previews. Read `docs/INVITE_GUARD.md`, `docs/GROUPS.md`, and `docs/SENDER_FILTER.md` before using any apply/confirm option or installing a schedule. Do not use a real mailbox for tests.

## Tests

```sh
.venv/bin/python -m unittest discover -s tests -v
```

Tests use synthetic data and mock Graph, browser, and filesystem interactions. They make no live mailbox or calendar changes.

## Privacy and safety

- Process only the minimum metadata needed for each rule.
- Keep account details, tokens, contacts, message logs, sender history, and browser profiles outside Git.
- Use a dedicated browser profile for CDP automation; never expose its debugging port beyond loopback.
- Begin with previews, choose a cutoff deliberately, and verify rollback/recovery behaviour before live deletion.

## Release status

Prototype; setup, policy review, platform-specific installation, and a dedicated secret scan are still required before enabling it broadly. This repository is licensed under the MIT License; see [LICENSE](LICENSE).
