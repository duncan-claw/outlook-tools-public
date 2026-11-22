# Sender filter: Inbox to Junk Email

The sender filter is an optional deterministic policy in `spam_guard/sender_filter.py`. It uses message ID, sender address/display name, received time, and folder metadata; it does not read subjects, bodies, links, or attachments. Sender data is untrusted input and must never be passed to an LLM.

## Policy overview

The first matching rule wins:

1. Keep malformed senders and the configured account's own address.
2. Apply the account owner's latest explicit learning action for that sender.
3. Keep senders found in configured contacts, Outlook Safe Senders, a local allowlist, or the sent-mail correspondent cache.
4. In `scored` mode, move a sender to Junk only when the deterministic identity score is suspicious, or when an unknown sender uses a configured free-mail domain.
5. Keep other senders.

A suspicious score requires multiple independent signals; no single feature is enough. The free-mail provider list is generic policy data in `sender_identity.py` and can be reviewed for a deployment's locale.

## Modes and reversibility

- `off` (default): no scheduled filtering.
- `dry-run`: report counts/reasons without moving messages.
- `scored`: apply the policy above.
- `strict`: move every sender not on the safe list; use only after a reviewed preview.

Every live candidate is re-read and re-evaluated before movement. Moves target Junk Email only, are verified, and are recorded in an external private ledger to prevent repeat processing. Manual rescue teaches the sender as safe; restore commands can return recorded items from Junk to Inbox. Each scheduled pass has a configured move cap.

Example preview:

```sh
.venv/bin/python -m spam_guard.sender_filter explain sender@example.com 'Example Sender'
.venv/bin/python -m spam_guard.sender_filter scan --report
```

Enabling a live policy affects new mail after its configured start time. Historical backlog scans require an explicit date range and `--apply`. Review Junk periodically; Outlook retention/purge behaviour is controlled by Microsoft and may vary by account.

## Safe-list sources and privacy

The filter can use contacts, a local allowlist, hashed sent-mail correspondents, and a hashed Outlook Safe Senders snapshot. The Outlook.com Safe Senders sync uses an internal Outlook web endpoint, so it is the most brittle integration; if its snapshot is absent, fail closed. Store configuration, allowlists, snapshots, sender-learning data, ledgers, and reports only under the private state directory configured by `OUTLOOK_TOOLS_STATE_DIR`.

## Commands

```sh
.venv/bin/python -m spam_guard.sender_filter explain sender@example.com 'Example Sender'
.venv/bin/python -m spam_guard.sender_filter scan --report
.venv/bin/python -m spam_guard.sender_filter enable off|dry-run|scored|strict
.venv/bin/python -m spam_guard.sender_filter restore --since <ISO-8601>
```

Read the command help and review the active local configuration before enabling or restoring any live messages.
