# Calendar and invitation-mail guard

## Calendar guard

`spam_guard/group_calendar_guard.py` uses delegated Microsoft Graph access to inspect calendar ownership, organiser, event type, start time, recurrence, and `isOrganizer`. It does not request titles, bodies, descriptions, attachments, or links.

The guard:

- Targets attendee events organised by an exact `groups.outlook.com` address. Optional contacts-based handling is a separate policy and should remain disabled until reviewed.
- Skips self-organised events, shared or ownerless calendars, malformed organisers, and events outside the configured history cutoff.
- Re-reads each candidate immediately before deletion and verifies the deletion response.
- Never accepts, declines, or sends cancellation responses.
- Defaults to preview mode. Live changes require `--apply` and a configured account and cutoff.

The cutoff is supplied with `OUTLOOK_HISTORY_CUTOFF` as an ISO-8601 datetime. If it is absent, the guard fails closed and no event qualifies for deletion. Set `OUTLOOK_TIMEZONE` to the account's intended IANA timezone. Historical-series expansion is bounded to a rolling future horizon; an endless recurrence cannot be enumerated completely.

## Groups invitation email

`spam_guard/invitation_mail.py` only considers Microsoft Graph event-message types from the exact `groups.outlook.com` domain. It moves matching messages to Junk Email, not Deleted Items. Ordinary email and other senders are outside this component's scope. The selection is based on sender and message type, never message content.

## Authentication and state

Set `OUTLOOK_PERSONAL_USER` and `OUTLOOK_PERSONAL_GRAPH_CLIENT_ID` locally. The Graph client ID is a public application identifier, not a client secret. OAuth tokens and audit state are stored below `OUTLOOK_TOOLS_STATE_DIR`, outside the repository. Do not commit configuration, token files, logs, or mailbox-derived data.

## Before enabling live actions

1. Run previews and review aggregate counts.
2. Confirm the account identity and cutoff are correct.
3. Test recovery for the target account and calendar before relying on it.
4. Keep optional contacts-based deletion disabled unless its scope and recovery path have explicit approval.
5. Install a scheduler only after reviewing the local configuration and the example plist.

The repository does not establish that any particular account has completed these steps; operators must verify their own setup.
