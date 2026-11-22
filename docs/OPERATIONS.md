# Operations

This project provides opt-in command-line tools; it does not install or start services automatically. The supplied LaunchAgent files are templates only and must be rendered locally with private paths and environment values before use.

## Suggested rollout

1. Read `INVITE_GUARD.md`, `GROUPS.md`, and `SENDER_FILTER.md`.
2. Set the account, app-registration ID, state directory, timezone, and history cutoff in the local environment.
3. Authenticate interactively once and confirm the Graph identity. Never put tokens in the repository or shell history.
4. Run each component in preview mode and inspect only aggregate results.
5. Test the target account's recovery path and choose the live policy deliberately.
6. Install only the schedules required for the selected features. Keep the Groups browser job isolated from Graph jobs.
7. Monitor fixed-code status and logs; configure log rotation and an operator-visible alert route for the deployment.

## Runtime boundaries

- Calendar and mail cleanup use Graph delegated access. Scheduled operations must never prompt for interactive sign-in.
- The calendar guard loads contacts before optional contacts-based handling; an authentication/contact failure must abort that policy.
- The sender filter moves Inbox messages only to Junk Email and defaults to off.
- The Groups leaver requires a dedicated signed-in browser profile and loopback CDP. Human login/MFA is required when the session expires.
- Local state (tokens, profiles, allowlists, status, audit, and logs) belongs outside Git under `OUTLOOK_TOOLS_STATE_DIR` and should be permission-restricted.

## Local LaunchAgent templates

Files in `deploy/*.plist.example` contain placeholders and are not directly installable. Replace placeholders locally, add the required environment variables in the local plist, and verify all program and state paths before bootstrapping. Keep the rendered plist and logs out of Git. Do not enable unknown-organiser deletion or strict sender filtering without a reviewed preview and recovery plan.

## Failure handling

The commands emit aggregate counts and fixed error codes. A nonzero status is not proof that any mutation occurred; inspect the local component status and audit record. Stop the relevant job before changing deployment configuration. For rollback, disable the job, restore recorded Junk moves when appropriate, and verify the result before resuming.
