# Outlook Groups browser automation

The Groups leaver is optional and separate from Graph calendar and mail cleanup. It uses a dedicated Chrome profile through Chrome DevTools Protocol on loopback; it does not rely on a normal browsing profile.

## Safety behaviour

- The account identity must match `OUTLOOK_PERSONAL_USER` before the tool can act. Missing or unreadable identity fails closed.
- A dry run is the default. Leaving memberships requires `--confirm`.
- Every discovered group is a target; do not add name-based exceptions without an explicit policy and tests.
- A group is counted as left only after the UI confirms departure. Missing controls and ambiguous confirmation are reported as failures and retried on a later pass.
- Login and MFA are human steps. The automation does not bypass authentication.
- Group names, IDs, URLs, and page content must not be written to logs or sent to an LLM.

## Setup

1. Configure `OUTLOOK_PERSONAL_USER`, `OUTLOOK_TOOLS_STATE_DIR`, and `OUTLOOK_GROUPS_CDP_PORT` locally.
2. Create a dedicated Chrome user-data directory beneath the private state directory and sign in to Outlook manually, including MFA.
3. Verify the account and discovery in dry-run mode.
4. Only then consider installing the optional Groups LaunchAgent from a locally rendered plist.

Keep the Chrome debugging endpoint bound to `127.0.0.1`; never expose it to a LAN or public interface. Protect the profile, status, screenshots, and logs with user-only filesystem permissions. Expired sessions require a human to sign in again.

## Operations

The runner has an overlap lock, a bounded runtime, and fixed-code/count-only status. A successful empty-list result means the account was verified and the Groups list was read successfully with zero memberships. It is distinct from login, identity, CDP, or discovery failure.

Use the local LaunchAgent label and state path chosen by your installation. Stop the job before changing code or removing its profile. Do not assume the profile or its authentication state is portable between machines.
