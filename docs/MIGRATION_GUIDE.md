# Migration guide

## Scope and safety model

`slack_migrate.py` is the single entry point. It delegates to the channel and DM migration modules while adding export inspection, shared configuration, relative-path support, environment-based secrets, and an explicit finalization guard.

The tool changes Microsoft Teams only during `channels` and `dm` migration phases. `inspect` and DM `audit` are local read-only operations. State files make message imports resumable. Keep the original exports and state directories backed up until the result has been accepted.

Never commit `config.json`, `dm_config.json`, Slack exports, ZIP files, or migration state. The included `.gitignore` protects those paths, but operators should still run `git status --ignored` before every publish.

## Prerequisites

- Python 3.10 or newer; examples use `uv` for an isolated environment.
- An Entra ID app registration with admin consent for the Microsoft Graph application permissions required by the phases you run:
  - `Teamwork.Migrate.All` for migration sessions and historical imports
  - `Team.Create` and `Channel.Create` for the channel workflow
  - `User.Read.All` for Slack-to-Entra user mapping
  - `TeamMember.ReadWrite.All` for adding the owner and mapped members
  - `Chat.Create` for the 1:1 DM workflow
  - `Files.ReadWrite.All` only when uploading attachments
- A real, licensed Teams user as `ownerUpn`.
- Tenant support for Teams migration mode.
- For DM export, a Slack user token belonging to the source participant and authorized to read that conversation.

Review Microsoft's current [Teams import overview](https://learn.microsoft.com/en-us/graph/teams-import-messages) and [permission matrix](https://learn.microsoft.com/en-us/microsoftteams/platform/graph-api/import-messages/import-external-messages-to-teams) before production use; cloud APIs and tenant policies change.

## Configuration

Copy the safe examples:

```powershell
Copy-Item slack_to_teams/config.example.json slack_to_teams/config.json
Copy-Item slack_to_teams/dm_config.example.json slack_to_teams/dm_config.json
```

Relative paths are resolved from the configuration file's directory. Prefer environment variables for secrets:

```powershell
$env:SLACK_MIGRATION_CLIENT_SECRET = "..."
$env:SLACK_MIGRATION_SLACK_TOKEN = "..."
```

The Slack token is needed only for live DM export or downloading Slack-hosted files. It is not needed to import an already-downloaded JSON export without attachments.

## Inventory and validation

The inspector reports counts and paths only; it does not print message bodies:

```powershell
uv run --with msal --with requests python slack_migrate.py inspect `
  --channel-export "slack_export/WORKSPACE_EXPORT_FOLDER" `
  --dm-export slack_dm_export
```

Before any remote operation, verify the intended source export and date range, channel filters, generated user map, target tenant/team/owner, DM participants, storage for attachments, and backups of exports and state.

## Channel migration runbook

Run each phase separately so there is a review point between creation, import, and irreversible completion:

1. `channels usermap` resolves Slack emails to Entra users. Review `state/user_map.json`.
2. `channels team` creates the target team in migration mode.
3. `channels channels` creates migration-mode channels.
4. `channels messages` imports roots, then threaded replies, with original authors and timestamps where mappings exist.
5. `channels files` optionally downloads and uploads attachments when `downloadFiles` is true.
6. Validate counts and spot-check authors, formatting, timestamps, threads, and file links in Teams.
7. `channels complete --confirm-complete` completes every channel and its migration-mode parent team. Microsoft documents parent-team completion as preventing additional imports into that team.
8. `channels members` adds mapped Entra users to the completed team.

`channels all --confirm-complete` exists for controlled reruns, but phased execution is recommended for the first production migration.

## 1:1 DM migration runbook

DMs are deliberately configured one conversation at a time. Confirm participant authorization and applicable retention/privacy requirements before proceeding.

```powershell
uv run --with msal --with requests python slack_migrate.py dm export
uv run --with msal --with requests python slack_migrate.py dm audit
uv run --with msal --with requests python slack_migrate.py dm create
uv run --with msal --with requests python slack_migrate.py dm start
uv run --with msal --with requests python slack_migrate.py dm messages
uv run --with msal --with requests python slack_migrate.py dm complete --confirm-complete
```

Microsoft's current chat migration API allows a later migration session to be started after completion. This tool still requires explicit confirmation because completion changes client visibility and ends the active session.

To migrate another authorized DM, make a separate private config and pass `--dm-config PATH`. Never reuse a state file across conversations.

## Resume and recovery

- Re-run the interrupted phase; imported Slack timestamps and created resource IDs are checkpointed.
- Do not edit state while a process is running; back it up before manual corrections.
- HTTP 429 and transient 5xx responses are retried with capped backoff.
- A failed message remains absent from its checkpoint and is retried on the next run.
- Do not run completion until expected/imported counts agree and the Teams result has been reviewed.

## Known limitations

- Unmapped Slack users use the fallback Teams identity with the original display name prepended.
- Slack formatting is translated to a practical subset of Teams HTML; complex blocks, apps, canvases, reactions, and huddles are not recreated exactly.
- Slack-hosted attachment URLs may be expired and may require a valid Slack token.
- The tool handles standard channels and authorized 1:1 DMs; private/shared channels and group DMs require separate design and permission review.
- Microsoft Graph migration-mode availability, permissions, and behavior are tenant-dependent.
