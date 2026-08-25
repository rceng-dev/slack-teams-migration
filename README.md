# Slack to Microsoft Teams migration toolkit

This repository consolidates the existing Slack migration work into one operator-facing command:

```powershell
uv run --with msal --with requests python slack_migrate.py --help
```

It supports Slack channel archives and explicitly authorized 1:1 DMs, preserves original timestamps through Microsoft Graph migration mode, resumes from checkpoints, maps Slack users to Entra ID users, retains threads, optionally uploads files, and adds mapped members after migration.

## What the original folders are

| Folder | Purpose | Git policy |
|---|---|---|
| `slack_export/` | Workspace channel export: `users.json`, `channels.json`, and channel/date message files. | Never committed; it contains company conversations and personal data. |
| `slack_dm_export/` | Per-conversation 1:1 DM exports. | Never committed; it contains private conversations. |
| `slack_to_teams/` | Microsoft Graph migration implementation, local credentials, virtual environment, and resume state. | Code and safe examples are committed; credentials, environment, and state are ignored. |

See [docs/MIGRATION_GUIDE.md](docs/MIGRATION_GUIDE.md) for setup, permissions, dry inspection, phased execution, recovery, and completion safeguards.

## Quick start

1. Copy `slack_to_teams/config.example.json` to `slack_to_teams/config.json` and fill in the non-secret tenant settings.
2. Set secrets for the current PowerShell session:

   ```powershell
   $env:SLACK_MIGRATION_CLIENT_SECRET = "your-entra-client-secret"
   $env:SLACK_MIGRATION_SLACK_TOKEN = "your-slack-user-token" # only for DM export/file download
   ```

3. Inspect local exports without printing message content:

   ```powershell
   uv run --with msal --with requests python slack_migrate.py inspect `
     --channel-export "slack_export/WORKSPACE_EXPORT_FOLDER" `
     --dm-export slack_dm_export
   ```

4. Run channel phases individually and review checkpoints before completion:

   ```powershell
   uv run --with msal --with requests python slack_migrate.py channels usermap
   uv run --with msal --with requests python slack_migrate.py channels team
   uv run --with msal --with requests python slack_migrate.py channels channels
   uv run --with msal --with requests python slack_migrate.py channels messages
   uv run --with msal --with requests python slack_migrate.py channels complete --confirm-complete
   uv run --with msal --with requests python slack_migrate.py channels members
   ```

The `--confirm-complete` flag is intentional: the channel workflow also completes its migration-mode parent team, after which that team cannot accept more historical imports through this legacy creation flow.

## Tests

```powershell
uv run --with msal --with requests python -m unittest discover -s tests -v
```
