"""One entry point for Slack export inspection and Microsoft Teams migration."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from slack_to_teams.slack_dm_to_teams import (
    load_json,
    phase_audit as dm_audit,
    phase_complete as dm_complete,
    phase_create as dm_create,
    phase_export as dm_export,
    phase_messages as dm_messages,
    phase_start as dm_start,
)
from slack_to_teams.slack_to_teams import (
    Config,
    Graph,
    PHASES,
    load_slack_channels,
    load_slack_users,
    phase_channels,
    phase_complete,
    phase_files,
    phase_members,
    phase_messages,
    phase_team,
    phase_usermap,
)


ROOT = Path(__file__).resolve().parent
DEFAULT_CONFIG = ROOT / "slack_to_teams" / "config.json"
DEFAULT_DM_CONFIG = ROOT / "slack_to_teams" / "dm_config.json"

CHANNEL_RUNNERS = {
    "usermap": phase_usermap,
    "team": phase_team,
    "channels": phase_channels,
    "messages": phase_messages,
    "files": phase_files,
    "complete": phase_complete,
    "members": phase_members,
}
DM_RUNNERS = {
    "create": dm_create,
    "start": dm_start,
    "messages": dm_messages,
    "complete": dm_complete,
}


def _json(path: Path):
    return json.loads(path.read_text(encoding="utf-8-sig"))


def inspect_exports(channel_export: Path, dm_export: Path) -> int:
    """Print metadata only; never print message bodies or credentials."""
    errors: list[str] = []
    report: dict[str, object] = {}

    if channel_export.is_dir():
        try:
            users = load_slack_users(channel_export)
            channels = load_slack_channels(channel_export)
            day_files = list(channel_export.glob("*/*.json"))
            report["channels"] = {
                "path": str(channel_export),
                "users": len(users),
                "declaredChannels": len(channels),
                "messageDayFiles": len(day_files),
            }
        except (OSError, KeyError, ValueError) as exc:
            errors.append(f"Channel export is invalid: {exc}")
    else:
        errors.append(f"Channel export not found: {channel_export}")

    if dm_export.is_dir():
        conversations = []
        for path in sorted(dm_export.glob("*/messages.json")):
            try:
                data = _json(path)
                messages = data.get("messages", data) if isinstance(data, dict) else data
                conversations.append({"folder": path.parent.name, "messages": len(messages)})
            except (OSError, TypeError, ValueError) as exc:
                errors.append(f"DM export is invalid ({path}): {exc}")
        report["directMessages"] = {
            "path": str(dm_export),
            "conversations": conversations,
        }
    else:
        errors.append(f"DM export not found: {dm_export}")

    report["errors"] = errors
    print(json.dumps(report, indent=2))
    return 1 if errors else 0


def run_channels(config_path: Path, phase: str, confirm_complete: bool) -> None:
    if phase in {"complete", "all"} and not confirm_complete:
        raise SystemExit(
            "This workflow completes the channel and parent-team migration. Re-run with "
            "--confirm-complete after verifying imported messages."
        )
    cfg = Config(config_path)
    graph = Graph(cfg)
    phases = PHASES if phase == "all" else [phase]
    for name in phases:
        print(f"\n=== channel phase: {name} ===", flush=True)
        CHANNEL_RUNNERS[name](cfg, graph)


def run_dm(config_path: Path, dm_config_path: Path, phase: str, confirm_complete: bool) -> None:
    cfg = Config(config_path)
    dm = load_json(dm_config_path)
    dm_base = dm_config_path.resolve().parent
    state_path = Config._resolve_path(dm_base, dm["statePath"])
    export_path = Config._resolve_path(dm_base, dm["exportPath"])
    dm["statePath"] = str(state_path)
    dm["exportPath"] = str(export_path)
    state_path.parent.mkdir(parents=True, exist_ok=True)

    if phase == "export":
        dm_export(cfg, dm)
        return
    if phase == "audit":
        dm_audit(cfg, dm)
        return
    if phase in {"complete", "all"} and not confirm_complete:
        raise SystemExit(
            "Completion makes the imported DM visible and ends its current session. Re-run with "
            "--confirm-complete after verifying imported messages."
        )

    graph = Graph(cfg)
    phases = ["create", "start", "messages", "complete"] if phase == "all" else [phase]
    for name in phases:
        print(f"\n=== DM phase: {name} ===", flush=True)
        DM_RUNNERS[name](cfg, dm, graph)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Inspect Slack exports and migrate channels or 1:1 DMs to Teams."
    )
    sub = parser.add_subparsers(dest="command", required=True)

    inspect_cmd = sub.add_parser("inspect", help="Safely summarize local Slack exports.")
    inspect_cmd.add_argument("--channel-export", type=Path, required=True)
    inspect_cmd.add_argument("--dm-export", type=Path, required=True)

    channels = sub.add_parser("channels", help="Run a channel migration phase.")
    channels.add_argument("phase", choices=PHASES + ["all"])
    channels.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    channels.add_argument("--confirm-complete", action="store_true")

    dm = sub.add_parser("dm", help="Run a 1:1 DM export or migration phase.")
    dm.add_argument(
        "phase",
        choices=["export", "audit", "create", "start", "messages", "complete", "all"],
    )
    dm.add_argument("--config", type=Path, default=DEFAULT_CONFIG)
    dm.add_argument("--dm-config", type=Path, default=DEFAULT_DM_CONFIG)
    dm.add_argument("--confirm-complete", action="store_true")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    if args.command == "inspect":
        raise SystemExit(inspect_exports(args.channel_export, args.dm_export))
    if args.command == "channels":
        run_channels(args.config, args.phase, args.confirm_complete)
        return
    run_dm(args.config, args.dm_config, args.phase, args.confirm_complete)


if __name__ == "__main__":
    main()
