"""Migrate one authorized Slack 1:1 DM into a Microsoft Teams 1:1 chat."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path
from typing import Any

import requests

try:
    from .slack_to_teams import (
        Config,
        Graph,
        _build_message_body,
        _ts_to_iso,
        load_state,
        must_ok,
        save_state,
    )
except ImportError:  # Allow direct execution from this directory.
    from slack_to_teams import (
        Config,
        Graph,
        _build_message_body,
        _ts_to_iso,
        load_state,
        must_ok,
        save_state,
    )


ROOT = Path(__file__).parent


def load_json(path: Path) -> dict:
    return json.loads(path.read_text(encoding="utf-8-sig"))


def slack_call(session: requests.Session, method: str, **params: Any) -> dict:
    response = session.get(
        f"https://slack.com/api/{method}", params=params, timeout=120
    )
    response.raise_for_status()
    data = response.json()
    if not data.get("ok"):
        raise RuntimeError(f"Slack {method} failed: {data.get('error')}")
    return data


def slack_pages(session: requests.Session, method: str, key: str, **params: Any) -> list[dict]:
    rows: list[dict] = []
    cursor = ""
    while True:
        data = slack_call(session, method, cursor=cursor, limit=200, **params)
        rows.extend(data.get(key, []))
        cursor = data.get("response_metadata", {}).get("next_cursor", "")
        if not cursor:
            return rows


def phase_export(cfg: Config, dm: dict) -> None:
    session = requests.Session()
    session.headers["Authorization"] = f"Bearer {cfg.slack_user_token}"
    auth = slack_call(session, "auth.test")
    if auth.get("user_id") != dm["slackSourceUserId"]:
        raise RuntimeError(
            f"Slack token belongs to {auth.get('user_id')}, expected {dm['slackSourceUserId']}"
        )

    users = slack_pages(session, "users.list", "members")
    user_ids = {u.get("id") for u in users}
    for required in (dm["slackSourceUserId"], dm["slackTargetUserId"]):
        if required not in user_ids:
            raise RuntimeError(f"Slack user missing: {required}")

    roots = slack_pages(
        session,
        "conversations.history",
        "messages",
        channel=dm["slackConversationId"],
    )
    messages_by_ts: dict[str, dict] = {m["ts"]: m for m in roots if m.get("ts")}
    for root in roots:
        if not root.get("reply_count") or not root.get("ts"):
            continue
        replies = slack_pages(
            session,
            "conversations.replies",
            "messages",
            channel=dm["slackConversationId"],
            ts=root["ts"],
        )
        for reply in replies:
            if reply.get("ts"):
                messages_by_ts[reply["ts"]] = reply

    messages = sorted(messages_by_ts.values(), key=lambda m: float(m["ts"]))
    referenced_user_ids = {
        dm["slackSourceUserId"],
        dm["slackTargetUserId"],
    }
    for message in messages:
        referenced_user_ids.update(
            re.findall(r"<@([A-Z0-9]+)>", message.get("text") or "")
        )
    users = [u for u in users if u.get("id") in referenced_user_ids]
    export_path = Path(dm["exportPath"])
    export_path.parent.mkdir(parents=True, exist_ok=True)
    export_path.write_text(
        json.dumps({"users": users, "messages": messages}, indent=2),
        encoding="utf-8",
    )
    print(f"Exported {len(messages)} Slack DM messages (content not displayed).")


def load_export(dm: dict) -> tuple[dict[str, dict], list[dict]]:
    data = load_json(Path(dm["exportPath"]))
    users = {u["id"]: u for u in data["users"]}
    return users, data["messages"]


def prepared_messages(cfg: Config, dm: dict) -> list[tuple[str, dict]]:
    users, messages = load_export(dm)
    user_map = {
        dm["slackSourceUserId"]: {
            "aadId": dm["teamsSourceUserId"],
            "displayName": dm["teamsSourceDisplayName"],
        },
        dm["slackTargetUserId"]: {
            "aadId": dm["teamsTargetUserId"],
            "displayName": dm["teamsTargetDisplayName"],
        },
    }
    prepared: list[tuple[str, dict]] = []
    last_ms = -1
    for message in messages:
        ts = str(message["ts"])
        raw_ms = int(float(ts) * 1000)
        assigned_ms = max(raw_ms, last_ms + 1)
        last_ms = assigned_ms
        body = _build_message_body(
            message,
            cfg,
            user_map,
            users,
            {},
            dm["teamsSourceUserId"],
            _ts_to_iso(assigned_ms / 1000),
        )
        if body is not None:
            prepared.append((ts, body))
    return prepared


def phase_audit(cfg: Config, dm: dict) -> None:
    data = load_json(Path(dm["exportPath"]))
    prepared = prepared_messages(cfg, dm)
    timestamps = [ts for ts, _ in prepared]
    if len(timestamps) != len(set(timestamps)):
        raise RuntimeError("Duplicate Slack timestamps remain after preparation.")
    print(
        "Local DM audit: "
        f"users={len(data['users'])}, raw={len(data['messages'])}, "
        f"importable={len(prepared)}, unique={len(set(timestamps))}"
    )


def phase_create(cfg: Config, dm: dict, graph: Graph) -> None:
    state_path = Path(dm["statePath"])
    state = load_state(state_path)
    if state.get("chatId"):
        print("Teams chat already checkpointed.")
        return

    # Verify both Entra objects before changing Teams.
    must_ok(graph.get(f"/users/{dm['teamsSourceUserId']}"))
    must_ok(graph.get(f"/users/{dm['teamsTargetUserId']}"))
    body = {
        "chatType": "oneOnOne",
        "members": [
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": (
                    "https://graph.microsoft.com/v1.0/users"
                    f"('{dm['teamsSourceUserId']}')"
                ),
            },
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": [dm["teamsTargetRole"]],
                "user@odata.bind": (
                    "https://graph.microsoft.com/v1.0/users"
                    f"('{dm['teamsTargetUserId']}')"
                ),
            },
        ],
    }
    response = graph.post("/chats", json_body=body)
    if response.status_code not in (200, 201, 202):
        raise RuntimeError(f"Create chat failed: {response.status_code} {response.text}")
    try:
        chat_id = response.json().get("id")
    except ValueError:
        chat_id = None
    if not chat_id:
        location = response.headers.get("Location", "")
        match = re.search(r"/chats\('([^']+)'\)", location)
        chat_id = match.group(1) if match else None
    if not chat_id:
        raise RuntimeError("Could not determine Teams chat ID from create response.")
    state["chatId"] = chat_id
    save_state(state_path, state)
    print("Teams 1:1 chat created or resolved and checkpointed.")


def phase_start(cfg: Config, dm: dict, graph: Graph) -> None:
    state_path = Path(dm["statePath"])
    state = load_state(state_path)
    if state.get("migrationStarted"):
        print("Chat migration already started.")
        return
    chat_id = state.get("chatId")
    if not chat_id:
        raise RuntimeError("Run create first.")
    prepared = prepared_messages(cfg, dm)
    if not prepared:
        raise RuntimeError("No importable messages found.")
    earliest = float(prepared[0][0]) - 60
    response = graph.post(
        f"/chats/{chat_id}/startMigration",
        json_body={"conversationCreationDateTime": _ts_to_iso(earliest)},
    )
    if response.status_code not in (200, 204):
        raise RuntimeError(f"Start migration failed: {response.status_code} {response.text}")
    state["migrationStarted"] = True
    state["expectedMessages"] = len(prepared)
    state.setdefault("imported", {})
    save_state(state_path, state)
    print(f"Chat migration started; expecting {len(prepared)} messages.")


def phase_messages(cfg: Config, dm: dict, graph: Graph) -> None:
    state_path = Path(dm["statePath"])
    state = load_state(state_path)
    if not state.get("migrationStarted"):
        raise RuntimeError("Run start first.")
    if state.get("completed"):
        print("Chat migration already completed.")
        return
    chat_id = state["chatId"]
    imported = state.setdefault("imported", {})
    prepared = prepared_messages(cfg, dm)
    for ts, body in prepared:
        if ts in imported:
            continue
        response = graph.post(f"/chats/{chat_id}/messages", json_body=body)
        if response.status_code not in (200, 201):
            raise RuntimeError(
                f"Message import failed at Slack ts={ts}: "
                f"{response.status_code} {response.text[:500]}"
            )
        imported[ts] = response.json().get("id")
        save_state(state_path, state)
    print(f"Imported/checkpointed {len(imported)}/{len(prepared)} messages.")


def phase_complete(cfg: Config, dm: dict, graph: Graph) -> None:
    state_path = Path(dm["statePath"])
    state = load_state(state_path)
    if state.get("completed"):
        print("Chat migration already completed.")
        return
    expected = int(state.get("expectedMessages", 0))
    imported = len(state.get("imported", {}))
    if not expected or imported != expected:
        raise RuntimeError(f"Refusing completion: expected={expected}, imported={imported}")
    response = graph.post(f"/chats/{state['chatId']}/completeMigration")
    if response.status_code not in (200, 204):
        raise RuntimeError(f"Complete migration failed: {response.status_code} {response.text}")
    state["completed"] = True
    save_state(state_path, state)
    print(f"Chat migration completed with {imported} messages.")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "phase",
        choices=["export", "audit", "create", "start", "messages", "complete"],
    )
    parser.add_argument("--config", default=str(ROOT / "config.json"))
    parser.add_argument("--dm-config", default=str(ROOT / "dm_config.json"))
    args = parser.parse_args()
    cfg = Config(Path(args.config))
    dm = load_json(Path(args.dm_config))
    Path(dm["statePath"]).parent.mkdir(parents=True, exist_ok=True)
    if args.phase == "export":
        phase_export(cfg, dm)
        return
    if args.phase == "audit":
        phase_audit(cfg, dm)
        return
    graph = Graph(cfg)
    phases = {
        "create": phase_create,
        "start": phase_start,
        "messages": phase_messages,
        "complete": phase_complete,
    }
    phases[args.phase](cfg, dm, graph)


if __name__ == "__main__":
    main()
