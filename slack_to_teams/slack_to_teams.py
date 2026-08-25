"""Slack -> Microsoft Teams migration via Microsoft Graph Import (migration) Mode.

Setup (one-time, in Azure):
  1. Register an app: Azure portal -> Entra ID -> App registrations -> New registration.
  2. API permissions (Application, admin-consented):
       Teamwork.Migrate.All
       Team.Create
       Channel.Create
       User.Read.All
       TeamMember.ReadWrite.All
       Chat.Create                 (only for 1:1 DM migration)
       Files.ReadWrite.All         (only if uploading attachments)
  3. Create a client secret. Copy tenantId, clientId, secret into config.json.
  4. The Teams tenant must allow migration; the team owner UPN must be a real licensed user.

Run order:
  python slack_to_teams.py usermap        # build slack -> AAD user map
  python slack_to_teams.py team           # create the team in migration mode
  python slack_to_teams.py channels       # create channels in migration mode
  python slack_to_teams.py messages       # post all messages preserving timestamps
  python slack_to_teams.py files          # (optional) upload attachments
  python slack_to_teams.py complete       # finalize migration on channels + team
  python slack_to_teams.py members        # add tenant users as team members

Or:
  python slack_to_teams.py all            # everything in order

State is written to <stateDir>/*.json so phases are idempotent and resumable.
"""

from __future__ import annotations

import argparse
import html
import json
import os
import re
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

import msal
import requests


GRAPH = "https://graph.microsoft.com/v1.0"
GRAPH_BETA = "https://graph.microsoft.com/beta"


# ------------- config + state ---------------------------------------------

class Config:
    def __init__(self, path: Path):
        data = json.loads(path.read_text(encoding="utf-8"))
        base_dir = path.resolve().parent
        self.tenant_id = data["tenantId"]
        self.client_id = data["clientId"]
        self.client_secret = os.environ.get(
            "SLACK_MIGRATION_CLIENT_SECRET", data.get("clientSecret", "")
        )
        self.owner_upn = data["ownerUpn"]
        self.export_path = self._resolve_path(base_dir, data["exportPath"])
        self.state_dir = self._resolve_path(base_dir, data["stateDir"])
        self.team_name = data["teamName"]
        self.team_description = data.get("teamDescription", "")
        self.channels_include = set(data.get("channelsInclude") or [])
        self.channels_exclude = set(data.get("channelsExclude") or [])
        self.skip_join_leave = bool(data.get("skipJoinLeaveMessages", True))
        self.fallback_user_name = data.get("fallbackUserDisplayName", "Slack Import")
        self.slack_user_token = os.environ.get(
            "SLACK_MIGRATION_SLACK_TOKEN", data.get("slackUserToken", "")
        )
        self.download_files = bool(data.get("downloadFiles", False))
        self.rate_sleep = float(data.get("rateLimitSleepSeconds", 0.0))
        if not self.client_secret or self.client_secret.startswith("REPLACE_"):
            raise ValueError(
                "Set SLACK_MIGRATION_CLIENT_SECRET or clientSecret in the private config."
            )
        self.state_dir.mkdir(parents=True, exist_ok=True)

    @staticmethod
    def _resolve_path(base_dir: Path, value: str) -> Path:
        path = Path(value).expanduser()
        return path if path.is_absolute() else (base_dir / path).resolve()


def load_state(path: Path) -> dict:
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    return {}


def save_state(path: Path, data: dict) -> None:
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    tmp.replace(path)


# ------------- auth + http -------------------------------------------------

class Graph:
    def __init__(self, cfg: Config):
        self._cfg = cfg
        self._app = msal.ConfidentialClientApplication(
            cfg.client_id,
            authority=f"https://login.microsoftonline.com/{cfg.tenant_id}",
            client_credential=cfg.client_secret,
        )
        self._token: str | None = None
        self._token_exp: float = 0
        self._sess = requests.Session()

    def _bearer(self) -> str:
        if self._token and time.time() < self._token_exp - 60:
            return self._token
        result = self._app.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )
        if "access_token" not in result:
            raise RuntimeError(f"Auth failed: {result}")
        self._token = result["access_token"]
        self._token_exp = time.time() + int(result.get("expires_in", 3600))
        return self._token

    def request(
        self,
        method: str,
        url: str,
        *,
        json_body: Any = None,
        params: dict | None = None,
        headers: dict | None = None,
        data: bytes | None = None,
        stream: bool = False,
    ) -> requests.Response:
        if not url.startswith("http"):
            url = GRAPH + url
        for attempt in range(8):
            h = {"Authorization": f"Bearer {self._bearer()}"}
            if headers:
                h.update(headers)
            if json_body is not None and "Content-Type" not in h:
                h["Content-Type"] = "application/json"
            r = self._sess.request(
                method, url, json=json_body, params=params, headers=h,
                data=data, stream=stream, timeout=120,
            )
            if r.status_code == 429 or 500 <= r.status_code < 600:
                wait = float(r.headers.get("Retry-After", 2 ** attempt))
                wait = min(wait, 60)
                print(f"  [retry {attempt+1}] {r.status_code} on {method} {url} -- sleeping {wait:.1f}s", flush=True)
                time.sleep(wait)
                continue
            if self._cfg.rate_sleep:
                time.sleep(self._cfg.rate_sleep)
            return r
        return r  # type: ignore[return-value]

    def get(self, url: str, **kw): return self.request("GET", url, **kw)
    def post(self, url: str, **kw): return self.request("POST", url, **kw)
    def put(self, url: str, **kw): return self.request("PUT", url, **kw)


def must_ok(r: requests.Response) -> dict:
    if not (200 <= r.status_code < 300):
        raise RuntimeError(f"{r.status_code} {r.request.method} {r.url}\n{r.text}")
    if r.text:
        try:
            return r.json()
        except ValueError:
            return {"_raw": r.text}
    return {}


# ------------- Slack export readers ---------------------------------------

def load_slack_users(export: Path) -> dict[str, dict]:
    rows = json.loads((export / "users.json").read_text(encoding="utf-8"))
    return {u["id"]: u for u in rows}


def load_slack_channels(export: Path) -> list[dict]:
    return json.loads((export / "channels.json").read_text(encoding="utf-8"))


def channel_dirs(export: Path) -> list[Path]:
    out = []
    for p in sorted(export.iterdir()):
        if p.is_dir() and any(p.glob("*.json")):
            out.append(p)
    return out


def iter_channel_messages(channel_dir: Path) -> Iterable[dict]:
    """Yield messages in chronological order across all date files."""
    for day_file in sorted(channel_dir.glob("*.json")):
        try:
            rows = json.loads(day_file.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            print(f"  ! skipping malformed {day_file}", flush=True)
            continue
        for m in rows:
            yield m


# ------------- Slack mrkdwn -> Teams HTML ---------------------------------

_LINK_RE = re.compile(r"<([^>|]+)\|?([^>]*)>")
_BOLD_RE = re.compile(r"(?<!\w)\*([^\*\n]+)\*(?!\w)")
_ITAL_RE = re.compile(r"(?<!\w)_([^_\n]+)_(?!\w)")
_STRK_RE = re.compile(r"(?<!\w)~([^~\n]+)~(?!\w)")
_CODE_RE = re.compile(r"`([^`\n]+)`")
_BLOCK_RE = re.compile(r"```([\s\S]+?)```")


def slack_to_html(text: str, slack_users: dict[str, dict], channel_name_by_id: dict[str, str]) -> str:
    if not text:
        return ""
    placeholders: dict[str, str] = {}

    def stash(s: str) -> str:
        key = f"\x00P{len(placeholders)}\x00"
        placeholders[key] = s
        return key

    # Code blocks first
    def repl_block(m):
        return stash(f"<pre><code>{html.escape(m.group(1))}</code></pre>")

    out = _BLOCK_RE.sub(repl_block, text)

    # Inline code
    def repl_code(m):
        return stash(f"<code>{html.escape(m.group(1))}</code>")

    out = _CODE_RE.sub(repl_code, out)

    # <@U123>, <#C123|name>, <url|text>, <url>
    def repl_link(m):
        target, label = m.group(1), m.group(2)
        if target.startswith("@"):
            uid = target[1:]
            u = slack_users.get(uid)
            name = (u or {}).get("profile", {}).get("real_name") or (u or {}).get("name") or uid
            return stash(f"<b>@{html.escape(name)}</b>")
        if target.startswith("#"):
            cid = target[1:].split("|")[0]
            name = label or channel_name_by_id.get(cid, cid)
            return stash(f"<b>#{html.escape(name)}</b>")
        if target.startswith("!"):
            # special: !channel, !here, !everyone
            tag = target[1:]
            return stash(f"<b>@{html.escape(tag)}</b>")
        url = target
        text_label = label or url
        return stash(f'<a href="{html.escape(url, quote=True)}">{html.escape(text_label)}</a>')

    out = _LINK_RE.sub(repl_link, out)

    # Now escape the rest, then markdown-ish
    out = html.escape(out)

    out = _BOLD_RE.sub(r"<b>\1</b>", out)
    out = _ITAL_RE.sub(r"<i>\1</i>", out)
    out = _STRK_RE.sub(r"<s>\1</s>", out)

    # Restore placeholders (escape double-escaped them)
    for k, v in placeholders.items():
        out = out.replace(html.escape(k), v).replace(k, v)

    out = out.replace("\n", "<br/>")
    return out


# ------------- Phase: usermap ---------------------------------------------

def phase_usermap(cfg: Config, g: Graph) -> None:
    state_path = cfg.state_dir / "user_map.json"
    state = load_state(state_path)
    slack_users = load_slack_users(cfg.export_path)
    print(f"Mapping {len(slack_users)} Slack users -> AAD users...", flush=True)
    found = 0
    for sid, u in slack_users.items():
        if sid in state and state[sid].get("aadId"):
            found += 1
            continue
        email = (u.get("profile") or {}).get("email")
        display = (u.get("profile") or {}).get("real_name") or u.get("name") or sid
        rec: dict[str, Any] = {
            "slackId": sid,
            "email": email,
            "displayName": display,
            "isBot": bool(u.get("is_bot")),
            "deleted": bool(u.get("deleted")),
            "aadId": None,
            "upn": None,
        }
        if email and not u.get("is_bot"):
            r = g.get(f"/users/{email}")
            if r.status_code == 200:
                d = r.json()
                rec["aadId"] = d.get("id")
                rec["upn"] = d.get("userPrincipalName")
                found += 1
            else:
                # try filter
                r = g.get("/users", params={"$filter": f"mail eq '{email}'", "$select": "id,userPrincipalName"})
                if r.status_code == 200 and r.json().get("value"):
                    d = r.json()["value"][0]
                    rec["aadId"] = d.get("id")
                    rec["upn"] = d.get("userPrincipalName")
                    found += 1
        state[sid] = rec
        if len(state) % 25 == 0:
            save_state(state_path, state)
    save_state(state_path, state)
    print(f"  matched {found}/{len(slack_users)} -- review {state_path} before proceeding.", flush=True)


# ------------- Phase: team + channels -------------------------------------

def _ts_to_iso(ts: float | int | str) -> str:
    f = float(ts)
    dt = datetime.fromtimestamp(f, tz=timezone.utc)
    return dt.strftime("%Y-%m-%dT%H:%M:%S.") + f"{dt.microsecond // 1000:03d}Z"


def _unique_message_times(channel_dir: Path) -> dict[str, str]:
    """Return stable, millisecond-unique Graph timestamps for Slack messages."""
    out: dict[str, str] = {}
    last_ms = -1
    for msg in iter_channel_messages(channel_dir):
        ts = msg.get("ts")
        if not ts:
            continue
        # Graph stores only millisecond precision and rejects duplicate imported
        # timestamps. Slack timestamps have microseconds, so advance collisions by
        # one millisecond while retaining chronological order.
        raw_ms = int(float(ts) * 1000)
        assigned_ms = max(raw_ms, last_ms + 1)
        out[str(ts)] = _ts_to_iso(assigned_ms / 1000)
        last_ms = assigned_ms
    return out


def phase_team(cfg: Config, g: Graph) -> None:
    state_path = cfg.state_dir / "team.json"
    state = load_state(state_path)
    if state.get("teamId"):
        print(f"Team already created: {state['teamId']}", flush=True)
        return

    owner = must_ok(g.get(f"/users/{cfg.owner_upn}"))

    # The imported team must predate its channels, which in turn must predate
    # their imported messages. Derive a safe historical timestamp from Slack.
    created_ts = time.time() - 5
    for channel in load_slack_channels(cfg.export_path):
        if channel.get("created"):
            created_ts = min(created_ts, float(channel["created"]))
    for channel_dir in channel_dirs(cfg.export_path):
        for msg in iter_channel_messages(channel_dir):
            if msg.get("ts"):
                created_ts = min(created_ts, float(msg["ts"]))
                break
    created_ts -= 120

    body = {
        "@microsoft.graph.teamCreationMode": "migration",
        "template@odata.bind": "https://graph.microsoft.com/v1.0/teamsTemplates('standard')",
        "displayName": cfg.team_name,
        "description": cfg.team_description,
        "createdDateTime": _ts_to_iso(created_ts),
    }
    print("Creating team in migration mode...", flush=True)
    r = g.post("/teams", json_body=body)
    if r.status_code not in (201, 202):
        raise RuntimeError(f"Create team failed: {r.status_code} {r.text}")
    location = r.headers.get("Location") or r.headers.get("Content-Location") or ""
    # Location like /teams('{id}')/operations('{opId}')
    m = re.search(r"/teams\('([^']+)'\)", location)
    if not m:
        # Some tenants return body
        try:
            tid = r.json().get("id")
        except ValueError:
            tid = None
        if not tid:
            raise RuntimeError(f"Could not parse team id from response: {r.headers} body={r.text}")
        team_id = tid
    else:
        team_id = m.group(1)

    state["teamId"] = team_id
    state["ownerAadId"] = owner["id"]
    save_state(state_path, state)
    print(f"  teamId={team_id}", flush=True)


def phase_channels(cfg: Config, g: Graph) -> None:
    team_state = load_state(cfg.state_dir / "team.json")
    team_id = team_state.get("teamId")
    if not team_id:
        raise RuntimeError("Run 'team' phase first.")

    state_path = cfg.state_dir / "channels.json"
    state = load_state(state_path)
    slack_channels = {c["name"]: c for c in load_slack_channels(cfg.export_path)}
    dirs = channel_dirs(cfg.export_path)

    for d in dirs:
        name = d.name
        if cfg.channels_include and name not in cfg.channels_include:
            continue
        if name in cfg.channels_exclude:
            continue
        if name in state and state[name].get("channelId"):
            print(f"  skip {name} (already created)", flush=True)
            continue

        meta = slack_channels.get(name) or {}
        created_ts = meta.get("created") or time.time() - 1
        # If actual messages predate the Slack channel's "created" (renamed/re-created
        # channels), Graph rejects all earlier message imports. Use the earliest
        # message ts in the export instead.
        earliest = None
        for m in iter_channel_messages(d):
            t = m.get("ts")
            if t:
                tf = float(t)
                if earliest is None or tf < earliest:
                    earliest = tf
        if earliest is not None:
            created_ts = min(float(created_ts), earliest) - 60
        # Teams names limited to 50 chars and a fixed charset
        base_name = re.sub(r"[^A-Za-z0-9 _\-]+", "_", name)[:50] or name[:50]
        teams_name = base_name

        body = {
            "@microsoft.graph.channelCreationMode": "migration",
            "displayName": teams_name,
            "description": (meta.get("purpose") or {}).get("value") or (meta.get("topic") or {}).get("value") or "",
            "membershipType": "standard",
            "createdDateTime": _ts_to_iso(created_ts),
        }
        cid = None
        for attempt in range(4):
            body["displayName"] = teams_name
            print(f"Creating channel {teams_name} ...", flush=True)
            r = g.post(f"/teams/{team_id}/channels", json_body=body)
            if r.status_code in (201, 202):
                cid = r.json().get("id")
                break
            if "ChannelNameAlreadyExist" in r.text:
                teams_name = f"{base_name}_v{attempt + 2}"[:50]
                continue
            print(f"  ! failed: {r.status_code} {r.text}", flush=True)
            break
        if not cid:
            continue
        state[name] = {"channelId": cid, "teamsName": teams_name, "slackMeta": meta}
        save_state(state_path, state)
        print(f"  -> {cid}", flush=True)


# ------------- Phase: messages --------------------------------------------

_SUBTYPES_SKIP = {
    "channel_join", "channel_leave", "channel_archive", "channel_unarchive",
    "channel_topic", "channel_purpose", "channel_name", "bot_add", "bot_remove",
}


def _build_message_body(
    msg: dict,
    cfg: Config,
    user_map: dict[str, dict],
    slack_users: dict[str, dict],
    channel_names_by_id: dict[str, str],
    fallback_aad_id: str,
    created_date_time: str,
) -> dict | None:
    text = msg.get("text") or ""
    if not text and not msg.get("attachments") and not msg.get("files") and not msg.get("blocks"):
        return None

    html_body = slack_to_html(text, slack_users, channel_names_by_id)

    extras: list[str] = []
    for a in msg.get("attachments") or []:
        if a.get("title") or a.get("text"):
            t = a.get("title", "")
            txt = a.get("text", "")
            extras.append(f"<blockquote>{html.escape(t)}<br/>{html.escape(txt)}</blockquote>")
    for f in msg.get("files") or []:
        nm = f.get("name") or f.get("title") or "file"
        url = f.get("permalink") or f.get("url_private") or ""
        if url:
            extras.append(f'<i>[file] <a href="{html.escape(url)}">{html.escape(nm)}</a></i>')
        else:
            extras.append(f"<i>[file] {html.escape(nm)}</i>")
    if extras:
        html_body = (html_body + "<br/>" if html_body else "") + "<br/>".join(extras)

    if not html_body:
        return None

    sid = msg.get("user") or msg.get("bot_id") or ""
    mapped = user_map.get(sid) or {}
    aad_id = mapped.get("aadId")
    display = (
        mapped.get("displayName")
        or (slack_users.get(sid, {}).get("profile") or {}).get("real_name")
        or cfg.fallback_user_name
    )

    if aad_id:
        from_block = {"user": {"id": aad_id, "displayName": display, "userIdentityType": "aadUser"}}
    else:
        # No AAD match: post as the owner (the only AAD identity we know exists)
        # but prepend the original Slack author name so attribution is preserved.
        from_block = {"user": {"id": fallback_aad_id, "displayName": cfg.fallback_user_name, "userIdentityType": "aadUser"}}
        html_body = f"<b>{html.escape(display)}:</b> " + html_body

    body = {
        "createdDateTime": created_date_time,
        "from": from_block,
        "body": {"contentType": "html", "content": html_body},
    }
    return body


def phase_messages(cfg: Config, g: Graph) -> None:
    team_state = load_state(cfg.state_dir / "team.json")
    chan_state = load_state(cfg.state_dir / "channels.json")
    team_id = team_state.get("teamId")
    if not team_id:
        raise RuntimeError("Run 'team' first.")

    user_map = load_state(cfg.state_dir / "user_map.json")
    slack_users = load_slack_users(cfg.export_path)
    slack_channels = load_slack_channels(cfg.export_path)
    channel_names_by_id = {c["id"]: c["name"] for c in slack_channels}

    fallback_aad_id = team_state.get("ownerAadId")
    if not fallback_aad_id:
        owner = must_ok(g.get(f"/users/{cfg.owner_upn}"))
        fallback_aad_id = owner["id"]
        team_state["ownerAadId"] = fallback_aad_id
        save_state(cfg.state_dir / "team.json", team_state)

    progress_path = cfg.state_dir / "messages_progress.json"
    progress = load_state(progress_path)

    for chan_name, info in chan_state.items():
        if cfg.channels_include and chan_name not in cfg.channels_include:
            continue
        if chan_name in cfg.channels_exclude:
            continue
        cid = info["channelId"]
        chan_dir = cfg.export_path / chan_name
        if not chan_dir.is_dir():
            continue

        chan_progress = progress.setdefault(chan_name, {
            "lastTs": "0",
            "tsToTeamsId": {},  # slack thread parent ts -> teams root message id
        })
        created_times = _unique_message_times(chan_dir)

        print(f"Importing messages for #{chan_name} ...", flush=True)
        # Pass 1: top-level messages (not thread replies)
        for msg in iter_channel_messages(chan_dir):
            if cfg.skip_join_leave and msg.get("subtype") in _SUBTYPES_SKIP:
                continue
            ts = msg.get("ts")
            if not ts:
                continue
            thread_ts = msg.get("thread_ts")
            is_reply = thread_ts and thread_ts != ts
            if is_reply:
                continue
            if ts in chan_progress["tsToTeamsId"]:
                continue
            body = _build_message_body(
                msg, cfg, user_map, slack_users, channel_names_by_id,
                fallback_aad_id, created_times[str(ts)],
            )
            if body is None:
                continue
            r = g.post(f"/teams/{team_id}/channels/{cid}/messages", json_body=body)
            if r.status_code not in (200, 201):
                print(f"  ! root msg failed ts={ts}: {r.status_code} {r.text[:300]}", flush=True)
                continue
            mid = r.json().get("id")
            chan_progress["tsToTeamsId"][ts] = mid
            chan_progress["lastTs"] = max(chan_progress["lastTs"], ts)
            if len(chan_progress["tsToTeamsId"]) % 25 == 0:
                save_state(progress_path, progress)

        save_state(progress_path, progress)

        # Pass 2: thread replies
        for msg in iter_channel_messages(chan_dir):
            if cfg.skip_join_leave and msg.get("subtype") in _SUBTYPES_SKIP:
                continue
            ts = msg.get("ts")
            thread_ts = msg.get("thread_ts")
            if not ts or not thread_ts or thread_ts == ts:
                continue
            parent_id = chan_progress["tsToTeamsId"].get(thread_ts)
            if not parent_id:
                # Parent not imported (was a skipped subtype, etc.) -- post inline as root.
                body = _build_message_body(
                    msg, cfg, user_map, slack_users, channel_names_by_id,
                    fallback_aad_id, created_times[str(ts)],
                )
                if body is None:
                    continue
                r = g.post(f"/teams/{team_id}/channels/{cid}/messages", json_body=body)
                if r.status_code in (200, 201):
                    chan_progress["tsToTeamsId"][ts] = r.json().get("id")
                continue
            if ts in chan_progress["tsToTeamsId"]:
                continue
            body = _build_message_body(
                msg, cfg, user_map, slack_users, channel_names_by_id,
                fallback_aad_id, created_times[str(ts)],
            )
            if body is None:
                continue
            r = g.post(
                f"/teams/{team_id}/channels/{cid}/messages/{parent_id}/replies",
                json_body=body,
            )
            if r.status_code not in (200, 201):
                print(f"  ! reply failed ts={ts}: {r.status_code} {r.text[:300]}", flush=True)
                continue
            chan_progress["tsToTeamsId"][ts] = r.json().get("id")
            if len(chan_progress["tsToTeamsId"]) % 25 == 0:
                save_state(progress_path, progress)

        save_state(progress_path, progress)
        print(f"  done #{chan_name}: {len(chan_progress['tsToTeamsId'])} messages", flush=True)


# ------------- Phase: files (optional) ------------------------------------

def phase_files(cfg: Config, g: Graph) -> None:
    """Download Slack-hosted files and upload to each channel's SharePoint Files folder.

    Requires either valid tokens still embedded in url_private (export tokens are typically
    short-lived) or a workspace token in cfg.slack_user_token. Files are linked back into
    Teams as plain links inside a follow-up message in the same thread.
    """
    if not cfg.download_files:
        print("downloadFiles=false in config; skipping files phase.", flush=True)
        return

    team_state = load_state(cfg.state_dir / "team.json")
    chan_state = load_state(cfg.state_dir / "channels.json")
    team_id = team_state["teamId"]
    progress = load_state(cfg.state_dir / "messages_progress.json")
    files_state = load_state(cfg.state_dir / "files.json")

    headers_slack = {"Authorization": f"Bearer {cfg.slack_user_token}"} if cfg.slack_user_token else {}

    for chan_name, info in chan_state.items():
        cid = info["channelId"]
        chan_dir = cfg.export_path / chan_name
        # Get drive folder for the channel
        r = g.get(f"/teams/{team_id}/channels/{cid}/filesFolder")
        if r.status_code != 200:
            print(f"  ! cannot resolve filesFolder for #{chan_name}: {r.status_code}", flush=True)
            continue
        ff = r.json()
        drive_id = ff["parentReference"]["driveId"]
        item_id = ff["id"]

        chan_progress = progress.get(chan_name, {})
        ts_to_id = chan_progress.get("tsToTeamsId", {})
        cstate = files_state.setdefault(chan_name, {})

        for msg in iter_channel_messages(chan_dir):
            ts = msg.get("ts")
            for f in msg.get("files") or []:
                fid = f.get("id")
                if not fid or cstate.get(fid):
                    continue
                url = f.get("url_private_download") or f.get("url_private")
                if not url:
                    continue
                name = f.get("name") or f.get("title") or fid
                try:
                    dl = requests.get(url, headers=headers_slack, timeout=120)
                except requests.RequestException as e:
                    print(f"  ! download err {fid}: {e}", flush=True)
                    continue
                if dl.status_code != 200 or dl.headers.get("Content-Type", "").startswith("text/html"):
                    print(f"  ! file {fid} not accessible ({dl.status_code})", flush=True)
                    continue
                # Upload to channel folder
                upload_path = name.replace("/", "_")
                up = g.put(
                    f"/drives/{drive_id}/items/{item_id}:/{upload_path}:/content",
                    data=dl.content,
                    headers={"Content-Type": f.get("mimetype", "application/octet-stream")},
                )
                if up.status_code not in (200, 201):
                    print(f"  ! upload failed {fid}: {up.status_code} {up.text[:200]}", flush=True)
                    continue
                webUrl = up.json().get("webUrl")
                cstate[fid] = {"webUrl": webUrl, "name": name}
                # Post a small follow-up linking the file in-thread
                parent = ts_to_id.get(msg.get("thread_ts") or ts) or ts_to_id.get(ts)
                if parent and webUrl:
                    body = {
                        "createdDateTime": _ts_to_iso(ts),
                        "body": {"contentType": "html",
                                 "content": f'<i>[uploaded file]</i> <a href="{html.escape(webUrl)}">{html.escape(name)}</a>'},
                    }
                    g.post(f"/teams/{team_id}/channels/{cid}/messages/{parent}/replies", json_body=body)
                save_state(cfg.state_dir / "files.json", files_state)


# ------------- Phase: complete + members ----------------------------------

def phase_complete(cfg: Config, g: Graph) -> None:
    team_state = load_state(cfg.state_dir / "team.json")
    chan_state = load_state(cfg.state_dir / "channels.json")
    team_id = team_state["teamId"]
    for chan_name, info in chan_state.items():
        cid = info["channelId"]
        if info.get("completed"):
            continue
        r = g.post(f"/teams/{team_id}/channels/{cid}/completeMigration")
        if r.status_code in (200, 204):
            info["completed"] = True
            save_state(cfg.state_dir / "channels.json", chan_state)
            print(f"  completed channel #{chan_name}", flush=True)
        else:
            print(f"  ! complete channel #{chan_name}: {r.status_code} {r.text[:200]}", flush=True)
    if not team_state.get("completed"):
        r = g.post(f"/teams/{team_id}/completeMigration")
        if r.status_code in (200, 204):
            team_state["completed"] = True
            save_state(cfg.state_dir / "team.json", team_state)
            print("  team migration completed.", flush=True)
            # Owner must be added immediately so the team isn't ownerless.
            owner_aad = team_state.get("ownerAadId")
            if owner_aad:
                g.post(f"/teams/{team_id}/members", json_body={
                    "@odata.type": "#microsoft.graph.aadUserConversationMember",
                    "roles": ["owner"],
                    "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{owner_aad}')",
                })
        else:
            print(f"  ! complete team: {r.status_code} {r.text[:200]}", flush=True)


def phase_members(cfg: Config, g: Graph) -> None:
    team_state = load_state(cfg.state_dir / "team.json")
    team_id = team_state["teamId"]
    user_map = load_state(cfg.state_dir / "user_map.json")
    added_path = cfg.state_dir / "members_added.json"
    added = load_state(added_path)
    for sid, rec in user_map.items():
        aad = rec.get("aadId")
        if not aad or added.get(sid):
            continue
        r = g.post(f"/teams/{team_id}/members", json_body={
            "@odata.type": "#microsoft.graph.aadUserConversationMember",
            "roles": [],
            "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{aad}')",
        })
        if r.status_code in (200, 201):
            added[sid] = True
            save_state(added_path, added)
        elif r.status_code == 409:
            added[sid] = True
            save_state(added_path, added)
        else:
            print(f"  ! add member {rec.get('displayName')}: {r.status_code} {r.text[:200]}", flush=True)


# ------------- CLI --------------------------------------------------------

PHASES = ["usermap", "team", "channels", "messages", "files", "complete", "members"]


def main() -> None:
    p = argparse.ArgumentParser()
    p.add_argument("phase", choices=PHASES + ["all"])
    p.add_argument("--config", default="config.json")
    args = p.parse_args()

    cfg_path = Path(args.config)
    if not cfg_path.is_absolute():
        cfg_path = Path(__file__).parent / cfg_path
    if not cfg_path.exists():
        print(f"Config not found: {cfg_path}. Copy config.example.json to config.json and fill it in.", file=sys.stderr)
        sys.exit(2)

    cfg = Config(cfg_path)
    g = Graph(cfg)

    runners = {
        "usermap": phase_usermap,
        "team": phase_team,
        "channels": phase_channels,
        "messages": phase_messages,
        "files": phase_files,
        "complete": phase_complete,
        "members": phase_members,
    }
    phases = PHASES if args.phase == "all" else [args.phase]
    for ph in phases:
        print(f"\n=== phase: {ph} ===", flush=True)
        runners[ph](cfg, g)


if __name__ == "__main__":
    main()
