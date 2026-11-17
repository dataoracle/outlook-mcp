"""Outlook mail MCP server (delegated Microsoft Graph access).

Read tools:   whoami, list_folders, list_messages, search_messages, get_message,
              get_thread, get_attachment, find_people.
Write tools (ENABLE_WRITE_TOOLS=true): create_draft, create_reply_draft,
              create_forward_draft, update_draft, delete_draft, mark_read,
              flag_messages, set_categories, move_messages, create_folder.
Send tools (ENABLE_SEND_TOOLS=true): send_draft, send_mail, reply, forward.
Calendar (ENABLE_CALENDAR_TOOLS=true): list_calendars, list_events, get_event,
              find_free_time; with write tools also create_event, update_event,
              respond_to_event, cancel_event (anything that notifies other people
              needs ENABLE_SEND_TOOLS and goes through the same send approval).
"""
from __future__ import annotations

import asyncio
import hashlib
import hmac
import html
import json
import logging
import os
import re
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Annotated
from urllib.parse import quote

import uvicorn
from fastmcp import Context, FastMCP
from fastmcp.exceptions import ToolError
from mcp_types import ElicitRequest, ElicitRequestFormParams, InputRequiredResult
from pydantic import Field
from starlette.requests import Request
from starlette.responses import JSONResponse

import compose
import events as cal
import formatting as fmt
from auth import AuthError, TokenManager, check_client_id, check_tenant
from compose import ComposeError
from extract import ExtractError, extract_text
from graph_client import GraphClient, GraphError
from vault import Vault, VaultError, load_key, vault_context

logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("outlook-mcp")


# ---------------------------------------------------------------- config
def _env(name: str, default: str | None = None) -> str:
    value = os.getenv(name, default)
    if value is None or value.strip() == "":
        raise SystemExit(f"Missing required environment variable {name}")
    return value.strip()


def _secret(env_name: str, cred_name: str) -> str:
    """Load a secret, preferring systemd credentials over environment variables.

    1. $CREDENTIALS_DIRECTORY/<cred_name>  (systemd LoadCredentialEncrypted=)
    2. file named by <ENV_NAME>_FILE       (e.g. Docker secrets)
    3. <ENV_NAME> environment variable      (development only)
    """
    cred_dir = os.getenv("CREDENTIALS_DIRECTORY")
    if cred_dir:
        path = Path(cred_dir) / cred_name
        if path.is_file():
            value = path.read_text(encoding="utf-8").rstrip("\r\n")
            if value:
                return value
    file_var = os.getenv(f"{env_name}_FILE")
    if file_var:
        value = Path(file_var).read_text(encoding="utf-8").rstrip("\r\n")
        if value:
            return value
    return _env(env_name)


def _flag(name: str, default: str) -> bool:
    return os.getenv(name, default).strip().lower() in ("1", "true", "yes", "on")


def _signature() -> str | None:
    path = os.getenv("SIGNATURE_FILE", "").strip()
    if not path:
        return None
    try:
        text = Path(path).read_text(encoding="utf-8").strip()
    except OSError as e:
        raise SystemExit(f"Cannot read SIGNATURE_FILE {path}: {e}") from None
    if not text:
        return None
    if path.lower().endswith((".html", ".htm")):
        return text
    return "<div>" + html.escape(text).replace("\n", "<br>\n") + "</div>"


TENANT = check_tenant(_env("MS_TENANT_ID"))
CLIENT_ID = check_client_id(_env("MS_CLIENT_ID"))
MCP_TOKEN = _secret("MCP_TOKEN", "mcp_token")                 # shared secret for mcp-remote
CLIENT_SECRET = _secret("MS_CLIENT_SECRET", "client_secret")  # Entra app secret (Web platform)
try:
    VAULT_KEY = load_key(_secret("VAULT_KEY", "vault_key"))
except VaultError as e:
    raise SystemExit(str(e)) from None
# systemd sets STATE_DIRECTORY=/var/lib/outlook-mcp (StateDirectory=)
_state_dir = os.getenv("STATE_DIRECTORY", ".").split(":")[0]
VAULT_PATH = os.getenv("VAULT_PATH") or str(Path(_state_dir) / "vault.json.enc")
HOST = os.getenv("HOST", "127.0.0.1")
PORT = int(os.getenv("PORT", "8767"))
ENABLE_WRITE = _flag("ENABLE_WRITE_TOOLS", "true")
ENABLE_SEND = _flag("ENABLE_SEND_TOOLS", "true")
SEND_CONFIRMATION = os.getenv("SEND_CONFIRMATION", "client").strip().lower()
if SEND_CONFIRMATION not in ("client", "elicit"):
    raise SystemExit("SEND_CONFIRMATION must be 'client' or 'elicit'")
ALLOWED_DOMAINS = [d.strip().lower().lstrip("@") for d in os.getenv("SEND_ALLOWED_DOMAINS", "").split(",")
                   if d.strip()]
KEEPALIVE_HOURS = float(os.getenv("KEEPALIVE_HOURS", "12"))
MAX_BODY_CHARS = int(os.getenv("MAX_BODY_CHARS", "20000"))
MAX_ATTACHMENT_MB = float(os.getenv("MAX_ATTACHMENT_MB", "25"))
SIGNATURE_HTML = _signature()
ENABLE_CALENDAR = _flag("ENABLE_CALENDAR_TOOLS", "true")
TIMEZONE = os.getenv("TIMEZONE", "").strip()      # IANA, e.g. Asia/Dubai; empty = mailbox setting
if TIMEZONE and cal.zone(TIMEZONE) is None:
    raise SystemExit(f"TIMEZONE {TIMEZONE!r} is not a known time zone (use an IANA name like Asia/Dubai)")
try:
    WORKING_DAYS = cal.parse_days(os.getenv("WORKING_DAYS", ""))   # empty = from Outlook
    _wh = os.getenv("WORKING_HOURS", "").strip()                   # e.g. 08:00-17:00
    WORKING_HOURS = tuple(cal.parse_hhmm(x) for x in _wh.split("-", 1)) if _wh else None
except (ValueError, IndexError):
    raise SystemExit("WORKING_DAYS must look like 'mon-fri' and WORKING_HOURS like '08:00-17:00'") from None

vault = Vault(VAULT_PATH, VAULT_KEY, vault_context(TENANT, CLIENT_ID))
tokens = TokenManager(vault, TENANT, CLIENT_ID, CLIENT_SECRET)
graph = GraphClient(tokens)

READ_ONLY = {"readOnlyHint": True, "openWorldHint": True}
WRITES = {"readOnlyHint": False, "destructiveHint": False, "idempotentHint": False, "openWorldHint": True}
SENDS = {"readOnlyHint": False, "destructiveHint": True, "idempotentHint": False, "openWorldHint": True}


def _instructions() -> str:
    parts = [
        "Access to the signed-in user's Outlook mailbox (delegated Microsoft Graph permissions). "
        "Use list_messages to browse a folder (inbox by default; filters for unread, sender, "
        "since/until), search_messages to find mail by keyword, sender, recipient or subject, "
        "get_message to read one message and list its attachments, get_thread for a whole "
        "conversation, get_attachment to read an attachment's text. Read tools return compact "
        "text (times in UTC); pass output='json' for structured data.",
    ]
    if ENABLE_WRITE:
        parts.append("create_draft / create_reply_draft / create_forward_draft save drafts "
                     "without sending; show the draft to the user.")
    if ENABLE_SEND:
        parts.append("Mail sent with send_draft, send_mail, reply or forward goes out as the signed-in "
                     "user and cannot be recalled: before every send, show the user the recipients, "
                     "subject and full text and get their explicit approval.")
    else:
        parts.append("This server cannot send mail: save drafts and let the user send them from Outlook.")
    if ENABLE_CALENDAR:
        parts.append("Calendar: list_events for a time range (needs_response=True for unanswered "
                     "invitations), get_event for details, find_free_time for open slots with or without "
                     "attendees. Calendar times are in the user's time zone (see whoami), not UTC.")
        if ENABLE_WRITE:
            parts.append("create_event without attendees only touches the user's calendar. create_event with "
                         "attendees, update_event on a meeting the user organises, respond_to_event (unless "
                         "send_response=False) and cancel_event send mail to other people: get the user's "
                         "explicit approval of the details first, as for sending mail.")
    return " ".join(parts)


mcp = FastMCP("outlook-mail", instructions=_instructions())

# ---------------------------------------------------------------- helpers
# Graph ids are base64 (standard or URL-safe alphabet).
ID_RE = re.compile(r"^[A-Za-z0-9+/=_\-]{8,600}$")
READ = {"annotations": READ_ONLY, "output_schema": None}

Recipients = Annotated[list[str] | str | None, Field(
    description="Email addresses: a list, or one string separated by ';' or ','. "
                "'Name <a@b.com>' is accepted. Plain names are refused.")]
Attachments = Annotated[list[dict] | None, Field(
    description="Optional small files (3 MB total): [{name, content (text) | content_base64, "
                "content_type?}].")]


def _id(value: str, label: str, source: str) -> str:
    value = (value or "").strip()
    if not ID_RE.match(value):
        raise ToolError(f"{label} {value[:60]!r} doesn't look like an Outlook id. Get it from {source} "
                        "(copy it whole: ids are long).")
    return value


def _q(value: str) -> str:
    return quote(value, safe="")


def _ids(values: list[str] | str, label: str = "message_ids") -> list[str]:
    items = values if isinstance(values, list) else re.split(r"[\s,;]+", values or "")
    out = [_id(v, label, "list_messages or search_messages") for v in items if (v or "").strip()]
    if not out:
        raise ToolError(f"Give at least one id in {label}.")
    if len(out) > 100:
        raise ToolError(f"At most 100 ids per call in {label}.")
    return out


def _clamp(n: int, lo: int, hi: int) -> int:
    return max(lo, min(int(n), hi))


def _time(value: str, label: str) -> datetime | None:
    if not (value or "").strip():
        return None
    try:
        return fmt.parse_time(value).astimezone(timezone.utc)
    except ValueError:
        raise ToolError(f"{label} must be an ISO 8601 date or date/time, e.g. 2026-11-13 or "
                        "2026-11-13T08:00:00Z") from None


async def _call(coro):
    """Run a data function and turn backend errors into tool errors."""
    try:
        return await coro
    except (GraphError, AuthError, ComposeError, ExtractError) as e:
        raise ToolError(str(e)) from e


async def _run(coro, render, output: str = "text") -> str:
    """Run a data function; return compact text (default) or compact JSON."""
    if output not in ("text", "json"):
        coro.close()
        raise ToolError("output must be 'text' or 'json'.")
    data = await _call(coro)
    if output == "json":
        return json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    return render(data)


_me_cache: dict = {}


async def _me() -> dict:
    """The signed-in user's id and all their addresses (for 'addressed to me')."""
    uid = tokens.user.get("id")
    if _me_cache.get("id") and _me_cache["id"] == uid:
        return _me_cache
    try:
        me = await graph.get("/me", {"$select": "id,displayName,mail,userPrincipalName,proxyAddresses"})
    except GraphError:
        me = await graph.get("/me", {"$select": "id,displayName,mail,userPrincipalName"})
    addrs = {a.lower() for a in (me.get("mail"), me.get("userPrincipalName")) if a}
    addrs |= {p.split(":", 1)[1].lower() for p in me.get("proxyAddresses") or []
              if p.lower().startswith("smtp:")}
    _me_cache.clear()
    _me_cache.update({"id": me.get("id"), "name": me.get("displayName"), "addresses": addrs,
                      "mail": me.get("mail") or me.get("userPrincipalName")})
    return _me_cache


# ---------------------------------------------------------------- folders
FOLDER_ALIASES = {"inbox": "inbox", "drafts": "drafts", "draft": "drafts", "sent": "sentitems",
                  "sentitems": "sentitems", "sent items": "sentitems", "deleted": "deleteditems",
                  "deleteditems": "deleteditems", "deleted items": "deleteditems", "trash": "deleteditems",
                  "junk": "junkemail", "junkemail": "junkemail", "junk email": "junkemail",
                  "spam": "junkemail", "archive": "archive", "outbox": "outbox"}
FOLDER_TTL = 300
_folders: dict = {"at": 0.0, "list": []}


async def _folder_tree() -> list[dict]:
    """All folders as [{folder_id, path, unread, total, children}], cached for 5 minutes."""
    if _folders["list"] and time.time() - _folders["at"] < FOLDER_TTL:
        return _folders["list"]
    sel = {"$select": "id,displayName,unreadItemCount,totalItemCount,childFolderCount", "$top": 100}
    out: list[dict] = []

    async def walk(path: str, prefix: str, depth: int) -> None:
        rows, _ = await graph.get_all(path, dict(sel), 500)
        for f in rows:
            p = f"{prefix}/{f.get('displayName')}" if prefix else f.get("displayName") or "?"
            out.append(fmt.folder(f, p))
            if f.get("childFolderCount") and depth < 4 and len(out) < 500:
                await walk(f"/me/mailFolders/{_q(f['id'])}/childFolders", p, depth + 1)

    await walk("/me/mailFolders", "", 0)
    _folders.update(at=time.time(), list=out)
    return out


async def _well_known_ids() -> dict[str, str]:
    """Map well-known names to their folder ids (cached with the tree)."""
    if "wk" in _folders and time.time() - _folders["at"] < FOLDER_TTL:
        return _folders["wk"]
    wk: dict[str, str] = {}
    for name in ("inbox", "drafts", "sentitems", "deleteditems", "junkemail", "archive"):
        try:
            wk[name] = (await graph.get(f"/me/mailFolders/{name}", {"$select": "id"}))["id"]
        except GraphError:
            pass
    _folders["wk"] = wk
    return wk


async def _folder(value: str) -> tuple[str, str]:
    """Resolve a folder given as a well-known name, an id, or a display name or path
    ('Projects/Alpha'). Returns (path segment for Graph, display label)."""
    v = (value or "inbox").strip()
    key = v.lower()
    if key in FOLDER_ALIASES:
        wk = FOLDER_ALIASES[key]
        return wk, fmt.WELL_KNOWN.get(wk, wk)
    tree = await _folder_tree()
    for f in tree:
        if f["folder_id"] == v:
            return _q(v), f["path"]
    matches = [f for f in tree if f["path"].lower() == key.strip("/")]
    if not matches:
        matches = [f for f in tree if f["path"].rsplit("/", 1)[-1].lower() == key]
    if len(matches) == 1:
        return _q(matches[0]["folder_id"]), matches[0]["path"]
    if len(matches) > 1:
        raise ToolError(f"Several folders are named {v!r}: " + "; ".join(m["path"] for m in matches)
                        + ". Give the full path or the folder id from list_folders.")
    if ID_RE.match(v) and len(v) > 40:
        return _q(v), "folder"
    raise ToolError(f"No folder {v!r}. Use a well-known name (inbox, drafts, sent, deleted, junk, "
                    "archive), a path like 'Projects/Alpha', or an id from list_folders.")


async def _folder_names() -> dict[str, str]:
    """folder id -> path, for showing where search results live."""
    try:
        return {f["folder_id"]: f["path"] for f in await _folder_tree()}
    except GraphError:
        return {}


# ---------------------------------------------------------------- read data functions
async def whoami_data() -> dict:
    me = await graph.get("/me", {"$select": "id,displayName,mail,userPrincipalName,jobTitle,department,"
                                            "officeLocation"})
    out: dict = {"user": fmt.person(me, ""), "auth": {k: v for k, v in tokens.info().items() if v}}
    try:
        inbox = await graph.get("/me/mailFolders/inbox", {"$select": "unreadItemCount,totalItemCount"})
        out["inbox"] = {"unread": inbox.get("unreadItemCount"), "total": inbox.get("totalItemCount")}
    except GraphError as e:
        out["inbox"] = {"error": str(e)}
    try:
        s = await graph.get("/me/mailboxSettings", {"$select": "timeZone,automaticRepliesSetting,language"})
        out["mailbox"] = {k: v for k, v in {
            "time_zone": s.get("timeZone"),
            "auto_reply": (s.get("automaticRepliesSetting") or {}).get("status"),
            "language": (s.get("language") or {}).get("locale"),
        }.items() if v}
    except GraphError:
        pass  # needs MailboxSettings.Read
    out["server"] = {
        "write_tools": ENABLE_WRITE, "send_tools": ENABLE_SEND,
        "send_confirmation": SEND_CONFIRMATION if ENABLE_SEND else None,
        "send_allowed_domains": ALLOWED_DOMAINS or None,
        "signature": bool(SIGNATURE_HTML), "calendar": ENABLE_CALENDAR,
    }
    if ENABLE_CALENDAR:
        out["calendar"] = {"time_zone": (await _tz()).key,
                           "time_zone_source": "TIMEZONE setting" if TIMEZONE else "mailbox settings"}
    return out


def whoami_text(d: dict) -> str:
    a, s = d["auth"], d["server"]
    u = d["user"]
    lines = [f"Signed in as {u.get('name')} <{u.get('email')}>"
             + "".join(f" · {u[k]}" for k in ("job_title", "department") if u.get(k))]
    lines.append(f"Auth: {a.get('status')}"
                 + (f" · signed in {fmt.short_time(a['signed_in_at'])}" if a.get("signed_in_at") else "")
                 + (f" · token rotated {fmt.short_time(a['refresh_token_rotated_at'])}"
                    if a.get("refresh_token_rotated_at") else "")
                 + (f" · problem: {a['problem']}" if a.get("problem") else ""))
    i = d.get("inbox") or {}
    lines.append(f"Inbox: {i['error']}" if i.get("error")
                 else f"Inbox: {i.get('unread')} unread of {i.get('total')}")
    m = d.get("mailbox")
    if m:
        lines.append("Mailbox: " + " · ".join(f"{k.replace('_', ' ')} {v}" for k, v in m.items()))
    c = d.get("calendar")
    if c:
        lines.append(f"Calendar: times shown in {c['time_zone']} (from {c['time_zone_source']})")
    modes = (["read"] + (["drafts/organise"] if s["write_tools"] else []) + (["send"] if s["send_tools"] else [])
             + (["calendar"] if s.get("calendar") else []))
    line = "Server: " + ", ".join(modes)
    if s["send_tools"]:
        line += f" · send confirmation: {s['send_confirmation']}"
        if s.get("send_allowed_domains"):
            line += " · send only to " + ", ".join(s["send_allowed_domains"])
    if s.get("signature"):
        line += " · a signature is appended to outgoing mail (don't add your own sign-off name)"
    lines.append(line)
    return "\n".join(lines)


async def list_folders_data() -> list[dict]:
    _folders["at"] = 0.0  # always fresh counts
    return await _folder_tree()


def _summary_list(rows: list[dict], preview_chars: int, names: dict | None = None) -> list[dict]:
    out = []
    for m in rows:
        d = fmt.summary(m, preview_chars)
        d.pop("conversation_id", None)
        if names is not None and m.get("parentFolderId") in names:
            d["folder"] = names[m["parentFolderId"]]
        out.append(d)
    return out


def _odata_str(value: str) -> str:
    return "'" + value.replace("'", "''") + "'"


async def list_messages_data(folder: str, limit: int, since: str, until: str, unread_only: bool,
                             sender: str, flagged_only: bool, has_attachments: bool, importance: str,
                             focused: str, preview_chars: int) -> dict:
    limit = _clamp(limit, 1, 200)
    seg, label = await _folder(folder)
    start, end = _time(since, "since"), _time(until, "until")
    # Graph requires the $orderby property to come first in $filter, so always filter on it.
    flt = [f"receivedDateTime ge {fmt.iso_z(start) if start else '1900-01-01T00:00:00Z'}"]
    if end:
        flt.append(f"receivedDateTime lt {fmt.iso_z(end)}")
    if unread_only:
        flt.append("isRead eq false")
    if sender.strip():
        s = sender.strip()
        if not compose.EMAIL_RE.match(s):
            raise ToolError("sender must be an email address here; to match a name, use "
                            "search_messages(sender=...).")
        flt.append(f"from/emailAddress/address eq {_odata_str(s)}")
    if flagged_only:
        flt.append("flag/flagStatus eq 'flagged'")
    if has_attachments:
        flt.append("hasAttachments eq true")
    if importance:
        if importance not in ("high", "normal", "low"):
            raise ToolError("importance must be high, normal or low.")
        flt.append(f"importance eq '{importance}'")
    if focused:
        if focused not in ("focused", "other"):
            raise ToolError("focused must be 'focused' or 'other' (or empty for both).")
        flt.append(f"inferenceClassification eq '{focused}'")
    params = {"$select": fmt.SUMMARY_FIELDS, "$filter": " and ".join(flt),
              "$orderby": "receivedDateTime desc", "$top": min(limit + 1, 100)}
    rows, reason = await graph.scan(f"/me/mailFolders/{seg}/messages", params, limit, max_scan=limit + 1)
    out: dict = {"folder": label, "count": len(rows),
                 "messages": _summary_list(rows, _clamp(preview_chars, 0, 1000))}
    if start:
        out["since"] = fmt.iso_z(start)
    if reason == "limit":
        oldest = rows[-1].get("receivedDateTime") if rows else None
        out["note"] = (f"more messages match: raise limit, or pass until={oldest} for the next page"
                       if oldest else "more messages match: raise limit")
    return out


def _list_header(d: dict, what: str) -> str:
    since = f" since {d['since'][:16]}Z" if d.get("since") else ""
    return f"{d['count']} {what}{since}, newest first (times UTC)"


def list_messages_text(d: dict) -> str:
    show = "to" if d["folder"] in ("Sent Items", "Drafts", "Outbox") else "from"
    body = fmt.render_summaries(d["messages"], show) or "No messages."
    return fmt.with_header(_list_header(d, f"messages in {d['folder']}"), body, d.get("note"))


_KQL_STRIP = re.compile(r'["\\]')


def _kql_value(v: str) -> str:
    v = _KQL_STRIP.sub(" ", v).strip()
    return f'\\"{v}\\"' if re.search(r"\s", v) else v


async def search_messages_data(query: str, sender: str, to: str, subject: str, since: str, until: str,
                               has_attachments: bool, folder: str, to_me: bool, unread_only: bool,
                               limit: int, preview_chars: int) -> dict:
    limit = _clamp(limit, 1, 100)
    start, end = _time(since, "since"), _time(until, "until")
    parts = []
    if (query or "").strip():
        parts.append(query.strip().replace("\\", " ").replace('"', '\\"'))
    for field, value in (("from", sender), ("to", to), ("subject", subject)):
        if (value or "").strip():
            parts.append(f"{field}:{_kql_value(value)}")
    if start:
        parts.append(f"received>={start.strftime('%Y-%m-%d')}")
    if end:
        parts.append(f"received<={end.strftime('%Y-%m-%d')}")
    if has_attachments:
        parts.append("hasAttachment:true")
    if not parts:
        raise ToolError("Give a query, sender, to or subject to search for "
                        "(use list_messages to browse a folder).")
    kql = " AND ".join(parts)
    path = "/me/messages"
    label = "all folders"
    if folder.strip():
        seg, label = await _folder(folder)
        path = f"/me/mailFolders/{seg}/messages"
    me = await _me() if to_me else None

    def keep(m: dict) -> bool:
        t = fmt.utc(m.get("receivedDateTime"))
        if start and t and t < start:
            return False
        if end and t and t >= end:
            return False
        if unread_only and m.get("isRead"):
            return False
        if me:
            rcpt = {((r.get("emailAddress") or {}).get("address") or "").lower()
                    for r in (m.get("toRecipients") or []) + (m.get("ccRecipients") or [])}
            if not rcpt & me["addresses"]:
                return False
        return True

    params = {"$search": f'"{kql}"', "$select": fmt.SUMMARY_FIELDS, "$top": 50}
    try:
        rows, reason = await graph.scan(path, params, limit, keep=keep, max_scan=500)
    except GraphError as e:
        if e.status != 400 or '\\"' not in kql:
            raise
        # Some tenants reject escaped phrase quotes: retry with the words grouped instead.
        kql = re.sub(r'(\w+):\\"([^"\\]*)\\"', r"\1:(\2)", kql).replace('\\"', "")
        params["$search"] = f'"{kql}"'
        rows, reason = await graph.scan(path, params, limit, keep=keep, max_scan=500)
    rows.sort(key=lambda m: m.get("receivedDateTime") or "", reverse=True)
    names = await _folder_names() if not folder.strip() else None
    out: dict = {"query": kql, "scope": label, "count": len(rows),
                 "messages": _summary_list(rows, _clamp(preview_chars, 0, 1000), names)}
    if reason in ("limit", "scan"):
        out["note"] = "more results exist: raise limit or narrow the search (e.g. since, sender)"
    return out


def search_text(d: dict) -> str:
    body = fmt.render_summaries(d["messages"], folders={"show": True}) or "No matches."
    return fmt.with_header(f"{d['count']} results for {d['query']} in {d['scope']}, newest first "
                           f"(times UTC)", body, d.get("note"))


ATT_SELECT = "id,name,contentType,size,isInline"


async def _attachments(mid: str) -> list[dict]:
    rows, _ = await graph.get_all(f"/me/messages/{_q(mid)}/attachments", {"$select": ATT_SELECT}, 100)
    return rows


async def get_message_data(message_id: str, body: str, max_chars: int, offset: int, mark_read: bool) -> dict:
    mid = _id(message_id, "message_id", "list_messages or search_messages")
    if body not in ("full", "new"):
        raise ToolError("body must be 'full' or 'new'.")
    m = await graph.get(f"/me/messages/{_q(mid)}", {"$select": fmt.FULL_FIELDS}, text_body=True)
    atts = await _attachments(mid) if m.get("hasAttachments") or "cid:" in str(m.get("body")) else []
    d = fmt.message(m, _clamp(max_chars, 200, 200000), body, atts, max(0, int(offset)))
    if mark_read and m.get("isRead") is False:
        if not ENABLE_WRITE:
            d["note"] = "not marked read: this server is read-only"
        else:
            await graph.patch(f"/me/messages/{_q(mid)}", {"isRead": True})
            d.pop("unread", None)
            d["note"] = "marked as read"
    return d


def get_message_text(d: dict) -> str:
    text = fmt.render_message(d)
    return text + (f"\n({d['note']})" if d.get("note") else "")


async def get_thread_data(message_id: str, conversation_id: str, limit: int, max_chars: int,
                          include_drafts: bool) -> dict:
    cid = (conversation_id or "").strip()
    if not cid:
        mid = _id(message_id, "message_id", "list_messages or search_messages")
        cid = (await graph.get(f"/me/messages/{_q(mid)}", {"$select": "conversationId"})).get("conversationId")
    else:
        cid = _id(cid, "conversation_id", "get_message")
    if not cid:
        raise ToolError("That message has no conversation id.")
    limit = _clamp(limit, 1, 200)
    # No $orderby: Graph rejects it together with a conversationId filter. Sort here.
    rows, more = await graph.get_all("/me/messages", {
        "$filter": f"conversationId eq {_odata_str(cid)}",
        "$select": fmt.FULL_FIELDS, "$top": 50}, 500, text_body=True)
    if not include_drafts:
        rows = [m for m in rows if not m.get("isDraft")]
    rows.sort(key=lambda m: m.get("receivedDateTime") or m.get("sentDateTime") or "")
    omitted = max(0, len(rows) - limit)
    rows = rows[-limit:]
    msgs = []
    for m in rows:
        atts = await _attachments(m["id"]) if m.get("hasAttachments") else []
        d = fmt.message(m, _clamp(max_chars, 200, 50000), "new", atts)
        d.pop("web_link", None)
        d.pop("internet_message_id", None)
        msgs.append(d)
    out: dict = {"conversation_id": cid, "subject": (rows[-1].get("subject") if rows else None),
                 "count": len(msgs), "messages": msgs}
    if omitted:
        out["note"] = f"{omitted} earlier messages in this conversation omitted: raise limit"
    return out


def thread_text(d: dict) -> str:
    header = (f"Conversation: {d.get('subject')} · {d['count']} messages, oldest first, each showing "
              f"only its new text (times UTC)")
    return fmt.with_header(header, fmt.render_thread(d["messages"]) or "No messages.", d.get("note"))


async def get_attachment_data(message_id: str, attachment_id: str, max_chars: int, offset: int) -> dict:
    mid = _id(message_id, "message_id", "list_messages or search_messages")
    aid = _id(attachment_id, "attachment_id", "get_message")
    base = f"/me/messages/{_q(mid)}/attachments/{_q(aid)}"
    meta = await graph.get(base, {"$select": ATT_SELECT})
    a = fmt.attachment(meta)
    out: dict = {"name": a["name"], "content_type": a.get("content_type"), "size": a.get("size"),
                 "kind": a.get("kind")}
    if a.get("kind") == "item":
        full = await graph.get(base, {"$expand": "microsoft.graph.itemAttachment/item"}, text_body=True)
        item = full.get("item") or {}
        if (item.get("@odata.type") or "").endswith("message") or "body" in item:
            text = fmt.render_message(fmt.message(item, 0))
        else:
            text = json.dumps({k: v for k, v in item.items() if not k.startswith("@")},
                              ensure_ascii=False)[:20000]
        how = "attached item"
    elif a.get("kind") == "reference":
        out["url"] = meta.get("sourceUrl")
        raise ToolError(f"{a['name']} is a link to a cloud file, not a copy in the mail"
                        f"{': ' + meta['sourceUrl'] if meta.get('sourceUrl') else ''}. "
                        "Open it with a OneDrive/SharePoint tool.")
    else:
        data = await graph.get_bytes(f"{base}/$value", int(MAX_ATTACHMENT_MB * 1024 * 1024))
        text, how = await asyncio.to_thread(extract_text, a["name"], a.get("content_type"), data)
    text = text or ""
    total = len(text)
    offset = max(0, int(offset))
    max_chars = _clamp(max_chars, 500, 500000)
    chunk = text[offset:offset + max_chars]
    out.update({"extracted_as": how, "text": chunk, "chars": total})
    if offset:
        out["offset"] = offset
    if offset + max_chars < total:
        out["note"] = (f"showing characters {offset}-{offset + len(chunk)} of {total}: pass "
                       f"offset={offset + max_chars} to continue")
    return out


def attachment_text(d: dict) -> str:
    head = f"{d['name']} ({fmt.human_size(d.get('size'))}, read as {d.get('extracted_as')})"
    return fmt.with_header(head, "\n" + (d.get("text") or "(no text)"), d.get("note"))


async def find_people_data(query: str, limit: int) -> dict:
    q = (query or "").strip()
    if not q or len(q) > 100:
        raise ToolError("Give a name or email address (1-100 characters).")
    limit = _clamp(limit, 1, 25)
    found: list[dict] = []
    missing: list[str] = []
    seen: set[str] = set()

    def add(p: dict) -> None:
        e = (p.get("email") or "").lower()
        if e and e not in seen:
            seen.add(e)
            found.append(p)

    # 1. People relevant to the user (People.Read): ranked by how much you work with them.
    try:
        r = await graph.get("/me/people", {"$search": f'"{_KQL_STRIP.sub(" ", q)}"', "$top": limit,
                                           "$select": "displayName,scoredEmailAddresses,jobTitle,"
                                                      "department,companyName,officeLocation"})
        for p in r.get("value") or []:
            add(fmt.person(p, "people"))
    except GraphError as e:
        if e.status not in (403, 400):
            raise
        missing.append("People.Read")
    # 2. Directory (User.ReadBasic.All).
    if len(found) < limit:
        try:
            term = re.sub(r'["\\:]', " ", q).strip()
            r = await graph.get("/users", {
                "$search": f'"displayName:{term}" OR "mail:{term}"', "$top": limit, "$count": "true",
                "$select": "displayName,mail,userPrincipalName,jobTitle,department,officeLocation"},
                headers={"ConsistencyLevel": "eventual"})
            for p in r.get("value") or []:
                add(fmt.person(p, "directory"))
        except GraphError as e:
            if e.status not in (403, 400):
                raise
            missing.append("User.ReadBasic.All")
    # 3. Your own mail: addresses you have exchanged mail with (needs only Mail.Read).
    if len(found) < limit:
        r = await graph.get("/me/messages", {"$search": f'"participants:{_kql_value(q)}"', "$top": 50,
                                              "$select": "from,toRecipients,ccRecipients,receivedDateTime,"
                                                         "isDraft"})
        ql = q.lower()
        hist: dict[str, dict] = {}
        for m in r.get("value") or []:
            if m.get("isDraft"):
                continue
            for rr in [m.get("from")] + (m.get("toRecipients") or []) + (m.get("ccRecipients") or []):
                a = fmt.address(rr)
                e = (a.get("email") or "").lower()
                if e and (ql in e or ql in (a.get("name") or "").lower()):
                    h = hist.setdefault(e, {"email": a.get("email"), "last_mail": m.get("receivedDateTime")})
                    if not h.get("name") and a.get("name") and a["name"].lower() != e:
                        h["name"] = a["name"]
                    if (m.get("receivedDateTime") or "") > (h.get("last_mail") or ""):
                        h["last_mail"] = m.get("receivedDateTime")
        for h in sorted(hist.values(), key=lambda h: h.get("last_mail") or "", reverse=True):
            add(fmt.person(h, "mail history"))
    out: dict = {"query": q, "count": len(found[:limit]), "people": found[:limit]}
    if missing and not found:
        out["note"] = (f"directory lookup needs {' and '.join(missing)}; only addresses from your "
                       "own mail were searched")
    return out


def people_text(d: dict) -> str:
    return fmt.with_header(f"{d['count']} people matching {d['query']!r}", fmt.render_people(d["people"]),
                           d.get("note"))


# ---------------------------------------------------------------- write data functions
def _recipient_payload(to, cc, bcc) -> dict:
    out = {}
    for key, value, label in (("toRecipients", to, "to"), ("ccRecipients", cc, "cc"),
                              ("bccRecipients", bcc, "bcc")):
        if value is not None:
            out[key] = compose.parse_recipients(value, label)
    return out


def _importance(value: str) -> str:
    if value not in ("low", "normal", "high"):
        raise ToolError("importance must be low, normal or high.")
    return value


async def _add_attachments(mid: str, attachments: list[dict] | None) -> int:
    payloads = compose.attachment_payloads(attachments)
    for p in payloads:
        await graph.post(f"/me/messages/{_q(mid)}/attachments", p)
    return len(payloads)


def _draft_summary(m: dict, verb: str, n_att: int = 0) -> str:
    d = fmt.summary(m, preview_chars=0)
    lines = [f"{verb} (not sent). draft_id {m.get('id')}",
             f"Subject: {d.get('subject')}"]
    for label, key in (("To", "to"), ("Cc", "cc")):
        if d.get(key):
            lines.append(f"{label}: " + "; ".join(fmt.addr_text(a) for a in d[key]))
    if m.get("bccRecipients"):
        lines.append("Bcc: " + "; ".join(fmt.addr_text(fmt.address(r)) for r in m["bccRecipients"]))
    if n_att:
        lines.append(f"Attachments added: {n_att}")
    if m.get("webLink"):
        lines.append(f"Open in Outlook: {m['webLink']}")
    return "\n".join(lines)


DRAFT_SELECT = "id,subject,toRecipients,ccRecipients,bccRecipients,webLink,isDraft,body,hasAttachments"


async def create_draft_data(to, subject, body, cc, bcc, format, importance, attachments) -> str:
    payload: dict = {
        "subject": (subject or "").strip()[:255],
        "body": {"contentType": "HTML",
                 "content": compose.wrap_html(compose.build_body(body, format, SIGNATURE_HTML))},
        "importance": _importance(importance),
        **_recipient_payload(to, cc, bcc),
    }
    compose.attachment_payloads(attachments)  # validate before creating anything
    m = await graph.post("/me/messages", payload)
    n = await _add_attachments(m["id"], attachments)
    return _draft_summary(m, "Draft saved", n)


async def _reply_draft(message_id: str, body: str, reply_all: bool, format: str, cc, attachments) -> dict:
    mid = _id(message_id, "message_id", "list_messages or search_messages")
    compose.attachment_payloads(attachments)
    new_html = compose.build_body(body, format, SIGNATURE_HTML)
    action = "createReplyAll" if reply_all else "createReply"
    draft = await graph.post(f"/me/messages/{_q(mid)}/{action}")
    patch: dict = {"body": {"contentType": "HTML",
                            "content": compose.prepend_to_reply(new_html, (draft.get("body") or {}).get("content", ""))}}
    if cc is not None:
        extra = compose.parse_recipients(cc, "cc")
        have = set(compose.addresses(draft.get("ccRecipients") or []))
        patch["ccRecipients"] = (draft.get("ccRecipients") or []) + [
            r for r in extra if r["emailAddress"]["address"].lower() not in have]
    m = await graph.patch(f"/me/messages/{_q(draft['id'])}", patch)
    m["_attachments"] = await _add_attachments(draft["id"], attachments)
    return m


async def _forward_draft(message_id: str, to, body: str, cc, format: str, attachments) -> dict:
    mid = _id(message_id, "message_id", "list_messages or search_messages")
    rcpt = _recipient_payload(to, cc, None)
    if not rcpt.get("toRecipients"):
        raise ToolError("Give at least one 'to' address to forward to.")
    compose.attachment_payloads(attachments)
    new_html = compose.build_body(body, format, SIGNATURE_HTML, allow_empty=True)
    draft = await graph.post(f"/me/messages/{_q(mid)}/createForward")
    patch = {"body": {"contentType": "HTML",
                      "content": compose.prepend_to_reply(new_html, (draft.get("body") or {}).get("content", ""))},
             **rcpt}
    m = await graph.patch(f"/me/messages/{_q(draft['id'])}", patch)
    m["_attachments"] = await _add_attachments(draft["id"], attachments)
    return m


async def create_reply_draft_data(message_id, body, reply_all, format, cc, attachments) -> str:
    m = await _reply_draft(message_id, body, reply_all, format, cc, attachments)
    return _draft_summary(m, "Reply draft saved", m.get("_attachments", 0))


async def create_forward_draft_data(message_id, to, body, cc, format, attachments) -> str:
    m = await _forward_draft(message_id, to, body, cc, format, attachments)
    return _draft_summary(m, "Forward draft saved", m.get("_attachments", 0))


async def _get_draft(draft_id: str) -> dict:
    did = _id(draft_id, "draft_id", "create_draft or list_messages(folder='drafts')")
    m = await graph.get(f"/me/messages/{_q(did)}", {"$select": DRAFT_SELECT})
    if not m.get("isDraft"):
        raise ToolError("That message is not a draft (it was already sent or received).")
    return m


async def update_draft_data(draft_id, subject, body, to, cc, bcc, format, importance, attachments) -> str:
    m = await _get_draft(draft_id)
    patch: dict = _recipient_payload(to, cc, bcc)
    if subject is not None:
        patch["subject"] = subject.strip()[:255]
    if body is not None:
        patch["body"] = {"contentType": "HTML",
                         "content": compose.wrap_html(compose.build_body(body, format, SIGNATURE_HTML))}
    if importance is not None:
        patch["importance"] = _importance(importance)
    compose.attachment_payloads(attachments)
    if not patch and not attachments:
        raise ToolError("Nothing to change: give subject, body, to, cc, bcc, importance or attachments.")
    if patch:
        m = await graph.patch(f"/me/messages/{_q(m['id'])}", patch)
    n = await _add_attachments(m["id"], attachments)
    return _draft_summary(m, "Draft updated", n)


async def delete_draft_data(draft_id: str) -> str:
    m = await _get_draft(draft_id)
    await graph.delete(f"/me/messages/{_q(m['id'])}")
    return f"Draft deleted: {m.get('subject') or '(no subject)'}"


async def _batch_patch(ids: list[str], body: dict, what: str) -> str:
    res = await graph.batch([{"method": "PATCH", "url": f"/me/messages/{_q(i)}", "body": body} for i in ids])
    failed = [(i, r) for i, r in zip(ids, res) if r["status"] >= 400]
    ok = len(ids) - len(failed)
    text = f"{what}: {ok} of {len(ids)} messages"
    if failed:
        text += "\nFailed: " + "; ".join(
            f"{i[:24]}… ({r['status']} {((r['body'] or {}).get('error') or {}).get('code', '')})"
            for i, r in failed)
    return text


async def mark_read_data(message_ids, read: bool) -> str:
    return await _batch_patch(_ids(message_ids), {"isRead": bool(read)},
                              "Marked read" if read else "Marked unread")


async def flag_messages_data(message_ids, status: str, due: str) -> str:
    flag: dict
    if status == "flagged":
        flag = {"flagStatus": "flagged"}
        due_dt = _time(due, "due")
        if due_dt:
            now = datetime.now(timezone.utc)
            flag["startDateTime"] = {"dateTime": fmt.iso_z(min(now, due_dt))[:-1], "timeZone": "UTC"}
            flag["dueDateTime"] = {"dateTime": fmt.iso_z(due_dt)[:-1], "timeZone": "UTC"}
    elif status == "complete":
        flag = {"flagStatus": "complete"}
    elif status == "clear":
        flag = {"flagStatus": "notFlagged"}
    else:
        raise ToolError("status must be flagged, complete or clear.")
    return await _batch_patch(_ids(message_ids), {"flag": flag}, f"Flag set to {status}")


async def set_categories_data(message_ids, categories, mode: str) -> str:
    ids = _ids(message_ids)
    cats = categories if isinstance(categories, list) else [c for c in re.split(r"[;,]", categories or "")]
    cats = [c.strip() for c in cats if c and c.strip()]
    if mode not in ("add", "remove", "replace"):
        raise ToolError("mode must be add, remove or replace.")
    if mode == "replace":
        return await _batch_patch(ids, {"categories": cats}, "Categories set")
    if not cats:
        raise ToolError("Give at least one category.")
    current = await graph.batch([{"method": "GET", "url": f"/me/messages/{_q(i)}?$select=categories"}
                                 for i in ids])
    reqs = []
    for i, r in zip(ids, current):
        have = (r["body"] or {}).get("categories") or []
        new = (have + [c for c in cats if c not in have]) if mode == "add" else [c for c in have if c not in cats]
        reqs.append({"method": "PATCH", "url": f"/me/messages/{_q(i)}", "body": {"categories": new}})
    res = await graph.batch(reqs)
    bad = sum(1 for r in res if r["status"] >= 400)
    return f"Categories {'added' if mode == 'add' else 'removed'}: {len(ids) - bad} of {len(ids)} messages"


async def move_messages_data(message_ids, folder: str) -> str:
    ids = _ids(message_ids)
    seg, label = await _folder(folder)
    dest = seg if seg in fmt.WELL_KNOWN else (await graph.get(f"/me/mailFolders/{seg}", {"$select": "id"}))["id"]
    res = await graph.batch([{"method": "POST", "url": f"/me/messages/{_q(i)}/move",
                              "body": {"destinationId": dest}} for i in ids])
    failed = [(i, r) for i, r in zip(ids, res) if r["status"] >= 400]
    _folders["at"] = 0.0
    text = f"Moved {len(ids) - len(failed)} of {len(ids)} messages to {label}"
    if failed:
        text += "\nFailed: " + "; ".join(f"{i[:24]}… ({r['status']})" for i, r in failed)
    return text


async def create_folder_data(name: str, parent: str) -> str:
    name = (name or "").strip()
    if not name or len(name) > 255 or "/" in name:
        raise ToolError("Give a folder name (no '/'); put it under another folder with parent.")
    if parent.strip():
        seg, plabel = await _folder(parent)
        path = f"/me/mailFolders/{seg}/childFolders"
    else:
        path, plabel = "/me/mailFolders", "the top level"
    f = await graph.post(path, {"displayName": name})
    _folders["at"] = 0.0
    return f"Folder {name!r} created under {plabel} (folder_id {f.get('id')})"


# ---------------------------------------------------------------- send
# With SEND_CONFIRMATION=elicit the server itself asks the human before each send:
# - 2026-07-28 clients: the tool returns an InputRequiredResult; the client shows
#   the prompt and calls the tool again with the answer and our sealed request_state;
# - older clients: a server-initiated elicitation request (ctx.elicit).
# The approval is bound to a fingerprint of recipients, subject, body and
# attachments: if the message changes between asking and sending, we ask again.
APPROVE_SCHEMA = {"type": "object", "required": ["value"],
                  "properties": {"value": {"type": "string", "title": "Decision", "enum": ["send", "cancel"]}}}


class Ask(Exception):
    """Carries an InputRequiredResult out of the data functions to the tool."""

    def __init__(self, result: InputRequiredResult):
        super().__init__("input required")
        self.result = result


def _fingerprint(m: dict) -> str:
    key = json.dumps([compose.addresses(m.get("toRecipients") or []),
                      compose.addresses(m.get("ccRecipients") or []),
                      compose.addresses(m.get("bccRecipients") or []),
                      m.get("subject"), (m.get("body") or {}).get("content"),
                      bool(m.get("hasAttachments"))], sort_keys=True)
    return hashlib.sha256(key.encode()).hexdigest()[:24]


def _prompt(m: dict, verb: str) -> str:
    d = fmt.summary(m, 0)
    body = fmt.body_text(m.get("body"))
    return "\n".join(x for x in (
        f"{verb} as you?",
        "To: " + "; ".join(fmt.addr_text(a) for a in d.get("to") or []),
        ("Cc: " + "; ".join(fmt.addr_text(a) for a in d["cc"])) if d.get("cc") else None,
        ("Bcc: " + "; ".join(fmt.addr_text(fmt.address(r)) for r in m["bccRecipients"]))
        if m.get("bccRecipients") else None,
        f"Subject: {d.get('subject')}",
        "Attachments: yes" if m.get("hasAttachments") else None,
        "",
        fmt.shorten(body, 1500) if len(body) > 1500 else body,
    ) if x is not None)


def _state_ref(ctx: Context | None) -> str | None:
    """The draft id we stored in request_state on an earlier leg, if any."""
    state = getattr(ctx, "request_state", None) if ctx is not None else None
    return state.split("|", 1)[0] if state and "|" in state else None


def _modern(ctx: Context) -> bool:
    check = getattr(ctx, "_is_modern_protocol", None)
    return bool(check and check())


async def _approve(ctx: Context | None, prompt: str, state: str, *, declined: str, no_ui: str,
                   no_ctx: str) -> None:
    """With SEND_CONFIRMATION=elicit, ask the human to approve `prompt`. `state`
    identifies exactly what is being approved (id|fingerprint). Raises Ask when a
    2026-07-28 client must show the prompt first; fails closed otherwise."""
    if SEND_CONFIRMATION != "elicit":
        return
    if ctx is None:
        raise ToolError(no_ctx)
    if _modern(ctx):
        answer = (ctx.input_responses or {}).get("approve")
        if answer is None or ctx.request_state != state:
            # First leg, or the content changed since the user was asked: ask (again).
            raise Ask(InputRequiredResult(
                input_requests={"approve": ElicitRequest(params=ElicitRequestFormParams(
                    message=prompt, requested_schema=APPROVE_SCHEMA))},
                request_state=state))
        action = getattr(answer, "action", None)
        value = (getattr(answer, "content", None) or {}).get("value")
    else:
        try:
            res = await ctx.elicit(prompt, ["send", "cancel"])
        except Exception as e:
            log.warning("Approval prompt could not be shown: %s", e)
            raise ToolError(no_ui) from None
        action, value = getattr(res, "action", None), getattr(res, "data", None)
    if action != "accept" or value != "send":
        raise ToolError(declined)


async def _confirm_send(ctx: Context | None, m: dict, verb: str) -> None:
    """Enforce SEND_ALLOWED_DOMAINS and, with SEND_CONFIRMATION=elicit, get the
    human's approval. Raises Ask when the client must ask the user first."""
    rcpt = (m.get("toRecipients") or []) + (m.get("ccRecipients") or []) + (m.get("bccRecipients") or [])
    if not rcpt:
        raise ToolError("The message has no recipients.")
    compose.check_allowed(rcpt, ALLOWED_DOMAINS)
    kept = f"; the draft {m['id']} is kept" if m.get("isDraft") else ""
    await _approve(
        ctx, _prompt(m, verb), f"{m.get('id') or 'new'}|{_fingerprint(m)}",
        declined="The user did not approve sending. Nothing was sent" + kept + ".",
        no_ui="This server requires the user to approve each send, and this client can't show the "
              "approval prompt. The message was NOT sent"
              + (f"; it is saved as draft {m['id']}" if m.get("isDraft") else "")
              + ". The user can send it from Outlook.",
        no_ctx="Send approval is required but there is no client session. Nothing was sent" + kept)


def _sent_line(m: dict, verb: str) -> str:
    d = fmt.summary(m, 0)
    to = "; ".join(fmt.addr_text(a) for a in (d.get("to") or []) + (d.get("cc") or []))
    return (f"{verb} at {fmt.short_time(datetime.now(timezone.utc).isoformat())} UTC: "
            f"{d.get('subject')!r} to {to}")


async def _send_existing_draft(ctx, m: dict, verb: str) -> str:
    await _confirm_send(ctx, m, verb)
    await graph.post(f"/me/messages/{_q(m['id'])}/send")
    return _sent_line(m, "Sent")


async def send_draft_data(ctx, draft_id: str) -> str:
    m = await _get_draft(draft_id)
    try:
        return await _send_existing_draft(ctx, m, "Send this email")
    except ComposeError as e:
        raise ToolError(f"{e} The draft was not sent.") from None


async def send_mail_data(ctx, to, subject, body, cc, bcc, format, importance, attachments) -> str:
    rcpt = _recipient_payload(to, cc, bcc)
    if not rcpt.get("toRecipients"):
        raise ToolError("Give at least one 'to' address.")
    msg: dict = {
        "subject": (subject or "").strip()[:255],
        "body": {"contentType": "HTML",
                 "content": compose.wrap_html(compose.build_body(body, format, SIGNATURE_HTML))},
        "importance": _importance(importance), **rcpt,
    }
    atts = compose.attachment_payloads(attachments)
    if atts:
        msg["attachments"] = atts
    preview = {**msg, "hasAttachments": bool(atts)}
    await _confirm_send(ctx, preview, "Send this email")
    await graph.post("/me/sendMail", {"message": msg, "saveToSentItems": True})
    return _sent_line(msg, "Sent")


async def _send_built_draft(ctx, draft_id: str, verb: str) -> str:
    m = await _get_draft(draft_id)
    try:
        return await _send_existing_draft(ctx, m, verb)
    except ComposeError as e:
        raise ToolError(f"{e} It was saved as draft {m['id']} and not sent.") from None
    except (ToolError, Ask):
        raise
    except (GraphError, AuthError) as e:
        raise ToolError(f"{e}. It was saved as draft {m['id']} and not sent.") from None


async def reply_data(ctx, message_id, body, reply_all, format, cc, attachments) -> str:
    # On the approval leg the client calls again: reuse the draft made on the first leg.
    draft_id = _state_ref(ctx) or (await _reply_draft(message_id, body, reply_all, format, cc, attachments))["id"]
    return await _send_built_draft(ctx, draft_id, "Send this reply")


async def forward_data(ctx, message_id, to, body, cc, format, attachments) -> str:
    draft_id = _state_ref(ctx) or (await _forward_draft(message_id, to, body, cc, format, attachments))["id"]
    return await _send_built_draft(ctx, draft_id, "Forward this email")


async def _call_send(coro) -> str | InputRequiredResult:
    try:
        return await _call(coro)
    except Ask as a:
        return a.result


# ---------------------------------------------------------------- read tools
@mcp.tool(**READ)
async def whoami(output: str = "text") -> str:
    """The mailbox this server acts as, the state of its sign-in, inbox counts,
    and what the server is allowed to do (drafts, sending, signature)."""
    return await _run(whoami_data(), whoami_text, output)


@mcp.tool(**READ)
async def list_folders(output: str = "text") -> str:
    """All mail folders with unread/total counts and their ids, nested."""
    return await _run(list_folders_data(), fmt.render_folders, output)


@mcp.tool(**READ)
async def list_messages(folder: str = "inbox", limit: int = 25, since: str = "", until: str = "",
                        unread_only: bool = False, sender: str = "", flagged_only: bool = False,
                        has_attachments: bool = False, importance: str = "", focused: str = "",
                        preview_chars: int = 160, output: str = "text") -> str:
    """Messages in one folder, newest first (limit 1-200): time, sender (recipients
    in Sent/Drafts), subject, UNREAD/!high/⚑/📎 marks, a preview and the message id.
    folder: inbox (default), drafts, sent, deleted, junk, archive, a folder name or
    path like 'Projects/Alpha', or an id from list_folders.
    since/until: ISO 8601 date or date/time (until is exclusive; use it to page back).
    sender: an exact email address. focused: 'focused' or 'other' (Focused Inbox).
    preview_chars=0 drops previews."""
    return await _run(list_messages_data(folder, limit, since, until, unread_only, sender, flagged_only,
                                         has_attachments, importance, focused, preview_chars),
                      list_messages_text, output)


@mcp.tool(**READ)
async def search_messages(query: str = "", sender: str = "", to: str = "", subject: str = "",
                          since: str = "", until: str = "", has_attachments: bool = False,
                          folder: str = "", to_me: bool = False, unread_only: bool = False,
                          limit: int = 25, preview_chars: int = 160, output: str = "text") -> str:
    """Search the whole mailbox (or one folder), newest first (limit 1-100).
    query: words to find in subject, body or attachments (KQL allowed, e.g.
    'invoice OR receipt', '"exact phrase"'). sender / to: a name or address.
    subject: words in the subject only. since/until: ISO 8601 date or date/time.
    to_me: only mail where the signed-in user is in To or Cc. Combine freely, e.g.
    sender='alice' query='budget' since='2026-11-01'."""
    return await _run(search_messages_data(query, sender, to, subject, since, until, has_attachments,
                                           folder, to_me, unread_only, limit, preview_chars),
                      search_text, output)


@mcp.tool(**READ)
async def get_message(message_id: str, body: str = "full", max_chars: int = MAX_BODY_CHARS, offset: int = 0,
                      mark_read: bool = False, output: str = "text") -> str:
    """One message in full: headers, recipients, attachments (with attachment ids),
    conversation id, Outlook link and the body as text. body='new' shows only the
    text this message added (quoted history removed). Long bodies are cut at
    max_chars; pass offset to continue. mark_read=True also marks it as read."""
    return await _run(get_message_data(message_id, body, max_chars, offset, mark_read),
                      get_message_text, output)


@mcp.tool(**READ)
async def get_thread(message_id: str = "", conversation_id: str = "", limit: int = 30,
                     max_chars: int = 4000, include_drafts: bool = False, output: str = "text") -> str:
    """A whole conversation (all folders, including your own sent replies), oldest
    first, each message showing only its new text without the quoted history.
    Give any message_id of the conversation, or its conversation_id. max_chars
    caps each message's text."""
    return await _run(get_thread_data(message_id, conversation_id, limit, max_chars, include_drafts),
                      thread_text, output)


@mcp.tool(**READ)
async def get_attachment(message_id: str, attachment_id: str, max_chars: int = 30000, offset: int = 0,
                         output: str = "text") -> str:
    """The text of an attachment (ids from get_message): PDF, Word (.docx), Excel
    (.xlsx), PowerPoint (.pptx), text/CSV/JSON/HTML, and attached emails. Long
    text is paged with offset/max_chars. Images and other binaries can't be read."""
    return await _run(get_attachment_data(message_id, attachment_id, max_chars, offset), attachment_text,
                      output)


@mcp.tool(**READ)
async def find_people(query: str, limit: int = 10, output: str = "text") -> str:
    """Find a person's email address by name (or part of an address): people you
    work with, the directory, and addresses from your own mail. Use the address
    with the compose tools, which don't accept plain names."""
    return await _run(find_people_data(query, limit), people_text, output)


# ---------------------------------------------------------------- write tools
async def create_draft(to: Recipients, subject: str, body: str, cc: Recipients = None,
                       bcc: Recipients = None, format: str = "text", importance: str = "normal",
                       attachments: Attachments = None) -> str:
    """Save a new email as a draft in Drafts, without sending. format: 'text'
    (as typed, line breaks kept), 'markdown' (bold, lists, links, tables) or
    'html'. Returns the draft_id and an Outlook link; send it later with
    send_draft once the user approves."""
    return await _call(create_draft_data(to, subject, body, cc, bcc, format, importance, attachments))


async def create_reply_draft(message_id: str, body: str, reply_all: bool = False, format: str = "text",
                             cc: Recipients = None, attachments: Attachments = None) -> str:
    """Save a reply to a message as a draft, without sending. The body goes above
    the quoted original; recipients and 'RE:' subject are filled in by Outlook
    (reply_all=True keeps everyone). cc adds people. Returns the draft_id."""
    return await _call(create_reply_draft_data(message_id, body, reply_all, format, cc, attachments))


async def create_forward_draft(message_id: str, to: Recipients, body: str = "", cc: Recipients = None,
                               format: str = "text", attachments: Attachments = None) -> str:
    """Save a forward of a message (with its attachments) as a draft, without
    sending. body is an optional note above the forwarded message."""
    return await _call(create_forward_draft_data(message_id, to, body, cc, format, attachments))


async def update_draft(draft_id: str, subject: str | None = None, body: str | None = None,
                       to: Recipients = None, cc: Recipients = None, bcc: Recipients = None,
                       format: str = "text", importance: str | None = None,
                       attachments: Attachments = None) -> str:
    """Change a draft. Only the fields given change; to/cc/bcc replace the whole
    list. body replaces the entire body, including any quoted original of a reply
    draft: to reword a reply, delete it and create a new reply draft instead.
    attachments are added."""
    return await _call(update_draft_data(draft_id, subject, body, to, cc, bcc, format, importance,
                                         attachments))


async def delete_draft(draft_id: str) -> str:
    """Permanently delete a draft (only drafts can be deleted this way)."""
    return await _call(delete_draft_data(draft_id))


async def mark_read(message_ids: list[str] | str, read: bool = True) -> str:
    """Mark messages as read (or unread with read=False). Up to 100 ids."""
    return await _call(mark_read_data(message_ids, read))


async def flag_messages(message_ids: list[str] | str, status: str = "flagged", due: str = "") -> str:
    """Follow-up flag: status 'flagged' (optional due date), 'complete' or 'clear'."""
    return await _call(flag_messages_data(message_ids, status, due))


async def set_categories(message_ids: list[str] | str, categories: list[str] | str,
                         mode: str = "add") -> str:
    """Add, remove or replace Outlook categories (colour labels) on messages,
    by category name, e.g. 'Red category' or a custom one."""
    return await _call(set_categories_data(message_ids, categories, mode))


async def move_messages(message_ids: list[str] | str, folder: str) -> str:
    """Move messages to a folder: archive, deleted (Deleted Items, recoverable),
    junk, inbox, a folder name or path, or a folder id. Ids stay valid after moving."""
    return await _call(move_messages_data(message_ids, folder))


async def create_folder(name: str, parent: str = "") -> str:
    """Create a mail folder at the top level, or under parent (name, path or id)."""
    return await _call(create_folder_data(name, parent))


# ---------------------------------------------------------------- send tools
async def send_draft(draft_id: str, ctx: Context) -> str | InputRequiredResult:
    """Send an existing draft as the signed-in user. Only after the user has seen
    the recipients, subject and text and approved sending it. Can't be recalled."""
    return await _call_send(send_draft_data(ctx, draft_id))


async def send_mail(to: Recipients, subject: str, body: str, ctx: Context, cc: Recipients = None,
                    bcc: Recipients = None, format: str = "text", importance: str = "normal",
                    attachments: Attachments = None) -> str | InputRequiredResult:
    """Send a new email immediately as the signed-in user (a copy goes to Sent
    Items). Only after the user approved these exact recipients, subject and text.
    Prefer create_draft + send_draft when the user should review it first."""
    return await _call_send(send_mail_data(ctx, to, subject, body, cc, bcc, format, importance, attachments))


async def reply(message_id: str, body: str, ctx: Context, reply_all: bool = False,
                format: str = "text", cc: Recipients = None, attachments: Attachments = None) -> str | InputRequiredResult:
    """Reply to a message and send it immediately (quoted original kept below).
    Only after the user approved the text and, for reply_all, the recipients."""
    return await _call_send(reply_data(ctx, message_id, body, reply_all, format, cc, attachments))


async def forward(message_id: str, to: Recipients, ctx: Context, body: str = "", cc: Recipients = None,
                  format: str = "text", attachments: Attachments = None) -> str | InputRequiredResult:
    """Forward a message (with its attachments) and send it immediately. Only
    after the user approved the recipients and note."""
    return await _call_send(forward_data(ctx, message_id, to, body, cc, format, attachments))


WRITE_TOOLS = (create_draft, create_reply_draft, create_forward_draft, update_draft, delete_draft,
               mark_read, flag_messages, set_categories, move_messages, create_folder)
SEND_TOOLS = (send_draft, send_mail, reply, forward)

if ENABLE_WRITE:
    for _fn in WRITE_TOOLS:
        mcp.tool(annotations=WRITES, output_schema=None)(_fn)
if ENABLE_SEND:
    for _fn in SEND_TOOLS:
        mcp.tool(annotations=SENDS, output_schema=None)(_fn)


# ---------------------------------------------------------------- calendar
_tz_cache: dict = {}


def _now() -> datetime:
    return datetime.now(timezone.utc)


async def _tz():
    """The time zone calendar times are shown and entered in: TIMEZONE, else the
    mailbox's setting (needs MailboxSettings.Read), else UTC."""
    if TIMEZONE:
        return cal.zone(TIMEZONE)
    uid = tokens.user.get("id")
    if _tz_cache.get("uid") != uid or not _tz_cache.get("tz"):
        tz = None
        try:
            s = await graph.get("/me/mailboxSettings", {"$select": "timeZone"})
            tz = cal.zone(s.get("timeZone"))
        except GraphError:
            pass
        _tz_cache.update(uid=uid, tz=tz or cal.zone("UTC"))
    return _tz_cache["tz"]


def _when(value: str, label: str, tz, now: datetime) -> datetime:
    try:
        return cal.parse_local(value, tz, now)
    except ValueError:
        raise ToolError(f"{label} must be a date or date/time like 2026-11-18, 2026-11-18T09:30 (in "
                        f"{tz.key}), 2026-11-18T05:30:00Z, 'today', 'tomorrow' or 'now'.") from None


async def _calendars() -> list[dict]:
    rows, _ = await graph.get_all("/me/calendars", {"$select": "id,name,isDefaultCalendar,canEdit,owner,color",
                                                     "$top": "100"}, limit=200)
    return rows


async def _calendar_path(calendar: str) -> str:
    """'' = the default calendar; otherwise a calendar name or id."""
    calendar = (calendar or "").strip()
    if not calendar:
        return "/me/calendar"
    rows = await _calendars()
    for c in rows:
        if c.get("id") == calendar or (c.get("name") or "").lower() == calendar.lower():
            return f"/me/calendars/{_q(c['id'])}"
    raise ToolError(f"No calendar named {calendar!r}. Calendars: "
                    + ", ".join(c.get("name") or "?" for c in rows))


def _event_id(value: str) -> str:
    return _id(value, "event_id", "list_events")


async def _get_event(event_id: str, full: bool = True) -> dict:
    return await graph.get(f"/me/events/{_q(_event_id(event_id))}",
                           {"$select": cal.EVENT_FULL_FIELDS if full else cal.EVENT_FIELDS}, text_body=full)


async def list_calendars_data() -> list[dict]:
    return [{k: v for k, v in {
        "name": c.get("name"), "calendar_id": c.get("id"), "default": c.get("isDefaultCalendar") or None,
        "can_edit": c.get("canEdit"), "owner": fmt.addr_text({"name": (c.get("owner") or {}).get("name"),
                                                              "email": (c.get("owner") or {}).get("address")}),
    }.items() if v is not None} for c in await _calendars()]


def calendars_text(rows: list[dict]) -> str:
    return "\n".join(f"{c['name']}" + (" (default)" if c.get("default") else "")
                     + ("" if c.get("can_edit") else " · read-only") + f" · owner {c.get('owner')}"
                     + f"\n{fmt._PAD}id {c['calendar_id']}" for c in rows) or "No calendars."


async def list_events_data(start: str, end: str, days: int, calendar: str, query: str,
                           needs_response: bool, limit: int) -> dict:
    tz, now = await _tz(), _now()
    s = _when(start, "start", tz, now) if start.strip() else cal.parse_local("today", tz, now)
    e = _when(end, "end", tz, now) if end.strip() else s + timedelta(days=_clamp(days, 1, 92))
    if cal.is_date(end) and end.strip():
        e += timedelta(days=1)  # an end date includes that whole day
    if e <= s:
        raise ToolError("end must be after start.")
    if e - s > timedelta(days=366):
        raise ToolError("At most one year per call.")
    limit = _clamp(limit, 1, 200)
    q = (query or "").strip().lower()

    def keep(ev: dict) -> bool:
        d = cal.event(ev, tz)
        if needs_response and not cal.needs_response(d):
            return False
        if q:
            hay = json.dumps([ev.get("subject"), (ev.get("location") or {}).get("displayName"),
                              ev.get("organizer"), ev.get("attendees")], ensure_ascii=False).lower()
            return all(w in hay for w in q.split())
        return True

    rows, why = await graph.scan(
        f"{await _calendar_path(calendar)}/calendarView",
        {"startDateTime": fmt.iso_z(s), "endDateTime": fmt.iso_z(e), "$select": cal.EVENT_FIELDS,
         "$orderby": "start/dateTime", "$top": "100"},
        limit=limit, keep=keep, max_scan=2000)
    return {"time_zone": tz.key, "start": s.isoformat(), "end": e.isoformat(),
            "events": [cal.event(ev, tz) for ev in rows], "more": why in ("limit", "scan"),
            "filter": " · ".join(x for x in (f"matching {query!r}" if q else "",
                                             "needing your response" if needs_response else "") if x)}


def events_text(d: dict) -> str:
    s, e = datetime.fromisoformat(d["start"]), datetime.fromisoformat(d["end"])
    last = e - timedelta(minutes=1)
    span = (cal.day_label(s.date()) if s.date() == last.date()
            else f"{cal.day_label(s.date())} to {cal.day_label(last.date())}")
    n = len(d["events"])
    head = (f"{n} event{'s' if n != 1 else ''}" + (f" {d['filter']}" if d.get("filter") else "")
            + f", {span} (times {d['time_zone']})")
    note = "more events match: narrow the range or raise limit" if d.get("more") else None
    return fmt.with_header(head, cal.render_events(d["events"]) or "Nothing scheduled.", note)


async def get_event_data(event_id: str, max_chars: int) -> dict:
    tz = await _tz()
    return {"time_zone": tz.key, **cal.event(await _get_event(event_id), tz,
                                             body_chars=_clamp(max_chars, 200, 100000))}


def event_text(d: dict) -> str:
    return cal.render_event(d, d["time_zone"])


async def find_free_time_data(start: str, end: str, duration_minutes: int, attendees,
                              working_hours_only: bool, tentative_is_busy: bool, max_slots: int) -> dict:
    tz, now = await _tz(), _now()
    s = _when(start, "start", tz, now) if start.strip() else now.astimezone(tz)
    e = _when(end, "end", tz, now) if end.strip() else s + timedelta(days=7)
    if cal.is_date(end) and end.strip():
        e += timedelta(days=1)
    s = max(s, now.astimezone(tz))
    if e <= s:
        raise ToolError("The range is in the past or empty: end must be after start (and after now).")
    if e - s > timedelta(days=42):
        raise ToolError("Search at most 6 weeks at a time.")
    minutes = _clamp(duration_minutes, 5, 1440)
    others = compose.addresses(compose.parse_recipients(attendees, "attendees")) if attendees else []
    me = (await _me()).get("mail")
    people = [me] + [a for a in dict.fromkeys(others) if a.lower() not in (await _me())["addresses"]]
    if len(people) > 20:
        raise ToolError("At most 19 attendees per call.")
    res = await graph.post("/me/calendar/getSchedule", {
        "schedules": people, "availabilityViewInterval": 15,
        "startTime": {"dateTime": s.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
        "endTime": {"dateTime": e.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S"), "timeZone": "UTC"},
    })
    busy_status = {"busy", "oof"} | ({"tentative"} if tentative_is_busy else set())
    busy, unavailable, mine = [], [], None
    for sched in res.get("value") or []:
        who = sched.get("scheduleId") or "?"
        if sched.get("error"):
            unavailable.append({"email": who, "reason": (sched["error"].get("message") or "unknown")[:120]})
            continue
        if who.lower() == (me or "").lower():
            mine = sched.get("workingHours")
        for item in sched.get("scheduleItems") or []:
            if item.get("status") in busy_status:
                a, b = cal.from_graph(item.get("start"), tz), cal.from_graph(item.get("end"), tz)
                if a and b:
                    busy.append((a, b))
    work = None
    if working_hours_only:
        wz = cal.zone(((mine or {}).get("timeZone") or {}).get("name")) or tz
        days = WORKING_DAYS or [cal.DAYS.index(x) for x in (mine or {}).get("daysOfWeek") or []
                                if x in cal.DAYS] or [0, 1, 2, 3, 4]
        if WORKING_HOURS:
            (h0, h1), wz = WORKING_HOURS, tz
        elif mine and mine.get("startTime") and mine.get("endTime"):
            h0, h1 = cal.parse_hhmm(mine["startTime"]), cal.parse_hhmm(mine["endTime"])
        else:
            h0, h1 = cal.parse_hhmm("08:00"), cal.parse_hhmm("17:00")
        work = {"days": days, "start": h0, "end": h1, "tz": wz}
    slots = cal.free_slots(s, e, busy, minutes, tz, work)
    max_slots = _clamp(max_slots, 1, 100)
    return {
        "time_zone": tz.key, "minutes": minutes, "participants": people,
        "working_hours": (f"{cal.days_text(work['days'])} {work['start']:%H:%M}–{work['end']:%H:%M}"
                          + (f" {work['tz'].key}" if work["tz"].key != tz.key else "")) if work else None,
        "slots": [{"start": a.isoformat(), "end": b.isoformat(), "minutes": int((b - a).total_seconds() // 60)}
                  for a, b in slots[:max_slots]],
        "more": len(slots) > max_slots, "unavailable": unavailable,
    }


# --- writing: anything that notifies other people counts as sending
def _notify_allowed(what: str) -> None:
    if not ENABLE_SEND:
        raise ToolError(f"This server is configured not to send anything (ENABLE_SEND_TOOLS=false), so it "
                        f"can't {what}.")


def _event_prompt(head: str, ev: dict, tz) -> str:
    d = cal.event(ev, tz, body_chars=1500)
    lines = [head, f"Subject: {d['subject']}", f"When: {cal.when_text(d)} ({tz.key})"]
    if ev.get("recurrence"):
        lines.append(f"Repeats: {cal.recurrence_text(ev['recurrence'])}")
    if d.get("location"):
        lines.append(f"Where: {d['location']}")
    if d.get("online_meeting"):
        lines.append("Teams meeting: yes")
    for kind in ("required", "optional"):
        who = [fmt.addr_text(a) for a in d.get("attendees") or [] if a["type"] == kind]
        if who:
            lines.append(f"{kind.capitalize()}: " + "; ".join(who))
    if d.get("body"):
        lines += ["", d["body"]]
    return "\n".join(lines)


def _digest(*parts) -> str:
    return hashlib.sha256(json.dumps(parts, sort_keys=True, default=str).encode()).hexdigest()[:24]


async def _approve_event(ctx, prompt: str, state: str, not_done: str) -> None:
    await _approve(ctx, prompt, state,
                   declined=f"The user did not approve. {not_done}",
                   no_ui="This server requires the user to approve anything sent to other people, and this "
                         f"client can't show the approval prompt. {not_done} The user can do it in Outlook.",
                   no_ctx=f"Approval is required but there is no client session. {not_done}")


def _attendees(required, optional) -> list[dict]:
    out = []
    for value, kind in ((required, "required"), (optional, "optional")):
        if value:
            for r in compose.parse_recipients(value, f"{kind} attendees"):
                out.append({"emailAddress": r["emailAddress"], "type": kind})
    return out


def _recurrence(repeat: str, start: datetime, until: str, count: int, tz, now) -> dict | None:
    repeat = (repeat or "").strip().lower()
    if not repeat or repeat == "none":
        return None
    day = cal.DAYS[start.weekday()]
    patterns = {
        "daily": {"type": "daily", "interval": 1},
        "weekdays": {"type": "weekly", "interval": 1, "daysOfWeek": cal.DAYS[:5], "firstDayOfWeek": "monday"},
        "weekly": {"type": "weekly", "interval": 1, "daysOfWeek": [day]},
        "biweekly": {"type": "weekly", "interval": 2, "daysOfWeek": [day]},
        "monthly": {"type": "absoluteMonthly", "interval": 1, "dayOfMonth": start.day},
    }
    if repeat not in patterns:
        raise ToolError("repeat must be one of: daily, weekdays, weekly, biweekly, monthly.")
    rng: dict = {"startDate": start.date().isoformat(), "recurrenceTimeZone": cal.graph_tz_name(tz)}
    if until.strip():
        rng.update(type="endDate", endDate=_when(until, "repeat_until", tz, now).date().isoformat())
    elif count:
        rng.update(type="numbered", numberOfOccurrences=_clamp(count, 1, 999))
    else:
        rng["type"] = "noEnd"
    return {"pattern": patterns[repeat], "range": rng}


def _show_as(value: str) -> str:
    if value not in cal.SHOW_AS:
        raise ToolError("show_as must be one of: " + ", ".join(cal.SHOW_AS))
    return value


def _body(body: str, format: str) -> dict:
    return {"contentType": "HTML", "content": compose.wrap_html(compose.build_body(body, format, None))}


def _event_line(verb: str, ev: dict, tz, extra: str = "") -> str:
    d = cal.event(ev, tz)
    out = f"{verb} {d['subject']!r}, {cal.when_text(d)} ({tz.key})" + extra
    if isinstance(d.get("online_meeting"), str):
        out += f"\nTeams link: {d['online_meeting']}"
    out += f"\nEvent id: {d['event_id']}"
    if d.get("web_link"):
        out += f"\nOpen in Outlook: {d['web_link']}"
    return out


async def create_event_data(ctx, subject, start, end, duration_minutes, all_day, attendees,
                            optional_attendees, location, body, format, online_meeting, show_as,
                            reminder_minutes, categories, sensitivity, calendar, repeat, repeat_until,
                            repeat_count) -> str:
    tz, now = await _tz(), _now()
    s = _when(start, "start", tz, now)
    if all_day:
        s = datetime.combine(s.date(), datetime.min.time(), tz)
        last = _when(end, "end", tz, now).date() if end.strip() else s.date()
        e = datetime.combine(last + timedelta(days=1), datetime.min.time(), tz)  # end date is inclusive
    else:
        e = _when(end, "end", tz, now) if end.strip() else s + timedelta(minutes=_clamp(duration_minutes or 30,
                                                                                      5, 1440 * 14))
    if e <= s:
        raise ToolError("end must be after start.")
    if sensitivity not in ("normal", "private", "personal", "confidential"):
        raise ToolError("sensitivity must be normal, private, personal or confidential.")
    ev: dict = {
        "subject": (subject or "").strip()[:255] or "(no subject)",
        "start": cal.to_graph(s, tz), "end": cal.to_graph(e, tz), "isAllDay": bool(all_day),
        "showAs": _show_as(show_as or ("free" if all_day else "busy")), "sensitivity": sensitivity,
    }
    if body.strip():
        ev["body"] = _body(body, format)
    if location.strip():
        ev["location"] = {"displayName": location.strip()[:255]}
    if online_meeting:
        ev.update(isOnlineMeeting=True, onlineMeetingProvider="teamsForBusiness")
    if reminder_minutes is not None:
        ev.update(isReminderOn=reminder_minutes >= 0, reminderMinutesBeforeStart=max(reminder_minutes, 0))
    if categories:
        ev["categories"] = [c.strip() for c in (categories if isinstance(categories, list)
                                                else categories.split(",")) if c.strip()]
    rec = _recurrence(repeat, s, repeat_until, repeat_count, tz, now)
    if rec:
        ev["recurrence"] = rec
    atts = _attendees(attendees, optional_attendees)
    if atts:
        _notify_allowed("send meeting invitations; create the event without attendees, and the user can "
                        "invite people from Outlook")
        compose.check_allowed(atts, ALLOWED_DOMAINS)
        ev.update(attendees=atts, responseRequested=True, allowNewTimeProposals=True)
        await _approve_event(ctx, _event_prompt("Send these meeting invitations as you?", ev, tz),
                             f"new|{_digest(ev)}", "Nothing was created or sent.")
    created = await graph.post(f"{await _calendar_path(calendar)}/events", ev)
    extra = f" · invitations sent to {len(atts)} attendee{'s' if len(atts) != 1 else ''}" if atts else ""
    return _event_line("Created", created, tz, extra)


MEETING_FIELDS = ("subject", "start", "end", "location", "body", "attendees", "isOnlineMeeting",
                  "onlineMeetingProvider")


async def update_event_data(ctx, event_id, subject, start, end, location, body, format, add_attendees,
                            add_optional_attendees, remove_attendees, online_meeting, show_as,
                            reminder_minutes, categories) -> str:
    tz, now = await _tz(), _now()
    ev = await _get_event(event_id, full=False)
    patch: dict = {}
    if subject is not None:
        patch["subject"] = subject.strip()[:255]
    if start or end:
        os_, oe = cal.from_graph(ev.get("start"), tz), cal.from_graph(ev.get("end"), tz)
        if ev.get("isAllDay"):
            ns = datetime.combine(_when(start, "start", tz, now).date(), datetime.min.time(), tz) if start \
                else os_
            ne = (datetime.combine(_when(end, "end", tz, now).date() + timedelta(days=1), datetime.min.time(),
                                   tz) if end else ns + (oe - os_))
        else:
            ns = _when(start, "start", tz, now) if start else os_
            ne = _when(end, "end", tz, now) if end else ns + (oe - os_)
        if ne <= ns:
            raise ToolError("end must be after start.")
        patch.update(start=cal.to_graph(ns, tz), end=cal.to_graph(ne, tz))
    if location is not None:
        patch["location"] = {"displayName": location.strip()[:255]}
    if body is not None:
        patch["body"] = _body(body, format)
    if online_meeting is not None:
        patch["isOnlineMeeting"] = bool(online_meeting)
        if online_meeting:
            patch["onlineMeetingProvider"] = "teamsForBusiness"
    current = ev.get("attendees") or []
    added = _attendees(add_attendees, add_optional_attendees)
    removed = set(compose.addresses(compose.parse_recipients(remove_attendees, "remove_attendees"))) \
        if remove_attendees else set()
    if added or removed:
        have = {(a.get("emailAddress") or {}).get("address", "").lower() for a in current}
        unknown = removed - have
        if unknown:
            raise ToolError("Not attendees of this event: " + ", ".join(sorted(unknown)))
        kept = [a for a in current if (a.get("emailAddress") or {}).get("address", "").lower() not in removed]
        patch["attendees"] = kept + [a for a in added
                                     if a["emailAddress"]["address"].lower() not in have]
    if show_as is not None:
        patch["showAs"] = _show_as(show_as)
    if reminder_minutes is not None:
        patch.update(isReminderOn=reminder_minutes >= 0, reminderMinutesBeforeStart=max(reminder_minutes, 0))
    if categories is not None:
        patch["categories"] = [c.strip() for c in (categories if isinstance(categories, list)
                                                   else categories.split(",")) if c.strip()]
    if not patch:
        raise ToolError("Nothing to change: give at least one field.")
    meeting_change = any(k in patch for k in MEETING_FIELDS)
    if meeting_change and not ev.get("isOrganizer"):
        raise ToolError("You're an attendee of this meeting, not its organiser: only show_as, reminder_minutes "
                        "and categories can be changed. To suggest another time, use respond_to_event with "
                        "propose_start.")
    attendees_after = patch.get("attendees", current)
    notify = meeting_change and bool(attendees_after or current)
    if notify:
        _notify_allowed("send meeting updates to attendees")
        compose.check_allowed(attendees_after, ALLOWED_DOMAINS)
        preview = {**ev, **patch}
        if "body" in patch:
            preview["body"] = patch["body"]
        changes = ", ".join(k for k in patch if k in MEETING_FIELDS and k != "onlineMeetingProvider")
        await _approve_event(ctx, _event_prompt(f"Send this meeting update ({changes}) to the attendees as you?",
                                                preview, tz),
                             f"{ev['id']}|{_digest(patch)}", "The meeting was not changed.")
    updated = await graph.patch(f"/me/events/{_q(ev['id'])}", patch)
    n = len(attendees_after)
    extra = f" · update sent to {n} attendee{'s' if n != 1 else ''}" if notify else ""
    return _event_line("Updated", {**ev, **updated} if updated else {**ev, **patch}, tz, extra)


RESPONSE_ACTIONS = {"accept": "accept", "tentative": "tentativelyAccept", "decline": "decline"}


async def respond_to_event_data(ctx, event_id, message_id, response, comment, send_response,
                                propose_start, propose_end) -> str:
    tz, now = await _tz(), _now()
    response = (response or "").strip().lower()
    if response not in RESPONSE_ACTIONS:
        raise ToolError("response must be accept, tentative or decline.")
    if not event_id and message_id:
        m = await graph.get(f"/me/messages/{_q(_id(message_id, 'message_id', 'list_messages'))}",
                            {"$select": "id,subject", "$expand": "microsoft.graph.eventMessage/event($select=id)"})
        event_id = (m.get("event") or {}).get("id")
        if not event_id:
            raise ToolError("That message isn't a meeting invitation (or the meeting no longer exists).")
    if not event_id:
        raise ToolError("Give event_id (from list_events) or the invitation's message_id.")
    ev = await _get_event(event_id, full=False)
    if ev.get("isOrganizer"):
        raise ToolError("You organise this meeting, so there is nothing to answer. Use update_event or "
                        "cancel_event.")
    payload: dict = {"sendResponse": bool(send_response)}
    if comment.strip():
        payload["comment"] = comment.strip()
    if propose_start:
        if response == "accept":
            raise ToolError("A new time can only be proposed with response 'tentative' or 'decline'.")
        if not send_response:
            raise ToolError("Proposing a new time needs send_response=True (the organiser must be told).")
        os_, oe = cal.from_graph(ev.get("start"), tz), cal.from_graph(ev.get("end"), tz)
        ps = _when(propose_start, "propose_start", tz, now)
        pe = _when(propose_end, "propose_end", tz, now) if propose_end else ps + (oe - os_)
        if pe <= ps:
            raise ToolError("propose_end must be after propose_start.")
        payload["proposedNewTime"] = {"start": cal.to_graph(ps, tz), "end": cal.to_graph(pe, tz)}
    verb = {"accept": "Accept", "tentative": "Tentatively accept", "decline": "Decline"}[response]
    if send_response:
        _notify_allowed("send meeting responses; use send_response=False to record the answer only in the "
                        "calendar")
        if payload.get("comment") or payload.get("proposedNewTime"):
            compose.check_allowed([ev.get("organizer") or {}], ALLOWED_DOMAINS)
        org = fmt.addr_text(fmt.address(ev.get("organizer")))
        head = f"{verb} this meeting and tell the organiser ({org}) as you?"
        if payload.get("proposedNewTime"):
            p = payload["proposedNewTime"]
            head += (f"\nProposed new time: "
                     f"{cal.when_text({'start': cal.from_graph(p['start'], tz).isoformat(), 'end': cal.from_graph(p['end'], tz).isoformat()})}")
        if payload.get("comment"):
            head += f"\nMessage: {payload['comment']}"
        prompt = _event_prompt(head, {k: v for k, v in ev.items() if k != "body"}, tz)
        await _approve_event(ctx, prompt, f"{ev['id']}|{_digest(response, payload)}",
                             "The invitation was not answered.")
    await graph.post(f"/me/events/{_q(ev['id'])}/{RESPONSE_ACTIONS[response]}", payload)
    past = {"accept": "Accepted", "tentative": "Tentatively accepted", "decline": "Declined"}[response]
    extra = (" · the organiser was notified" + (" with a proposed new time" if payload.get("proposedNewTime")
                                                 else "")) if send_response else " · the organiser was not notified"
    if response == "decline":
        extra += " · removed from your calendar"
    return _event_line(past, ev, tz, extra)


async def cancel_event_data(ctx, event_id, comment) -> str:
    tz = await _tz()
    ev = await _get_event(event_id, full=False)
    if not ev.get("isOrganizer"):
        raise ToolError("You're an attendee of this meeting, not its organiser. To drop out, use "
                        "respond_to_event with response='decline'.")
    atts = ev.get("attendees") or []
    if atts:
        _notify_allowed("send meeting cancellations")
        if comment.strip():
            compose.check_allowed(atts, ALLOWED_DOMAINS)
        head = f"Cancel this meeting and notify {len(atts)} attendee{'s' if len(atts) != 1 else ''} as you?"
        if comment.strip():
            head += f"\nMessage: {comment.strip()}"
        await _approve_event(ctx, _event_prompt(head, ev, tz), f"{ev['id']}|{_digest('cancel', comment)}",
                             "The meeting was not cancelled.")
        await graph.post(f"/me/events/{_q(ev['id'])}/cancel", {"Comment": comment.strip()})
        return _event_line("Cancelled", ev, tz, f" · cancellation sent to {len(atts)} attendee"
                                                f"{'s' if len(atts) != 1 else ''}")
    await graph.delete(f"/me/events/{_q(ev['id'])}")
    return _event_line("Deleted", ev, tz, " · it was only in your calendar (no attendees)")


# ---------------------------------------------------------------- calendar tools
Attendees = Annotated[list[str] | str | None, Field(
    description="Email addresses: a list, or one string separated by ';' or ','. Plain names are refused: "
                "use find_people first.")]
When = Annotated[str, Field(description="Date/time in the user's time zone (see whoami), e.g. "
                                        "2026-11-18T09:30, or a date 2026-11-18, 'today', 'tomorrow', 'now'. "
                                        "A Z or +hh:mm suffix is converted.")]


async def list_calendars(output: str = "text") -> str:
    """The user's calendars (name, id, owner). Most tools use the default
    calendar unless you pass a calendar name."""
    return await _run(list_calendars_data(), calendars_text, output)


async def list_events(start: str = "", end: str = "", days: int = 7, calendar: str = "", query: str = "",
                      needs_response: bool = False, limit: int = 50, output: str = "text") -> str:
    """Events in a time range, oldest first, with recurring meetings expanded.
    start defaults to today 00:00; end defaults to start + days (an end date
    includes that whole day). query keeps events whose subject, location,
    organiser or attendees contain all its words. needs_response=True keeps
    invitations the user hasn't answered. Times are in the user's time zone."""
    return await _run(list_events_data(start, end, days, calendar, query, needs_response, limit),
                      events_text, output)


async def get_event(event_id: str, max_chars: int = 8000, output: str = "text") -> str:
    """One event in full: time, recurrence, location, Teams link, organiser,
    each attendee's response, and the description."""
    return await _run(get_event_data(event_id, max_chars), event_text, output)


async def find_free_time(start: When = "", end: When = "", duration_minutes: int = 30,
                         attendees: Attendees = None, working_hours_only: bool = True,
                         tentative_is_busy: bool = True, max_slots: int = 20, output: str = "text") -> str:
    """Free time for the user, and for attendees if given (their free/busy from
    Outlook), at least duration_minutes long. Defaults: from now for 7 days,
    within the user's Outlook working hours. Use it before proposing or creating
    a meeting. People outside the organisation are usually not visible and are
    reported as such."""
    return await _run(find_free_time_data(start, end, duration_minutes, attendees, working_hours_only,
                                          tentative_is_busy, max_slots), cal.render_slots, output)


async def create_event(subject: str, start: When, ctx: Context, end: When = "", duration_minutes: int = 30,
                       all_day: bool = False, attendees: Attendees = None, optional_attendees: Attendees = None,
                       location: str = "", body: str = "", format: str = "text", online_meeting: bool = False,
                       show_as: str = "", reminder_minutes: int | None = None,
                       categories: list[str] | str | None = None, sensitivity: str = "normal",
                       calendar: str = "", repeat: str = "", repeat_until: str = "",
                       repeat_count: int = 0) -> str | InputRequiredResult:
    """Create an appointment, or a meeting when attendees are given.
    WITH ATTENDEES, INVITATIONS ARE SENT IMMEDIATELY: only after the user approved
    the subject, time, attendees and text. Without attendees nothing is sent.
    end, or duration_minutes (default 30). all_day: start/end are dates, end
    inclusive. online_meeting=True adds a Teams link. show_as: busy (default),
    free, tentative, oof, workingElsewhere. reminder_minutes: -1 = no reminder.
    repeat: daily, weekdays, weekly, biweekly, monthly (with repeat_until date or
    repeat_count)."""
    return await _call_send(create_event_data(
        ctx, subject, start, end, duration_minutes, all_day, attendees, optional_attendees, location, body,
        format, online_meeting, show_as, reminder_minutes, categories, sensitivity, calendar, repeat,
        repeat_until, repeat_count))


async def update_event(event_id: str, ctx: Context, subject: str | None = None, start: When = "",
                       end: When = "", location: str | None = None, body: str | None = None,
                       format: str = "text", add_attendees: Attendees = None,
                       add_optional_attendees: Attendees = None, remove_attendees: Attendees = None,
                       online_meeting: bool | None = None, show_as: str | None = None,
                       reminder_minutes: int | None = None,
                       categories: list[str] | str | None = None) -> str | InputRequiredResult:
    """Change an event. Only the fields given change; a new start without end
    keeps the duration. If the user organises a meeting with attendees, changes
    to time, place, text or attendees SEND AN UPDATE to everyone: only after the
    user approved the change. As an attendee only show_as, reminder_minutes and
    categories can be changed. Pass an occurrence's id for one occurrence, or the
    series id for the whole series."""
    return await _call_send(update_event_data(
        ctx, event_id, subject, start, end, location, body, format, add_attendees, add_optional_attendees,
        remove_attendees, online_meeting, show_as, reminder_minutes, categories))


async def respond_to_event(response: str, ctx: Context, event_id: str = "", message_id: str = "",
                           comment: str = "", send_response: bool = True, propose_start: When = "",
                           propose_end: When = "") -> str | InputRequiredResult:
    """Answer a meeting invitation: response 'accept', 'tentative' or 'decline'.
    Give event_id (from list_events) or the invitation email's message_id. By
    default the organiser is told (with the optional comment): only after the user
    approved this answer. send_response=False records it silently. With
    tentative/decline, propose_start (and propose_end) suggests another time.
    Use the series id to answer every occurrence of a recurring meeting.
    Declining removes the meeting from the calendar."""
    return await _call_send(respond_to_event_data(ctx, event_id, message_id, response, comment,
                                                  send_response, propose_start, propose_end))


async def cancel_event(event_id: str, ctx: Context, comment: str = "") -> str | InputRequiredResult:
    """Cancel a meeting the user organises (SENDS A CANCELLATION to all attendees,
    with the optional comment: only after the user approved), or delete an
    appointment that has no attendees. Pass the series id to cancel the whole
    series. For meetings organised by someone else, decline with respond_to_event."""
    return await _call_send(cancel_event_data(ctx, event_id, comment))


CALENDAR_READ_TOOLS = (list_calendars, list_events, get_event, find_free_time)
CALENDAR_WRITE_TOOLS = (create_event, update_event, respond_to_event, cancel_event)

if ENABLE_CALENDAR:
    for _fn in CALENDAR_READ_TOOLS:
        mcp.tool(**READ)(_fn)
    if ENABLE_WRITE:
        for _fn in CALENDAR_WRITE_TOOLS:
            mcp.tool(annotations=SENDS, output_schema=None)(_fn)


# ---------------------------------------------------------------- http + auth
@mcp.custom_route("/health", methods=["GET"])
async def health(_: Request):
    return JSONResponse({"status": "ok", "auth": tokens.status()})


class BearerAuth:
    """Minimal ASGI middleware: every request except /health needs the token."""

    def __init__(self, app, token: str):
        self.app = app
        self.expected = f"Bearer {token}".encode()

    async def __call__(self, scope, receive, send):
        if scope["type"] == "http" and scope.get("path") != "/health":
            supplied = dict(scope.get("headers") or []).get(b"authorization", b"")
            if not hmac.compare_digest(supplied, self.expected):
                await send({"type": "http.response.start", "status": 401,
                            "headers": [(b"content-type", b"text/plain")]})
                await send({"type": "http.response.body", "body": b"Unauthorized"})
                return
        await self.app(scope, receive, send)


def build_app():
    return BearerAuth(mcp.http_app(path="/mcp"), MCP_TOKEN)


async def _keepalive_loop() -> None:
    """Refresh periodically so the stored refresh token keeps rotating even when
    the server is idle; otherwise it would expire after ~90 days of no use."""
    if KEEPALIVE_HOURS <= 0:
        return
    await asyncio.sleep(60)
    while True:
        await tokens.keepalive()
        await asyncio.sleep(KEEPALIVE_HOURS * 3600)


async def main() -> None:
    log.info("Starting on %s:%d (write tools %s, send tools %s, confirmation %s, vault %s, sign-in %s)",
             HOST, PORT, "enabled" if ENABLE_WRITE else "disabled", "enabled" if ENABLE_SEND else "disabled",
             SEND_CONFIRMATION, VAULT_PATH, tokens.status())
    server = uvicorn.Server(uvicorn.Config(build_app(), host=HOST, port=PORT, proxy_headers=True))
    keepalive = asyncio.create_task(_keepalive_loop())
    try:
        await server.serve()
    finally:
        keepalive.cancel()
        await graph.aclose()
        await tokens.aclose()


if __name__ == "__main__":
    asyncio.run(main())
