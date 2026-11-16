"""Building outgoing mail: recipients, bodies, signature, attachments, send policy."""
from __future__ import annotations

import base64
import binascii
import html
import re

import markdown as md

EMAIL_RE = re.compile(r"^[A-Za-z0-9.!#$%&'*+/=?^_`{|}~\-]+@[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?"
                      r"(?:\.[A-Za-z0-9](?:[A-Za-z0-9\-]{0,61}[A-Za-z0-9])?)+$")
# "Name <email>" as copied from a header or a list line.
_NAMED_RE = re.compile(r"^\s*\"?([^\"<]*?)\"?\s*<([^>]+)>\s*$")

MAX_RECIPIENTS = 100
MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024   # Graph's limit for attachments in a single request
MAX_BODY_BYTES = 3 * 1024 * 1024


class ComposeError(ValueError):
    pass


def _split(text: str) -> list[str]:
    """Split on ; , or newlines outside quotes and <>, then rejoin pieces that were
    an unquoted 'Last, First <a@b.com>' name."""
    parts, cur, quoted, angle = [], "", False, False
    for ch in text:
        if ch == '"':
            quoted = not quoted
        elif ch == "<" and not quoted:
            angle = True
        elif ch == ">" and not quoted:
            angle = False
        if ch in ";,\n" and not quoted and not angle:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    parts.append(cur)
    out: list[str] = []
    pending = ""
    for p in (p.strip() for p in parts):
        if not p:
            continue
        if "@" not in p:
            pending = f"{pending}, {p}" if pending else p
            continue
        if pending and "<" in p and "@" not in p.split("<", 1)[0]:
            p = f"{pending}, {p}"
            pending = ""
        elif pending:
            out.append(pending)
            pending = ""
        out.append(p)
    if pending:
        out.append(pending)
    return out


def parse_recipients(value: list[str] | str | None, label: str) -> list[dict]:
    """Accept a list or a ';'/','-separated string of addresses, each either
    'a@b.com' or 'Name <a@b.com>'. Names alone are refused: an agent must not guess."""
    if value is None:
        return []
    items = [p for v in (value if isinstance(value, list) else [value]) for p in _split(v or "")]
    out, seen = [], set()
    for raw in items:
        raw = (raw or "").strip()
        if not raw:
            continue
        name, addr = "", raw
        m = _NAMED_RE.match(raw)
        if m:
            name, addr = m.group(1).strip(), m.group(2).strip()
        addr = addr.strip().strip("'\"").removeprefix("mailto:")
        if not EMAIL_RE.match(addr):
            raise ComposeError(f"{label}: {raw[:80]!r} is not an email address. Use find_people to "
                               "get the address; names alone are not accepted.")
        if addr.lower() in seen:
            continue
        seen.add(addr.lower())
        e = {"address": addr}
        if name:
            e["name"] = name
        out.append({"emailAddress": e})
    if len(out) > MAX_RECIPIENTS:
        raise ComposeError(f"{label}: more than {MAX_RECIPIENTS} recipients.")
    return out


def addresses(recipients: list[dict]) -> list[str]:
    return [((r.get("emailAddress") or {}).get("address") or "").lower() for r in recipients]


def check_allowed(all_recipients: list[dict], allowed_domains: list[str]) -> None:
    """Refuse to send outside SEND_ALLOWED_DOMAINS (when configured)."""
    if not allowed_domains:
        return
    bad = sorted({a for a in addresses(all_recipients)
                  if not any(a.endswith("@" + d) or a.endswith("." + d) for d in allowed_domains)})
    if bad:
        raise ComposeError("Sending is restricted to " + ", ".join(allowed_domains)
                           + " on this server; not allowed: " + ", ".join(bad)
                           + ". Save it as a draft instead and send it from Outlook if intended.")


def to_html(content: str, fmt: str) -> str:
    content = (content or "").strip("\n")
    if fmt == "text":
        return "<div>" + html.escape(content).replace("\n", "<br>\n") + "</div>"
    if fmt == "markdown":
        return md.markdown(content, extensions=["extra", "sane_lists", "nl2br"])
    if fmt == "html":
        return content
    raise ComposeError("format must be 'text', 'markdown' or 'html'.")


def build_body(content: str, fmt: str, signature_html: str | None, allow_empty: bool = False) -> str:
    if not (content or "").strip() and not allow_empty:
        raise ComposeError("The message body is empty.")
    body = to_html(content, fmt) if (content or "").strip() else ""
    if signature_html:
        body += "<br>\n" + signature_html
    if len(body.encode()) > MAX_BODY_BYTES:
        raise ComposeError("The message body is too large.")
    return body


def wrap_html(body: str) -> str:
    return f"<html><body>{body}</body></html>"


def prepend_to_reply(new_html: str, draft_html: str) -> str:
    """Put our text above the quoted original that createReply/createForward
    produced, keeping the original intact."""
    m = re.search(r"<body[^>]*>", draft_html or "", re.IGNORECASE)
    if not m:
        return wrap_html(new_html + "<br>\n" + (draft_html or ""))
    i = m.end()
    return draft_html[:i] + new_html + "<br>\n" + draft_html[i:]


def attachment_payloads(items: list[dict] | None) -> list[dict]:
    """Small file attachments supplied inline by the agent:
    {name, content (text) | content_base64, content_type?}."""
    out, total = [], 0
    for a in items or []:
        if not isinstance(a, dict):
            raise ComposeError("Each attachment must be an object with name and content or content_base64.")
        name = (a.get("name") or "").strip()
        if not name or len(name) > 255 or "/" in name or "\\" in name:
            raise ComposeError("Each attachment needs a plain file name (no path).")
        if a.get("content_base64"):
            try:
                data = base64.b64decode(a["content_base64"], validate=True)
            except (binascii.Error, ValueError):
                raise ComposeError(f"{name}: content_base64 is not valid base64.") from None
        elif a.get("content") is not None:
            data = str(a["content"]).encode("utf-8")
        else:
            raise ComposeError(f"{name}: give content (text) or content_base64.")
        total += len(data)
        if total > MAX_ATTACHMENT_BYTES:
            raise ComposeError("Attachments over 3 MB in total can't be added through this server; "
                               "attach them in Outlook instead.")
        out.append({"@odata.type": "#microsoft.graph.fileAttachment", "name": name,
                    "contentType": a.get("content_type") or "application/octet-stream",
                    "contentBytes": base64.b64encode(data).decode()})
    return out
