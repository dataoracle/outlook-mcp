"""Turn Graph mail objects into compact output for the model.

Two layers, as in msteams-mcp:
- normalisers (summary, message, attachment, folder, ...) turn Graph JSON into
  small dicts, keeping only what is useful; these are also the output="json" form;
- renderers (render_*) turn those dicts into compact text, the default output.
  Ids appear where a follow-up call needs them; links only in single-item tools.
"""
from __future__ import annotations

import re
from datetime import datetime, timezone
from html.parser import HTMLParser
from urllib.parse import parse_qs, unquote, urlparse

_BLOCK = {"p", "div", "tr", "h1", "h2", "h3", "h4", "h5", "h6", "blockquote", "pre", "ul", "ol",
          "table", "hr"}
_SKIP = {"style", "script", "head", "title"}


class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out: list[str] = []
        self._skip = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag in _SKIP:
            self._skip += 1
        elif tag == "br":
            self.out.append("\n")
        elif tag == "li":
            self.out.append("\n- ")
        elif tag == "td" or tag == "th":
            self.out.append(" | ")
        elif tag == "img":
            self.out.append(f"[{a.get('alt') or 'image'}]")
        elif tag in _BLOCK:
            self.out.append("\n")

    def handle_endtag(self, tag):
        if tag in _SKIP:
            self._skip = max(0, self._skip - 1)
        elif tag in _BLOCK:
            self.out.append("\n")

    def handle_data(self, data):
        if not self._skip:
            self.out.append(data)


def html_to_text(content: str) -> str:
    p = _TextExtractor()
    p.feed(content or "")
    p.close()
    return tidy("".join(p.out))


# Whitespace in mail bodies, including the invisible kinds that str.strip() keeps.
_WS = " \t\u00a0\u2007\u202f\u200b\u2060\ufeff\u200c\u200d"
_SAFELINK_RE = re.compile(r"https://[a-z0-9.\-]*safelinks\.protection\.outlook\.com/[^\s<>\]\)\"']+",
                          re.IGNORECASE)
_CID_RE = re.compile(r"\[cid:[^\]]*\]", re.IGNORECASE)


def _unwrap_safelink(m: re.Match) -> str:
    try:
        target = parse_qs(urlparse(m.group(0)).query).get("url", [""])[0]
    except ValueError:
        target = ""
    return unquote(target) if target else m.group(0)


def clean_links(text: str) -> str:
    """Unwrap Defender Safe Links (hundreds of characters each) to the real URL,
    and replace inline-image references with [image]."""
    text = _SAFELINK_RE.sub(_unwrap_safelink, text or "")
    return _CID_RE.sub("[image]", text)


def tidy(text: str) -> str:
    """Trim every line, and keep at most one blank line between paragraphs."""
    lines = [ln.strip(_WS) for ln in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n")]
    out: list[str] = []
    for ln in lines:
        if ln or (out and out[-1]):
            out.append(ln)
    while out and not out[-1]:
        out.pop()
    return "\n".join(out)


def body_text(body: dict | None) -> str:
    body = body or {}
    content = body.get("content") or ""
    if (body.get("contentType") or "").lower() == "html":
        return clean_links(html_to_text(content))
    return clean_links(tidy(content))


def parse_time(value: str) -> datetime:
    """Parse ISO 8601 (with Z or offset, up to 7 fractional digits). No offset = UTC.
    A bare date (2026-11-13) means its 00:00 UTC."""
    s = value.strip().replace("Z", "+00:00").replace("z", "+00:00")
    if re.fullmatch(r"\d{4}-\d\d-\d\d", s):
        s += "T00:00:00"
    m = re.match(r"^(.*T\d\d:\d\d:\d\d)(\.\d+)?(.*)$", s)
    if m and m.group(2):
        s = m.group(1) + m.group(2)[:7] + m.group(3)
    dt = datetime.fromisoformat(s)
    return dt if dt.tzinfo else dt.replace(tzinfo=timezone.utc)


def iso_z(dt: datetime) -> str:
    return dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def utc(value: str | None) -> datetime | None:
    try:
        return parse_time(value).astimezone(timezone.utc) if value else None
    except ValueError:
        return None


def short_time(value: str | None) -> str:
    """'2026-11-13T10:03:21Z' -> '2026-11-13 10:03' (UTC)."""
    t = utc(value)
    return t.strftime("%Y-%m-%d %H:%M") if t else (value or "?")


def shorten(text: str, n: int) -> str:
    text = " ".join((text or "").split())
    return text if len(text) <= n else text[: n - 1].rstrip() + "…"


def human_size(n: int | None) -> str:
    n = int(n or 0)
    if n < 1024:
        return f"{n} B"
    if n < 1024 * 1024:
        return f"{n / 1024:.0f} KB"
    return f"{n / 1024 / 1024:.1f} MB"


# ---------------------------------------------------------------- normalisers
def address(r: dict | None) -> dict:
    e = (r or {}).get("emailAddress") or {}
    return {k: v for k, v in {"name": e.get("name"), "email": e.get("address")}.items() if v}


def addr_text(a: dict) -> str:
    name, email = a.get("name"), a.get("email")
    if name and email and name.lower() != email.lower():
        return f"{name} <{email}>"
    return email or name or "?"


def _names(addrs: list[dict], cap: int) -> str:
    names = [a.get("name") or a.get("email") or "?" for a in addrs]
    return ", ".join(names[:cap]) + (f" +{len(names) - cap}" if len(names) > cap else "")


def _flags(m: dict) -> dict:
    out: dict = {}
    if m.get("isRead") is False and not m.get("isDraft"):
        out["unread"] = True
    if m.get("isDraft"):
        out["draft"] = True
    if m.get("importance") not in (None, "normal"):
        out["importance"] = m["importance"]
    flag = (m.get("flag") or {}).get("flagStatus")
    if flag and flag != "notFlagged":
        out["flag"] = flag
        due = ((m.get("flag") or {}).get("dueDateTime") or {}).get("dateTime")
        if due:
            out["flag_due"] = due[:10]
    if m.get("hasAttachments"):
        out["attachments"] = True
    if m.get("categories"):
        out["categories"] = m["categories"]
    mtype = (m.get("@odata.type") or "").rsplit(".", 1)[-1]
    if mtype and mtype != "message":
        out["kind"] = mtype  # eventMessageRequest, eventMessageResponse, ...
    return out


SUMMARY_FIELDS = ("id,receivedDateTime,sentDateTime,from,toRecipients,ccRecipients,subject,"
                  "bodyPreview,isRead,isDraft,importance,flag,hasAttachments,categories,"
                  "conversationId,parentFolderId")


def summary(m: dict, preview_chars: int = 160) -> dict:
    out: dict = {"message_id": m.get("id"),
                 "received": m.get("receivedDateTime") or m.get("sentDateTime"),
                 "from": address(m.get("from") or m.get("sender")),
                 "to": [address(r) for r in m.get("toRecipients") or []],
                 "subject": m.get("subject") or "(no subject)"}
    if m.get("ccRecipients"):
        out["cc"] = [address(r) for r in m["ccRecipients"]]
    if preview_chars > 0 and m.get("bodyPreview"):
        out["preview"] = shorten(clean_links(m["bodyPreview"]), preview_chars)
    if m.get("conversationId"):
        out["conversation_id"] = m["conversationId"]
    out.update(_flags(m))
    return out


FULL_FIELDS = (SUMMARY_FIELDS + ",bccRecipients,replyTo,body,uniqueBody,webLink,internetMessageId,"
               "createdDateTime,lastModifiedDateTime,inferenceClassification")


def attachment(a: dict) -> dict:
    kind = (a.get("@odata.type") or "").rsplit(".", 1)[-1].removesuffix("Attachment") or "file"
    out = {"attachment_id": a.get("id"), "name": a.get("name") or "(unnamed)",
           "content_type": a.get("contentType"), "size": a.get("size"), "kind": kind}
    if a.get("isInline"):
        out["inline"] = True
    if a.get("sourceUrl"):  # reference attachment (a OneDrive/SharePoint link)
        out["url"] = a["sourceUrl"]
    return {k: v for k, v in out.items() if v not in (None, "")}


def message(m: dict, max_chars: int, body: str = "full", attachments: list[dict] | None = None,
            offset: int = 0) -> dict:
    """A message in full. body='new' uses uniqueBody (only what this message added,
    without the quoted history), falling back to the full body. offset/max_chars
    page through a long body."""
    out = summary(m, preview_chars=0)
    out.pop("conversation_id", None)
    out["conversation_id"] = m.get("conversationId")
    if m.get("bccRecipients"):
        out["bcc"] = [address(r) for r in m["bccRecipients"]]
    if m.get("replyTo"):
        out["reply_to"] = [address(r) for r in m["replyTo"]]
    if m.get("sentDateTime") and m.get("sentDateTime") != m.get("receivedDateTime"):
        out["sent"] = m["sentDateTime"]
    src = m.get("uniqueBody") if body == "new" and (m.get("uniqueBody") or {}).get("content") else m.get("body")
    full_text = body_text(src)
    text = full_text[offset:] if offset else full_text
    if offset:
        out["offset"] = offset
    out["body"] = text
    if max_chars and len(text) > max_chars:
        out["body"] = text[:max_chars].rstrip() + "…"
        out["truncated_at"] = offset + max_chars
        out["body_chars"] = len(full_text)
    if body == "new" and src is m.get("uniqueBody"):
        out["body_kind"] = "new text only (quoted history removed)"
    if attachments is not None:
        out["attachments"] = [attachment(a) for a in attachments]
    if m.get("internetMessageId"):
        out["internet_message_id"] = m["internetMessageId"]
    if m.get("webLink"):
        out["web_link"] = m["webLink"]
    return out


WELL_KNOWN = {"inbox": "Inbox", "drafts": "Drafts", "sentitems": "Sent Items",
              "deleteditems": "Deleted Items", "junkemail": "Junk Email", "archive": "Archive",
              "outbox": "Outbox"}


def folder(f: dict, path: str) -> dict:
    return {k: v for k, v in {
        "folder_id": f.get("id"), "path": path,
        "unread": f.get("unreadItemCount"), "total": f.get("totalItemCount"),
        "children": f.get("childFolderCount"),
    }.items() if v not in (None, "")}


def person(p: dict, source: str) -> dict:
    email = (p.get("mail") or p.get("userPrincipalName")
             or ((p.get("scoredEmailAddresses") or [{}])[0].get("address"))
             or ((p.get("emailAddresses") or [{}])[0].get("address")) or p.get("email"))
    return {k: v for k, v in {
        "name": p.get("displayName") or p.get("name"),
        "email": email,
        "job_title": p.get("jobTitle"),
        "department": p.get("department"),
        "company": p.get("companyName"),
        "office": p.get("officeLocation"),
        "source": source,
        "last_mail": p.get("last_mail"),
    }.items() if v}


# ---------------------------------------------------------------- renderers
_PAD = "  "


def _marks(d: dict) -> str:
    marks = []
    if d.get("unread"):
        marks.append("UNREAD")
    if d.get("draft"):
        marks.append("DRAFT")
    if d.get("importance"):
        marks.append(f"!{d['importance']}")
    if d.get("flag") == "flagged":
        marks.append("⚑" + (f" due {d['flag_due']}" if d.get("flag_due") else ""))
    elif d.get("flag") == "complete":
        marks.append("⚑done")
    if d.get("attachments") is True:
        marks.append("📎")
    if d.get("kind"):
        marks.append(f"({d['kind']})")
    if d.get("categories"):
        marks.append("[" + ", ".join(d["categories"]) + "]")
    return " ".join(marks)


def render_summaries(msgs: list[dict], show: str = "from", folders: dict | None = None) -> str:
    """One block per message, a date line per day:
    HH:MM Sender <email> · Subject · MARKS
      preview
      id <message id>"""
    lines: list[str] = []
    day = None
    for d in msgs:
        t = utc(d.get("received"))
        d_day = t.strftime("%Y-%m-%d") if t else "?"
        if d_day != day:
            lines.append(d_day)
            day = d_day
        hm = t.strftime("%H:%M") if t else "--:--"
        who = (("→ " + (_names(d.get("to") or [], 3) or "(no recipients)")) if show == "to"
               else addr_text(d.get("from") or {}))
        head = " · ".join(x for x in (f"{hm} {who}", d.get("subject"), _marks(d)) if x)
        if folders and d.get("folder"):
            head += f" · in {d['folder']}"
        lines.append(head)
        if d.get("preview"):
            lines.append(_PAD + d["preview"])
        lines.append(f"{_PAD}id {d.get('message_id')}")
    return "\n".join(lines)


def _addr_line(label: str, addrs: list[dict] | None) -> str | None:
    return f"{label}: " + "; ".join(addr_text(a) for a in addrs) if addrs else None


def render_attachments(atts: list[dict]) -> str:
    out = []
    for a in atts:
        bits = [a["name"], human_size(a.get("size"))]
        if a.get("inline"):
            bits.append("inline")
        if a.get("kind") not in (None, "file"):
            bits.append(a["kind"])
        line = f"- {' · '.join(bits)} · attachment_id {a.get('attachment_id')}"
        if a.get("url"):
            line += f" · {a['url']}"
        out.append(line)
    return "\n".join(out)


def render_message(d: dict) -> str:
    lines = [f"Subject: {d.get('subject')}"]
    for label, key in (("From", "from"),):
        if d.get(key):
            lines.append(f"{label}: {addr_text(d[key])}")
    for label, key in (("To", "to"), ("Cc", "cc"), ("Bcc", "bcc"), ("Reply-To", "reply_to")):
        line = _addr_line(label, d.get(key))
        if line:
            lines.append(line)
    date = f"Date: {short_time(d.get('received'))} UTC"
    marks = _marks({k: v for k, v in d.items() if k != "attachments"})
    lines.append(date + (f" · {marks}" if marks else ""))
    lines.append(f"Message id: {d.get('message_id')}")
    if d.get("conversation_id"):
        lines.append(f"Conversation id: {d['conversation_id']}")
    atts = d.get("attachments") or []
    visible = [a for a in atts if not a.get("inline")]
    if atts:
        lines.append(f"Attachments ({len(visible)}"
                     + (f", plus {len(atts) - len(visible)} inline" if len(atts) > len(visible) else "")
                     + "):")
        lines.append(render_attachments(visible or atts))
    if d.get("web_link"):
        lines.append(f"Open in Outlook: {d['web_link']}")
    if d.get("body_kind"):
        lines.append(f"({d['body_kind']})")
    lines.append("")
    lines.append(d.get("body") or "(empty body)")
    if d.get("truncated_at"):
        lines.append(f"\n(body cut at {d['truncated_at']} of {d['body_chars']} characters: "
                     f"pass offset={d['truncated_at']} to continue)")
    return "\n".join(lines)


def render_thread(msgs: list[dict]) -> str:
    blocks = []
    for d in msgs:
        head = f"── {short_time(d.get('received'))} {addr_text(d.get('from') or {})}"
        marks = _marks({k: v for k, v in d.items() if k != "attachments"})
        if marks:
            head += f" · {marks}"
        lines = [head]
        to = _addr_line("To", d.get("to"))
        if to:
            lines.append(to)
        cc = _addr_line("Cc", d.get("cc"))
        if cc:
            lines.append(cc)
        atts = [a for a in d.get("attachments") or [] if not a.get("inline")]
        if atts:
            lines.append("Attachments: " + "; ".join(a["name"] for a in atts))
        lines.append(f"id {d.get('message_id')}")
        lines.append(d.get("body") or "(empty)")
        if d.get("truncated_at"):
            lines.append(f"(cut at {d['truncated_at']} characters: get_message shows the rest)")
        blocks.append("\n".join(lines))
    return "\n\n".join(blocks)


def render_folders(folders: list[dict]) -> str:
    out = []
    for f in folders:
        depth = f["path"].count("/")
        name = f["path"].rsplit("/", 1)[-1]
        counts = f"{f.get('unread', 0)} unread / {f.get('total', 0)}"
        out.append(f"{'  ' * depth}• {name} · {counts} · id {f.get('folder_id')}")
    return "\n".join(out) or "No folders."


def render_people(people: list[dict]) -> str:
    return "\n".join(
        "• " + " · ".join(x for x in (
            f"{p.get('name') or '?'} <{p.get('email', '?')}>", p.get("job_title"), p.get("department"),
            p.get("company"), p.get("office"),
            f"last mail {short_time(p['last_mail'])}" if p.get("last_mail") else None,
            f"({p['source']})" if p.get("source") else None) if x)
        for p in people) or "No matches."


def with_header(header: str, body: str, note: str | None = None) -> str:
    return "\n".join(x for x in (header, body, f"({note})" if note else None) if x)
