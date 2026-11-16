"""Offline tests for outlook-mcp. Microsoft Graph and Entra ID are mocked by a
small in-memory mailbox; nothing leaves the machine.

Run:  python tests/test_mail.py      (or: pytest tests/)
Pass -v to print sample output.
"""
import asyncio
import io
import json
import os
import sys
import tempfile
import zipfile
from pathlib import Path
from urllib.parse import parse_qsl, unquote

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

import httpx  # noqa: E402

from vault import Vault, load_key, new_key_text, vault_context  # noqa: E402

TENANT, CLIENT = "contoso.onmicrosoft.com", "11111111-2222-3333-4444-555555555555"
_work = Path(tempfile.mkdtemp())
_key = new_key_text()
(_work / "sig.txt").write_text("Antonio Rivas\nData Platform", encoding="utf-8")
os.environ.update(MS_TENANT_ID=TENANT, MS_CLIENT_ID=CLIENT, MCP_TOKEN="test-token",
                  MS_CLIENT_SECRET="test-secret", VAULT_KEY=_key, KEEPALIVE_HOURS="0",
                  VAULT_PATH=str(_work / "vault.json.enc"), SIGNATURE_FILE=str(_work / "sig.txt"),
                  SEND_ALLOWED_DOMAINS="contoso.example,partner.example")
Vault(_work / "vault.json.enc", load_key(_key), vault_context(TENANT, CLIENT)).save(
    {"refresh_token": "RT", "user": {"id": "me-id", "name": "Me", "upn": "me@contoso.example"}})

import compose  # noqa: E402
import formatting as fmt  # noqa: E402
import server  # noqa: E402
from fastmcp import Client  # noqa: E402

VERBOSE = "-v" in sys.argv
requests: list[tuple[str, str, dict, object]] = []


# ---------------------------------------------------------------- fake mailbox
def addr(name, email):
    return {"emailAddress": {"name": name, "address": email}}


ME = addr("Me", "me@contoso.example")
ALICE = addr("Alice Johnson", "alice@contoso.example")
BOB = addr("Bob Smith", "bob@partner.example")
EVE = addr("Eve Outsider", "eve@gmail.example")
SAFE = ("https://emea01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.com%2Freport%3Fq%3D1"
        "&data=05%7C02%7Cme%40contoso.example%7Cabc&sdata=xyz&reserved=0")

FOLDERS = {"inbox": "FOLDER-INBOX-000000000000000000000000000001",
           "drafts": "FOLDER-DRAFTS-00000000000000000000000000002",
           "sentitems": "FOLDER-SENT-000000000000000000000000000003",
           "deleteditems": "FOLDER-DELETED-0000000000000000000000000004",
           "archive": "FOLDER-ARCHIVE-0000000000000000000000000005"}
PROJECTS = "FOLDER-PROJECTS-000000000000000000000000006"
ALPHA = "FOLDER-ALPHA-00000000000000000000000000000007"


def docx_bytes(paragraphs):
    ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
    body = "".join(f"<w:p><w:r><w:t>{p}</w:t></w:r></w:p>" for p in paragraphs)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("word/document.xml", f"<w:document {ns}><w:body>{body}</w:body></w:document>")
    return buf.getvalue()


def xlsx_bytes():
    ns = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
    rns = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("xl/workbook.xml", f'<workbook {ns} {rns}><sheets><sheet name="Q4" sheetId="1" '
                                      f'r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>')
        z.writestr("xl/sharedStrings.xml", f'<sst {ns}><si><t>Region</t></si><si><t>Revenue</t></si>'
                                           f'<si><t>EMEA</t></si></sst>')
        z.writestr("xl/worksheets/sheet1.xml",
                   f'<worksheet {ns}><sheetData>'
                   f'<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>'
                   f'<row r="2"><c r="A2" t="s"><v>2</v></c><c r="C2"><v>40000</v></c></row>'
                   f'</sheetData></worksheet>')
    return buf.getvalue()


def pdf_bytes(text):
    """A minimal one-page PDF with a text layer."""
    stream = f"BT /F1 12 Tf 72 720 Td ({text}) Tj ET".encode()
    objs = [b"<< /Type /Catalog /Pages 2 0 R >>",
            b"<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R "
            b"/Resources << /Font << /F1 5 0 R >> >> >>",
            b"<< /Length " + str(len(stream)).encode() + b" >>\nstream\n" + stream + b"\nendstream",
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"]
    out, offsets = b"%PDF-1.4\n", []
    for i, o in enumerate(objs, 1):
        offsets.append(len(out))
        out += f"{i} 0 obj\n".encode() + o + b"\nendobj\n"
    xref = len(out)
    out += f"xref\n0 {len(objs) + 1}\n0000000000 65535 f \n".encode()
    out += b"".join(f"{o:010d} 00000 n \n".encode() for o in offsets)
    out += f"trailer\n<< /Size {len(objs) + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF".encode()
    return out


def mk(mid, received, frm, subject, body, to=(ME,), cc=(), folder="inbox", read=True, conv="CONV-A",
       draft=False, unique=None, atts=(), importance="normal", flag="notFlagged"):
    return {"id": mid, "receivedDateTime": received, "sentDateTime": received, "from": frm,
            "toRecipients": list(to), "ccRecipients": list(cc), "bccRecipients": [], "replyTo": [],
            "subject": subject, "bodyPreview": body[:255], "isRead": read, "isDraft": draft,
            "importance": importance, "flag": {"flagStatus": flag}, "hasAttachments": bool(atts),
            "categories": [], "conversationId": conv, "parentFolderId": FOLDERS.get(folder, folder),
            "body": {"contentType": "text", "content": body},
            "uniqueBody": {"contentType": "text", "content": unique if unique is not None else body},
            "webLink": f"https://outlook.office.com/mail/{mid}", "internetMessageId": f"<{mid}@x>",
            "_atts": list(atts)}


def att(aid, name, ctype, data, inline=False):
    return {"@odata.type": "#microsoft.graph.fileAttachment", "id": aid, "name": name, "contentType": ctype,
            "size": len(data), "isInline": inline, "_data": data}


def reset_mailbox():
    global MAILBOX, SENT, NEXT
    NEXT = [0]
    SENT = []
    MAILBOX = {m["id"]: m for m in [
        mk("MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA", "2026-11-13T10:01:00Z", ALICE, "Q4 numbers",
           f"Hi,\r\n\r\n\r\nCan you check the Q4 numbers?   \r\nReport: {SAFE}\r\n[cid:image001.png@01DA]\r\n"
           "Thanks", read=False, importance="high",
           atts=[att("ATT-DOCX-000000000001", "Notes.docx",
                     "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                     docx_bytes(["Budget is 40k", "Deadline Friday"])),
                 att("ATT-XLSX-000000000002", "Q4.xlsx",
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", xlsx_bytes()),
                 att("ATT-PDF-0000000000003", "Contract.pdf", "application/pdf",
                     pdf_bytes("Termination notice is 90 days")),
                 att("ATT-PNG-0000000000004", "image001.png", "image/png", b"\x89PNG....", inline=True)]),
        mk("MSG-0002-BBBBBBBBBBBBBBBBBBBBBBBB", "2026-11-13T09:00:00Z", BOB, "Vendor call",
           "The vendor call moved to Thursday.", to=(ALICE,), cc=(ME,), conv="CONV-B"),
        mk("MSG-0003-CCCCCCCCCCCCCCCCCCCCCCCC", "2026-11-12T08:00:00Z", BOB, "Newsletter",
           "Big newsletter about budget", to=(addr("All", "all@partner.example"),), conv="CONV-C"),
        mk("MSG-0004-DDDDDDDDDDDDDDDDDDDDDDDD", "2026-11-13T11:00:00Z", ME, "RE: Q4 numbers",
           "Will do.\n\nFrom: Alice\nCan you check the Q4 numbers?", to=(ALICE,), folder="sentitems",
           unique="Will do."),
        mk("MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE", "2026-11-13T11:30:00Z", ME, "RE: Q4 numbers",
           "half-written", to=(ALICE,), folder="drafts", draft=True),
        mk("MSG-0006-FFFFFFFFFFFFFFFFFFFFFFFF", "2026-11-10T07:00:00Z", ALICE, "Alpha kickoff",
           "Kickoff notes", folder=ALPHA, conv="CONV-D"),
    ]}


def new_id(prefix="NEW"):
    NEXT[0] += 1
    return f"{prefix}-{NEXT[0]:04d}-ZZZZZZZZZZZZZZZZZZZZZZZZ"


def public(m):
    return {k: v for k, v in m.items() if not k.startswith("_")}


def j(data, status=200):
    return httpx.Response(status, json=data)


def forbidden():
    return j({"error": {"code": "Authorization_RequestDenied", "message": "Insufficient privileges"}}, 403)


def apply_filter(rows, flt):
    for clause in flt.split(" and "):
        field, op, value = clause.split(" ", 2)
        value = value.strip("'")
        if field == "receivedDateTime":
            rows = [m for m in rows if (m[field] >= value if op == "ge" else m[field] < value)]
        elif field == "isRead":
            rows = [m for m in rows if m["isRead"] == (value == "true")]
        elif field == "from/emailAddress/address":
            rows = [m for m in rows if m["from"]["emailAddress"]["address"] == value]
        elif field == "conversationId":
            rows = [m for m in rows if m["conversationId"] == value]
        elif field == "hasAttachments":
            rows = [m for m in rows if m["hasAttachments"]]
        elif field == "flag/flagStatus":
            rows = [m for m in rows if m["flag"]["flagStatus"] == value]
        elif field == "importance":
            rows = [m for m in rows if m["importance"] == value]
        else:
            raise AssertionError(f"unexpected filter {clause}")
    return rows


def search(rows, kql):
    kql = kql.strip('"').replace('\\"', '"')
    for part in kql.split(" AND "):
        if ":" in part and part.split(":", 1)[0] in ("from", "to", "subject", "participants"):
            field, value = part.split(":", 1)
            value = value.strip('"').lower()
            if field == "from":
                rows = [m for m in rows if value in json.dumps(m["from"]).lower()]
            elif field == "to":
                rows = [m for m in rows if value in json.dumps(m["toRecipients"]).lower()]
            elif field == "subject":
                rows = [m for m in rows if value in m["subject"].lower()]
            else:
                rows = [m for m in rows if value in json.dumps([m["from"], m["toRecipients"],
                                                                 m["ccRecipients"]]).lower()]
        elif part.startswith("received"):
            continue
        else:
            rows = [m for m in rows if part.lower() in (m["subject"] + m["body"]["content"]).lower()]
    return rows


def graph(req: httpx.Request) -> httpx.Response:
    p = unquote(req.url.path).removeprefix("/v1.0")
    q = dict(req.url.params)
    body = json.loads(req.content) if req.content else None
    requests.append((req.method, p, q, body))
    assert 'IdType="ImmutableId"' in req.headers.get("Prefer", ""), req.headers
    if p == "/$batch":
        out = []
        for r in body["requests"]:
            sub = httpx.Request(r["method"], "https://graph.microsoft.com/v1.0" + r["url"],
                                headers=r.get("headers") or {},
                                content=json.dumps(r["body"]).encode() if "body" in r else b"")
            res = graph(sub)
            out.append({"id": r["id"], "status": res.status_code,
                        "body": res.json() if res.content else {}})
        return j({"responses": out})
    if p == "/me":
        return j({"id": "me-id", "displayName": "Me", "mail": "me@contoso.example",
                  "userPrincipalName": "me@contoso.onmicrosoft.com",
                  "proxyAddresses": ["SMTP:me@contoso.example", "smtp:antonio@contoso.example"]})
    if p in ("/me/people", "/users"):
        return forbidden()
    if p == "/me/mailboxSettings":
        return j({"timeZone": "Arabian Standard Time", "automaticRepliesSetting": {"status": "disabled"}})
    if p == "/me/mailFolders":
        if req.method == "POST":
            return j({"id": new_id("FOLDER"), "displayName": body["displayName"]}, 201)
        return j({"value": [
            {"id": FOLDERS["inbox"], "displayName": "Inbox", "unreadItemCount": 1, "totalItemCount": 3,
             "childFolderCount": 0},
            {"id": FOLDERS["drafts"], "displayName": "Drafts", "unreadItemCount": 0, "totalItemCount": 1,
             "childFolderCount": 0},
            {"id": PROJECTS, "displayName": "Projects", "unreadItemCount": 0, "totalItemCount": 0,
             "childFolderCount": 1}]})
    if p == f"/me/mailFolders/{PROJECTS}/childFolders":
        return j({"value": [{"id": ALPHA, "displayName": "Alpha", "unreadItemCount": 0, "totalItemCount": 1,
                             "childFolderCount": 0}]})
    if p.startswith("/me/mailFolders/") and p.count("/") == 3:
        name = p.rsplit("/", 1)[1]
        if name in FOLDERS:
            return j({"id": FOLDERS[name], "unreadItemCount": 1, "totalItemCount": 3})
        return j({"id": name})
    if p.startswith("/me/mailFolders/") and p.endswith("/messages"):
        fid = p.split("/")[3]
        fid = FOLDERS.get(fid, fid)
        rows = [m for m in MAILBOX.values() if m["parentFolderId"] == fid]
        if "$search" in q:
            rows = search(rows, q["$search"])
        else:
            assert q["$filter"].startswith("receivedDateTime ge "), q  # Graph's $orderby rule
            assert q["$orderby"] == "receivedDateTime desc"
            rows = apply_filter(rows, q["$filter"])
        rows.sort(key=lambda m: m["receivedDateTime"], reverse=True)
        return j({"value": [public(m) for m in rows]})
    if p == "/me/messages" and req.method == "POST":
        m = mk(new_id("DRAFT"), "2026-11-14T08:00:00Z", ME, body.get("subject", ""), "", to=(),
               folder="drafts", draft=True)
        m.update({k: v for k, v in body.items() if k in ("toRecipients", "ccRecipients", "bccRecipients",
                                                          "importance")})
        m["body"] = body["body"]
        MAILBOX[m["id"]] = m
        return j(public(m), 201)
    if p == "/me/messages":
        rows = list(MAILBOX.values())
        if "$search" in q:
            rows = search(rows, q["$search"])
            rows.sort(key=lambda m: m["receivedDateTime"], reverse=True)
        if "$filter" in q:
            assert "$orderby" not in q
            rows = apply_filter(rows, q["$filter"])
        return j({"value": [public(m) for m in rows]})
    if p == "/me/sendMail":
        SENT.append(body)
        return httpx.Response(202)
    parts = p.split("/")
    if p.startswith("/me/messages/"):
        mid = parts[3]
        m = MAILBOX.get(mid)
        if m is None:
            return j({"error": {"code": "ErrorItemNotFound", "message": "not found"}}, 404)
        rest = parts[4:]
        if not rest:
            if req.method == "GET":
                return j(public(m))
            if req.method == "PATCH":
                m.update(body)
                return j(public(m))
            if req.method == "DELETE":
                del MAILBOX[mid]
                return httpx.Response(204)
        if rest[0] in ("createReply", "createReplyAll", "createForward"):
            fwd = rest[0] == "createForward"
            d = mk(new_id("DRAFT"), "2026-11-14T08:00:00Z", ME,
                   ("FW: " if fwd else "RE: ") + m["subject"], "", folder="drafts", draft=True,
                   to=() if fwd else (m["from"],),
                   cc=m["ccRecipients"] if rest[0] == "createReplyAll" else ())
            d["body"] = {"contentType": "html", "content":
                         f"<html><head></head><body><hr><b>From:</b> {m['from']['emailAddress']['name']}"
                         f"<br>{m['body']['content']}</body></html>"}
            d["hasAttachments"] = fwd and m["hasAttachments"]
            MAILBOX[d["id"]] = d
            return j(public(d), 201)
        if rest[0] == "send":
            SENT.append(public(m))
            m.update(isDraft=False, parentFolderId=FOLDERS["sentitems"])
            return httpx.Response(202)
        if rest[0] == "move":
            m["parentFolderId"] = FOLDERS.get(body["destinationId"], body["destinationId"])
            return j(public(m), 201)
        if rest[0] == "attachments":
            if len(rest) == 1:
                if req.method == "POST":
                    m["_atts"].append({**body, "id": new_id("ATT"), "size": 10})
                    m["hasAttachments"] = True
                    return j(body, 201)
                return j({"value": [{k: v for k, v in a.items() if k != "_data"} for a in m["_atts"]]})
            a = next((a for a in m["_atts"] if a["id"] == rest[1]), None)
            if a is None:
                return j({"error": {"code": "ErrorItemNotFound", "message": "no attachment"}}, 404)
            if len(rest) == 3 and rest[2] == "$value":
                return httpx.Response(200, content=a["_data"])
            return j({k: v for k, v in a.items() if k != "_data"})
    return j({"error": {"code": "NotFound", "message": p}}, 404)


def token(req: httpx.Request) -> httpx.Response:
    form = dict(parse_qsl(req.content.decode()))
    assert form["client_secret"] == "test-secret"
    return j({"access_token": "AT", "expires_in": 3600, "refresh_token": "RT"})


server.tokens._http = httpx.AsyncClient(transport=httpx.MockTransport(token))
server.graph._http = httpx.AsyncClient(base_url="https://graph.microsoft.com/v1.0",
                                       transport=httpx.MockTransport(graph))
reset_mailbox()


async def _call(tool, handler=None, mode="auto", **args):
    kw = {"elicitation_handler": handler} if handler else {}
    async with Client(server.mcp, mode=mode, **kw) as c:
        res = await c.call_tool(tool, args, raise_on_error=False)
    text = res.content[0].text
    if VERBOSE:
        print(f"\n--- {tool}({', '.join(f'{k}={v!r}' for k, v in args.items())})\n{text}")
    return text


def call(tool, handler=None, mode="auto", **args):
    return asyncio.run(_call(tool, handler, mode, **args))


# ------------------------------------------------------------------ tests
def test_whoami():
    out = call("whoami")
    assert out.startswith("Signed in as Me <me@contoso.example>"), out
    assert "Inbox: 1 unread of 3" in out and "time zone Arabian Standard Time" in out, out
    assert "send confirmation: client" in out and "send only to contoso.example, partner.example" in out, out
    assert "signature is appended" in out, out


def test_list_messages_inbox():
    requests.clear()
    out = call("list_messages")
    lines = out.splitlines()
    assert lines[0] == "3 messages in Inbox, newest first (times UTC)", out
    assert lines[1] == "2026-11-13", out
    assert lines[2] == "10:01 Alice Johnson <alice@contoso.example> · Q4 numbers · UNREAD !high 📎", out
    # Safe Links are unwrapped in the preview, inline image refs shortened.
    assert "https://example.com/report?q=1" in lines[3] and "safelinks" not in out, lines[3]
    assert "[image]" in lines[3], lines[3]
    assert lines[4] == "  id MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA", out


def test_list_messages_filters_and_paging():
    requests.clear()
    out = call("list_messages", unread_only=True, since="2026-11-13")
    flt = [r for r in requests if r[1].endswith("/messages")][0][2]["$filter"]
    assert flt == "receivedDateTime ge 2026-11-13T00:00:00Z and isRead eq false", flt
    assert out.startswith("1 messages in Inbox since 2026-11-13T00:00Z"), out
    out = call("list_messages", limit=2, preview_chars=0)
    assert out.endswith("(more messages match: raise limit, or pass until=2026-11-13T09:00:00Z "
                        "for the next page)"), out
    out = call("list_messages", until="2026-11-13T09:00:00Z")
    assert "Newsletter" in out and "Vendor call" not in out, out
    assert "sender must be an email" in call("list_messages", sender="Alice")


def test_list_messages_sent_shows_recipients_and_folder_paths():
    out = call("list_messages", folder="sent")
    assert "11:00 → Alice Johnson · RE: Q4 numbers" in out, out
    out = call("list_messages", folder="Projects/Alpha")
    assert out.startswith("1 messages in Projects/Alpha") and "Alpha kickoff" in out, out
    assert "No folder 'Nope'" in call("list_messages", folder="Nope")


def test_list_folders():
    out = call("list_folders")
    assert "• Inbox · 1 unread / 3" in out and "  • Alpha · 0 unread / 1 · id " + ALPHA in out, out


def test_search_by_sender_and_keyword():
    requests.clear()
    out = call("search_messages", sender="Bob Smith", query="budget", since="2026-11-01")
    q = [r for r in requests if r[1] == "/me/messages"][0][2]["$search"]
    assert q == '"budget AND from:\\"Bob Smith\\" AND received>=2026-11-01"', q
    assert out.splitlines()[0].startswith("1 results for budget AND from:"), out
    assert "Newsletter" in out and "in Inbox" in out, out


def test_search_to_me_filters_on_my_addresses():
    out = call("search_messages", sender="bob", to_me=True)
    # The vendor-call mail has me in Cc; the newsletter went to a list.
    assert "Vendor call" in out and "Newsletter" not in out, out
    assert "Give a query" in call("search_messages")


def test_get_message_full():
    out = call("get_message", message_id="MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA")
    assert "From: Alice Johnson <alice@contoso.example>" in out and "To: Me <me@contoso.example>" in out
    assert "Attachments (3, plus 1 inline):" in out, out
    assert "- Notes.docx · " in out and "attachment_id ATT-DOCX-000000000001" in out, out
    assert "Hi,\n\nCan you check the Q4 numbers?\nReport: https://example.com/report?q=1" in out, out
    assert "Open in Outlook: https://outlook.office.com/mail/" in out
    out = call("get_message", message_id="MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA", max_chars=200, offset=10)
    assert "Hi," not in out
    assert "doesn't look like an Outlook id" in call("get_message", message_id="abc")


def test_get_message_mark_read():
    reset_mailbox()
    out = call("get_message", message_id="MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA", mark_read=True)
    assert "(marked as read)" in out and "UNREAD" not in out, out
    assert MAILBOX["MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA"]["isRead"] is True
    reset_mailbox()


def test_get_thread_oldest_first_new_text_only():
    out = call("get_thread", message_id="MSG-0004-DDDDDDDDDDDDDDDDDDDDDDDD")
    blocks = out.split("\n\n")
    assert out.startswith("Conversation: RE: Q4 numbers · 2 messages"), out
    assert "── 2026-11-13 10:01 Alice Johnson" in out and "── 2026-11-13 11:00 Me" in out
    assert out.index("10:01 Alice") < out.index("11:00 Me"), out
    assert "half-written" not in out  # drafts excluded
    assert blocks[-1].endswith("Will do."), blocks[-1]


def test_get_attachment_formats():
    mid = "MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA"
    out = call("get_attachment", message_id=mid, attachment_id="ATT-DOCX-000000000001")
    assert "read as docx" in out and "Budget is 40k\nDeadline Friday" in out, out
    out = call("get_attachment", message_id=mid, attachment_id="ATT-XLSX-000000000002")
    assert "--- sheet Q4\nRegion\tRevenue\nEMEA\t\t40000" in out, out
    out = call("get_attachment", message_id=mid, attachment_id="ATT-PDF-0000000000003")
    assert "--- page 1" in out and "Termination notice is 90 days" in out, out
    out = call("get_attachment", message_id=mid, attachment_id="ATT-PNG-0000000000004")
    assert "binary format" in out, out


def test_find_people_falls_back_to_mail_history():
    out = call("find_people", query="bob")
    assert "• Bob Smith <bob@partner.example>" in out and "(mail history)" in out, out
    assert "People.Read and User.ReadBasic.All" not in out  # found something, so no note


def test_create_draft_markdown_with_signature():
    reset_mailbox()
    out = call("create_draft", to="Alice Johnson <alice@contoso.example>; bob@partner.example",
               subject="Plan", body="**Agenda**\n\n- one\n- two", format="markdown")
    assert out.startswith("Draft saved (not sent). draft_id DRAFT-"), out
    assert "To: Alice Johnson <alice@contoso.example>; bob@partner.example" in out, out
    d = [m for m in MAILBOX.values() if m["subject"] == "Plan"][0]
    html = d["body"]["content"]
    assert "<strong>Agenda</strong>" in html and "<li>one</li>" in html, html
    assert "Antonio Rivas<br>" in html, html
    out = call("create_draft", to="Alice", subject="x", body="y")
    assert "is not an email address" in out, out


def test_reply_draft_keeps_quoted_original():
    reset_mailbox()
    out = call("create_reply_draft", message_id="MSG-0002-BBBBBBBBBBBBBBBBBBBBBBBB",
               body="Thursday works.\nSee you", reply_all=True, cc="carol@contoso.example")
    assert out.startswith("Reply draft saved (not sent)"), out
    assert "To: Bob Smith <bob@partner.example>" in out and "carol@contoso.example" in out, out
    d = [m for m in MAILBOX.values() if m["subject"] == "RE: Vendor call"][0]
    html = d["body"]["content"]
    assert html.index("Thursday works.<br>") < html.index("<b>From:</b> Bob Smith"), html


def test_update_and_delete_draft():
    reset_mailbox()
    out = call("update_draft", draft_id="MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE", subject="RE: Q4 numbers (v2)",
               attachments=[{"name": "data.csv", "content": "a,b\n1,2"}])
    assert "Draft updated" in out and "Attachments added: 1" in out, out
    out = call("update_draft", draft_id="MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA", subject="x")
    assert "not a draft" in out, out
    out = call("delete_draft", draft_id="MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE")
    assert out == "Draft deleted: RE: Q4 numbers (v2)", out
    assert "MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE" not in MAILBOX


def test_send_mail_respects_allowed_domains():
    reset_mailbox()
    out = call("send_mail", to=["alice@contoso.example"], cc="eve@gmail.example", subject="Hi", body="Hello")
    assert "not allowed: eve@gmail.example" in out and not SENT, out
    out = call("send_mail", to=["alice@contoso.example"], subject="Hi", body="Hello")
    assert out.startswith("Sent at ") and "'Hi' to alice@contoso.example" in out, out
    assert SENT[0]["saveToSentItems"] is True and SENT[0]["message"]["subject"] == "Hi"


def test_reply_and_send_draft():
    reset_mailbox()
    out = call("reply", message_id="MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA", body="On it")
    assert out.startswith("Sent at ") and "'RE: Q4 numbers' to Alice Johnson <alice@contoso.example>" in out
    assert SENT and "On it" in SENT[0]["body"]["content"]
    # A forward to an outside domain is refused but kept as a draft.
    out = call("forward", message_id="MSG-0001-AAAAAAAAAAAAAAAAAAAAAAAA", to="eve@gmail.example")
    assert "not allowed" in out and "saved as draft DRAFT-" in out and len(SENT) == 1, out
    out = call("send_draft", draft_id="MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE")
    assert out.startswith("Sent at ") and len(SENT) == 2, out
    assert "not a draft" in call("send_draft", draft_id="MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE")


def test_send_confirmation_elicit():
    reset_mailbox()
    server.SEND_CONFIRMATION = "elicit"
    prompts = []

    async def approve(message, response_type, params, ctx):
        prompts.append(message)
        return "send"

    async def decline(message, response_type, params, ctx):
        from fastmcp.client.elicitation import ElicitResult
        return ElicitResult(action="decline")

    try:
        out = call("send_draft", decline, draft_id="MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE")
        assert "did not approve" in out and "draft MSG-0005" in out and not SENT, out
        try:  # a client that can't show prompts: the call fails, nothing is sent
            out = call("send_draft", draft_id="MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE")
            assert "NOT sent" in out, out
        except Exception as e:
            assert "Elicitation not supported" in str(e), e
        assert not SENT
        out = call("send_draft", approve, draft_id="MSG-0005-EEEEEEEEEEEEEEEEEEEEEEEE")
        assert out.startswith("Sent at ") and len(SENT) == 1, out
        assert prompts[0].startswith("Send this email as you?\nTo: Alice Johnson <alice@contoso.example>"), prompts
    finally:
        server.SEND_CONFIRMATION = "client"


def test_send_confirmation_elicit_reply_makes_one_draft():
    reset_mailbox()
    server.SEND_CONFIRMATION = "elicit"

    async def approve(message, response_type, params, ctx):
        assert message.startswith("Send this reply as you?\nTo: Bob Smith"), message
        return "send"

    try:
        before = len(MAILBOX)
        out = call("reply", approve, message_id="MSG-0002-BBBBBBBBBBBBBBBBBBBBBBBB", body="Thursday works")
        assert out.startswith("Sent at ") and len(SENT) == 1, out
        assert len(MAILBOX) == before + 1, "the approval leg must reuse the first leg's draft"
    finally:
        server.SEND_CONFIRMATION = "client"


def test_send_confirmation_elicit_legacy_protocol():
    reset_mailbox()
    server.SEND_CONFIRMATION = "elicit"

    async def decline(message, response_type, params, ctx):
        from fastmcp.client.elicitation import ElicitResult
        return ElicitResult(action="decline")

    async def approve(message, response_type, params, ctx):
        return "send"

    try:
        out = call("send_mail", decline, mode="legacy", to="alice@contoso.example", subject="Hi", body="x")
        assert "did not approve" in out and not SENT, out
        out = call("send_mail", approve, mode="legacy", to="alice@contoso.example", subject="Hi", body="x")
        assert out.startswith("Sent at ") and len(SENT) == 1, out
        out = call("send_mail", mode="legacy", to="alice@contoso.example", subject="Hi", body="x")
        assert "NOT sent" in out and len(SENT) == 1, out
    finally:
        server.SEND_CONFIRMATION = "client"


def test_organise_tools():
    reset_mailbox()
    ids = ["MSG-0002-BBBBBBBBBBBBBBBBBBBBBBBB", "MSG-0003-CCCCCCCCCCCCCCCCCCCCCCCC"]
    assert call("mark_read", message_ids=ids, read=False) == "Marked unread: 2 of 2 messages"
    assert MAILBOX[ids[0]]["isRead"] is False
    out = call("flag_messages", message_ids=ids[0], due="2026-11-20")
    assert out == "Flag set to flagged: 1 of 1 messages" and MAILBOX[ids[0]]["flag"]["dueDateTime"]
    assert call("set_categories", message_ids=ids, categories="Finance") == \
        "Categories added: 2 of 2 messages"
    assert MAILBOX[ids[1]]["categories"] == ["Finance"]
    out = call("move_messages", message_ids=ids, folder="archive")
    assert out == "Moved 2 of 2 messages to Archive" and MAILBOX[ids[0]]["parentFolderId"] == FOLDERS["archive"]
    out = call("create_folder", name="Invoices", parent="Projects")
    assert out.startswith("Folder 'Invoices' created under Projects"), out
    reset_mailbox()


def test_compose_helpers():
    r = compose.parse_recipients("a@x.com, \"Bo, B\" <b@y.org>; Smith, John <j@z.net>; A@x.com", "to")
    assert compose.addresses(r) == ["a@x.com", "b@y.org", "j@z.net"], r
    assert r[1]["emailAddress"]["name"] == "Bo, B" and r[2]["emailAddress"]["name"] == "Smith, John", r
    try:
        compose.parse_recipients("John Smith", "to")
        raise AssertionError("a plain name must be refused")
    except compose.ComposeError:
        pass
    assert compose.parse_recipients(["Bo B <b@y.org>"], "to") == [
        {"emailAddress": {"address": "b@y.org", "name": "Bo B"}}]
    assert "&lt;b&gt;" in compose.to_html("<b>", "text")
    assert fmt.tidy("a\n\n\n\nb  \n\n") == "a\n\nb"


if __name__ == "__main__":
    tests = [(n, f) for n, f in sorted(globals().items()) if n.startswith("test_")]
    for name, fn in tests:
        fn()
        print(f"ok  {name}")
    print(f"{len(tests)} passed")
