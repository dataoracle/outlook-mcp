"""Extract readable text from mail attachments.

Supported: plain text formats (txt, csv, json, xml, md, ics, ...), HTML, PDF
(pypdf), Word .docx, Excel .xlsx, PowerPoint .pptx (OOXML read with the
standard library) and attached .eml messages. Anything else is reported as
binary. Attachments are untrusted input: archive members are read with a
size cap, and nothing is ever written to disk.
"""
from __future__ import annotations

import email
import email.policy
import io
import re
import zipfile
from xml.etree import ElementTree as ET

from formatting import html_to_text, tidy

MAX_XML_BYTES = 40 * 1024 * 1024   # uncompressed size cap per OOXML part (zip-bomb guard)
MAX_XLSX_ROWS = 2000               # rows per sheet

_TEXT_EXT = {"txt", "csv", "tsv", "json", "xml", "md", "log", "ics", "vcf", "yaml", "yml", "ini",
             "cfg", "conf", "sql", "py", "js", "ts", "sh", "ps1", "java", "cs", "go", "rb"}


class ExtractError(Exception):
    pass


def _ext(name: str) -> str:
    return name.rsplit(".", 1)[-1].lower() if "." in (name or "") else ""


def _decode(data: bytes) -> str:
    for enc in ("utf-8-sig", "utf-16") if data[:2] in (b"\xff\xfe", b"\xfe\xff") else ("utf-8-sig",):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            pass
    return data.decode("cp1252", errors="replace")


def _read_member(z: zipfile.ZipFile, name: str) -> bytes:
    info = z.getinfo(name)
    if info.file_size > MAX_XML_BYTES:
        raise ExtractError(f"{name} inside the file is too large to read ({info.file_size} bytes).")
    with z.open(info) as f:
        data = f.read(MAX_XML_BYTES + 1)
    if len(data) > MAX_XML_BYTES:
        raise ExtractError(f"{name} inside the file is too large to read.")
    return data


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _docx(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        root = ET.fromstring(_read_member(z, "word/document.xml"))
    paras: list[str] = []
    for p in root.iter():
        if _local(p.tag) != "p":
            continue
        parts = []
        for el in p.iter():
            t = _local(el.tag)
            if t == "t" and el.text:
                parts.append(el.text)
            elif t == "tab":
                parts.append("\t")
            elif t in ("br", "cr"):
                parts.append("\n")
        paras.append("".join(parts))
    return tidy("\n".join(paras))


def _pptx(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        slides = sorted((n for n in z.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)),
                        key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[-1]).group(1)))
        out = []
        for i, name in enumerate(slides, 1):
            root = ET.fromstring(_read_member(z, name))
            paras = []
            for p in root.iter():
                if _local(p.tag) == "p":
                    text = "".join(el.text or "" for el in p.iter() if _local(el.tag) == "t")
                    if text.strip():
                        paras.append(text)
            out.append(f"--- slide {i}\n" + "\n".join(paras))
    return tidy("\n".join(out))


def _col_index(ref: str) -> int:
    n = 0
    for ch in re.match(r"[A-Z]*", ref).group(0):
        n = n * 26 + (ord(ch) - 64)
    return max(n - 1, 0)


def _xlsx(data: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        names = z.namelist()
        shared: list[str] = []
        if "xl/sharedStrings.xml" in names:
            for si in ET.fromstring(_read_member(z, "xl/sharedStrings.xml")):
                shared.append("".join(el.text or "" for el in si.iter() if _local(el.tag) == "t"))
        # Sheet names in workbook order, mapped to their part via the relationships.
        wb = ET.fromstring(_read_member(z, "xl/workbook.xml"))
        rels = {}
        if "xl/_rels/workbook.xml.rels" in names:
            for r in ET.fromstring(_read_member(z, "xl/_rels/workbook.xml.rels")):
                rels[r.get("Id")] = r.get("Target", "")
        sheets = []
        for s in wb.iter():
            if _local(s.tag) == "sheet":
                rid = next((v for k, v in s.attrib.items() if _local(k) == "id"), None)
                target = rels.get(rid, "")
                target = target.lstrip("/") if target.startswith("/") else "xl/" + target
                sheets.append((s.get("name"), target))
        out = []
        for sheet_name, part in sheets:
            if part not in names:
                continue
            root = ET.fromstring(_read_member(z, part))
            rows, cut = [], False
            for row in root.iter():
                if _local(row.tag) != "row":
                    continue
                if len(rows) >= MAX_XLSX_ROWS:
                    cut = True
                    break
                cells: dict[int, str] = {}
                for c in row:
                    if _local(c.tag) != "c":
                        continue
                    kind, value = c.get("t"), None
                    for el in c:
                        if _local(el.tag) == "v":
                            value = el.text
                        elif _local(el.tag) == "is":
                            value = "".join(x.text or "" for x in el.iter() if _local(x.tag) == "t")
                    if value is None:
                        continue
                    if kind == "s":
                        try:
                            value = shared[int(value)]
                        except (ValueError, IndexError):
                            pass
                    cells[_col_index(c.get("r", "A"))] = value
                if cells:
                    rows.append("\t".join(cells.get(i, "") for i in range(max(cells) + 1)).rstrip())
            out.append(f"--- sheet {sheet_name}" + (f" (first {MAX_XLSX_ROWS} rows)" if cut else "")
                       + "\n" + "\n".join(rows))
    return "\n".join(out)


def _pdf(data: bytes) -> str:
    try:
        from pypdf import PdfReader
    except ImportError as e:  # pragma: no cover
        raise ExtractError("PDF support needs the pypdf package.") from e
    try:
        reader = PdfReader(io.BytesIO(data))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise ExtractError("The PDF is password-protected.") from None
        pages = []
        for i, page in enumerate(reader.pages, 1):
            pages.append(f"--- page {i}\n" + tidy(page.extract_text() or ""))
    except ExtractError:
        raise
    except Exception as e:
        raise ExtractError(f"Could not read the PDF ({type(e).__name__}).") from None
    text = "\n".join(pages)
    if not re.search(r"\w", re.sub(r"--- page \d+", "", text)):
        return text + "\n(no text layer: probably a scanned document)"
    return text


def _eml(data: bytes) -> str:
    msg = email.message_from_bytes(data, policy=email.policy.default)
    head = [f"{h}: {msg[h]}" for h in ("From", "To", "Cc", "Date", "Subject") if msg[h]]
    part = msg.get_body(preferencelist=("plain", "html"))
    body = ""
    if part is not None:
        content = part.get_content()
        body = html_to_text(content) if part.get_content_type() == "text/html" else tidy(content)
    names = [p.get_filename() for p in msg.iter_attachments() if p.get_filename()]
    if names:
        head.append("Attachments: " + "; ".join(names))
    return "\n".join(head) + "\n\n" + body


def extract_text(name: str, content_type: str | None, data: bytes) -> tuple[str, str]:
    """Return (text, how) for an attachment, or raise ExtractError."""
    ext = _ext(name)
    ct = (content_type or "").lower().split(";")[0].strip()
    try:
        if ext == "pdf" or ct == "application/pdf":
            return _pdf(data), "pdf"
        if ext == "docx" or ct.endswith("wordprocessingml.document"):
            return _docx(data), "docx"
        if ext == "xlsx" or ct.endswith("spreadsheetml.sheet"):
            return _xlsx(data), "xlsx"
        if ext == "pptx" or ct.endswith("presentationml.presentation"):
            return _pptx(data), "pptx"
        if ext == "eml" or ct == "message/rfc822":
            return _eml(data), "eml"
        if ext in ("htm", "html") or ct == "text/html":
            return html_to_text(_decode(data)), "html"
        if ext in _TEXT_EXT or ct.startswith("text/") or ct in ("application/json", "application/xml"):
            return _decode(data), "text"
    except (zipfile.BadZipFile, ET.ParseError, KeyError) as e:
        raise ExtractError(f"{name} looks damaged or isn't really a {ext or ct} file "
                           f"({type(e).__name__}).") from None
    if ext in ("doc", "xls", "ppt"):
        raise ExtractError(f"{name} is a legacy binary Office file (.{ext}); only .docx/.xlsx/.pptx "
                           "can be read.")
    raise ExtractError(f"{name} ({ct or 'unknown type'}) is a binary format this server can't "
                       "turn into text.")
