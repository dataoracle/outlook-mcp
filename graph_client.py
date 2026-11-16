"""Minimal Microsoft Graph client using delegated tokens from TokenManager."""
from __future__ import annotations

import asyncio
import logging

import httpx

from auth import TokenManager

log = logging.getLogger("outlook-mcp.graph")

GRAPH = "https://graph.microsoft.com/v1.0"

# Every request asks for immutable ids: a message keeps its id when it is moved
# to another folder (or sent from Drafts), so ids handed to the agent stay valid.
PREFER_IDS = 'IdType="ImmutableId"'
# Reads ask Graph to convert HTML bodies to plain text (cheaper for the model).
PREFER_TEXT = 'outlook.body-content-type="text"'


class GraphError(Exception):
    """A Graph call failed in a way worth explaining to the user."""

    def __init__(self, message: str, status: int = 0, code: str = ""):
        super().__init__(message)
        self.status = status
        self.code = code


class GraphClient:
    def __init__(self, tokens: TokenManager, timeout: float = 30,
                 transport: httpx.AsyncBaseTransport | None = None):
        self.tokens = tokens
        self._http = httpx.AsyncClient(
            base_url=GRAPH,
            headers={"Accept": "application/json"},
            timeout=timeout,
            follow_redirects=False,
            transport=transport,
        )

    async def aclose(self) -> None:
        await self._http.aclose()

    async def _send(self, method: str, path: str, *, params: dict | None = None,
                    json: dict | None = None, headers: dict | None = None,
                    text_body: bool = False) -> httpx.Response:
        # Absolute URLs only come from @odata.nextLink; never send the token elsewhere.
        if path.startswith("http") and not path.startswith(GRAPH + "/"):
            raise GraphError(f"Refusing to follow a link outside Microsoft Graph: {path[:100]}")
        prefer = PREFER_IDS + (", " + PREFER_TEXT if text_body else "")
        force, retried_401 = False, False
        for attempt in range(4):
            token = await self.tokens.access_token(force=force)
            force = False
            try:
                r = await self._http.request(
                    method, path, params=params, json=json,
                    headers={"Prefer": prefer, **(headers or {}), "Authorization": f"Bearer {token}"},
                )
            except httpx.HTTPError as e:
                raise GraphError(f"Could not reach Microsoft Graph: {e}") from e
            if r.status_code == 401 and not retried_401:
                retried_401, force = True, True
                continue
            if r.status_code in (429, 503, 504) and attempt < 3:
                try:
                    delay = float(r.headers.get("Retry-After", "2"))
                except ValueError:
                    delay = 2.0
                log.warning("Graph %s on %s, retrying in %.0fs", r.status_code, path[:80], delay)
                await asyncio.sleep(min(max(delay, 1.0), 15.0))
                continue
            break
        if r.status_code >= 400:
            raise self._error(r, method, path)
        return r

    async def request(self, method: str, path: str, *, params: dict | None = None,
                      json: dict | None = None, headers: dict | None = None,
                      text_body: bool = False) -> dict:
        r = await self._send(method, path, params=params, json=json, headers=headers,
                             text_body=text_body)
        if r.status_code in (202, 204) or not r.content:
            return {}
        return r.json()

    @staticmethod
    def _error(r: httpx.Response, method: str, path: str) -> GraphError:
        try:
            err = r.json().get("error") or {}
        except ValueError:
            err = {}
        code = err.get("code") or ""
        msg = (err.get("message") or r.text[:300] or "").strip()
        where = f"{method} {path.split('?')[0][:120]}"
        if r.status_code == 401:
            text = f"401 Unauthorized on {where}: {msg}"
        elif r.status_code == 403:
            text = (f"403 Forbidden on {where}: {msg} (the signed-in user, or the app's "
                    f"delegated permissions, don't allow this)")
        elif r.status_code == 404:
            text = (f"404 Not found on {where}: {msg} (check the id: the message may have "
                    f"been deleted, or the id was cut off)")
        else:
            text = f"Microsoft Graph returned {r.status_code} {code} on {where}: {msg}"
        return GraphError(text, r.status_code, code)

    async def get(self, path: str, params: dict | None = None, headers: dict | None = None,
                  text_body: bool = False) -> dict:
        return await self.request("GET", path, params=params, headers=headers, text_body=text_body)

    async def post(self, path: str, json: dict | None = None, text_body: bool = False) -> dict:
        return await self.request("POST", path, json=json if json is not None else {},
                                  text_body=text_body)

    async def patch(self, path: str, json: dict) -> dict:
        return await self.request("PATCH", path, json=json)

    async def delete(self, path: str) -> None:
        await self.request("DELETE", path)

    async def get_bytes(self, path: str, max_bytes: int) -> bytes:
        """Download raw content (e.g. an attachment's $value), refusing anything larger
        than max_bytes."""
        r = await self._send("GET", path, headers={"Accept": "*/*"})
        data = r.content
        if len(data) > max_bytes:
            raise GraphError(f"The content is {len(data) // 1024} KB, over the "
                             f"{max_bytes // 1024} KB limit.", 413, "TooLarge")
        return data

    async def batch(self, requests: list[dict]) -> list[dict]:
        """JSON $batch, 20 requests per call. Each request: {method, url, body?}.
        Returns the responses in request order: {status, body}."""
        out: list[dict] = []
        for i in range(0, len(requests), 20):
            chunk = requests[i:i + 20]
            payload = {"requests": [
                {"id": str(n), "method": r["method"], "url": r["url"],
                 "headers": {"Prefer": PREFER_IDS,
                             **({"Content-Type": "application/json"} if r.get("body") is not None else {})},
                 **({"body": r["body"]} if r.get("body") is not None else {})}
                for n, r in enumerate(chunk)]}
            res = await self.request("POST", "/$batch", json=payload)
            by_id = {x.get("id"): x for x in res.get("responses") or []}
            for n in range(len(chunk)):
                x = by_id.get(str(n)) or {}
                out.append({"status": int(x.get("status") or 0), "body": x.get("body") or {}})
        return out

    async def get_all(self, path: str, params: dict | None = None, limit: int = 100,
                      headers: dict | None = None, text_body: bool = False) -> tuple[list[dict], bool]:
        """Follow @odata.nextLink until `limit` items. Returns (items, more_available)."""
        items: list[dict] = []
        url: str | None = path
        while url and len(items) < limit:
            page = await self.request("GET", url, params=params, headers=headers, text_body=text_body)
            items.extend(page.get("value") or [])
            url = page.get("@odata.nextLink")
            params = None  # the next link already carries the query
        return items[:limit], bool(url) or len(items) > limit

    async def scan(self, path: str, params: dict | None = None, limit: int = 100, *,
                   keep=None, stop=None, max_scan: int = 1000,
                   headers: dict | None = None, text_body: bool = False) -> tuple[list[dict], str]:
        """Page through a collection in server order, keeping items that pass
        `keep`, until one of these happens (the second value says which):
          'stop'  - `stop(item)` was true (e.g. an ordered cutoff was passed);
          'limit' - `limit` items were kept and at least one more would be kept;
          'scan'  - `max_scan` items were examined (more may exist);
          'end'   - the collection ran out."""
        kept: list[dict] = []
        seen = 0
        url: str | None = path
        while url:
            page = await self.request("GET", url, params=params, headers=headers, text_body=text_body)
            params = None  # the next link already carries the query
            for item in page.get("value") or []:
                if stop and stop(item):
                    return kept, "stop"
                if keep is None or keep(item):
                    if len(kept) >= limit:
                        return kept, "limit"
                    kept.append(item)
                seen += 1
                if seen >= max_scan:
                    return kept, "scan"
            url = page.get("@odata.nextLink")
        return kept, "end"
