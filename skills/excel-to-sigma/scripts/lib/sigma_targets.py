"""Resolve the target folder and write-enabled connection by name (read-only Sigma REST GETs).

``migrate.py resolve-targets --folder-name NAME --connection-name NAME`` prints the ids the
drafter takes (``--folder`` / ``--connection``), so nobody pipes ``sigma api files list`` into an
ad hoc filter. Names match exactly, ignoring case; a missing or ambiguous name fails and lists the
candidates. A connection must be write-enabled (input tables write back through it).
"""
from __future__ import annotations

import urllib.parse

import sigma_auth

MAX_PAGES = 50


class TargetError(ValueError):
    """A folder or connection name that does not resolve to exactly one usable target."""


def sigma_api(method, path, body=None, base_url=None):
    return sigma_auth.api(method, path, body, base_url=base_url)


def _list(api, path, params):
    entries, page = [], None
    for _ in range(MAX_PAGES):
        query = dict(params, **({"page": page} if page else {}))
        data = api("GET", f"{path}?{urllib.parse.urlencode(query)}") or {}
        entries += data.get("entries") or []
        page = data.get("nextPage")
        if not (data.get("hasMore") and page):
            return entries
    raise TargetError(f"{path}: more than {MAX_PAGES} pages; narrow the search")


def _one(kind, name, entries, label):
    want = name.casefold().strip()
    hits = [e for e in entries if str(e.get("name", "")).casefold().strip() == want]
    if len(hits) == 1:
        return hits[0]
    if hits:
        raise TargetError(f"{len(hits)} {kind}s named {name!r}: "
                          + "; ".join(label(e) for e in hits) + " — pass the id instead")
    near = [e for e in entries if want in str(e.get("name", "")).casefold()]
    raise TargetError(f"no {kind} named {name!r}"
                      + (": did you mean " + "; ".join(label(e) for e in near[:10]) + "?" if near else ""))


def find_folder(api, name):
    folders = _list(api, "/v2/files", {"typeFilters": "folder", "limit": 1000})
    hit = _one("folder", name, [f for f in folders if f.get("type", "folder") == "folder"],
               lambda e: f"{e.get('name')} ({e.get('id')}, in {e.get('path') or '?'})")
    return {"id": hit["id"], "name": hit["name"], "path": hit.get("path")}


def find_connection(api, name):
    connections = _list(api, "/v2/connections", {"limit": 1000})
    hit = _one("connection", name, connections,
               lambda e: f"{e.get('name')} ({e.get('connectionId')}, {e.get('type') or '?'})")
    if hit.get("writeAccess") is not True:
        raise TargetError(f"connection {hit['name']!r} ({hit['connectionId']}) is not write-enabled "
                          "(writeAccess is not true); input tables need a write-enabled connection")
    return {"id": hit["connectionId"], "name": hit["name"], "type": hit.get("type"), "writeAccess": True}


def resolve_targets(api, folder_name=None, connection_name=None):
    out = {}
    if folder_name:
        out["folder"] = find_folder(api, folder_name)
    if connection_name:
        out["connection"] = find_connection(api, connection_name)
    return out
