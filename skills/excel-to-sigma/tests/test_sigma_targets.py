#!/usr/bin/env python3
"""migrate.py resolve-targets: folder and write-enabled connection looked up by name (read-only)."""
from __future__ import annotations

import contextlib
import importlib.util
import io
import json
import sys
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "scripts" / "lib"))

import sigma_targets as st  # noqa: E402

spec = importlib.util.spec_from_file_location("migrate_targets", ROOT / "scripts" / "migrate.py")
MIGRATE = importlib.util.module_from_spec(spec)
spec.loader.exec_module(MIGRATE)

FOLDERS = [
    {"id": "f-1", "name": "Excel Demo — Rehearsal", "type": "folder", "path": "My Documents"},
    {"id": "f-2", "name": "Excel Demo", "type": "folder", "path": "My Documents"},
    {"id": "f-3", "name": "Shared", "type": "folder", "path": "Workspaces"},
]
CONNECTIONS = [
    {"connectionId": "c-ro", "name": "OAuth Snowflake", "type": "snowflake", "writeAccess": False},
    {"connectionId": "c-rw", "name": "Snowflake ymb68310", "type": "snowflake", "writeAccess": True},
]


def fake_api(pages):
    """Serve each path's entries over two pages, so pagination is exercised; record every call."""
    calls = []

    def api(method, path, body=None, base_url=None):
        calls.append((method, path))
        assert method == "GET" and body is None, (method, body)
        key = "/v2/files" if path.startswith("/v2/files") else "/v2/connections"
        entries = pages[key]
        if "page=p2" in path:
            return {"entries": entries[2:], "hasMore": False, "nextPage": None}
        return {"entries": entries[:2], "hasMore": len(entries) > 2, "nextPage": "p2" if len(entries) > 2 else None}
    return api, calls


class ResolveTargetsTests(unittest.TestCase):
    def setUp(self):
        self.api, self.calls = fake_api({"/v2/files": FOLDERS, "/v2/connections": CONNECTIONS})

    def test_exact_names_resolve_to_ids_with_read_only_calls(self):
        out = st.resolve_targets(self.api, folder_name="excel demo — rehearsal", connection_name="Snowflake ymb68310")
        self.assertEqual(out["folder"], {"id": "f-1", "name": "Excel Demo — Rehearsal", "path": "My Documents"})
        self.assertEqual(out["connection"], {"id": "c-rw", "name": "Snowflake ymb68310", "type": "snowflake",
                                             "writeAccess": True})
        self.assertTrue(all(m == "GET" for m, _ in self.calls))
        self.assertTrue(any("page=p2" in p for _, p in self.calls))   # walked every page
        self.assertTrue(any("typeFilters=folder" in p for _, p in self.calls))

    def test_connection_must_be_write_enabled(self):
        with self.assertRaisesRegex(st.TargetError, "OAuth Snowflake.*not write-enabled"):
            st.resolve_targets(self.api, connection_name="OAuth Snowflake")

    def test_missing_or_ambiguous_names_list_candidates(self):
        with self.assertRaisesRegex(st.TargetError, "no folder named 'Excel'.*Excel Demo — Rehearsal"):
            st.resolve_targets(self.api, folder_name="Excel")
        twin = FOLDERS + [{"id": "f-9", "name": "Excel Demo", "type": "folder", "path": "Workspaces/FP&A"}]
        api, _ = fake_api({"/v2/files": twin, "/v2/connections": CONNECTIONS})
        with self.assertRaisesRegex(st.TargetError, "2 folders named 'Excel Demo'.*f-2.*f-9"):
            st.resolve_targets(api, folder_name="Excel Demo")

    def test_cli_prints_ids_and_the_draft_flags(self):
        stdout = io.StringIO()
        with mock.patch.object(st, "sigma_api", self.api), contextlib.redirect_stdout(stdout):
            code = MIGRATE.main(["resolve-targets", "--folder-name", "Excel Demo — Rehearsal",
                                 "--connection-name", "Snowflake ymb68310"])
        self.assertEqual(code, 0)
        result = json.loads(stdout.getvalue())
        self.assertEqual((result["folderId"], result["connectionId"]), ("f-1", "c-rw"))
        self.assertEqual(result["draftFlags"], "--connection c-rw --folder f-1")
        err = io.StringIO()
        with mock.patch.object(st, "sigma_api", self.api), contextlib.redirect_stdout(io.StringIO()), \
                contextlib.redirect_stderr(err):
            self.assertEqual(MIGRATE.main(["resolve-targets", "--connection-name", "OAuth Snowflake"]), 1)
        self.assertIn("not write-enabled", err.getvalue())

    def test_router_names_resolve_targets_for_the_folder_and_connection(self):
        route = MIGRATE.ROUTES["forecast"]
        self.assertTrue(any("resolve-targets" in d for d in route["decisions"]), route["decisions"])


if __name__ == "__main__":
    unittest.main()
