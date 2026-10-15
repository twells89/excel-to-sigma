#!/usr/bin/env python3
"""Offline smoke for the current workbook contents/dry-run API."""
from __future__ import annotations

import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
LIB = ROOT / "scripts" / "lib"
sys.path.insert(0, str(LIB))

import code_rep  # noqa: E402
import workbook_wire  # noqa: E402


class TestWorkbookWire(unittest.TestCase):
    def test_wire_wraps_contents_and_flattens_elements(self):
        draft = {
            "name": "Forecast Entry",
            "folderId": "00000000-0000-0000-0000-000000000001",
            "description": "test",
            "schemaVersion": 1,
            "pages": [{
                "id": "entryPage",
                "name": "Entry",
                "elements": [
                    {"id": "inputTable", "kind": "input-table", "name": "Entry"},
                    {"id": "spine", "kind": "table", "name": "Spine"},
                ],
            }],
        }
        body = workbook_wire.wire_workbook(draft)

        self.assertIn("contents", body)
        self.assertNotIn("document", body)
        self.assertNotIn("pages", body)
        self.assertNotIn("elements", body)
        self.assertNotIn("layout", body)
        self.assertEqual(body["name"], "Forecast Entry")
        self.assertEqual(body["folderId"], draft["folderId"])

        doc = body["contents"]
        self.assertEqual(doc["schemaVersion"], 1)
        self.assertEqual(doc["kind"], "workbook")
        self.assertEqual([p["id"] for p in doc["pages"]], ["entryPage"])
        self.assertNotIn("elements", doc["pages"][0])
        self.assertEqual([e["id"] for e in doc["elements"]], ["inputTable", "spine"])
        self.assertIn("<Page", doc["layout"])
        self.assertIn('id="entryPage"', doc["layout"])
        self.assertIn('elementId="inputTable"', doc["layout"])
        self.assertIn("<Element", doc["layout"])
        self.assertNotIn("<LayoutElement", doc["layout"])

        # Round-trip through the shared adapter stays in the current envelope.
        again = code_rep.wrap_contents(
            code_rep.document(body), extra=code_rep.metadata(body))
        self.assertEqual(again["contents"]["elements"][0]["id"], "inputTable")

    def test_already_wrapped_is_idempotent_enough(self):
        draft = {
            "name": "X",
            "document": {
                "schemaVersion": 1,
                "kind": "workbook",
                "pages": [{"id": "p1", "name": "P"}],
                "elements": [{"id": "e1", "kind": "table"}],
                "layout": '<Page id="p1"><Element elementId="e1"/></Page>',
            },
        }
        body = workbook_wire.wire_workbook(draft)
        self.assertEqual(body["contents"]["elements"][0]["id"], "e1")
        self.assertIn("<Element", body["contents"]["layout"])

    def test_current_contents_is_idempotent_enough(self):
        body = workbook_wire.wire_workbook({
            "name": "X",
            "folderId": "folder",
            "contents": {
                "schemaVersion": 1,
                "kind": "workbook",
                "pages": [{"id": "p1", "name": "P"}],
                "elements": [{"id": "e1", "kind": "table"}],
                "layout": '<Page id="p1"><Element elementId="e1"/></Page>',
            },
        })
        self.assertEqual(body["contents"]["elements"][0]["id"], "e1")
        self.assertNotIn("document", body)

    def test_dry_run_request_does_not_mutate_create_body(self):
        body = workbook_wire.wire_workbook({
            "name": "X",
            "folderId": "folder",
            "pages": [{"id": "p1", "name": "P", "elements": []}],
        })
        checked = workbook_wire.dry_run_request(body)
        self.assertTrue(checked["dryRun"])
        self.assertNotIn("dryRun", body)

    def test_dry_run_success_contract_has_no_valid_true(self):
        warnings = workbook_wire.assert_dry_run_success({
            "dryRun": True,
            "warnings": [{"summary": "heads up"}],
        })
        self.assertEqual(warnings[0]["summary"], "heads up")
        with self.assertRaises(ValueError):
            workbook_wire.assert_dry_run_success({
                "dryRun": True,
                "valid": False,
                "errors": [{"summary": "bad formula"}],
            })


if __name__ == "__main__":
    unittest.main()
