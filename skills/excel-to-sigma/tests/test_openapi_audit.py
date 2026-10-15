#!/usr/bin/env python3
from __future__ import annotations

import copy
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts" / "lib"))

from openapi_audit import audit_openapi, resolve_ref  # noqa: E402


def fixture_spec():
    contents = {
        "type": "object",
        "required": ["schemaVersion", "kind", "pages", "elements", "layout"],
        "properties": {
            "schemaVersion": {"enum": [1]},
            "kind": {"enum": ["workbook"]},
            "pages": {"type": "array"},
            "elements": {
                "type": "array",
                "items": {
                    "oneOf": [
                        {
                            "properties": {
                                "kind": {"enum": ["input-table"]},
                                "source": {},
                                "inputMode": {"enum": ["edit", "explore", "view"]},
                                "columns": {
                                    "properties": {
                                        "values": {},
                                        "valuesFrom": {},
                                        "range": {},
                                        "maxFileNum": {},
                                        "maxFileSizeMb": {},
                                        "acceptedFileTypes": {},
                                    }
                                },
                            }
                        },
                        {
                            "properties": {
                                "kind": {"enum": ["code", "code-output", "pivot-table",
                                                  "bar-chart", "line-chart", "scatter-chart",
                                                  "control"]},
                                "code": {},
                                "connectionId": {},
                                "output": {},
                            }
                        },
                    ]
                },
            },
            "layout": {"type": "string"},
            "automatedActions": {
                "properties": {
                    "trigger": {"enum": ["schedule", "webhook"]},
                    "effect": {"enum": ["call-agent"]},
                }
            },
            "actions": {
                "properties": {
                    "trigger": {},
                    "effects": {
                        "enum": ["insert-rows", "update-rows", "delete-rows", "call-api",
                                 "call-stored-procedure", "run-python-element"]
                    },
                }
            },
        },
    }
    return {
        "openapi": "3.1.0",
        "info": {"version": "fixture"},
        "paths": {
            "/v2/workbooks": {
                "post": {
                    "requestBody": {
                        "content": {
                            "application/json": {
                                "schema": {
                                    "required": ["folderId", "name", "contents"],
                                    "properties": {
                                        "folderId": {"type": "string"},
                                        "name": {"type": "string"},
                                        "contents": {"$ref": "#/components/schemas/Contents"},
                                        "dryRun": {"type": "boolean"},
                                    },
                                }
                            }
                        }
                    },
                    "responses": {
                        "200": {
                            "content": {"application/json": {"schema": {
                                "properties": {"dryRun": {}, "warnings": {}}
                            }}}
                        },
                        "400": {
                            "content": {"application/json": {"schema": {
                                "properties": {
                                    "dryRun": {}, "valid": {}, "errors": {}, "warnings": {}
                                }
                            }}}
                        },
                    },
                }
            },
            "/v2/workbooks/{workbookId}": {
                "get": {
                    "parameters": [{"name": "includeContents", "in": "query"}],
                    "responses": {"200": {"content": {"application/json": {"schema": {
                        "properties": {"contents": {}, "documentVersion": {}}
                    }}}}},
                }
            },
            "/v2/workbooks/{workbookId}/contents": {
                "put": {
                    "requestBody": {"content": {"application/json": {"schema": {
                        "properties": {"contents": {}, "documentVersion": {}}
                    }}}}
                }
            },
        },
        "components": {
            "schemas": {
                "Contents": contents,
                "Unrelated": {
                    "properties": {"futureMagic": {}},
                    "enum": ["not-reachable"],
                },
            }
        },
    }


class TestOpenApiAudit(unittest.TestCase):
    def test_complete_reachable_contract_passes(self):
        report = audit_openapi(fixture_spec())
        self.assertEqual(report["status"], "PASS", report)
        self.assertEqual(report["missing"], [])

    def test_missing_reachable_field_fails(self):
        spec = fixture_spec()
        del spec["components"]["schemas"]["Contents"]["properties"]["layout"]
        report = audit_openapi(spec)
        self.assertEqual(report["status"], "DRIFT")
        self.assertIn("workbookContents", report["missing"])

    def test_unreferenced_component_does_not_supply_capability(self):
        spec = fixture_spec()
        del spec["components"]["schemas"]["Contents"]["properties"]["actions"]
        spec["components"]["schemas"]["Unrelated"]["enum"] = [
            "insert-rows", "update-rows", "delete-rows", "call-api",
            "call-stored-procedure", "run-python-element",
        ]
        report = audit_openapi(spec)
        self.assertIn("actions", report["missing"])

    def test_json_pointer_decoding(self):
        document = {"components": {"schemas": {"A/B": {"value~name": 1}}}}
        self.assertEqual(
            resolve_ref(document, "#/components/schemas/A~1B/value~0name"), 1)

    def test_ref_cycles_are_tolerated(self):
        spec = fixture_spec()
        spec["components"]["schemas"]["Contents"]["properties"]["self"] = {
            "$ref": "#/components/schemas/Contents"
        }
        self.assertEqual(audit_openapi(copy.deepcopy(spec))["status"], "PASS")


if __name__ == "__main__":
    unittest.main()
