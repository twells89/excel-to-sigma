#!/usr/bin/env python3
"""The forecast stages come from one function: today's live specs must not change (goldens)."""
from __future__ import annotations

import copy
import importlib.util
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
DEMO_KIT = ROOT.parents[1] / "docs" / "demo-kit"
GOLDEN = ROOT / "tests" / "fixtures" / "forecast-golden"
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(SCRIPTS))

import demo_workflows as dw  # noqa: E402
from forecast_plan import draft_plan  # noqa: E402


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def canonical(spec):
    return json.dumps(spec, indent=1, sort_keys=True) + "\n"


@unittest.skipUnless(DEMO_KIT.is_dir(), "demo kit not present (standalone skill copy)")
class LiveSpecsUnchanged(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        tmp = Path(cls._tmp.name)
        cls.build = load("bfm_stages", SCRIPTS / "build-forecast-model.py")
        src = load("make_demo_forecast_stages", DEMO_KIT / "make-demo-forecast.py").build_demo_workbook(tmp)
        name = "Vantora FY27 Rolling Forecast — golden"
        cls.what_if, _ = draft_plan(src, intent="what-if", connection="conn-1", folder="fold-1", name=name)
        cls.read_only, _ = draft_plan(src, intent="read-only", connection="conn-1", folder="fold-1", name=name)
        cls.manifest = dw.build_workflow_manifest(DEMO_KIT / "demo-forecast-macros.bas")
        cls.els = dict(cls.build.LOCAL_DM_IDS)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()

    def specs(self):
        pct = copy.deepcopy(self.what_if)
        first = next(ln for ln in pct["lines"] if ln["kind"] == "growth" and ln["section"] != "Revenue")
        first.update({"kind": "pct_of_total", "of": "Revenue", "rate": 0.1})
        b = self.build.build_wb_spec
        return {"plain-what-if": b(copy.deepcopy(self.what_if), "dm-1", self.els),
                "app-what-if": b(copy.deepcopy(self.what_if), "dm-1", self.els, workflows=self.manifest),
                "app-read-only": b(copy.deepcopy(self.read_only), "dm-1", self.els, workflows=self.manifest),
                "app-pct": b(pct, "dm-1", self.els, workflows=self.manifest)}

    def test_live_specs_unchanged(self):
        specs = self.specs()
        if os.environ.get("WRITE_FORECAST_GOLDENS"):
            GOLDEN.mkdir(parents=True, exist_ok=True)
            for key, spec in specs.items():
                (GOLDEN / f"{key}.json").write_text(canonical(spec))
        for key, spec in specs.items():
            with self.subTest(key=key):
                self.assertEqual(canonical(spec), (GOLDEN / f"{key}.json").read_text())


if __name__ == "__main__":
    unittest.main()
