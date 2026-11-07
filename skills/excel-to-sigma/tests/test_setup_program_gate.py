#!/usr/bin/env python3
"""build-forecast-model.py --emit-manual-commit-program: a plan without input tables skips the program,
and the program is built and checked from the local spec before anything is posted."""
from __future__ import annotations

import contextlib
import copy
import importlib.util
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
sys.path.insert(0, str(SCRIPTS / "lib"))
sys.path.insert(0, str(SCRIPTS))

spec = importlib.util.spec_from_file_location("bfm_setup_gate", SCRIPTS / "build-forecast-model.py")
BUILD = importlib.util.module_from_spec(spec)
spec.loader.exec_module(BUILD)
SAMPLE = json.loads((SCRIPTS / "sample-model-plan.json").read_text())


class SetupProgramGateTests(unittest.TestCase):
    def run_main(self, plan, extra, tmp):
        plan_path = Path(tmp) / "plan.json"
        plan_path.write_text(json.dumps(plan))
        gate = (Path(tmp) / "source.xlsx", {"status": "DRAFT", "sourceSha256": "0" * 64})
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err), \
                mock.patch.object(BUILD, "gate_plan_source", return_value=gate), \
                mock.patch.object(BUILD.demo_workflows, "source_history", return_value=[]):
            code = BUILD.main([str(plan_path), "--out-dir", str(Path(tmp) / "out"), *extra])
        return code, out.getvalue(), err.getvalue()

    def test_read_only_plan_skips_the_program_with_a_note(self):
        plan = {**copy.deepcopy(SAMPLE), "intent": "read-only"}
        with tempfile.TemporaryDirectory() as tmp:
            program = Path(tmp) / "setup-program.py"
            code, out, err = self.run_main(plan, ["--emit-manual-commit-program", str(program)], tmp)
            self.assertEqual(code, 0, out + err)
            self.assertNotIn("PLAN ERROR", err)
            self.assertIn("SETUP PROGRAM: skipped", out)
            self.assertIn("no input tables", out)
            self.assertFalse(program.exists())
            self.assertTrue((Path(tmp) / "out" / "wb_spec.json").is_file())

    def test_a_program_failure_on_post_stops_before_any_api_call(self):
        plan = {**copy.deepcopy(SAMPLE), "intent": "what-if"}
        with tempfile.TemporaryDirectory() as tmp:
            program = Path(tmp) / "setup-program.py"
            boom = BUILD.demo_workflows.WorkflowError("forced program failure")
            with mock.patch.object(BUILD, "api") as api, \
                    mock.patch.object(BUILD, "env", return_value={"SIGMA_BASE_URL": "https://x",
                                                                  "SIGMA_API_TOKEN": "t"}), \
                    mock.patch.object(BUILD.demo_workflows, "emit_manual_commit_program", side_effect=boom):
                code, out, err = self.run_main(plan, ["--post", "--emit-manual-commit-program", str(program)], tmp)
            self.assertEqual(code, 1)
            api.assert_not_called()
            self.assertIn("forced program failure", err)
            self.assertIn("nothing was written or posted", err)
            self.assertFalse(program.exists())
            self.assertFalse((Path(tmp) / "out" / "dm_spec.json").exists())


if __name__ == "__main__":
    unittest.main()
