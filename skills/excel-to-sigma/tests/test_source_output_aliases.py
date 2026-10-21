"""CLI output destinations must never overwrite an OOXML source workbook."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

from openpyxl import Workbook
from openpyxl.worksheet.table import Table


SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"


class SourceOutputAliasTests(unittest.TestCase):
    def setUp(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.folder = Path(temporary.name)
        self.source = self.folder / "source.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Inputs"
        sheet.append(["Business ID", "Demand"])
        sheet.append(["A", 10])
        sheet.append(["B", 20])
        sheet.add_table(Table(displayName="tblInputs", ref="A1:B3"))
        workbook.save(self.source)
        workbook.close()
        self.original = self.source.read_bytes()
        self.source_hash = hashlib.sha256(self.original).hexdigest()
        self.decisions = self.folder / "decisions.json"
        self.decisions.write_text(json.dumps({"sourceSha256": self.source_hash,
            "features": [{"featureId": "demand-input", "sheet": "Inputs",
                "table": "tblInputs", "role": "input", "reviewed": True,
                "keyColumns": ["Business ID"],
                "literalColumns": ["Business ID", "Demand"],
                "columnTypes": {"Business ID": "text", "Demand": "number"}}]}))

    def alias(self, kind, name):
        if kind == "direct":
            return self.source
        path = self.folder / name
        if kind == "symlink":
            path.symlink_to(self.source)
        elif kind == "hardlink":
            os.link(self.source, path)
        else:
            raise AssertionError(kind)
        return path

    def run_presentation(self, out, plan_out=None, *extra):
        argv = [sys.executable, str(SCRIPTS / "build-source-presentation.py"),
                str(self.source), "--out", str(out)]
        if plan_out is not None:
            argv += ["--plan-out", str(plan_out)]
        return subprocess.run([*argv, *extra], capture_output=True, text=True)

    def run_inputs(self, out, decisions=None):
        return subprocess.run([sys.executable, str(SCRIPTS / "source-inputs.py"),
            "plan", "--source", str(self.source), "--decisions", str(decisions or self.decisions),
            "--connection", "connection", "--folder", "folder", "--out", str(out)],
            capture_output=True, text=True)

    def run_native(self, command, out, **paths):
        argv = [sys.executable, str(SCRIPTS / "native-calculations.py"), command]
        for name, path in paths.items():
            argv += ["--" + name.replace("_", "-"), str(path)]
        return subprocess.run([*argv, "--out", str(out)], capture_output=True, text=True)

    def run_regional(self, script, out, **paths):
        argv = [sys.executable, str(SCRIPTS / script)]
        for name, path in paths.items():
            argv += ["--" + name.replace("_", "-"), str(path)]
        return subprocess.run([*argv, "--out", str(out)], capture_output=True, text=True)

    def test_presentation_rejects_each_source_alias_for_both_outputs_before_any_write(self):
        for kind in ("direct", "symlink", "hardlink"):
            for slot in ("out", "plan-out"):
                with self.subTest(kind=kind, slot=slot):
                    alias = self.alias(kind, f"presentation-{kind}-{slot}.xlsx")
                    other = self.folder / f"presentation-other-{kind}-{slot}.json"
                    result = (self.run_presentation(alias, other) if slot == "out"
                              else self.run_presentation(other, alias))
                    self.assertNotEqual(result.returncode, 0)
                    self.assertIn("aliases input artifact: source", result.stderr)
                    self.assertFalse(other.exists(), "the safe output was written before alias rejection")
                    self.assertEqual(self.source.read_bytes(), self.original)

    def test_presentation_rejects_outputs_aliasing_each_other_and_late_errors(self):
        shared = self.folder / "shared.json"
        shared.write_text("original")
        alias = self.folder / "shared-hardlink.json"
        os.link(shared, alias)
        result = self.run_presentation(shared, alias)
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("output paths alias", result.stderr)
        self.assertEqual(shared.read_text(), "original")
        plan_out = self.folder / "plan-if-error.json"
        out = self.folder / "result-if-error.json"
        result = self.run_presentation(out, plan_out, "--preview",
                                       "--bindings", str(self.folder / "missing-bindings.json"))
        self.assertNotEqual(result.returncode, 0)
        self.assertFalse(plan_out.exists())
        self.assertFalse(out.exists())
        self.assertEqual(self.source.read_bytes(), self.original)

    def test_presentation_rejects_every_declared_input_artifact_alias(self):
        for option in ("bindings", "manifest", "approval", "staging-receipt", "openapi", "contents"):
            protected = self.folder / f"{option}.json"
            protected.write_text("{}")
            for kind in ("direct", "symlink", "hardlink"):
                for slot in ("out", "plan-out"):
                    with self.subTest(option=option, kind=kind, slot=slot):
                        alias = protected if kind == "direct" else self.folder / f"{option}-{kind}-{slot}.json"
                        if kind == "symlink":
                            alias.symlink_to(protected)
                        elif kind == "hardlink":
                            os.link(protected, alias)
                        other = self.folder / f"safe-{option}-{kind}-{slot}.json"
                        extra = [f"--{option}", str(protected)]
                        if option == "contents":
                            extra += ["--bindings", str(self.decisions)]
                        result = (self.run_presentation(alias, other, *extra)
                                  if slot == "out" else
                                  self.run_presentation(other, alias, *extra))
                        self.assertNotEqual(result.returncode, 0)
                        self.assertIn("aliases input artifact", result.stderr)
                        self.assertFalse(other.exists())
                        self.assertEqual(protected.read_text(), "{}")

    def test_source_inputs_rejects_direct_symlink_and_hardlink_aliases(self):
        for kind in ("direct", "symlink", "hardlink"):
            with self.subTest(kind=kind):
                alias = self.alias(kind, f"inputs-{kind}.xlsx")
                result = self.run_inputs(alias)
                self.assertNotEqual(result.returncode, 0)
                self.assertIn("aliases input artifact: source", result.stderr)
                self.assertEqual(self.source.read_bytes(), self.original)

    def test_source_inputs_rejects_decisions_and_axis_values_aliases(self):
        axis = self.folder / "axis.json"
        axis.write_text("[1, 2]")
        for protected, command in ((self.decisions, "plan"), (axis, "unpivot")):
            original = protected.read_bytes()
            for kind in ("direct", "symlink", "hardlink"):
                with self.subTest(command=command, kind=kind):
                    alias = protected if kind == "direct" else self.folder / f"{command}-{kind}.json"
                    if kind == "symlink":
                        alias.symlink_to(protected)
                    elif kind == "hardlink":
                        os.link(protected, alias)
                    if command == "plan":
                        result = self.run_inputs(alias)
                    else:
                        result = subprocess.run([sys.executable, str(SCRIPTS / "source-inputs.py"),
                            "unpivot", "--source", str(self.source), "--sheet", "Inputs",
                            "--table", "tblInputs", "--key-column", "Business ID",
                            "--value-column", "Demand", "--axis-name", "axis",
                            "--value-name", "value", "--axis-values", str(axis),
                            "--out", str(alias)], capture_output=True, text=True)
                    self.assertNotEqual(result.returncode, 0)
                    self.assertIn("aliases input artifact", result.stderr)
                    self.assertEqual(protected.read_bytes(), original)

    def test_native_cli_rejects_source_and_immutable_input_aliases(self):
        input_plan = self.folder / "native-input-plan.json"
        current = self.folder / "native-current.json"
        observed = self.folder / "native-observed.json"
        for path in (input_plan, current, observed):
            path.write_text("{}")
        cases = (("plan-source", "plan", {"source": self.source}, self.source),
                 ("plan-input", "plan", {"source": self.source, "input_plan": input_plan}, input_plan),
                 ("backup-current", "backup", {"current": current, "observed_rows": observed}, current),
                 ("backup-rows", "backup", {"current": current, "observed_rows": observed}, observed))
        for label, command, paths, protected in cases:
            original = protected.read_bytes()
            for kind in ("direct", "symlink", "hardlink"):
                with self.subTest(label=label, kind=kind):
                    output = (protected if kind == "direct" else
                              self.folder / f"native-{label}-{kind}.json")
                    if kind == "symlink":
                        output.symlink_to(protected)
                    elif kind == "hardlink":
                        os.link(protected, output)
                    result = self.run_native(command, output, **paths)
                    self.assertNotEqual(result.returncode, 0)
                    self.assertIn("aliases input artifact", result.stderr)
                    self.assertEqual(protected.read_bytes(), original)
        distinct = self.folder / "native-plan.json"
        result = self.run_native("plan", distinct, source=self.source)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(json.loads(distinct.read_text())["sourceSha256"], self.source_hash)
        self.assertEqual(self.source.read_bytes(), self.original)

    def test_regional_emitters_reject_every_immutable_input_alias(self):
        base = self.folder / "base-input-plan.json"
        candidate = self.folder / "candidate.json"
        native = self.folder / "native-plan.json"
        oracle = self.folder / "excel-oracles.jsonl"
        for path in (base, candidate, native, oracle):
            path.write_text("{}")
        cases = (
            ("regional-canary-inputs.py",
             {"source": self.source, "base_input_plan": base, "candidate": candidate}),
            ("regional-grouped-canary.py",
             {"source": self.source, "input_plan": base,
              "native_plan": native, "excel_oracles": oracle}),
        )
        for script, paths in cases:
            for name, protected in paths.items():
                original = protected.read_bytes()
                for kind in ("direct", "symlink", "hardlink"):
                    with self.subTest(script=script, input=name, kind=kind):
                        output = (protected if kind == "direct" else
                                  self.folder / f"{script}-{name}-{kind}.json")
                        if kind == "symlink":
                            output.symlink_to(protected)
                        elif kind == "hardlink":
                            os.link(protected, output)
                        result = self.run_regional(script, output, **paths)
                        self.assertNotEqual(result.returncode, 0)
                        self.assertIn("aliases input artifact", result.stderr)
                        self.assertEqual(protected.read_bytes(), original)

    def test_distinct_outputs_succeed_without_mutating_source(self):
        plan, result = self.folder / "presentation-plan.json", self.folder / "presentation-result.json"
        presentation = self.run_presentation(result, plan)
        self.assertEqual(presentation.returncode, 0, presentation.stderr)
        self.assertEqual(json.loads(plan.read_text()), json.loads(result.read_text()))
        inputs_out = self.folder / "source-input-plan.json"
        inputs = self.run_inputs(inputs_out)
        self.assertEqual(inputs.returncode, 0, inputs.stderr)
        self.assertEqual(json.loads(inputs_out.read_text())["inputPlanVersion"], 2)
        self.assertEqual(self.source.read_bytes(), self.original)


if __name__ == "__main__":
    unittest.main()
