"""Focused provenance and hostile binding checks for the isolated B4 fixture."""
import ast
import copy
import inspect
import json
from pathlib import Path
import sys
import types
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).resolve().parents[1] / "scripts"
sys.path[:0] = [str(SCRIPTS), str(SCRIPTS / "lib")]
from lib.regional_canary_inputs import (build_canary_input_plan, validate_canary_input_plan,
    prepare_input_create_program, prepare_input_seed_program, reconcile_input_rows,
    confirm_phase_input_stability)
from lib.regional_grouped_canary import (prepare_contents_patch, confirm_contents_patch,
    prepare_grouping_program, prepare_scaffold_program, observe_b4_query,
    REVIEWED_CANDIDATE_DIGEST, _verified_namespaces)
from lib.native_exact_plan import digest

ROOT = Path(__file__).resolve().parents[5]
WORK = ROOT / "work"
SOURCE = WORK / "broader-eight-source-replay-20261020/source-originals/01 Regional Sales Forecast.xlsx"
BASE = WORK / "complete-source-record-bindings-plans-20261020/01-regional-sales-forecast-xlsx.json"
CANDIDATE = WORK / "regional-grouped-canary-20261020/candidate.json"
SDK_FORMULA = WORK / "native-grouped-docs-20261020/formula.py"


def check_formula_api_contract(program, sdk_path=SDK_FORMULA):
    """Resolve emitted formula calls against actual cached SDK declarations.

    The API inventory comes from the MCP SDK, independently of the emitter.
    This verifies names and argument binding, not live MCP admission/runtime.
    """
    declarations = {node.name: node for node in ast.parse(sdk_path.read_text()).body
                    if isinstance(node, ast.FunctionDef)}
    for call in ast.walk(ast.parse(program)):
        if not (isinstance(call, ast.Call) and isinstance(call.func, ast.Attribute)
                and ast.unparse(call.func.value) == "sigma.formula"):
            continue
        name = call.func.attr
        if name not in declarations:
            raise AssertionError(f"sigma.formula.{name} absent from actual SDK {sdk_path}")
        declaration = copy.deepcopy(declarations[name])
        # Compile declarations only, without importing SDK dependencies or
        # invoking its functions. Deferred annotations preserve exact signatures.
        declaration.body = [ast.Pass()]
        declaration.decorator_list = []
        module = ast.Module(body=[ast.ImportFrom(module="__future__",
            names=[ast.alias(name="annotations")], level=0), declaration], type_ignores=[])
        namespace = {}
        exec(compile(ast.fix_missing_locations(module), str(sdk_path), "exec"), namespace)
        inspect.signature(namespace[name]).bind(*call.args,
            **{keyword.arg: keyword.value for keyword in call.keywords})


class RegionalFormulaReadbackTest(unittest.TestCase):
    def packet(self):
        formula = 'If(Not([Not( column]), "Not(\\\"literal\\\")", Not([Flag]))'
        contents = {"elements": [{"id": "owned", "name": "Rows", "columns": [
            {"id": "h0", "name": "h0", "formula": formula},
            {"id": "foreign", "name": "Unchanged", "formula": "Not([Flag])"}]}],
            "layout": "exact layout"}
        packet = {"reviewedCandidateDigest": REVIEWED_CANDIDATE_DIGEST,
            "status": "REST_FORMULA_PATCH_PENDING_EXECUTION", "documentId": "document",
            "sourceSha256": "source", "ifMatchDocumentVersion": 4,
            "elementId": "owned", "contents": contents, "afterContentsHash": digest(contents),
            "changes": [{"logicalColumnId": "logical-h0", "actualColumnId": "h0",
                         "formula": formula}]}
        readback = {"workbookId": "document", "documentVersion": 5,
                    "contents": copy.deepcopy(contents)}
        return packet, readback

    def test_accepts_only_recorded_not_call_spacing_and_keeps_actual_lineage_hash(self):
        packet, readback = self.packet()
        actual = 'If(Not ([Not( column]), "Not(\\\"literal\\\")", Not ([Flag]))'
        readback["contents"]["elements"][0]["columns"][0]["formula"] = actual
        confirmed = confirm_contents_patch(packet, readback)
        self.assertEqual(confirmed["contentsHash"], digest(readback["contents"]))
        self.assertEqual(confirmed["expectedContentsHash"], packet["afterContentsHash"])
        self.assertEqual(confirmed["actualContentsHash"], digest(readback["contents"]))
        normalization, = confirmed["formulaNormalizations"]
        self.assertEqual(normalization["actualColumnId"], "h0")
        self.assertEqual(normalization["expectedFormula"], packet["changes"][0]["formula"])
        self.assertEqual(normalization["actualFormula"], actual)
        self.assertEqual(normalization["rule"], "not-call-single-ascii-space-before-open-paren")

    def test_exact_readback_still_has_no_normalization(self):
        packet, readback = self.packet()
        confirmed = confirm_contents_patch(packet, readback)
        self.assertEqual(confirmed["formulaNormalizations"], [])
        self.assertEqual(confirmed["contentsHash"], packet["afterContentsHash"])

    def test_rejects_literal_reference_function_operator_and_other_whitespace_edits(self):
        for change in (lambda f: f.replace('literal', 'changed'),
                       lambda f: f.replace('[Flag]', '[Other]'),
                       lambda f: f.replace('[Not( column]', '[Not ( column]'),
                       lambda f: f.replace('Not(', 'NotSomething('),
                       lambda f: f.replace('Not([Flag])', 'Not([Flag] = True)'),
                       lambda f: f.replace('Not(', 'Not  ('),
                       lambda f: f.replace('If(', 'If ('),
                       lambda f: f.replace('Not(', 'Not\t(')):
            packet, readback = self.packet()
            readback["contents"]["elements"][0]["columns"][0]["formula"] = change(
                packet["changes"][0]["formula"])
            with self.subTest(formula=readback["contents"]["elements"][0]["columns"][0]["formula"]):
                with self.assertRaises(ValueError):
                    confirm_contents_patch(packet, readback)

    def test_rejects_unowned_formula_or_document_shape_changes(self):
        for change in (lambda c: c["elements"][0]["columns"][1].__setitem__("formula", "Not ([Flag])"),
                       lambda c: c.__setitem__("layout", "different"),
                       lambda c: c["elements"][0].__setitem__("name", "Renamed"),
                       lambda c: c["elements"][0]["columns"][0].__setitem__("id", "other")):
            packet, readback = self.packet()
            change(readback["contents"])
            with self.assertRaises(ValueError):
                confirm_contents_patch(packet, readback)


@unittest.skipUnless(SOURCE.exists() and BASE.exists() and CANDIDATE.exists(),
                     "source-bound Regional canary packet is not installed")
class RegionalCanaryInputsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.base = json.loads(BASE.read_text())
        cls.candidate = json.loads(CANDIDATE.read_text())
        cls.plan = build_canary_input_plan(SOURCE, cls.base, cls.candidate)

    def setup_document(self):
        elements = []
        bindings = {}
        queries = {}
        for feature, spec in self.plan["features"].items():
            prefix = "test-only-sales" if feature == "sales-forecast" else "test-only-summary"
            columns = [{"id": f"{prefix}-c{i}", "name": c["name"], "type": c["type"]}
                       for i, c in enumerate(spec["columns"])]
            columns.append({"id": "ID"})  # Actual Sigma system-column serialization.
            element = {"id": prefix, "kind": "input-table", "name": spec["name"],
                       "source": {"kind": "empty", "connectionId": "test-only-connection"},
                       "columns": columns}
            elements.append(element)
            bindings[feature] = {"documentId": "test-only-document",
                "canaryInputPlanHash": self.plan["canaryInputPlanHash"],
                "elementId": prefix, "connectionId": "test-only-connection",
                "pageId": "test-only-page", "columnIds": {c["name"]: c["id"] for c in columns[:-1]},
                "rowIdColumn": "ID", "rowIdColumnId": columns[-1]["id"],
                "rowIdQueryColumnId": f"{prefix}-query-id-alias"}
            queries[feature] = {"elementId": prefix, "documentVersion": 7,
                "snapshotId": f"test-only-query-{feature}",
                "rowIdQueryColumnId": bindings[feature]["rowIdQueryColumnId"],
                "rows": [{**row, "ID": f"test-only-physical-{feature}-{i}"}
                         for i, row in enumerate(spec["rows"])]}
        current = {"workbookId": "test-only-document", "documentVersion": 7,
                   "contents": {"kind": "workbook", "pages": [{"id": "test-only-page"}],
                                "elements": elements, "layout": ""}}
        return current, bindings, queries

    def observed(self, current, bindings, queries):
        return {feature: {"binding": bindings[feature], "query": queries[feature],
                          "receipt": reconcile_input_rows(self.plan, current, feature=feature,
                              binding=bindings[feature], query=queries[feature])}
                for feature in bindings}

    def add_query_lineage(self, current, bindings, queries, feature, *, alias="ID"):
        binding, query = bindings[feature], queries[feature]
        binding["rowIdQueryColumnId"] = alias
        query["rowIdQueryColumnId"] = alias
        spec = self.plan["features"][feature]
        columns = [{"id": binding["columnIds"][c["name"]], "name": c["name"]}
                   for c in spec["columns"]] + [{"id": alias, "name": "ID"}]
        selection = ', '.join(f'"{c["id"]}" AS "{c["name"]}"' for c in columns)
        request = {"query": {"type": "workbook", "workbookId": current["workbookId"],
            "sql": f'SELECT {selection} FROM "workbook"."{binding["elementId"]}"'}}
        ddl_columns = [f'  "{binding["columnIds"][c["name"]]}" {c["type"]}, -- '
            f'"{c["name"]}" | Formula: [New Input Table/{binding["columnIds"][c["name"]]}]'
            for c in spec["columns"]]
        ddl_columns.append(f'  "{alias}" text -- "ID" | Formula: [New Input Table/ID]')
        ddl = '\n'.join(['-- Type: "input-table"',
            f'-- Workbook ID: "{current["workbookId"]}"',
            f'-- Element ID: "{binding["elementId"]}"',
            f'CREATE TABLE "workbook"."{binding["elementId"]}" (', *ddl_columns, ');'])
        labels = [c["name"] for c in columns]
        query["queryRequest"] = request
        query["columns"] = columns
        binding["rowIdQueryLineage"] = {"documentId": current["workbookId"],
            "documentVersion": current["documentVersion"], "elementId": binding["elementId"],
            "snapshotId": query["snapshotId"], "ddl": ddl, "columns": columns,
            "queryRequest": request, "count": len(query["rows"]),
            "resultHash": digest({"columns": labels,
                "rows": [[row[name] for name in labels] for row in query["rows"]]})}

    def test_actual_physical_id_query_is_accepted_only_with_exact_fresh_lineage(self):
        current, bindings, queries = self.setup_document()
        feature = "sales-forecast"
        self.add_query_lineage(current, bindings, queries, feature)
        observed = self.observed(current, bindings, queries)
        self.assertEqual(observed[feature]["receipt"]["rowIdQueryColumnId"], "ID")
        for attack in (lambda b, q: b.pop("rowIdQueryLineage"),
                       lambda b, q: b["rowIdQueryLineage"].__setitem__("documentVersion", 6),
                       lambda b, q: b["rowIdQueryLineage"].__setitem__("snapshotId", "stale"),
                       lambda b, q: b["rowIdQueryLineage"].__setitem__("ddl", b["rowIdQueryLineage"]["ddl"].replace('"ID" text', '"invented" text')),
                       lambda b, q: q["queryRequest"]["query"].__setitem__("sql", 'SELECT "foreign" FROM "workbook"."foreign"'),
                       lambda b, q: b.__setitem__("rowIdColumnId", "FOREIGN"),
                       lambda b, q: q["rows"][0].__setitem__("ID", "wrong-row"),
                       lambda b, q: q["rows"][0].pop("Source Key"),
                       lambda b, q: q["rows"][1].__setitem__("ID", q["rows"][0]["ID"])):
            bad_bindings, bad_queries = copy.deepcopy(bindings), copy.deepcopy(queries)
            attack(bad_bindings[feature], bad_queries[feature])
            with self.assertRaises(ValueError):
                self.observed(current, bad_bindings, bad_queries)

    def test_query_alias_transition_preserves_physical_inputs_and_records_lineage(self):
        before_current, bindings, queries = self.setup_document()
        before = self.observed(before_current, bindings, queries)
        after_current = copy.deepcopy(before_current)
        after_current["documentVersion"] = 8
        after_bindings, after_queries = copy.deepcopy(bindings), copy.deepcopy(queries)
        for feature in bindings:
            after_queries[feature]["documentVersion"] = 8
            after_queries[feature]["snapshotId"] = f"fresh-{feature}"
            self.add_query_lineage(after_current, after_bindings, after_queries, feature)
        after = self.observed(after_current, after_bindings, after_queries)
        confirmed = confirm_phase_input_stability(self.plan, SOURCE, self.base, self.candidate,
            before_current, before, after_current, after)
        self.assertEqual(set(confirmed["queryAliasLineage"]), set(bindings))
        self.assertEqual(confirmed["queryAliasLineage"]["sales-forecast"]["afterQueryColumnId"], "ID")
        # An arbitrary new distinct alias is not sufficient to authorize an
        # alias transition, even though legacy distinct-ID receipts remain readable.
        undocumented = copy.deepcopy(after)
        for field in ("binding", "query"):
            undocumented["sales-forecast"][field]["rowIdQueryColumnId"] = "invented"
        undocumented["sales-forecast"]["binding"].pop("rowIdQueryLineage")
        undocumented["sales-forecast"]["receipt"] = reconcile_input_rows(self.plan, after_current,
            feature="sales-forecast", binding=undocumented["sales-forecast"]["binding"],
            query=undocumented["sales-forecast"]["query"])
        with self.assertRaises(ValueError):
            confirm_phase_input_stability(self.plan, SOURCE, self.base, self.candidate,
                before_current, before, after_current, undocumented)
        # Freshly captured bad row values must still fail source reconciliation
        # or physical-row conservation, even with a matching result evidence hash.
        for attack in (lambda rows: rows[0].__setitem__("ID", "wrong-physical-row"),
                       lambda rows: rows[1].__setitem__("ID", rows[0]["ID"]),
                       lambda rows: rows[0].__setitem__("Source Key", "unknown-source-key"),
                       lambda rows: rows[0].__setitem__("Source Ordinal", 999),
                       lambda rows: rows[0].__setitem__("Units", 1234567)):
            bad_bindings, bad_queries = copy.deepcopy(after_bindings), copy.deepcopy(after_queries)
            attack(bad_queries["sales-forecast"]["rows"])
            self.add_query_lineage(after_current, bad_bindings, bad_queries, "sales-forecast")
            with self.assertRaises(ValueError):
                bad_after = self.observed(after_current, bad_bindings, bad_queries)
                confirm_phase_input_stability(self.plan, SOURCE, self.base, self.candidate,
                    before_current, before, after_current, bad_after)

    def namespace(self, element, *, system_alias=None):
        columns = {c["name"]: {"id": c["id"], "name": c["name"]}
                   for c in element["columns"] if "name" in c}
        if system_alias:
            columns["Row ID"] = {"id": "ID", "name": "ID"}
        return {"elementId": element["id"], "name": element["name"], "columns": columns}

    @unittest.skipUnless((WORK / "regional-grouped-canary-20261020/root-after-grouping-source6-R2.json").exists(),
                         "actual persisted grouping source unavailable")
    def test_actual_grouped_flat_unnamed_columns_require_exact_parent_and_native_ddl_lineage(self):
        packet = WORK / "regional-grouped-canary-20261020"
        current = json.loads((packet / "root-after-grouping-source6-R2.json").read_text())
        observed = json.loads((packet / "root-after-grouping-observed-inputs6-R2.json").read_text())
        raw_ddl = json.loads((packet / "root-grouped-flat-describe6-R2.json").read_text())
        by_id = {e["id"]: e for e in current["contents"]["elements"]}
        sales = by_id["GCVsg9yMKO"]
        summary = by_id["CCnm2Z8VYG"]
        rows = by_id["yDvaQ-qY-V"]
        flat = by_id["ZmE2J1tt_Z"]
        self.assertTrue(all("name" not in c for c in flat["columns"]))
        row_columns = {c["name"]: {"id": c["id"], "name": c["name"]}
                       for c in rows["columns"] if c["id"] not in ("EaPQHGxkAi", "1zF2g5Jr80")}
        namespaces = {
            "Sales": self.namespace(sales, system_alias=True),
            "Summary Regions": self.namespace(summary, system_alias=True),
            "Rows": {"elementId": rows["id"], "name": rows["name"], "columns": row_columns},
            "Grouped": {"elementId": flat["id"], "name": flat["name"],
                "columns": {"Group Key": {"id": "XFx9OejojL", "name": "Group Key"},
                            "Revenue Sum": {"id": "ThCTKsVs2J", "name": "Revenue Sum"},
                            "First Error Ordinal": {"id": "UVtagwDpPJ", "name": "First Error Ordinal"}},
                "nativeDescribe": raw_ddl},
        }
        self.assertEqual(_verified_namespaces(current, namespaces, observed,
            element_kind="summary-table", candidate=self.candidate)["Grouped"]["id"], flat["id"])
        # A valid DDL receipt cannot launder a changed parent, grouping or flat formula.
        for mutate in (
            lambda c, n: c["contents"]["elements"][3]["source"].__setitem__("groupingId", "foreign"),
            lambda c, n: c["contents"]["elements"][2]["groupings"][0]["calculations"].__setitem__(0, "foreign"),
            lambda c, n: c["contents"]["elements"][3]["columns"][1].__setitem__("formula", "1234567"),
            lambda c, n: c["contents"]["elements"][2]["columns"][1].__setitem__("formula", "Sum([Other])"),
            lambda c, n: n["Grouped"]["columns"]["Revenue Sum"].__setitem__("id", "foreign"),
            lambda c, n: n["Grouped"]["nativeDescribe"]["structuredContent"].__setitem__("ddl", "forged"),
        ):
            forged_current, forged_namespaces = copy.deepcopy(current), copy.deepcopy(namespaces)
            mutate(forged_current, forged_namespaces)
            with self.subTest(mutate=mutate), self.assertRaises(ValueError):
                _verified_namespaces(forged_current, forged_namespaces, observed,
                                     element_kind="summary-table", candidate=self.candidate)

    def grouping_fixture(self):
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        row_element = {"id": "test-only-row-table", "kind": "table", "name": "Regional Sales Rows",
            "source": {"kind": "table", "elementId": bindings["sales-forecast"]["elementId"]},
            "columns": [{"id": f"test-only-row-c{i}", "name": c["name"], "formula": c["formula"]}
                        for i, c in enumerate(self.candidate["elements"][0]["columns"])]}
        current["contents"]["elements"].append(row_element)
        confirmed = {"status": "CONFIRMED_FORMULA_PATCH_STRUCTURE_ONLY",
            "sourceSha256": self.candidate["sourceSha256"], "documentId": current["workbookId"],
            "documentVersion": 7, "contentsHash": digest(current["contents"])}
        grouped = prepare_grouping_program(self.candidate, current, confirmed,
            row_element_id=row_element["id"], page_id="test-only-page", grid_row=15,
            source_path=SOURCE, base_input_plan=self.base,
            canary_input_plan=self.plan, observed_inputs=observed)
        return row_element, bindings, grouped["program"]

    def test_grouping_receiver_guards_reject_drift_before_any_mutation(self):
        row_element, bindings, program = self.grouping_fixture()
        effects = []
        class Column:
            def __init__(self, meta): self.meta = meta
            def id(self): return self.meta["id"]
            def name(self): return self.meta["name"]
            def formula(self): return self.meta["formula"]
        class QueryableSource:
            def __init__(self, identity): self.identity = identity
            def ref(self): return self.identity
        class Grouping:
            def set_calculation(self, name, formula):
                effects.append(("calculation", name))
                return Column({"id": name, "name": name})
            def to_source(self): return QueryableSource("grouped")
        class LevelTable:
            def __init__(self, meta):
                self.meta = meta
                self.source = QueryableSource(meta["source"]["elementId"])
            def id(self): return self.meta["id"]
            def name(self): return self.meta["name"]
            def data_source(self): return self.source
            def list_columns(self): return [Column(c) for c in self.meta["columns"]]
            def get_column_by_name(self, name):
                return next(c for c in self.list_columns() if c.name() == name)
            def get_column(self, identity):
                return next(c for c in self.list_columns() if c.id() == identity)
            def set_groupings(self, keys):
                effects.append(("grouping", keys))
                return [Grouping()]
            def set_name(self, name): effects.append(("name", name))
            def set_frame(self, *frame): effects.append(("frame", frame))
        modules = {name: types.ModuleType(name) for name in ("sigma.levelTable", "sigma.data.source")}
        modules["sigma.levelTable"].LevelTable = LevelTable
        modules["sigma.data.source"].QueryableSource = QueryableSource
        for attack in ("none", "receiver-type", "element-id", "name", "source-type", "source-id",
                       "column-id", "column-name", "column-formula", "column-order"):
            effects.clear()
            table = LevelTable(copy.deepcopy(row_element))
            if attack == "receiver-type": table = object()
            elif attack == "element-id": table.meta["id"] = "foreign"
            elif attack == "name": table.meta["name"] = "Foreign"
            elif attack == "source-type": table.source = object()
            elif attack == "source-id": table.source = QueryableSource("foreign")
            elif attack.startswith("column-"):
                if attack == "column-order": table.meta["columns"].reverse()
                else: table.meta["columns"][0][attack.split('-')[1]] = "foreign"
            page = types.SimpleNamespace(get_element=lambda identity: table,
                add_table=lambda source: effects.append(("flat-table", source)) or LevelTable(copy.deepcopy(row_element)))
            sigma = types.SimpleNamespace(workbook=types.SimpleNamespace(get_page=lambda identity: page),
                data=types.SimpleNamespace(source=types.SimpleNamespace(element=QueryableSource)),
                formula=types.SimpleNamespace(sum=lambda c: ("sum", c), min=lambda c: ("min", c)))
            with self.subTest(attack=attack), patch.dict(sys.modules, modules):
                if attack == "none":
                    exec(program, {"sigma": sigma, "print": lambda value: None})
                    self.assertTrue(effects)
                else:
                    with self.assertRaises(RuntimeError):
                        exec(program, {"sigma": sigma, "print": lambda value: None})
                    self.assertEqual(effects, [])

    @unittest.skipUnless(SDK_FORMULA.exists(), "actual cached MCP SDK declarations unavailable")
    def test_sdk_receiver_contract_narrows_generic_get_element_and_scaffolds_use_add_table(self):
        docs = WORK / "native-grouped-docs-20261020"
        table_docs = (docs / "levelTable.md").read_text()
        self.assertIn("Canonical module: /sdk/sigma/levelTable.py", table_docs)
        self.assertIn("def get_column(self, id: str) -> FormulaColumn:", table_docs)
        self.assertIn("def data_source(self) -> Source \\| None:", table_docs)
        source_docs = (docs / "source.md").read_text()
        self.assertIn("Canonical module: /sdk/sigma/data/source.py", source_docs)
        self.assertIn("## Queryable Source", source_docs)
        self.assertIn("def ref(self) -> str:", source_docs)
        page = (WORK / "sdk-input-docs/page.md").read_text()
        self.assertIn("def get_element(self, element_id: ElementId) -> Element:", page)
        _, _, program = self.grouping_fixture()
        tree = ast.parse(program)
        self.assertTrue(any(isinstance(n, ast.ImportFrom) and n.module == "sigma.levelTable"
            and any(a.name == "LevelTable" for a in n.names) for n in tree.body))
        guard_index = next(i for i, n in enumerate(tree.body) if isinstance(n, ast.If)
            and ast.unparse(n.test) == "not isinstance(rows, LevelTable)"
            and any(isinstance(child, ast.Raise) for child in n.body))
        method_indices = [i for i, statement in enumerate(tree.body)
            if any(isinstance(n, ast.Call) and isinstance(n.func, ast.Attribute)
                   and ast.unparse(n.func.value) == "rows" and
                   n.func.attr in ("list_columns", "get_column", "get_column_by_name", "set_groupings")
                   for n in ast.walk(statement))]
        self.assertTrue(method_indices and all(i > guard_index for i in method_indices))
        # Both scaffold phases obtain a documented LevelTable directly; neither
        # re-fetches the summary/source receiver through generic get_element.
        self.assertIn("-> LevelTable:", page)
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        for kind, feature in (("calculated-table", "sales-forecast"), ("summary-table", "summary-regions")):
            scaffold = prepare_scaffold_program(self.candidate, element_kind=kind,
                page_id="test-only-page", source_element_id=bindings[feature]["elementId"],
                grid_row=1, current=current, source_path=SOURCE, base_input_plan=self.base,
                canary_input_plan=self.plan, observed_inputs=observed)
            self.assertNotIn(".get_element(", scaffold)
            self.assertIn("t = p.add_table(", scaffold)

    def test_projection_exact_two_inputs_and_pending_denominator(self):
        self.assertEqual(self.plan["status"], "CANARY_QUALIFICATION_ONLY")
        self.assertEqual({k: len(v["rows"]) for k, v in self.plan["features"].items()},
                         {"sales-forecast": 144, "summary-regions": 4})
        self.assertEqual(len(self.plan["pendingFormulaCellIds"]), 214)
        self.assertEqual(self.plan["omittedSourceInputFeatures"],
                         ["regions", "summary-months", "summary-products"])
        sales = self.plan["features"]["sales-forecast"]
        self.assertEqual(sales["sourceCells"]["0"]["Units"], "Forecast Input!D2")
        self.assertEqual(sales["rows"][0]["Source Key"], "Forecast Input!row:2")
        self.assertEqual(sales["rows"][0]["Source Ordinal"], 1)
        self.assertNotIn("Source Key", sales["sourceCells"]["0"])
        self.assertEqual(self.plan["features"]["summary-regions"]["sourceCells"]["0"],
                         {"Region": "Summary!A4"})
        forged = copy.deepcopy(self.plan)
        forged["features"]["sales-forecast"]["rows"][0]["Units"] = 1234567
        from lib.native_exact_plan import digest
        forged["canaryInputPlanHash"] = digest({k: v for k, v in forged.items()
                                                if k != "canaryInputPlanHash"})
        with self.assertRaisesRegex(ValueError, "independently reviewed projection"):
            validate_canary_input_plan(forged, SOURCE, self.base, self.candidate)

    def test_create_seed_and_reconcile_are_bounded_and_observed(self):
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        scaffold = prepare_scaffold_program(self.candidate, element_kind="calculated-table",
            page_id="test-only-page", source_element_id=bindings["sales-forecast"]["elementId"],
            grid_row=1, current=current, source_path=SOURCE, base_input_plan=self.base,
            canary_input_plan=self.plan, observed_inputs=observed)
        compile(scaffold, "<row-scaffold>", "exec")
        self.assertEqual(scaffold.count("t.add_column("), len(self.candidate["elements"][0]["columns"]))
        empty = copy.deepcopy(current)
        empty["contents"]["elements"] = []
        for feature in bindings:
            create = prepare_input_create_program(self.plan, empty, feature=feature,
                page_id="test-only-page", connection_id="test-only-connection", grid_row=1)
            compile(create["program"], "<input-create>", "exec")
            self.assertNotIn("upsert_rows", create["program"])
            empty_query = {"elementId": bindings[feature]["elementId"],
                           "documentVersion": 7, "snapshotId": "test-only-empty", "rows": [],
                           "rowIdQueryColumnId": bindings[feature]["rowIdQueryColumnId"]}
            seed = prepare_input_seed_program(self.plan, current, feature=feature,
                binding=bindings[feature], empty_query=empty_query)
            compile(seed["program"], "<input-seed>", "exec")
            self.assertIn("row_ids = table.upsert_rows([", seed["program"])
            self.assertEqual(seed["expectedRows"], len(self.plan["features"][feature]["rows"]))
            result = reconcile_input_rows(self.plan, current, feature=feature,
                                          binding=bindings[feature], query=queries[feature])
            self.assertEqual(len(result["rows"]), seed["expectedRows"])
        wrong = copy.deepcopy(queries["summary-regions"])
        wrong["rows"][0]["Region"] = "Wrong"
        with self.assertRaisesRegex(ValueError, "value differs"):
            reconcile_input_rows(self.plan, current, feature="summary-regions",
                                 binding=bindings["summary-regions"], query=wrong)
        for forged_system in ({"id": "FOREIGN"}, {"id": "ID", "name": "Foreign ID"}):
            wrong_doc = copy.deepcopy(current)
            wrong_doc["contents"]["elements"][0]["columns"][-1] = forged_system
            with self.assertRaisesRegex(ValueError, "foreign or overlapping columns|physical column"):
                reconcile_input_rows(self.plan, wrong_doc, feature="sales-forecast",
                                     binding=bindings["sales-forecast"],
                                     query=queries["sales-forecast"])
        foreign_doc = copy.deepcopy(current)
        foreign_doc["contents"]["elements"][0]["columns"].append(
            {"id": "foreign", "name": "Foreign", "type": "text"})
        with self.assertRaisesRegex(ValueError, "foreign or overlapping columns"):
            reconcile_input_rows(self.plan, foreign_doc, feature="sales-forecast",
                                 binding=bindings["sales-forecast"],
                                 query=queries["sales-forecast"])
        wrong_alias = copy.deepcopy(queries["sales-forecast"])
        wrong_alias["rowIdQueryColumnId"] = "foreign-query-alias"
        with self.assertRaisesRegex(ValueError, "fresh physical input query"):
            reconcile_input_rows(self.plan, current, feature="sales-forecast",
                                 binding=bindings["sales-forecast"], query=wrong_alias)

    @unittest.skipUnless(SDK_FORMULA.exists(), "actual cached MCP SDK formula declarations unavailable")
    def test_scaffolds_resolve_formula_calls_against_actual_sdk(self):
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        for kind, feature in (("calculated-table", "sales-forecast"),
                              ("summary-table", "summary-regions")):
            with self.subTest(kind=kind):
                program = prepare_scaffold_program(self.candidate, element_kind=kind,
                    page_id="test-only-page", source_element_id=bindings[feature]["elementId"],
                    grid_row=1, current=current, source_path=SOURCE, base_input_plan=self.base,
                    canary_input_plan=self.plan, observed_inputs=observed)
                check_formula_api_contract(program)
        # Demonstrate that the independent contract rejects the observed live
        # failure and incorrect invocation of the real constant constructor.
        with self.assertRaisesRegex(AssertionError, "const_number absent"):
            check_formula_api_contract("sigma.formula.const_number(0)")
        with self.assertRaises(TypeError):
            check_formula_api_contract("sigma.formula.const_num()")

    def test_native_phase_postcheck_rejects_editable_metadata_drift(self):
        before_current, bindings, before_queries = self.setup_document()
        before = self.observed(before_current, bindings, before_queries)
        after_current = copy.deepcopy(before_current)
        after_current["documentVersion"] = 8
        after_queries = copy.deepcopy(before_queries)
        for feature, query in after_queries.items():
            query["documentVersion"] = 8
            query["snapshotId"] = f"test-only-post-{feature}"
        after = self.observed(after_current, bindings, after_queries)
        receipt = confirm_phase_input_stability(self.plan, SOURCE, self.base, self.candidate,
            before_current, before, after_current, after)
        self.assertEqual(receipt["status"], "CONFIRMED_CANARY_INPUTS_STABLE_AFTER_EDIT")
        changed = copy.deepcopy(after_queries)
        changed["sales-forecast"]["rows"][0]["Source Ordinal"] = 999
        with self.assertRaisesRegex(ValueError, "metadata value differs"):
            self.observed(after_current, bindings, changed)

    def test_patch_checks_input_source_namespaces_and_b4_row_id(self):
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        row_logical = self.candidate["elements"][0]
        row_element = {"id": "test-only-row-table", "kind": "table", "name": "Rows",
                       "source": {"kind": "table", "elementId": bindings["sales-forecast"]["elementId"]},
                       "columns": [{"id": f"test-only-row-c{i}", "name": c["name"], "formula": "0"}
                                   for i, c in enumerate(row_logical["columns"])]}
        current["contents"]["elements"].append(row_element)
        namespaces = {"Sales": self.namespace(current["contents"]["elements"][0], system_alias=True),
                      "Summary Regions": self.namespace(current["contents"]["elements"][1], system_alias=True),
                      "Rows": self.namespace(row_element)}
        receipt = {"sourceSha256": self.plan["sourceSha256"],
                   "canaryInputPlanHash": self.plan["canaryInputPlanHash"],
                   "documentId": current["workbookId"], "documentVersion": 7,
                   "elements": {row_logical["id"]: {"elementId": row_element["id"],
                       "name": row_element["name"],
                       "columns": {logical["id"]: actual["id"] for logical, actual in
                                   zip(row_logical["columns"], row_element["columns"])}}},
                   "namespaces": namespaces}
        kwargs = dict(element_kind="calculated-table", source_path=SOURCE,
                      base_input_plan=self.base, canary_input_plan=self.plan,
                      observed_inputs=observed)
        packet = prepare_contents_patch(self.candidate, current, receipt, **kwargs)
        self.assertEqual(len(packet["changes"]), len(row_logical["columns"]))
        self.assertEqual(confirm_contents_patch(packet, {"workbookId": current["workbookId"],
            "documentVersion": 8, "contents": packet["contents"]})["runtimeStatus"],
                         "UNVERIFIED_CANDIDATE")
        wrong_source = copy.deepcopy(current)
        wrong_source["contents"]["elements"][-1]["source"]["elementId"] = "wrong-sales"
        with self.assertRaisesRegex(ValueError, "sourced|source"):
            prepare_contents_patch(self.candidate, wrong_source, receipt, **kwargs)
        wrong_namespace = copy.deepcopy(receipt)
        wrong_namespace["namespaces"]["Sales"]["name"] = "Wrong Sales"
        with self.assertRaisesRegex(ValueError, "namespace"):
            prepare_contents_patch(self.candidate, current, wrong_namespace, **kwargs)
        wrong_column = copy.deepcopy(receipt)
        wrong_column["namespaces"]["Sales"]["columns"]["Units"]["id"] = "foreign"
        with self.assertRaisesRegex(ValueError, "lineage"):
            prepare_contents_patch(self.candidate, current, wrong_column, **kwargs)
        for forged in (lambda c: next(col for col in c["elements"][0]["columns"]
                                      if col["name"] == "Revenue Number").__setitem__("formula", "1234567"),
                       lambda c: next(col for col in c["elements"][2]["columns"]
                                      if col["name"] == "B4 Number").__setitem__("formula", "1234567"),
                       lambda c: c["constructionBudget"].__setitem__("formulaColumns", 29),
                       lambda c: c["sourceMappings"].__setitem__("Summary!B4", {"targetElement": "foreign"})):
            changed = copy.deepcopy(self.candidate)
            forged(changed)
            from lib.native_exact_plan import digest
            changed["candidateDigest"] = digest(changed)  # attacker-controlled self-hash is irrelevant
            with self.assertRaisesRegex(ValueError, "independently reviewed graph"):
                validate_canary_input_plan(self.plan, SOURCE, self.base, changed)
            with self.assertRaisesRegex(ValueError, "independently reviewed graph"):
                prepare_contents_patch(changed, current, receipt, **kwargs)
            with self.assertRaisesRegex(ValueError, "independently reviewed graph"):
                prepare_scaffold_program(changed, element_kind="calculated-table",
                    page_id="test-only-page", source_element_id=bindings["sales-forecast"]["elementId"],
                    grid_row=1, current=current, source_path=SOURCE,
                    base_input_plan=self.base, canary_input_plan=self.plan,
                    observed_inputs=observed)

    def test_grouping_and_summary_b4_physical_row_readback(self):
        current, bindings, queries = self.setup_document()
        observed = self.observed(current, bindings, queries)
        row_logical, _, summary_logical = self.candidate["elements"]
        row_element = {"id": "test-only-row-table", "kind": "table", "name": "Regional Sales Rows",
            "source": {"kind": "table", "elementId": bindings["sales-forecast"]["elementId"]},
            "columns": [{"id": f"test-only-row-c{i}", "name": c["name"], "formula": c["formula"]}
                        for i, c in enumerate(row_logical["columns"])]}
        current["contents"]["elements"].append(row_element)
        from lib.native_exact_plan import digest
        confirmed = {"status": "CONFIRMED_FORMULA_PATCH_STRUCTURE_ONLY",
            "sourceSha256": self.candidate["sourceSha256"], "documentId": current["workbookId"],
            "documentVersion": 7, "contentsHash": digest(current["contents"])}
        grouped = prepare_grouping_program(self.candidate, current, confirmed,
            row_element_id=row_element["id"], page_id="test-only-page", grid_row=15,
            source_path=SOURCE, base_input_plan=self.base,
            canary_input_plan=self.plan, observed_inputs=observed)
        compile(grouped["program"], "<grouping>", "exec")
        if SDK_FORMULA.exists():
            check_formula_api_contract(grouped["program"])
        self.assertIn("sigma.formula.min(error_ordinal)", grouped["program"])
        self.assertIn("grouping.to_source()", grouped["program"])
        forged = copy.deepcopy(self.candidate)
        next(c for c in forged["elements"][0]["columns"] if c["name"] == "Revenue Number")["formula"] = "1234567"
        with self.assertRaisesRegex(ValueError, "independently reviewed graph"):
            prepare_grouping_program(forged, current, confirmed,
                row_element_id=row_element["id"], page_id="test-only-page", grid_row=15,
                source_path=SOURCE, base_input_plan=self.base,
                canary_input_plan=self.plan, observed_inputs=observed)
        grouped_element = {"id": "test-only-grouped", "kind": "table",
            "name": "Regional Sales Region Quarter",
            "source": {"kind": "grouping", "elementId": row_element["id"]},
            "columns": [{"id": f"test-only-group-c{i}", "name": name}
                        for i, name in enumerate(("Group Key", "Revenue Sum", "First Error Ordinal"))]}
        summary_element = {"id": "test-only-summary-table", "kind": "table", "name": "Regional Summary B4",
            "source": {"kind": "table", "elementId": bindings["summary-regions"]["elementId"]},
            "columns": [{"id": f"test-only-summary-c{i}", "name": c["name"], "formula": "0"}
                        for i, c in enumerate(summary_logical["columns"])]}
        current["contents"]["elements"] += [grouped_element, summary_element]
        namespaces = {"Sales": self.namespace(current["contents"]["elements"][0], system_alias=True),
            "Summary Regions": self.namespace(current["contents"]["elements"][1], system_alias=True),
            "Rows": self.namespace(row_element), "Grouped": self.namespace(grouped_element)}
        receipt = {"sourceSha256": self.plan["sourceSha256"],
            "canaryInputPlanHash": self.plan["canaryInputPlanHash"],
            "documentId": current["workbookId"], "documentVersion": 7,
            "elements": {summary_logical["id"]: {"elementId": summary_element["id"],
                "name": summary_element["name"],
                "columns": {logical["id"]: actual["id"] for logical, actual in
                            zip(summary_logical["columns"], summary_element["columns"])}}},
            "namespaces": namespaces}
        kwargs = dict(element_kind="summary-table", source_path=SOURCE,
            base_input_plan=self.base, canary_input_plan=self.plan, observed_inputs=observed)
        patch = prepare_contents_patch(self.candidate, current, receipt, **kwargs)
        self.assertEqual(patch["b4Source"]["sourceRowId"],
                         observed["summary-regions"]["receipt"]["rows"]["0"]["rowId"])
        readback = {"workbookId": current["workbookId"], "documentVersion": 8,
                    "contents": patch["contents"]}
        post_queries = copy.deepcopy(queries)
        for feature, source_query in post_queries.items():
            source_query["documentVersion"] = 8
            source_query["snapshotId"] = f"test-only-post-query-{feature}"
        post_inputs = self.observed(readback, bindings, post_queries)
        query = {"elementId": summary_element["id"], "documentVersion": 8,
                 "snapshotId": "test-only-summary-query",
                 "rows": [{"Source Key": "Summary!row:4",
                           "Source Row ID": patch["b4Source"]["sourceRowId"],
                           "B4 Number": 1544984.02, "B4 Text": "", "B4 Logical": False,
                           "B4 Kind": "n", "B4 Error": ""}]}
        observer_kwargs = dict(source_path=SOURCE, base_input_plan=self.base,
            canary_input_plan=self.plan, observed_inputs=post_inputs)
        result = observe_b4_query(self.candidate, patch, readback, query, **observer_kwargs)
        self.assertEqual(result["status"], "B4_QUERY_OBSERVED_UNVERIFIED")
        wrong = copy.deepcopy(query)
        wrong["rows"][0]["Source Row ID"] = "test-only-other-row"
        with self.assertRaisesRegex(ValueError, "physical row"):
            observe_b4_query(self.candidate, patch, readback, wrong, **observer_kwargs)
        forged = copy.deepcopy(self.candidate)
        next(c for c in forged["elements"][2]["columns"] if c["name"] == "B4 Number")["formula"] = "1234567"
        with self.assertRaisesRegex(ValueError, "independently reviewed graph"):
            observe_b4_query(forged, patch, readback, query, **observer_kwargs)


if __name__ == "__main__":
    unittest.main()
