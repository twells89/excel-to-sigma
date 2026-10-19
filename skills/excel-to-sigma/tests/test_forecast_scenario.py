"""Synthetic source/capture rows test the planner, never published Sigma behavior."""
import copy
import ast
import builtins
import json
from pathlib import Path
import subprocess
import sys
import symtable
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/"scripts/lib"))
import forecast_scenario as scenario


class SaveScenarioTests(unittest.TestCase):
    def setUp(self):
        tmp=tempfile.TemporaryDirectory();self.addCleanup(tmp.cleanup)
        self.folder=Path(tmp.name)
        self.ids={key:digit*32 for key,digit in {"op":"1","saved":"2","sub":"3","usage":"4","other":"5"}.items()}
        self.definition={"version":1,"workflow":"saveScenario","context":"forecast-edit",
            "sourceHash":"a"*64,"planHash":"b"*64,"procedureHash":"c"*64,"decisionId":"reviewed",
            "sourceSemantics":{"appendBeforeNamePrompt":True,"allowBlankName":True,"allowWhitespaceName":True,
                               "growthSources":["Assumptions!B5","Assumptions!B6"],
                               "blankNamePolicy":copy.deepcopy(scenario.BLANK_NAME_POLICY)},
            "authoringTarget":{"documentId":"doc","version":16,"contentsHash":"e"*64},
            "primary":{"elementId":"primary","pageId":"edit","rowIdColumnId":"id_primary",
                       "expectedRowIds":sorted([self.ids["sub"],self.ids["usage"],self.ids["other"]]),
                       "columnIds":{"Line Item":"line","Monthly Growth":"growth"},"identityReadbackHash":"read"},
            "fields":[{"sourceCell":"Assumptions!B5","feature":"forecastBusinessInput","keyColumn":"Line Item",
                       "key":"Subscription Revenue","valueColumn":"Monthly Growth","rowId":self.ids["sub"],
                       "beforeField":"Before Subscription Growth","targetField":"Target Subscription Growth",
                       "savedField":"Subscription Growth"},
                      {"sourceCell":"Assumptions!B6","feature":"forecastBusinessInput","keyColumn":"Line Item",
                       "key":"Usage Revenue","valueColumn":"Monthly Growth","rowId":self.ids["usage"],
                       "beforeField":"Before Usage Growth","targetField":"Target Usage Growth",
                       "savedField":"Usage Growth"}],
            "bindingHash":"bound","productionStatus":"BLOCKED"}
        self.definition["definitionHash"]=scenario.digest(self.definition)
        self.target={"documentId":"doc","version":17,"operationsElementId":"operations",
            "operationsRowIdColumnId":"id_ops","savedElementId":"saved","savedRowIdColumnId":"id_saved",
            "primaryElementId":"primary","primaryRowIdColumnId":"id_primary",
            "controlId":"operation-control","nameControlId":"name-control"}
        self.raw_time="2026-10-19 15:04:05.723"
        self.op={"ID":self.ids["op"],"Workflow":"saveScenario","Scenario Name":"",
            "Context":"forecast-edit","Source Hash":"a"*64,"Plan Hash":"b"*64,
            "Source Procedure Hash":"c"*64,"State":"PREPARED","Created At":self.raw_time,
            "Subscription Input Row ID":self.ids["sub"],"Usage Input Row ID":self.ids["usage"],
            "Before Subscription Growth":0.04,"Before Usage Growth":0.02,
            "Target Subscription Growth":0.04,"Target Usage Growth":0.02}
        self.saved={"ID":self.ids["saved"],"Operation ID":self.ids["op"],"Scenario Name":"",
            "Saved At":self.raw_time,"Subscription Growth":0.04,"Usage Growth":0.02,
            "Source Hash":"a"*64,"Plan Hash":"b"*64,"Source Procedure Hash":"c"*64}
        self.primary=[{"ID":self.ids["sub"],"Line Item":"Subscription Revenue","Method":"growth","Monthly Growth":0.04},
                      {"ID":self.ids["usage"],"Line Item":"Usage Revenue","Method":"growth","Monthly Growth":0.02},
                      {"ID":self.ids["other"],"Line Item":"Other Revenue","Method":"manual","Monthly Growth":0.03}]
        op_types={"Created At":"datetime",**{key:"number" for key in self.op if "Growth" in key}}
        saved_types={"Saved At":"datetime","Subscription Growth":"number","Usage Growth":"number"}
        def audit_columns(row,typ,prefix):
            return ([{"id":prefix,"name":"ID","class":"InputTableIDColumn","kind":"row_id"}]+
                    [{"id":prefix+"_"+str(i),"name":name,"class":"LiteralEditableColumn",
                      "type":typ.get(name,"text")} for i,name in enumerate(row) if name!="ID"])
        self.columns={"operations":audit_columns(self.op,op_types,"id_ops"),
                      "saved":audit_columns(self.saved,saved_types,"id_saved"),
                      "primary":[{"id":"id_primary","name":"ID","class":"FormulaColumn","formula":"[Forecast/ID]"},
                                 {"id":"line","name":"Line Item","class":"LiteralEditableColumn","type":"text"},
                                 {"id":"method","name":"Method","class":"LiteralEditableColumn","type":"text"},
                                 {"id":"growth","name":"Monthly Growth","class":"LiteralEditableColumn","type":"number"}]}
        # Sparse Operations may have hundreds of assigned columns whose unrelated values are null/omitted.
        self.columns["operations"] += [{"id":"extra_"+str(i),"name":"Extra "+str(i),
                                        "class":"LiteralEditableColumn","type":"text"} for i in range(214)]
        self.definition["audit"]={role:{"elementId":role,"rowIdColumnId":"id_ops" if role=="operations" else "id_saved",
            "columns":{name:{"id":col["id"],"type":col["type"],"class":col["class"]}
                       for name in scenario._audit_fields(self.definition["fields"])[role]
                       for col in self.columns[role] if col["name"]==name}}
            for role in ("operations","saved")}
        self.definition["auditMappingHash"]="f"*64
        self.definition["definitionHash"]=scenario.digest({k:v for k,v in self.definition.items() if k!="definitionHash"})
        self.bindings={role:{"elementId":role,"pageId":"edit" if role=="primary" else "hidden",
                             "name":{"primary":"Forecast Input","operations":"Workflow Operations",
                                     "saved":"Saved Scenarios"}[role],"columns":self.columns[role]}
                       for role in self.columns}
        elements=[]
        for role,b in self.bindings.items():
            elements.append({"id":role,"kind":"input-table","name":b["name"],
                "columns":[({"id":"ID"} if col["name"]=="ID" and role!="primary" else
                            {k:col[k] for k in ("id","name","type","formula") if k in col})
                           for col in b["columns"]]})
        self.contents={"kind":"workbook","pages":[{"id":"edit","name":"Edit"},
                         {"id":"hidden","name":"Hidden","visibility":"hidden"}],
            "elements":elements+[{"id":"operation-control","kind":"control","controlType":"text",
                                  "controlId":"saveScenario-operation"},
                                 {"id":"name-control","kind":"control","controlType":"text",
                                  "controlId":"saveScenario-name"}],
            "layout":"<Page id='edit'><Element id='primary' /><Element id='name-control' /></Page>"
                     "<Page id='hidden'><Element id='operations' /><Element id='saved' />"
                     "<Element id='operation-control' /></Page>"}
        self.source={"workbookId":"doc","documentVersion":17,"contents":self.contents}
        self.write("source-before.json",self.source);self.write("source-after.json",self.source)
        self.write("bindings.json",{role:{"elementId":role,"columns":self.columns[role]} for role in ("operations","saved")})
        self.write("timezone.json",{"timezone":"UTC"})
        self.write("request-provenance.json",{"qualification":"actual calls retained; controller review pending",
            "operations":{"queryId":"query-ops","actualExportInvocation":"call-ops","actualDownloadInvocation":"download-ops"},
            "saved":{"queryId":"query-saved","actualExportInvocation":"call-saved","actualDownloadInvocation":"download-saved"}})
        self.write("primary-binding.json",{"elementId":"primary","rowIdColumnId":"id_primary",
                                          "columns":self.columns["primary"]})

    def write(self,name,value):
        path=self.folder/name
        path.write_text(json.dumps(value),encoding="utf-8")
        return path

    def captures(self,*,op=None,saved=None,primary=None,name=None):
        op=copy.deepcopy(self.op if op is None else op)
        saved=copy.deepcopy(self.saved if saved is None else saved)
        primary=copy.deepcopy(self.primary if primary is None else primary)
        if name is not None:
            op["Scenario Name"]=name;saved["Scenario Name"]=name
        results={}
        for role,row,job in (("operations",op,"query-ops"),("saved",saved,"query-saved")):
            self.write(role+"-request.json",{"method":"POST","path":"/v2/workbooks/doc/export",
                "body":{"elementId":role,"format":{"type":"json"},"rowLimit":1000}})
            self.write(role+"-job.json",{"queryId":job,"jobComplete":True})
            self.write(role+"-download.json",{"method":"GET","path":"/v2/query/"+job+"/download"})
            self.write(role+"-rows.json",row if isinstance(row,list) else [row])
            results[role]=scenario.adapt_save_rest_capture(self.definition,role=role,target=self.target,
                bindings_path=self.folder/"bindings.json",request_path=self.folder/(role+"-request.json"),
                job_path=self.folder/(role+"-job.json"),download_request_path=self.folder/(role+"-download.json"),
                request_provenance_path=self.folder/"request-provenance.json",
                rows_path=self.folder/(role+"-rows.json"),timezone_path=self.folder/"timezone.json",
                source_before_path=self.folder/"source-before.json",source_after_path=self.folder/"source-after.json")
        columns=self.columns["primary"]
        ids=[col["id"] for col in columns]
        sql='SELECT '+', '.join('"'+cid+'"' for cid in ids)+' FROM "workbook"."primary"'
        content={"columns":ids,"rows":[[row[col["name"]] for col in columns] for row in primary]}
        self.write("primary-query.json",{"request":{"type":"workbook","workbookId":"doc","sql":sql},
            "executionId":"query-primary","actualInvocation":"sigma query invocation",
            "result":{"structuredContent":content,"content":[{"text":json.dumps(content)}]}})
        results["primary"]=scenario.adapt_save_primary_capture(self.definition,target=self.target,
            binding_path=self.folder/"primary-binding.json",query_path=self.folder/"primary-query.json",
            source_before_path=self.folder/"source-before.json",source_after_path=self.folder/"source-after.json")
        return results

    def plan(self,captures,*,name="",control=None):
        return scenario.plan_save_completion(self.definition,self.ids["op"],captures["operations"],
            captures["saved"],captures["primary"],self.target,name_control=name,
            control_value=self.ids["op"] if control is None else control)

    def test_blank_whitespace_terminal_retry_and_disabled_renderer(self):
        for name in ("","   "):
            captures=self.captures(name=name)
            plan=self.plan(captures,name=name)
            self.assertEqual(plan["nextAction"],"MARK_CONFIRMED")
            current={"workbookId":"doc","documentVersion":17,"contents":self.contents}
            rendered=scenario.render_save_completion_candidate(self.definition,current,self.bindings,plan,
                operations=captures["operations"],saved=captures["saved"],primary=captures["primary"],phase="mark")
            self.assertIn("const_date('2026-10-19T15:04:05.723+00:00')",rendered["program"])
            self.assertIn("update_rows(t_operations",rendered["program"])
            self.assertIn("sequence.set_disabled(True)",rendered["program"])
            self.assertEqual(rendered["productionStatus"],"BLOCKED")
            terminal=self.captures(op={**self.op,"State":"CONFIRMED"},name=name)
            clear=self.plan(terminal,name=name)
            self.assertEqual(clear["nextAction"],"CLEAR_CONTROL")
            rendered=scenario.render_save_completion_candidate(self.definition,current,self.bindings,clear,
                operations=terminal["operations"],saved=terminal["saved"],primary=terminal["primary"],phase="clear")
            self.assertIn("set_control_value(op, '')",rendered["program"])
            self.assertNotIn("update_rows(t_operations",rendered["program"])
            self.assertEqual(self.plan(terminal,name=name,control="")["nextAction"],"NONE")

    def test_wrong_operation_name_provenance_growth_time_and_duplicate_saved_fail(self):
        cases=[("wrong workflow",{"op":{"Workflow":"applyScenario"}}),
               ("wrong name",{"name":"Different"}),
               ("source",{"op":{"Source Hash":"d"*64}}),
               ("plan",{"saved":{"Plan Hash":"d"*64}}),
               ("procedure",{"saved":{"Source Procedure Hash":"d"*64}}),
               ("wrong input ID",{"op":{"Subscription Input Row ID":self.ids["other"]}}),
               ("stale growth",{"primary":[{**self.primary[0],"Monthly Growth":0.05},*self.primary[1:]]}),
               ("text numeric",{"saved":{"Subscription Growth":"0.04"}}),
               ("boolean numeric",{"op":{"Before Usage Growth":True}}),
               ("nonfinite numeric",{"saved":{"Usage Growth":float("inf")}}),
               ("wrong timestamp",{"saved":{"Saved At":"2026-10-19 15:04:05.72"}}),
               ("invalid timestamp",{"op":{"Created At":"2026-10-32 15:04:05.723"}})]
        for label,changes in cases:
            with self.subTest(label=label),self.assertRaises(ValueError):
                op={**self.op,**changes.get("op",{})}
                saved={**self.saved,**changes.get("saved",{})}
                captures=self.captures(op=op,saved=saved,primary=changes.get("primary"))
                self.plan(captures,name=changes.get("name",""))
        with self.assertRaises(ValueError):
            captures=self.captures(saved=[self.saved,{**self.saved,"ID":self.ids["other"]}]);self.plan(captures)
        with self.assertRaises(ValueError):
            captures=self.captures(op={**self.op,"ID":"not-physical"});self.plan(captures)

    def test_capture_provenance_and_control_reject_tampering(self):
        captures=self.captures();self.assertEqual(self.plan(captures)["nextAction"],"MARK_CONFIRMED")
        with self.assertRaises(ValueError):self.plan(captures,control="")
        changed=copy.deepcopy(captures)
        changed["saved"]["rows"][0]["Scenario Name"]="tampered"
        with self.assertRaises(ValueError):self.plan(changed)
        self.write("saved-download.json",{"method":"GET","path":"/v2/query/wrong/download"})
        with self.assertRaises(ValueError):self.plan(captures)

    def test_same_id_retry_and_deliberate_new_operation(self):
        first=self.captures()
        self.assertEqual(self.plan(first),self.plan(first))
        next_id="6"*32;next_saved_id="7"*32
        old={**self.op,"State":"CONFIRMED"}
        new={**self.op,"ID":next_id,"Scenario Name":"  "}
        older_saved=self.saved
        newer_saved={**self.saved,"ID":next_saved_id,"Operation ID":next_id,"Scenario Name":"  "}
        captures=self.captures(op=[old,new],saved=[older_saved,newer_saved])
        next_plan=scenario.plan_save_completion(self.definition,next_id,captures["operations"],
            captures["saved"],captures["primary"],self.target,name_control="  ",control_value=next_id)
        self.assertEqual(next_plan["nextAction"],"MARK_CONFIRMED")
        self.assertEqual(next_plan["savedPhysicalRowId"],next_saved_id)
        with self.assertRaises(ValueError):
            scenario.plan_save_completion(self.definition,next_id,captures["operations"],
                captures["saved"],captures["primary"],self.target,name_control="  ",
                control_value=self.ids["op"])

    def test_cli_plans_only_from_retained_capture_files(self):
        captures=self.captures()
        self.write("definition.json",self.definition)
        self.write("target.json",self.target)
        for role,capture in captures.items():self.write(role+"-capture.json",capture)
        script=Path(__file__).resolve().parents[1]/"scripts/forecast-scenario.py"
        output=self.folder/"completion.json"
        command=[sys.executable,str(script),"plan","--definition",str(self.folder/"definition.json"),
            "--operation-id",self.ids["op"],"--operations",str(self.folder/"operations-capture.json"),
            "--saved",str(self.folder/"saved-capture.json"),"--primary",str(self.folder/"primary-capture.json"),
            "--target",str(self.folder/"target.json"),"--name-control","","--control-value",self.ids["op"],
            "--out",str(output)]
        subprocess.run(command,check=True,capture_output=True,text=True)
        self.assertEqual(json.loads(output.read_text())["nextAction"],"MARK_CONFIRMED")
        self.write("saved-download.json",{"method":"GET","path":"/v2/query/wrong/download"})
        failed=subprocess.run(command,capture_output=True,text=True)
        self.assertNotEqual(failed.returncode,0)

    def test_actual_blank_null_or_omitted_name_normalizes_only_reviewed_field(self):
        for variant in ("omitted","null"):
            op=copy.deepcopy(self.op);saved=copy.deepcopy(self.saved)
            if variant=="omitted":
                op.pop("Scenario Name");saved.pop("Scenario Name")
            else:
                op["Scenario Name"]=None;saved["Scenario Name"]=None
            with self.subTest(variant=variant):
                captures=self.captures(op=op,saved=saved)
                for role in ("operations","saved"):
                    self.assertEqual(captures[role]["restEvidence"]["nameNormalization"][0]["rawForm"],variant)
                plan=scenario.plan_save_completion(self.definition,self.ids["op"],
                    captures["operations"],captures["saved"],captures["primary"],self.target,
                    name_control=None,control_value=self.ids["op"])
                self.assertEqual(plan["nameControl"],"")
                self.assertIsNone(plan["nameControlObserved"])
                program=scenario.render_save_completion_candidate(self.definition,
                    {"workbookId":"doc","documentVersion":17,"contents":self.contents},
                    self.bindings,plan,operations=captures["operations"],saved=captures["saved"],
                    primary=captures["primary"],phase="mark")["program"]
                self.assertIn("as_text(coalesce(name, const_text('')))",program)
                self.assertIn("as_text(coalesce(lookup(c_operations_",program)
                self.assertIn("as_text(coalesce(lookup(c_saved_",program)
                self.assertNotIn("set_control_value(name",program)
        bad=copy.deepcopy(self.op);bad.pop("Created At")
        with self.assertRaises(ValueError):self.captures(op=bad)

    def test_generated_mark_and_clear_have_no_unbound_formula_helper(self):
        for state,phase in (("PREPARED","mark"),("CONFIRMED","clear")):
            captures=self.captures(op={**self.op,"State":state})
            plan=self.plan(captures)
            program=scenario.render_save_completion_candidate(self.definition,
                {"workbookId":"doc","documentVersion":17,"contents":self.contents},self.bindings,plan,
                operations=captures["operations"],saved=captures["saved"],
                primary=captures["primary"],phase=phase)["program"]
            tree=ast.parse(program)
            imported={alias.asname or alias.name for node in ast.walk(tree)
                      if isinstance(node,ast.ImportFrom) and node.module=="sigma.formula"
                      for alias in node.names}
            self.assertIn("as_text",imported)
            symbols=symtable.symtable(program,"generated-"+phase,"exec")
            unbound={s.get_name() for s in symbols.get_symbols() if s.is_referenced()
                     and not (s.is_imported() or s.is_assigned() or s.get_name() in dir(builtins)
                              or s.get_name()=="sigma")}
            self.assertEqual(unbound,set(),unbound)

    def test_nullable_row_name_requires_approved_declared_audit_schema(self):
        original_columns=copy.deepcopy(self.columns)
        original_contents=copy.deepcopy(self.contents)
        for role in ("operations","saved"):
            for variant in ("missing","number","select","foreign-id","duplicate-name",
                            "public-duplicate-name","missing-growth"):
                with self.subTest(role=role,variant=variant):
                    self.columns=copy.deepcopy(original_columns)
                    self.contents=copy.deepcopy(original_contents)
                    self.source={"workbookId":"doc","documentVersion":17,"contents":self.contents}
                    private=next(c for c in self.columns[role] if c["name"]=="Scenario Name")
                    public=next(c for c in next(e for e in self.contents["elements"]
                        if e["id"]==role)["columns"] if c.get("name")=="Scenario Name")
                    if variant=="missing":
                        self.columns[role].remove(private)
                        next(e for e in self.contents["elements"] if e["id"]==role)["columns"].remove(public)
                    elif variant=="number":
                        private["type"]="number";public["type"]="number"
                    elif variant=="select":
                        private["class"]="SingleSelectEditableColumn"
                    elif variant=="foreign-id":
                        private["id"]="foreign_name_id";public["id"]="foreign_name_id"
                    elif variant=="duplicate-name":
                        self.columns[role].append({**private,"id":"duplicate_name_id"})
                        next(e for e in self.contents["elements"] if e["id"]==role)["columns"].append(
                            {**public,"id":"duplicate_name_id"})
                    elif variant=="public-duplicate-name":
                        next(e for e in self.contents["elements"] if e["id"]==role)["columns"].append(
                            {**public,"id":"duplicate_public_name_id"})
                    else:
                        growth="Target Subscription Growth" if role=="operations" else "Subscription Growth"
                        missing=next(c for c in self.columns[role] if c["name"]==growth)
                        self.columns[role].remove(missing)
                        public_growth=next(c for c in next(e for e in self.contents["elements"]
                            if e["id"]==role)["columns"] if c.get("name")==growth)
                        next(e for e in self.contents["elements"] if e["id"]==role)["columns"].remove(public_growth)
                    self.write("bindings.json",{r:{"elementId":r,"columns":self.columns[r]}
                                                 for r in ("operations","saved")})
                    self.write("source-before.json",self.source);self.write("source-after.json",self.source)
                    op=copy.deepcopy(self.op);saved=copy.deepcopy(self.saved)
                    (op if role=="operations" else saved).pop("Scenario Name")
                    with self.assertRaisesRegex(ValueError,"audit|Scenario Name|schema|column"):
                        captures=self.captures(op=op,saved=saved);self.plan(captures)

    def test_actual_primary_method_select_class_requires_exact_options_in_renderer(self):
        # The observed native Forecast Business Inputs Method is a select, not a literal.
        method=next(c for c in self.columns["primary"] if c["name"]=="Method")
        method["class"]="SingleSelectEditableColumn"
        method["allowedValues"]=["growth","manual","flat"]
        self.write("primary-binding.json",{"elementId":"primary","rowIdColumnId":"id_primary",
                                           "columns":self.columns["primary"]})
        captures=self.captures();plan=self.plan(captures)
        current={"workbookId":"doc","documentVersion":17,"contents":self.contents}
        program=scenario.render_save_completion_candidate(self.definition,current,self.bindings,plan,
            operations=captures["operations"],saved=captures["saved"],primary=captures["primary"],
            phase="mark")["program"]
        self.assertIn("SingleSelectEditableColumn",program)
        self.assertIn("allowed_values() != ['growth', 'manual', 'flat']",program)
        del method["allowedValues"]
        with self.assertRaisesRegex(ValueError,"select|allowed"):
            scenario.render_save_completion_candidate(self.definition,current,self.bindings,plan,
                operations=captures["operations"],saved=captures["saved"],primary=captures["primary"],
                phase="mark")

    def test_sparse_operations_render_checks_complete_schema_offline_but_emits_bounded_program(self):
        current={"workbookId":"doc","documentVersion":17,"contents":self.contents}
        for state,phase in (("PREPARED","mark"),("CONFIRMED","clear")):
            with self.subTest(phase=phase):
                captures=self.captures(op={**self.op,"State":state})
                plan=self.plan(captures)
                program=scenario.render_save_completion_candidate(self.definition,current,self.bindings,plan,
                    operations=captures["operations"],saved=captures["saved"],
                    primary=captures["primary"],phase=phase)["program"]
                self.assertLess(len(program.encode("utf-8")),35000)
                self.assertNotIn("get_column('extra_0')",program)
                for role in ("operations","saved"):
                    for approved in self.definition["audit"][role]["columns"].values():
                        self.assertIn("get_column("+repr(approved["id"])+")",program)
                for cid in ("id_ops","id_saved","id_primary","line","method","growth"):
                    self.assertIn("get_column("+repr(cid)+")",program)
                self.assertIn("sequence.set_disabled(True)",program)

    def test_cli_accepts_actual_null_name_observation(self):
        op=copy.deepcopy(self.op);saved=copy.deepcopy(self.saved)
        op.pop("Scenario Name");saved.pop("Scenario Name")
        captures=self.captures(op=op,saved=saved)
        self.write("definition.json",self.definition);self.write("target.json",self.target)
        for role,capture in captures.items():self.write(role+"-capture.json",capture)
        script=Path(__file__).resolve().parents[1]/"scripts/forecast-scenario.py"
        output=self.folder/"null-completion.json"
        command=[sys.executable,str(script),"plan","--definition",str(self.folder/"definition.json"),
            "--operation-id",self.ids["op"],"--operations",str(self.folder/"operations-capture.json"),
            "--saved",str(self.folder/"saved-capture.json"),"--primary",str(self.folder/"primary-capture.json"),
            "--target",str(self.folder/"target.json"),"--name-control-null",
            "--control-value",self.ids["op"],"--out",str(output)]
        subprocess.run(command,check=True,capture_output=True,text=True)
        self.assertEqual(json.loads(output.read_text())["nameControlObserved"],None)
        self.assertEqual(json.loads(output.read_text())["nextAction"],"MARK_CONFIRMED")

    def test_primary_query_must_use_definition_reviewed_growth_column_id(self):
        changed=copy.deepcopy(self.columns["primary"])
        changed[-1]["id"]="foreign_growth"
        self.columns["primary"]=changed
        next(item for item in self.contents["elements"] if item["id"]=="primary")["columns"][-1]["id"]="foreign_growth"
        self.write("source-before.json",self.source);self.write("source-after.json",self.source)
        self.write("primary-binding.json",{"elementId":"primary","rowIdColumnId":"id_primary",
                                           "columns":changed})
        with self.assertRaisesRegex(ValueError,"reviewed|column"):
            captures=self.captures();self.plan(captures)

    def test_renderer_requires_exact_fresh_bracketed_current_source(self):
        captures=self.captures();plan=self.plan(captures)
        current={"workbookId":"doc","documentVersion":17,"contents":copy.deepcopy(self.contents)}
        current["contents"]["elements"].append({"id":"foreign","kind":"text","name":"Changed"})
        with self.assertRaisesRegex(ValueError,"source|contents"):
            scenario.render_save_completion_candidate(self.definition,current,self.bindings,plan,
                operations=captures["operations"],saved=captures["saved"],
                primary=captures["primary"],phase="mark")

    def test_definition_uses_reviewed_mapping_and_complete_bound_identities(self):
        binding={"sourceHash":"a"*64,"planHash":"b"*64,"target":{"documentId":"doc","version":17},
            "contentsHash":"e"*64,
            "tables":{"forecastBusinessInput":{"target":{"elementId":"primary","pageId":"edit",
                "columnIds":{"Line Item":"line","Monthly Growth":"growth"}},
                "systemColumn":{"id":"id_primary"},"readbackHash":"read",
                "identities":[{"rowId":self.ids["sub"],"key":{"Line Item":"Subscription Revenue"}},
                              {"rowId":self.ids["usage"],"key":{"Line Item":"Usage Revenue"}},
                              {"rowId":self.ids["other"],"key":{"Line Item":"Other Revenue"}}]}}}
        binding["bindingHash"]=scenario.digest(binding)
        audit_mapping={"version":1,"sourceHash":"a"*64,"planHash":"b"*64,
            "target":{"documentId":"doc","version":17},"tables":{
                role:{"elementId":role,"rowIdColumnId":self.definition["audit"][role]["rowIdColumnId"],
                      "columns":self.columns[role]}
                for role in ("operations","saved")}}
        audit_mapping["mappingHash"]=scenario.digest(audit_mapping)
        decision={"decisionId":"reviewed","page":"forecast-edit","assumptions":[
            {"sourceCell":field["sourceCell"],"feature":"forecastBusinessInput",
             "keyColumn":"Line Item","key":field["key"],"valueColumn":"Monthly Growth"}
            for field in self.definition["fields"]]}
        source='''Sub SaveScenario_Click()
Set newRow = t.ListRows.Add
newRow.Range(1, 1).Value = InputBox("Scenario name?")
newRow.Range(1, 2).Value = Sheets("Assumptions").Range("B5").Value
newRow.Range(1, 3).Value = Sheets("Assumptions").Range("B6").Value
End Sub'''
        manifest={"sourceHash":"a"*64,"planHash":"b"*64,"target":{"documentId":"doc","version":17},"planProjection":{
            "nativeActionArchitecture":{"sourceModules":[{"procedures":[{"procedureId":"save",
                "procedureSha256":"c"*64,"sourceBody":source,"guards":[]}]}],
                "workflows":{"saveScenario":{"status":"STAGED_CANDIDATE","procedureId":"save",
                "procedureSha256":"c"*64,"decision":decision}}}}}
        with patch.object(scenario,"validate_native_action_plan",return_value=True):
            definition=scenario.build_save_definition(manifest,binding,audit_mapping)
        self.assertEqual(definition["primary"]["expectedRowIds"],sorted([self.ids["sub"],self.ids["usage"],self.ids["other"]]))
        self.assertEqual([f["savedField"] for f in definition["fields"]],["Subscription Growth","Usage Growth"])
        self.assertEqual(definition["sourceSemantics"]["allowBlankName"],True)
        self.assertEqual(definition["audit"]["saved"]["columns"]["Scenario Name"]["id"],
                         self.definition["audit"]["saved"]["columns"]["Scenario Name"]["id"])
        incomplete=copy.deepcopy(audit_mapping)
        incomplete["tables"]["saved"]["columns"]=[c for c in incomplete["tables"]["saved"]["columns"]
            if c["name"]!="Source Procedure Hash"]
        incomplete["mappingHash"]=scenario.digest({k:v for k,v in incomplete.items() if k!="mappingHash"})
        with patch.object(scenario,"validate_native_action_plan",return_value=True),self.assertRaisesRegex(ValueError,"audit"):
            scenario.build_save_definition(manifest,binding,incomplete)
        manifest["planProjection"]["nativeActionArchitecture"]["sourceModules"][0]["procedures"][0]["sourceBody"] = source.replace(
            "Set newRow = t.ListRows.Add", "If name = '' Then Exit Sub\nSet newRow = t.ListRows.Add")
        with patch.object(scenario,"validate_native_action_plan",return_value=True),self.assertRaises(ValueError):
            scenario.build_save_definition(manifest,binding,audit_mapping)
        binding["target"]["documentId"]="foreign";binding["bindingHash"]=scenario.digest({k:v for k,v in binding.items() if k!="bindingHash"})
        manifest["planProjection"]["nativeActionArchitecture"]["sourceModules"][0]["procedures"][0]["sourceBody"]=source
        with patch.object(scenario,"validate_native_action_plan",return_value=True),self.assertRaisesRegex(ValueError,"target|document"):
            scenario.build_save_definition(manifest,binding,audit_mapping)


if __name__=="__main__":unittest.main()
