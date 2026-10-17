"""Failure injection for sequence construction; no Sigma runtime evidence."""
import ast
import copy
from pathlib import Path
import sys
import unittest
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
import native_actions
import test_native_public_actions as fixtures


class NativeDisabledConstructionTests(unittest.TestCase):
    def programs(self):
        helper=fixtures.PublicNativeActionsTests();self.addCleanup(helper.doCleanups)
        manifest,receipt,binding,current=helper.fixture()
        return {'authored':native_actions._native_action_authoring_plan(manifest,receipt,binding,current_contents=current)['program'],
                'executable':native_actions.emit_native_action_program(manifest,receipt,binding,current_contents=current)['program']}

    def factories(self,nodes):
        for index,node in enumerate(nodes):
            if (isinstance(node,ast.Assign) and len(node.targets)==1 and isinstance(node.targets[0],ast.Name)
                    and isinstance(node.value,ast.Call) and isinstance(node.value.func,ast.Attribute)
                    and node.value.func.attr in ('on_click','on_select')):
                yield index,node.targets[0].id,node

    def test_every_new_authored_and_executable_sequence_is_immediately_disabled(self):
        for kind,program in self.programs().items():
            nodes=ast.parse(program).body;factories=list(self.factories(nodes));self.assertTrue(factories)
            for index,name,node in factories:
                with self.subTest(kind=kind,sequence=name):
                    first=nodes[index+1]
                    self.assertIsInstance(first,ast.Expr)
                    self.assertEqual(ast.unparse(first.value.func),name+'.set_disabled')
                    self.assertEqual(ast.literal_eval(first.value.args[0]),True)

    def test_condition_failure_leaves_real_generated_sequence_disabled(self):
        class Sequence:
            def __init__(self):self.disabled=False
            def set_disabled(self,value):self.disabled=value
            def set_condition(self,_value):raise RuntimeError('injected condition failure')
        class Button:
            def on_click(self,*_args,**_kwargs):self.sequence=Sequence();return self.sequence
        nodes=ast.parse(self.programs()['executable']).body
        for index,name,factory in self.factories(nodes):
            if name not in ('prepare_sequence','new_sequence'):continue
            condition_index=next(i for i in range(index+1,len(nodes)) if isinstance(nodes[i],ast.Expr)
                and isinstance(nodes[i].value,ast.Call) and isinstance(nodes[i].value.func,ast.Attribute)
                and ast.unparse(nodes[i].value.func)==name+'.set_condition')
            snippet=ast.Module(body=copy.deepcopy(nodes[index:condition_index+1]),type_ignores=[])
            # Effect/Formula builders are inert arguments in this constructor
            # test; run the actual factory/disable/condition statement order.
            environment={node.id:None for node in ast.walk(snippet) if isinstance(node,ast.Name)}
            for node in ast.walk(snippet):
                if isinstance(node,ast.Call) and isinstance(node.func,ast.Name):environment[node.func.id]=lambda *_args,**_kwargs:None
            host=Button();environment[factory.value.func.value.id]=host
            with self.subTest(sequence=name):
                with self.assertRaisesRegex(RuntimeError,'injected condition failure'):
                    exec(compile(snippet,'generated-constructor','exec'),environment)
                self.assertTrue(host.sequence.disabled,'new sequence remained enabled after condition failure')

if __name__=='__main__':unittest.main()
