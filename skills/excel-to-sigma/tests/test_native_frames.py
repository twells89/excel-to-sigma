"""Native page-grid frames must be checked before mutable SDK calls."""
import ast
import copy
from pathlib import Path
import sys
import unittest
SCRIPTS=Path(__file__).resolve().parents[1]/'scripts'
sys.path[:0]=[str(SCRIPTS),str(SCRIPTS/'lib')]
from lib.native_frames import validate_grid_frame, validate_native_grid_program
import test_regional_inherited_columns as inherited
class NativeFrameTest(unittest.TestCase):
    def current(self):
        return {'contents':{'layout':'<Page type="grid" id="p" gridTemplateColumns="repeat(24, 1fr)"><Element elementId="rows" gridColumn="1 / 25" gridRow="130 / 142"/><Element elementId="flat" gridColumn="1 / 25" gridRow="145 / 157"/></Page>'}}
    def test_documented_one_based_bounds_and_exact_last_column(self):
        self.assertEqual(validate_grid_frame(1,145,24,12),[1,145,24,12])
        for frame in [(26,130,24,12),(0,130,24,12),(2,130,24,12),(1,0,24,12),(1,1,0,12),(True,1,24,12),(1,1,24.0,12)]:
            with self.subTest(frame=frame),self.assertRaisesRegex(ValueError,'BLOCKED'):validate_grid_frame(*frame)
    def test_invalid_frame_inside_false_body_is_still_blocked(self):
        with self.assertRaisesRegex(ValueError,'BLOCKED'):
            validate_native_grid_program('if False:\n    flat.set_frame(26,130,24,12)\n')
        with self.assertRaisesRegex(ValueError,'BLOCKED'):
            validate_native_grid_program('flat.set_frame(x,130,24,12)\n')
    def test_actual_page_overlap_and_missing_layout_are_blocked(self):
        for program in ['flat.set_frame(1,130,24,12)','flat.set_frame(1,140,24,12)','flat.set_frame(1,155,24,12)']:
            with self.subTest(program=program),self.assertRaisesRegex(ValueError,'BLOCKED'):
                validate_native_grid_program(program,current=self.current(),page_id='p')
        with self.assertRaisesRegex(ValueError,'BLOCKED'):
            validate_native_grid_program('flat.set_frame(1,160,24,12)',current={'contents':{}},page_id='p')
    def test_remaining_stacked_native_plan_has_no_overlap(self):
        programs='\n'.join('t'+str(i)+'.set_frame(1,'+str(row)+',24,12)' for i,row in enumerate([160,175,190]))
        self.assertEqual(len(validate_native_grid_program(programs,current=self.current(),page_id='p')),3)
        with self.assertRaisesRegex(ValueError,'BLOCKED'):
            validate_native_grid_program('a.set_frame(1,160,24,12)\nb.set_frame(1,165,24,12)',current=self.current(),page_id='p')
    def test_product_grouping_generator_emits_valid_full_width(self):
        inherited.InheritedColumnReceiptTest.setUpClass();t=inherited.InheritedColumnReceiptTest();t.setUp();confirmation,ddl,sdk=t.grouping_state()
        packet=t.group(confirmation,ddl,sdk)
        calls=[n for n in ast.walk(ast.parse(packet['program'])) if isinstance(n,ast.Call) and isinstance(n.func,ast.Attribute) and n.func.attr=='set_frame']
        self.assertEqual([[ast.literal_eval(a) for a in call.args] for call in calls],[[1,130,24,12]])
    def test_grouping_generator_rejects_actual_occupied_target_before_packet(self):
        import re
        inherited.InheritedColumnReceiptTest.setUpClass()
        t=inherited.InheritedColumnReceiptTest();t.setUp()
        page_id=t.current['contents']['pages'][0]['id']
        old=t.current['contents']['layout']
        pattern=r'(<Page[^>]*id="'+re.escape(page_id)+r'")/>'
        updated,count=re.subn(pattern,r'\1><Element elementId="occupied" gridColumn="1 / 25" gridRow="130 / 142"/></Page>',old)
        self.assertEqual(count,1)
        t.current['contents']['layout']=updated
        t.current['contents']['elements'].append({'id':'occupied','kind':'table','name':'Existing table','columns':[]})
        confirmation,ddl,sdk=t.grouping_state()
        with self.assertRaisesRegex(ValueError,'overlaps'):
            t.group(confirmation,ddl,sdk)
if __name__=='__main__':unittest.main()
