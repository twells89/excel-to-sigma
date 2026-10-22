"""Only one declared scaffold delta may become a trusted assigned-ID receipt."""
import copy
import unittest
from pathlib import Path
import sys
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from native_exact_plan import confirm_exact_native_scaffold

class ScaffoldConservationTests(unittest.TestCase):
    def fixture(self):
        before={'kind':'workbook','schemaVersion':1,'pages':[{'id':'page','name':'Owned Page'},{'id':'foreign-page','name':'Foreign'}],
          'elements':[{'id':'foreign','kind':'text','body':'Preserve'}],
          'layout':'<Page id="page"><Element elementId="foreign" gridColumn="1 / 2" gridRow="1 / 2" /></Page><Page id="foreign-page" />',
          'privateState':{'owner':'keep','nested':[1,2]},'events':[{'id':'keep-event'}]}
        proposal={'status':'REVIEWABLE_MCP_SCAFFOLD_PENDING_EXECUTION','documentId':'doc','documentVersion':7,'sourceSha256':'source',
          'nativePlanHash':'native','manifestPlanHash':'manifest','logicalElementId':'logical-new','name':'Native Exact 1','sourceElementId':'actual-seed',
          'pageId':'page','frame':{'gridColumn':'1 / 25','gridRow':'13 / 25'},
          'logicalColumns':[{'id':'logical-column','name':'native_key'}], 'beforeContents':before,
          'beforeEnvelope':{'workbookId':'doc','owner':'unchanged'}}
        saved={'workbookId':'doc','documentVersion':8,'owner':'unchanged','contents':copy.deepcopy(before)}
        saved['contents']['elements'].append({'id':'assigned-new','kind':'table','name':'Native Exact 1',
          'source':{'kind':'table','elementId':'actual-seed'},'columns':[{'id':'assigned-column','name':'native_key','formula':'0'}]})
        saved['contents']['layout']=before['layout'].replace('</Page>','<Element elementId="assigned-new" gridColumn="1 / 25" gridRow="13 / 25" /></Page>',1)
        return proposal,saved
    def test_exact_declared_addition_and_xml_lexical_normalization_are_allowed(self):
        proposal,saved=self.fixture()
        self.assertEqual(confirm_exact_native_scaffold(proposal,saved)['elementId'],'assigned-new')
        saved['contents']['layout']=saved['contents']['layout'].replace('elementId="assigned-new" gridColumn="1 / 25" gridRow="13 / 25"','gridRow="13 / 25" elementId="assigned-new" gridColumn="1 / 25"')
        self.assertEqual(confirm_exact_native_scaffold(proposal,saved)['elementId'],'assigned-new')
    def test_nonowned_envelope_pages_private_state_and_elements_are_conserved(self):
        def extra_element(x):x['contents']['elements'].append({'id':'other-new','kind':'text','body':'Injected'})
        def duplicate_element(x):x['contents']['elements'].append(copy.deepcopy(x['contents']['elements'][0]))
        mutations={'envelope':lambda x:x.update(owner='changed'),'extra-envelope':lambda x:x.update(unrelated=True),
          'page':lambda x:x['contents']['pages'][0].update(name='Changed'),
          'foreign-page':lambda x:x['contents']['pages'][1].update(name='Changed'),
          'privateState':lambda x:x['contents']['privateState'].update(owner='Changed'),
          'extra-element':extra_element,'duplicate-element':duplicate_element,
          'unknown-top-level':lambda x:x['contents'].update(unrelated=True),
          'foreign-element':lambda x:x['contents']['elements'][0].update(body='Changed')}
        for name,mutate in mutations.items():
            proposal,saved=self.fixture();mutate(saved)
            with self.subTest(name=name),self.assertRaises(ValueError):confirm_exact_native_scaffold(proposal,saved)
    def test_columns_source_and_new_element_are_exactly_declared(self):
        def extra_column(x):x['contents']['elements'][-1]['columns'].append({'id':'extra','name':'extra','formula':'0'})
        mutations={'extra-column':extra_column,'changed-formula':lambda x:x['contents']['elements'][-1]['columns'][0].update(formula='99'),
          'extra-column-property':lambda x:x['contents']['elements'][-1]['columns'][0].update(unrelated=True),
          'source':lambda x:x['contents']['elements'][-1]['source'].update(elementId='wrong'),
          'extra-element-property':lambda x:x['contents']['elements'][-1].update(unrelated=True)}
        for name,mutate in mutations.items():
            proposal,saved=self.fixture();mutate(saved)
            with self.subTest(name=name),self.assertRaises(ValueError):confirm_exact_native_scaffold(proposal,saved)
    def test_only_declared_placement_is_added_on_declared_page(self):
        mutations={'missing':lambda s:s.split('<Element elementId="assigned-new"')[0]+'</Page><Page id="foreign-page" />',
          'wrong-frame':lambda s:s.replace('gridRow="13 / 25"','gridRow="1 / 13"'),
          'foreign-placement':lambda s:s.replace('gridColumn="1 / 2"','gridColumn="2 / 3"'),
          'extra-layout-element':lambda s:s.replace('<Page id="foreign-page" />','<Page id="foreign-page"><Element elementId="other" /></Page>'),
          'unknown-attribute':lambda s:s.replace('elementId="assigned-new"','elementId="assigned-new" style="bad"')}
        for name,mutate in mutations.items():
            proposal,saved=self.fixture();saved['contents']['layout']=mutate(saved['contents']['layout'])
            with self.subTest(name=name),self.assertRaises(ValueError):confirm_exact_native_scaffold(proposal,saved)

if __name__=='__main__':unittest.main()
