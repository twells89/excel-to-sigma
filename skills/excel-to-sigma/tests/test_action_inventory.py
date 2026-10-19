import copy,hashlib,sys,unittest
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'scripts/lib'))
from action_inventory import validate_action_inventory,verify_action_preservation,digest,configuration,empty_action_capture
class ActionInventoryTests(unittest.TestCase):
    def fixture(self):
        doc={'workbookId':'doc','documentVersion':7,'contents':{'elements':[{'id':'button','kind':'button'}]}}
        element={'elementId':'button','complete':True,'inspection':'list_action_sequences','sequences':[{
            'id':'seq1','name':'Prepare','trigger':{'kind':'click'},'disabled':False,
            'condition':{'kind':'opaque','representation':'BoolFormula','constructionArtifactHash':'known','bindingHash':'bound-assigned-ids'},
            'blocks':[{'kind':'InsertRowEffect','fields':{'tableId':'ops','values':{'state':'PREPARED'},'returns':'block0-rowid'}}], 'successToast':None}]}
        element['semanticHash']=digest(configuration(element))
        inventory={'documentId':'doc','version':7,'evidenceId':'actual-read','origin':'sigma-sdk-readonly','complete':True,'elements':{'button':element}}
        return doc,inventory
    def reviewed(self,inventory):
        result=copy.deepcopy(inventory);element=result['elements']['button'];program="# exact missing-sequence restoration, independently reviewed\nprint('restore')\n"
        element['restorationArtifact']={'elementId':'button','documentId':'doc','reviewer':'independent','sourceObservationHash':digest(element),
            'semanticHash':element['semanticHash'],'program':program,'programHash':hashlib.sha256(program.encode()).hexdigest(),
            'constructionArtifactHash':'known','bindingHash':'bound-assigned-ids','guardPolicy':'missing-sequences-only; exact-current-target-and-bindings; no-constructor-replay',
            'restoreProof':{'status':'PASSED_ISOLATED_RESTORATION','evidenceId':'actual-restore','documentId':'isolated','semanticHash':element['semanticHash'],'sdkReadbackEvidenceId':'restored-read','programHash':hashlib.sha256(program.encode()).hexdigest(),'bindingHash':'bound-assigned-ids','beforeMissingSequences':True}}
        return result
    def test_rest_omission_never_proves_absence_and_opaque_actions_require_actual_restoration(self):
        doc,inventory=self.fixture()
        with self.assertRaises(ValueError):validate_action_inventory(empty_action_capture(doc,evidence_id='template'),doc)
        self.assertEqual(validate_action_inventory(inventory,doc)['status'],'BLOCKED_ACTION_RESTORATION')
        with self.assertRaises(ValueError):validate_action_inventory(inventory,doc,require_recovery=True)
        ready=self.reviewed(inventory);self.assertEqual(validate_action_inventory(ready,doc,require_recovery=True)['status'],'READY')
        bad=copy.deepcopy(ready);bad['elements']['button']['sequences'][0]['blocks'][0]['fields'].pop('values')
        with self.assertRaises(ValueError):validate_action_inventory(bad,doc)
    def test_post_rest_sdk_capture_detects_lost_actions_and_changed_conditions_despite_same_rest(self):
        doc,inventory=self.fixture();before=self.reviewed(inventory)
        after=copy.deepcopy(before);after['version']=8;after['evidenceId']='post';fresh={**doc,'documentVersion':8}
        after['elements']['button']['sequences'][0]['id']='new-assigned-sequence'
        self.assertEqual(verify_action_preservation(before,after,fresh)['status'],'VERIFIED_SDK_ACTION_SEMANTICS')
        lost=copy.deepcopy(after);lost['elements']['button']['sequences']=[]
        with self.assertRaises(ValueError):verify_action_preservation(before,lost,fresh)
        changed=copy.deepcopy(after);changed['elements']['button']['sequences'][0]['disabled']=True
        with self.assertRaises(ValueError):verify_action_preservation(before,changed,fresh)
if __name__=='__main__':unittest.main()
