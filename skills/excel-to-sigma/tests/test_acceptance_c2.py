"""C2 contract regressions. Synthetic receipts test admission, never runtime proof."""
import copy
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts/lib'))
import test_acceptance as acceptance_fixture
import acceptance as a


def synthetic_validation_branches(case,target,session):
    """Unit-only complete shape for validator tests; never publish as evidence."""
    evidence=[]
    for index,branch in enumerate(case['validationContract']['branches']):
        column=branch['target']['columnId'];attempt=branch['attempt']['value']
        baseline='baseline' if branch['attempt']['type']=='text' else 0.5
        if baseline==attempt:baseline=0.25
        before=[{'ID':'physical-id','Key':'source-row',column:baseline}]
        after=copy.deepcopy(before)
        saved=branch['expectedOutcome']=='SAVED'
        if saved:after[0][column]=attempt
        execution_id='native-'+str(index)
        proof={**branch,'observedOutcome':branch['expectedOutcome'],
            'origin':'sigma-browser','executionId':execution_id,'publishedVersion':target['version'],
            'sourceHash':target['sourceHash'],'planHash':target['planHash'],
            'session':copy.deepcopy(session),'steps':['Edit data',branch['id'],'Save or Cancel'],
            'screenshots':[{'sha256':str(index).zfill(64)}],
            'references':[{'kind':'browser','id':execution_id}],
            'receipt':{'status':'COMPLETED','executionId':execution_id,
                       'documentId':target['documentId'],'version':target['version']},
            'writeCount':1 if saved else 0,
            'rows':{'rowIdColumn':'ID','keyColumns':['Key'],'before':before,'after':after,
                    'reopened':copy.deepcopy(after),'restored':copy.deepcopy(before)}}
        if saved:
            proof['rows']['allowedChanges']=[{'key':{'Key':'source-row'},'column':column,'value':attempt}]
        evidence.append(proof)
    return evidence


class SourceValidationRoleTests(unittest.TestCase):
    def role_fixture(self, role='Editor'):
        helper = acceptance_fixture.AcceptanceTests()
        matrix = helper.matrix()
        matrix.pop('matrixHash')
        case = matrix['cases'][0]
        case.update(id='sales-published-'+role.lower(), mode='browser', count=2,
                    role=copy.deepcopy(a.ROLE_MEMBERS[role]), inputMode='published-explore',
                    editEntitlement=role!='Viewer', persistence=role!='Viewer')
        matrix = a.seal_matrix(matrix)
        capture = helper.capture(matrix)
        capture.update(caseId=case['id'], origin='sigma-browser', published=True,
                       publishedVersion=6, steps=['published native Edit data', 'Save', 'reopen'],
                       screenshots=[{'sha256':'c'*64}],
                       session={'memberId':case['role']['memberId'], 'accountType':case['role']['accountType'],
                                'effectiveGrant':case['role']['grant'], 'identityEvidenceId':'role-banner',
                                'adminBypass':False, 'editInputTables':role!='Viewer'},
                       dataEntry={'inputMode':'published-explore','editDataAvailable':role!='Viewer'})
        capture['actual']=[{'key':'A','value':100.01},{'key':'B','value':'Open'}]
        if role!='Viewer':
            before=[{'ID':'physical-row','Key':'one','Value':1}]
            changed=[{'ID':'physical-row','Key':'one','Value':2}]
            capture['dataEntry']['validSaved']=True
            capture['dataEntry']['writeCount']=1
            capture['persistence']={'before':before,'changed':changed,'reopened':copy.deepcopy(changed),
                'restored':copy.deepcopy(before),'rowIdColumn':'ID','keyColumns':['Key'],
                'allowedChanges':[{'key':{'Key':'one'},'column':'Value','value':2}],
                'reopenEvidenceId':'reopen-query','restoreEvidenceId':'restore-query'}
        else:
            capture['dataEntry']['writeCount']=0
        return matrix,capture

    def test_writable_roles_save_without_fabricated_invalid_or_cancel_and_viewer_no_write(self):
        for role in ('Editor','Explorer','Viewer'):
            with self.subTest(role=role):
                matrix,capture=self.role_fixture(role)
                self.assertEqual(a.append_evidence([],matrix,capture)[0]['status'],'PASSED')
        for role in ('Editor','Explorer'):
            matrix,capture=self.role_fixture(role)
            for mutation in ('missing-save','zero-write','no-persistence','wrong-role','service-query','wrong-row'):
                bad=copy.deepcopy(capture)
                if mutation=='missing-save':bad['dataEntry'].pop('validSaved')
                elif mutation=='zero-write':bad['dataEntry']['writeCount']=0
                elif mutation=='no-persistence':bad.pop('persistence')
                elif mutation=='wrong-role':bad['session']['memberId']='owner'
                elif mutation=='service-query':bad['origin']='sigma-query'
                else:bad['persistence']['changed'][0]['ID']='other-row'
                with self.subTest(role=role,mutation=mutation),self.assertRaises(ValueError):
                    a.append_evidence([],matrix,bad)
        matrix,capture=self.role_fixture('Viewer')
        capture['dataEntry']['writeCount']=1
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)
        matrix,capture=self.role_fixture('Viewer')
        for changed in ({'effectiveGrant':'explore'},{'accountType':'build'},{'editInputTables':True}):
            bad=copy.deepcopy(capture);bad['session'].update(changed)
            with self.subTest(viewer=changed),self.assertRaises(ValueError):a.append_evidence([],matrix,bad)

    def test_production_recomputes_all_source_rules_targets_and_classes(self):
        manifests,targets,receipts=acceptance_fixture.AcceptanceTests().bound_matrix_inputs()
        matrix=a.build_acceptance_matrix(manifests,targets,receipts=receipts)
        self.assertTrue(a.validate_production_matrix(matrix))
        self.assertEqual(len(matrix['cases']),102)
        for workflow in manifests:
            case=next(c for c in matrix['cases'] if c['id']==workflow+'-zero-and-invalid-entry')
            contract=case['validationContract']
            self.assertEqual(contract['sourceHash'],manifests[workflow]['sourceHash'])
            self.assertEqual(len(contract['rules']),matrix['workflows'][workflow]['sourceFeatures']['validationRules'])
            self.assertEqual(sum(b['id']=='cancel-no-effect' for b in contract['branches']),1)
            for rule in contract['rules']:
                self.assertTrue(rule['target']['elementId'])
                self.assertTrue(rule['target']['columnIds'])
        sales=next(c for c in matrix['cases'] if c['id']=='sales-zero-and-invalid-entry')
        self.assertNotIn('AsOfDate',str(sales['validationContract']))
        self.assertIn('Probability',str(sales['validationContract']))
        inventory=next(c for c in matrix['cases'] if c['id']=='inventory-zero-and-invalid-entry')
        lead=[b for b in inventory['validationContract']['branches'] if b['ruleId']=='validation-3']
        self.assertEqual({b['attempt']['value'] for b in lead if b['attempt']['type']=='number'},
                         {1,90,0,91,1.5})
        for mutation in ('drop-rule','drop-branch','wrong-target','changed-outcome'):
            bad=copy.deepcopy(matrix);bad.pop('matrixHash')
            contract=next(c for c in bad['cases'] if c['id']=='inventory-zero-and-invalid-entry')['validationContract']
            if mutation=='drop-rule':contract['rules'].pop()
            elif mutation=='drop-branch':contract['branches'].pop()
            elif mutation=='wrong-target':contract['rules'][0]['target']['elementId']='foreign'
            else:contract['branches'][0]['expectedOutcome']='REJECTED_NO_WRITE'
            with self.subTest(mutation=mutation),self.assertRaises(ValueError):
                a.validate_production_matrix(a.seal_matrix(bad))

    def validation_fixture(self):
        manifests,targets,receipts=acceptance_fixture.AcceptanceTests().bound_matrix_inputs()
        matrix=a.build_acceptance_matrix(manifests,targets,receipts=receipts)
        case=next(c for c in matrix['cases'] if c['id']=='sales-zero-and-invalid-entry')
        target=matrix['workflows']['sales']
        case['oracle']={'origin':'independent-oracle','reviewer':'source-reviewer',
            'sourceHash':target['sourceHash'],'checks':[{'key':'source-validation-and-cancel',
            'expected':case['validationContract']['expectedOutcomes'],'comparison':'exact'}]}
        matrix=a.seal_matrix({k:v for k,v in matrix.items() if k!='matrixHash'})
        capture={'caseId':case['id'],'evidenceId':'validation-root','executionId':'browser-root',
            'origin':'sigma-browser','sourceHash':target['sourceHash'],'planHash':target['planHash'],
            'documentId':target['documentId'],'version':target['version'],'published':True,'publishedVersion':target['version'],
            'steps':['native validation route'],'screenshots':['browser-root'],
            'session':{'memberId':'owner','accountType':'build','effectiveGrant':'edit',
                       'editInputTables':True,'identityEvidenceId':'owner-banner','adminBypass':False},
            'dataEntry':{'inputMode':'published-explore','editDataAvailable':True},
            'receipt':{'status':'COMPLETED','executionId':'browser-root','documentId':target['documentId'],'version':target['version']},
            'references':['browser-root'],
            'actual':[{'key':'source-validation-and-cancel','value':case['validationContract']['expectedOutcomes']}]}
        return matrix,case,capture

    def test_validation_capture_needs_actual_linked_branch_proof(self):
        matrix,case,capture=self.validation_fixture()
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)
        capture['validationBranches']=[]
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)
        capture['origin']='sigma-query'
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)

    def test_full_source_validation_vector_and_cancel_rejects_false_runtime_or_target(self):
        matrix,case,capture=self.validation_fixture()
        capture['validationBranches']=synthetic_validation_branches(case,matrix['workflows']['sales'],capture['session'])
        capture['references'] += [item['references'][0] for item in capture['validationBranches']]
        self.assertEqual(a.append_evidence([],matrix,capture)[0]['status'],'PASSED')
        scenarios={
            'foreign-target':lambda bad:bad['validationBranches'][0]['target'].update(columnId='foreign'),
            'different-source-rule':lambda bad:bad['validationBranches'][0].update(ruleId='foreign'),
            'foreign-source-hash':lambda bad:bad['validationBranches'][0].update(sourceHash='f'*64),
            'substituted-api':lambda bad:bad['validationBranches'][0].update(origin='sigma-query'),
            'unlinked-reference':lambda bad:bad['references'].pop(),
            'owner-for-editor':lambda bad:bad['validationBranches'][0]['session'].update(memberId='foreign'),
            'no-browser-proof':lambda bad:bad['validationBranches'][0].update(screenshots=[]),
            'invalid-persisted':lambda bad:bad['validationBranches'][1]['rows']['after'][0].update(
                {bad['validationBranches'][1]['target']['columnId']:'foreign'}),
            'cancel-wrote':lambda bad:bad['validationBranches'][-1].update(writeCount=1),
            'same-execution-reused':lambda bad:bad['validationBranches'][1].update(executionId='native-0'),
            'invented-oracle':lambda bad:bad['actual'][0].update(value=[{'id':'generic','outcome':'SAVED'}]),
        }
        for label,mutate in scenarios.items():
            bad=copy.deepcopy(capture);mutate(bad)
            with self.subTest(label=label),self.assertRaises(ValueError):a.append_evidence([],matrix,bad)


if __name__=='__main__':unittest.main()
