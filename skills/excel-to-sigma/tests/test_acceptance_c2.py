"""C2 contract regressions. Synthetic receipts test admission, never runtime proof."""
import copy
import re
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts/lib'))
import test_acceptance as acceptance_fixture
import acceptance as a


def synthetic_qualified_blank_reviews(manifest, outcome='REJECTED_NO_WRITE'):
    """Unit-only decisions for acceptance mechanics; never source observations."""
    source=a.extract_workflow_repairs(manifest['source']['path'])
    reviews={}
    for rule in source['validations']:
        fields=rule['fields']
        if source['workflow']=='research':
            match=re.fullmatch(r'C(\d+):C(\d+)',rule['range'])
            fields=['Model Inputs!C'+str(row) for row in range(int(match[1]),int(match[2])+1)]
        for field in fields:
            branch_id=rule['id']+'-'+field+'-blank'
            review={'status':'REVIEWED_QUALIFIED','branchId':branch_id,
                'sourceHash':manifest['sourceHash'], 'ruleHash':rule['sourceXml']['ruleHash'],
                'policyHash':a.POLICY_HASH,'route':'edit-to-empty-commit',
                'expectedOutcome':outcome,
                'basis':'qualified-source-contract','reviewer':'unit-only-independent-reviewer',
                'reviewedAt':'2026-10-20','decisionId':'unit-'+branch_id,
                'sourceEvidenceId':'unit-only-source-decision',
                'rationale':'Synthetic unit test decision only; no actual Excel route claim.'}
            review['recordHash']=a.source_digest(review)
            reviews[branch_id]=review
    return reviews


def synthetic_validation_bundle(case,target,session):
    """Unit-only assigned query shapes; physical IDs here are never live evidence."""
    contract=case['validationContract'];table=contract['table'];feature=table['feature']
    spec=a._manifest_architecture(target['inputContract']['manifest'])['inputs'][feature]
    row_id='PHYSICAL_ROW_ID'
    rows=[]
    for index,source in enumerate(spec['rows']):
        row={row_id:'unit-physical-'+str(index)}
        for column in table['columns']:
            value=source[column['name']]
            if column['id'] not in table['keyColumns']:
                if column['type']=='number':value=0.37
                elif column['type']=='text':value='unit-baseline'
            row[column['id']]=value
        rows.append(row)
    binding={key:table[key] for key in ('documentId','version','pageId','elementId','connectionId',
        'tableIdentity','targetBindingHash','identityEvidenceId')}
    references=[]
    def readback(name,actual):
        execution='query-'+name
        references.append({'kind':'query','id':execution})
        return {'origin':'sigma-query','executionId':execution,'evidenceId':'unit-'+execution,
            'sourceHash':contract['sourceHash'],'planHash':contract['planHash'],
            'binding':copy.deepcopy(binding),'rowIdColumn':row_id,
            'columns':copy.deepcopy(table['columns']),'rowsHash':a._digest(actual),
            'receipt':{'status':'COMPLETED','executionId':execution,
                       'evidenceId':'unit-'+execution,
                       'documentId':target['documentId'],'version':target['version'],
                       'sourceHash':contract['sourceHash'],'planHash':contract['planHash'],
                       'binding':copy.deepcopy(binding),'rowIdColumn':row_id,
                       'columns':copy.deepcopy(table['columns']),'rowsHash':a._digest(actual)}}
    inventory={**readback('inventory',rows),'rows':copy.deepcopy(rows)}
    inventory['physicalIdMap']=[{'sourceKey':[row[key] for key in table['keyColumns']],
                                'physicalId':row[row_id]} for row in rows]
    evidence=[]
    for index,branch in enumerate(contract['branches']):
        column=branch['target']['columnId'];attempt=branch['attempt']['value']
        before=copy.deepcopy(rows);after=copy.deepcopy(before)
        saved=branch['expectedOutcome']=='SAVED'
        if saved:after[0][column]=attempt
        execution_id='native-'+str(index)
        references.append({'kind':'browser','id':execution_id})
        proof={**branch,'observedOutcome':branch['expectedOutcome'],
            'origin':'sigma-browser','executionId':execution_id,'publishedVersion':target['version'],
            'sourceHash':target['sourceHash'],'planHash':target['planHash'],
            'inputMode':table['editMode'],'session':copy.deepcopy(session),
            'steps':['Edit data',branch['id'],'Save or Cancel'],
            'screenshots':[{'sha256':str(index).zfill(64)}],
            'references':[{'kind':'browser','id':execution_id}],
            'receipt':{'status':'COMPLETED','executionId':execution_id,
                       'documentId':target['documentId'],'version':target['version']},
            'writeCount':1 if saved else 0,
            'rows':{'rowIdColumn':row_id,'keyColumns':copy.deepcopy(table['keyColumns']),
                    'before':before,'after':after,'reopened':copy.deepcopy(after),
                    'restored':copy.deepcopy(before)}}
        proof['rows']['readbacks']={phase:readback(str(index)+'-'+phase,proof['rows'][phase])
            for phase in ('before','after','reopened','restored')}
        if saved:
            proof['rows']['allowedChanges']=[{'key':{key:before[0][key] for key in table['keyColumns']},
                                               'column':column,'value':attempt}]
        evidence.append(proof)
    return inventory,evidence,references


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
        blanks=[b for c in matrix['cases'] if c['id'].endswith('-zero-and-invalid-entry')
                for b in c['validationContract']['branches'] if b['id'].endswith('-blank')]
        self.assertEqual(len(blanks),25)
        self.assertEqual({b['expectedOutcome'] for b in blanks},{'UNRESOLVED_SOURCE_ROUTE'})
        research=next(c for c in matrix['cases'] if c['id']=='research-zero-and-invalid-entry')
        self.assertEqual(research['validationContract']['rules'][0]['blankSemantics']['classification'],
                         'SATISFIES_NUMERIC_RULE_AS_ZERO')
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
        for mutation in ('drop-rule','drop-branch','wrong-target','changed-outcome','forged-classification'):
            bad=copy.deepcopy(matrix);bad.pop('matrixHash')
            contract=next(c for c in bad['cases'] if c['id']=='inventory-zero-and-invalid-entry')['validationContract']
            if mutation=='drop-rule':contract['rules'].pop()
            elif mutation=='drop-branch':contract['branches'].pop()
            elif mutation=='wrong-target':contract['rules'][0]['target']['elementId']='foreign'
            elif mutation=='forged-classification':contract['rules'][0]['blankSemantics']['classification']='ALLOW_BLANK'
            else:contract['branches'][0]['expectedOutcome']='REJECTED_NO_WRITE'
            with self.subTest(mutation=mutation),self.assertRaises(ValueError):
                a.validate_production_matrix(a.seal_matrix(bad))

    def validation_fixture(self, reviewed=True, blank_outcome='REJECTED_NO_WRITE'):
        manifests,targets,receipts=acceptance_fixture.AcceptanceTests().bound_matrix_inputs()
        if reviewed:
            targets['sales']['sourceBlankReviews']=synthetic_qualified_blank_reviews(
                manifests['sales'],blank_outcome)
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

    def test_unreviewed_blank_cannot_pass_with_self_selected_runtime_outcome(self):
        matrix,case,capture=self.validation_fixture(reviewed=False)
        capture['validationInventory'],capture['validationBranches'],refs=synthetic_validation_bundle(
            case,matrix['workflows']['sales'],capture['session'])
        capture['references']+=refs
        for observed in capture['validationBranches']:
            if observed['id'].endswith('-blank'):
                observed['observedOutcome']='REJECTED_NO_WRITE'
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)
        report=a.build_workflow_reports(matrix,[])['sales']
        self.assertEqual(report['verdict'],'BLOCKED')
        self.assertIn('source blank edit/clear route remains unresolved',report['completionBlockers'])

    def test_qualified_review_must_bind_source_rule_policy_and_branch(self):
        manifests,targets,receipts=acceptance_fixture.AcceptanceTests().bound_matrix_inputs()
        matrix,case,_=self.validation_fixture()
        review=next(iter(matrix['workflows']['sales']['sourceBlankReviews'].values()))
        for mutation in ('ruleHash','policyHash','sourceHash','branchId','route','recordHash','sourceEvidenceId'):
            bad=copy.deepcopy(review)
            bad[mutation]='foreign'
            targets['sales']['sourceBlankReviews']={review['branchId']:bad}
            with self.subTest(mutation=mutation),self.assertRaises(ValueError):
                a.build_acceptance_matrix(manifests,targets,receipts=receipts)

    def test_unresolved_custom_rule_keeps_blocking_blank_branch(self):
        manifests,targets,receipts=acceptance_fixture.AcceptanceTests().bound_matrix_inputs()
        source=a.extract_workflow_repairs(manifests['sales']['source']['path'])
        rule=source['validations'][0]
        rule.update(sourceType='custom',status='unresolved',emission=None)
        rule['blankSemantics']['classification']='UNRESOLVED'
        contract=a._validation_contract('sales',source,manifests['sales'],receipts['sales'],targets['sales'])
        branches=[b for b in contract['branches'] if b['ruleId']==rule['id'] and b['id'].endswith('-blank')]
        self.assertEqual(len(branches),1)
        self.assertEqual(branches[0]['expectedOutcome'],'UNRESOLVED_SOURCE_ROUTE')
        self.assertEqual(branches[0]['declaredBlankClass'],'UNRESOLVED')

    def test_qualified_saved_blank_requires_persisted_null_not_numeric_zero(self):
        matrix,case,capture=self.validation_fixture(blank_outcome='SAVED')
        capture['validationInventory'],capture['validationBranches'],refs=synthetic_validation_bundle(
            case,matrix['workflows']['sales'],capture['session'])
        capture['references']+=refs
        self.assertEqual(a.append_evidence([],matrix,capture)[0]['status'],'PASSED')
        changed=copy.deepcopy(capture)
        branch=next(item for item in changed['validationBranches'] if item['id'].endswith('-blank'))
        branch['rows']['allowedChanges'][0]['value']=0
        with self.assertRaises(ValueError):a.append_evidence([],matrix,changed)

    def test_validation_capture_needs_actual_linked_branch_proof(self):
        matrix,case,capture=self.validation_fixture()
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)
        capture['validationBranches']=[]
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)
        capture['origin']='sigma-query'
        with self.assertRaises(ValueError):a.append_evidence([],matrix,capture)

    def test_full_source_validation_vector_and_cancel_rejects_false_runtime_or_target(self):
        matrix,case,capture=self.validation_fixture()
        capture['validationInventory'],capture['validationBranches'],refs=synthetic_validation_bundle(case,matrix['workflows']['sales'],capture['session'])
        capture['references'] += refs
        self.assertEqual(a.append_evidence([],matrix,capture)[0]['status'],'PASSED')
        reordered=copy.deepcopy(capture)
        for item in reordered['validationBranches']:
            for phase in ('after','reopened'):
                item['rows'][phase].reverse()
                row_hash=a._digest(item['rows'][phase])
                item['rows']['readbacks'][phase]['rowsHash']=row_hash
                item['rows']['readbacks'][phase]['receipt']['rowsHash']=row_hash
        self.assertEqual(a.append_evidence([],matrix,reordered)[0]['status'],'PASSED')
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
            'wrong-blank-route':lambda bad:next(item for item in bad['validationBranches']
                if item['id'].endswith('-blank')).update(sourceRoute='delete-key'),
        }
        for label,mutate in scenarios.items():
            bad=copy.deepcopy(capture);mutate(bad)
            with self.subTest(label=label),self.assertRaises(ValueError):a.append_evidence([],matrix,bad)

    def test_c2_reviewer_reproductions_are_rejected(self):
        matrix,case,capture=self.validation_fixture()
        capture['validationInventory'],capture['validationBranches'],refs=synthetic_validation_bundle(case,matrix['workflows']['sales'],capture['session'])
        capture['references'] += refs
        self.assertEqual(a.append_evidence([],matrix,capture)[0]['status'],'PASSED')

        def foreign_rows(bad):
            for item in bad['validationBranches']:
                item['rows']['rowIdColumn']='ID'
                item['rows']['keyColumns']=['Key']
                for phase in ('before','after','reopened','restored'):
                    item['rows'][phase]=[{'ID':'not-an-assigned-source-row','Key':'foreign',
                                          item['target']['columnId']:0.37}]
                if item['rows'].get('allowedChanges'):
                    item['rows']['allowedChanges'][0]['key']={'Key':'foreign'}
        bad=copy.deepcopy(capture);foreign_rows(bad)
        with self.assertRaises(ValueError):a.append_evidence([],matrix,bad)

        numeric=copy.deepcopy(capture)
        branch=next(item for item in numeric['validationBranches'] if item['id']=='validation-4-Probability-upper')
        column=branch['target']['columnId']
        branch['attempt']['value']=True
        with self.assertRaises(ValueError):a.append_evidence([],matrix,numeric)
        numeric=copy.deepcopy(capture)
        branch=next(item for item in numeric['validationBranches'] if item['id']=='validation-4-Probability-upper')
        column=branch['target']['columnId']
        branch['rows']['allowedChanges'][0]['value']=True
        for phase in ('before','restored'):branch['rows'][phase][0][column]=False
        for phase in ('after','reopened'):branch['rows'][phase][0][column]=True
        with self.assertRaises(ValueError):a.append_evidence([],matrix,numeric)

        wrong_route=copy.deepcopy(capture)
        wrong_route['dataEntry']['inputMode']='published-all-users'
        wrong_route['session']['effectiveGrant']='view'
        for item in wrong_route['validationBranches']:
            item['session']=copy.deepcopy(wrong_route['session'])
        with self.assertRaises(ValueError):a.append_evidence([],matrix,wrong_route)

        for label,mutate in {
            'foreign-storage':lambda bad:bad['validationInventory']['binding'].update(tableIdentity='foreign'),
            'foreign-query':lambda bad:bad['validationInventory'].update(origin='foreign-service'),
            'missing-row':lambda bad:bad['validationBranches'][0]['rows']['before'].pop(),
            'missing-column':lambda bad:bad['validationBranches'][0]['rows']['before'][0].pop('c1'),
            'changed-physical-id':lambda bad:bad['validationBranches'][0]['rows']['before'][0].update(PHYSICAL_ROW_ID='foreign'),
            'changed-other-cell':lambda bad:bad['validationBranches'][0]['rows']['after'][1].update(c3='undisclosed'),
            'foreign-source-key':lambda bad:bad['validationInventory']['rows'][0].update(c0='foreign'),
            'foreign-row-id-column':lambda bad:bad['validationInventory'].update(rowIdColumn='c0'),
            'foreign-snapshot-receipt':lambda bad:bad['validationBranches'][0]['rows']['readbacks']['before']['binding'].update(elementId='foreign'),
        }.items():
            bad=copy.deepcopy(capture);mutate(bad)
            # Keep claimed hashes consistent so coverage checks, rather than only
            # receipt-hash checks, reject the altered row populations and cells.
            if label=='foreign-storage':
                bad['validationInventory']['receipt']['binding']['tableIdentity']='foreign'
            for item in bad['validationBranches']:
                for phase,readback in item['rows']['readbacks'].items():
                    row_hash=a._digest(item['rows'][phase])
                    readback['rowsHash']=row_hash;readback['receipt']['rowsHash']=row_hash
            inventory_hash=a._digest(bad['validationInventory']['rows'])
            bad['validationInventory']['rowsHash']=inventory_hash
            bad['validationInventory']['receipt']['rowsHash']=inventory_hash
            with self.subTest(label=label),self.assertRaises(ValueError):a.append_evidence([],matrix,bad)


if __name__=='__main__':unittest.main()
