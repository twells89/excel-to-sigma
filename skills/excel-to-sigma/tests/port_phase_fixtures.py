"""Explicitly synthetic phase/controller observations for offline unit tests only."""
import copy,hashlib
import custom_ports as p
from action_inventory import empty_action_capture,digest,configuration
from input_sync import build_artifact_records

def reviewed(record):
    record['review']={'status':'APPROVED_STAGING','owner':'unit-reviewer','captureHash':p._digest({k:v for k,v in record.items() if k!='review'})};return record

def context(m,r,c):
    codes={e['id']:e for e in c['contents']['elements'] if e['kind']=='code'}
    supported=copy.deepcopy(c);supported['contents']['elements']=[e for e in supported['contents']['elements'] if e['id'] not in codes]
    capture=empty_action_capture(supported,evidence_id='SYNTHETIC_UNIT_ONLY');capture.update(origin='sigma-sdk-readonly',complete=True)
    for element in supported['contents']['elements']:
        record=capture['elements'][element['id']];record.update(complete=True,inspection='sdk-type-no-action-events')
        for action in element.get('actions',[]):
            record['sequences'].append({'id':action['id'],'name':'unit-run','trigger':'on-click','condition':None,'disabled':False,'successToast':None,
                'blocks':[{'kind':'RunPythonElement','fields':{'codeElementId':effect['codeElementId'],'name':effect.get('name')}} for effect in action['effects']]})
        if record['sequences']:
            record['semanticHash']=digest(configuration(record));program='# SYNTHETIC missing-only test artifact\n';ph=hashlib.sha256(program.encode()).hexdigest()
            record['restorationArtifact']={'elementId':element['id'],'documentId':'doc','reviewer':'unit-reviewer','sourceObservationHash':digest(record),'semanticHash':record['semanticHash'],
                'program':program,'programHash':ph,'constructionArtifactHash':'unit-only','bindingHash':'unit-only','guardPolicy':'missing-sequences-only; exact-current-target-and-bindings; no-constructor-replay',
                'restoreProof':{'status':'PASSED_ISOLATED_RESTORATION','evidenceId':'SYNTHETIC_ONLY','documentId':'unit-isolated','semanticHash':record['semanticHash'],
                    'sdkReadbackEvidenceId':'unit-only','programHash':ph,'bindingHash':'unit-only','beforeMissingSequences':True}}
    return reviewed({'origin':'sigma-controller-actual','evidenceId':'SYNTHETIC_UNIT_ONLY','documentId':c['workbookId'],'version':c['documentVersion'],
        'contentsHash':p._digest(c['contents']),'receiptHash':p._digest(r),'inputValuesHash':'a'*64,'schemaDisplayHash':'b'*64,'rawCaptureHash':'c'*64,'sdkActions':capture,
        'codeObservations':{eid:{'origin':'sigma-code-rest-ui-readonly','evidenceId':'SYNTHETIC_UNIT_ONLY','documentId':c['workbookId'],'version':c['documentVersion'],
            'elementHash':p._digest(e),'codeHash':hashlib.sha256(e['code'].encode()).hexdigest(),'inspection':'public-code-schema-no-action-events'} for eid,e in codes.items()}})

def restage(m,r,c):
    m=copy.deepcopy(m);r=copy.deepcopy(r);m['target']={'documentId':c['workbookId'],'version':c['documentVersion']}
    for target in r['targetBindings'].values():target['version']=c['documentVersion']
    r['artifacts']=build_artifact_records(m,m['planProjection']['sharedInputArchitecture'],r['targetBindings']);return m,r

def installed(current,proposal):return {**current,'documentVersion':current['documentVersion']+1,'contents':copy.deepcopy(proposal['dryRunRequest']['contents'])}

def outputs(m,r,c,previous,keys):
    elements={e['id']:e for e in c['contents']['elements']};records={}
    for key in keys:
        eid=p.PORT_PREFIX[key]+'Python';ch=hashlib.sha256(elements[eid]['code'].encode()).hexdigest();execution='SYNTHETIC-'+key
        namespace='Code: '+p.OUTPUT_NAMES[key]
        observed=[{'outputColumn':name,'identifier':'"'+name+'"','rowKey':name,'formula':'['+namespace+'/'+name+']',
            'type':'text' if typ in {'STRING','VARCHAR'} else 'boolean' if typ=='BOOLEAN' else 'integer' if typ=='INTEGER' else 'number'} for name,typ in p.OUTPUT_SCHEMAS[key]]
        metadata={'origin':'sigma-python-output-metadata-actual','evidenceId':'SYNTHETIC_UNIT_ONLY','executionId':execution,'documentId':c['workbookId'],
            'version':c['documentVersion'],'outputName':p.OUTPUT_NAMES[key],'sourceNamespace':namespace,'rawCaptureHash':'d'*64,'observedFields':observed}
        records[key]={'codeElementId':eid,'codeHash':ch,'outputName':p.OUTPUT_NAMES[key],'connectionId':'conn','metadata':metadata,'metadataHash':p._digest(metadata),
            'columnMap':{f['outputColumn']:{k:v for k,v in f.items() if k!='type'} for f in observed},
            'execution':{'origin':'sigma-run-python-actual','status':'COMPLETED','businessStatus':'FAILURE','executionId':execution,'evidenceId':'SYNTHETIC_UNIT_ONLY',
                'documentId':c['workbookId'],'version':c['documentVersion'],'codeElementId':eid,'codeHash':ch}}
    return reviewed({'origin':'sigma-python-output-actual','evidenceId':'SYNTHETIC_UNIT_ONLY','documentId':c['workbookId'],'version':c['documentVersion'],
        'sourceHash':m['sourceHash'],'planHash':m['planHash'],'modelBinding':previous['modelBinding'],'contentsHash':p._digest(c['contents']),'receiptHash':p._digest(r),
        'installedProposalHash':previous['proposalHash'],'installation':{'status':'VERIFIED_COMPONENT_CONSTRUCTION','proposalHash':previous['proposalHash'],'documentId':c['workbookId'],
            'version':c['documentVersion'],'evidenceId':'SYNTHETIC_UNIT_ONLY','contentsHash':p._digest(previous['dryRunRequest']['contents'])},'outputs':records})

def complete(m,r,c,sources):
    previous=p.prepare_port_update(m,r,c,sources,phase_context=context(m,r,c));c=installed(c,previous);m,r=restage(m,r,c)
    for phase in ('solver-outputs','guard-output'):
        previous=p.prepare_port_update(m,r,c,sources,phase=phase,previous_phase=previous,phase_context=context(m,r,c),output_evidence=outputs(m,r,c,previous,p.PHASE_KEYS[phase]))
        c=installed(c,previous);m,r=restage(m,r,c)
    return m,r,c,previous

def native_children(m,r,c,evidence):
    c=copy.deepcopy(c);evidence=copy.deepcopy(evidence);c['documentVersion']+=1
    for key,record in evidence['outputs'].items():
        eid='actual-native-'+key;columns=[];entries=[];ids={}
        for i,(name,pair) in enumerate(record['columnMap'].items()):
            cid=eid+'-column-'+str(i);ids[name]=cid;columns.append({'id':cid,'formula':pair['formula']})
            typ=next(f['type'] for f in record['metadata']['observedFields'] if f['outputColumn']==pair['outputColumn'])
            entries.append({'columnId':cid,'label':pair['outputColumn'],'formula':pair['formula'],'type':{'type':typ}})
        child={'id':eid,'kind':'table','source':{'kind':'code-output','elementId':record['codeElementId'],'output':record['outputName']},'columns':columns,'order':[col['id'] for col in columns]}
        if key!='allocation':child['name']={'goalSeek':'Goal Seek Result','applyGuard':'Goal Seek Apply Guard'}[key]
        c['contents']['elements'].append(child);c['contents']['layout']=c['contents']['layout'].replace('</Page>','<Element elementId="'+eid+'" gridColumn="1 / 25" gridRow="300 / 314"/></Page>',1)
        typed={'entries':entries,'total':len(entries),'nextPage':None}
        record['nativeChildObservation']={'origin':'sigma-native-code-output-child-actual','evidenceId':'SYNTHETIC_UNIT_ONLY','documentId':c['workbookId'],
            'version':c['documentVersion'],'elementId':eid,'elementHash':p._digest(child),'typedColumns':typed,'typedColumnsHash':p._digest(typed),'columnIds':ids,'rawCaptureHash':'f'*64}
        record['metadata']['version']=c['documentVersion'];record['metadataHash']=p._digest(record['metadata'])
    m,r=restage(m,r,c);evidence.update(version=c['documentVersion'],contentsHash=p._digest(c['contents']),receiptHash=p._digest(r));reviewed(evidence)
    return m,r,c,evidence
