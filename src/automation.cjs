'use strict';
const fs=require('node:fs/promises');
const path=require('node:path');
const {URL}=require('node:url');
const {sha256}=require('./docx.cjs');
const {prepare,evaluate}=require('./increment3.cjs');
const {accountResponse}=require('./openai.cjs');
const {createLedger}=require('./budget.cjs');
const {jobStore,atomicJson,uuid,cleanArtifacts}=require('./job-store.cjs');
const {LinearClient}=require('./linear.cjs');
const {request:network}=require('./http.cjs');
const {readCredential}=require('./credentials.cjs');
const {requireCondition:check}=require('./errors.cjs');
const {selectModel,downloadModel,requireOperationalModel}=require('./template-catalog.cjs');
const {prepareDynamicMock,evaluateLocal}=require('./dynamic-template.cjs');
const {prepareDynamicApi,evaluateDynamicApi,authorizeDynamicApi}=require('./dynamic-api.cjs');
const {prepareProduction,authorizeProduction}=require('./production.cjs');
const {prepareTemplateFill,evaluateTemplateFill,workflowVersion:fieldWorkflowVersion}=require('./template-fill.cjs');
const {workflowVersion,prepareNativeDocument,validateReading,prepareNativeEditing,downloadNativeDocument,auditNativeDocument}=require('./ai-document.cjs');
const hash=value=>sha256(Buffer.from(JSON.stringify(value)));
const canonicalUrl=url=>{const u=new URL(url);u.search='';u.hash='';return u.href;};
function parseOrder(issue,catalog=null){
  const blocks=[...(issue.description||'').matchAll(/```cartorio-json\s*([\s\S]*?)```/g)];
  if(!blocks.length&&catalog){
    const text=(issue.title||'')+'\n'+(issue.description||'');
    let found=catalog.filter(m=>{const number=/^MODELO-(\d+)$/.exec(m.id);return new RegExp('(?:^|[^A-Za-z0-9_-])'+m.id.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'(?![A-Za-z0-9_-])','i').test(text)||number&&new RegExp('\\bMODELO\\s+0*'+Number(number[1])+'\\b','i').test(text);});
    if(!found.length){
      const normalize=s=>s.normalize('NFD').replace(/\p{M}/gu,'').toLowerCase().replace(/\b(?:de|do|da|dos|das|e)\b/g,' ').replace(/[^a-z0-9]+/g,' ').trim().replace(/\s+/g,' '),description=' '+normalize(text)+' ';
      found=catalog.filter(m=>m.origin!=='approved-fictional-derivative'&&[...(m.aliases||[]),m.name,m.source?.filename?.replace(/\.docx?$/i,'').replace(/^MODELO[\s_-]*\d+[\s_-]*/i,'')].filter(Boolean).some(name=>{const alias=normalize(name);return !!alias&&description.includes(' '+alias+' ');}));
    }
    if(found.length===1&&/\b(?:vers[aã]o|c[oó]pia)\s+anonimizada\b/iu.test(text)&&found[0].origin!=='approved-fictional-derivative'){
      const original=found[0];found=catalog.filter(m=>m.origin==='approved-fictional-derivative'&&m.id===original.id+'-API-FICTICIO'&&m.version===original.version);
    }
    check(found.length===1,found.length?'AMBIGUOUS_TEMPLATE_REQUEST':'EXACT_TEMPLATE_NOT_FOUND');
    const m=found[0],h=hash({issueId:issue.id,title:issue.title,description:issue.description,inputs:inputPdfs(issue).map(p=>canonicalUrl(p.url))});
    return {requestId:h.slice(0,8)+'-'+h.slice(8,12)+'-4'+h.slice(13,16)+'-a'+h.slice(17,20)+'-'+h.slice(20,32),actType:m.actType,templateId:m.id,templateVersion:m.version};
  }
  check(blocks.length===1,'REQUEST_DESCRIPTION_REQUIRED');
  let value;try{value=JSON.parse(blocks[0][1]);}catch{check(false,'REQUEST_JSON_INVALID');}
  check(value&&Object.keys(value).every(k=>['requestId','actType','templateId','templateVersion'].includes(k))&&uuid(value.requestId)&&['actType','templateId','templateVersion'].every(k=>typeof value[k]==='string'&&/^[A-Za-z0-9_.-]{1,80}$/.test(value[k])),'REQUEST_FIELDS_INVALID');return value;
}
function inputPdf(issue){
  const resources=issue.attachments.nodes.filter(a=>!a.metadata?.cartorioSuperseded&&/\.pdf$/i.test(a.title||''));
  const links=[...(issue.description||'').matchAll(/\[([^\]\n]+\.pdf)\]\((https:\/\/[^\s)]+)\)/gi)].map(m=>({id:'description-'+hash(canonicalUrl(m[2])),title:m[1],url:m[2],issue:{id:issue.id}}));
  const unique=new Map([...resources,...links].map(a=>[canonicalUrl(a.url),a]));
  check(unique.size===1,unique.size?'MULTIPLE_INPUT_PDFS':'INPUT_PDF_MISSING');return [...unique.values()][0];
}
function inputPdfs(issue){
 const resources=issue.attachments.nodes.filter(a=>!a.metadata?.cartorioSuperseded&&(/\.pdf$/i.test(a.title||'')||a.metadata?.contentType==='application/pdf'||/\.pdf$/i.test(new URL(a.url).pathname)));
 const links=[...(issue.description||'').matchAll(/\[([^\]\n]+\.pdf)\]\((https:\/\/[^\s)]+)\)/gi)].map(m=>({id:'description-'+hash(canonicalUrl(m[2])),title:m[1],url:m[2],issue:{id:issue.id}}));
 const unique=new Map([...resources,...links].map(a=>[canonicalUrl(a.url),a]));check(unique.size>0,'INPUT_PDF_MISSING');return [...unique.values()].sort((a,b)=>canonicalUrl(a.url).localeCompare(canonicalUrl(b.url)));
}
function requestSnapshot(issue,catalog=null){if(catalog)return hash({order:parseOrder(issue,catalog),inputs:inputPdfs(issue).map(p=>({id:p.id,url:canonicalUrl(p.url)})),title:issue.title,description:issue.description});const input=inputPdf(issue);return hash({order:parseOrder(issue),inputId:input.id,inputUrl:canonicalUrl(input.url),description:issue.description});}
async function readJson(file,fallback=null){try{return JSON.parse(await fs.readFile(file,'utf8'));}catch(e){check(e.code==='ENOENT','LOCAL_JSON_INVALID');return fallback;}}
async function runBatch({bundle,stateRoot,artifactRoot,trigger='manual',onlyIssueIds=null,linear=null,transport=network,secretReader=async p=>readCredential(p,/openai/i.test(p)?'openai':'linear'),approval=null,recoveryOnly=false,checkpoint=async()=>{}}){
  const config=bundle.config;
  const production=config.pipeline==='production',snapshotOf=i=>requestSnapshot(i,production?bundle.catalog:null);
  const nativeDocument=production&&config.documentEditingMode==='openai-native';
  const templateFields=production&&config.documentEditingMode==='template-fields';
  check(!config.linear.readyStateIds.includes(config.linear.reviewStateId),'REVIEW_STATE_CANNOT_BE_ELIGIBLE');
  check(config.concurrency===1&&(config.developmentOnly===true||production&&config.developmentOnly===false),'AUTOMATION_POLICY_REVIEW_REQUIRED');
  if(config.developmentOnly&&config.allowedDevelopmentWorkspaceId)check(config.linear.workspaceId===config.allowedDevelopmentWorkspaceId,'DEVELOPMENT_WORKSPACE_FORBIDDEN');
  check(config.openaiMode==='mock'||config.openaiMode==='live','INVALID_EXECUTION_MODE');
  if(trigger!=='manual')check(config.automaticEnabled===true,'AUTOMATIC_EXECUTION_DISABLED');
  if(config.openaiMode==='live')check(config.paidApiEnabled===true,'PAID_API_NOT_AUTHORIZED');
  if(recoveryOnly)check(production&&trigger==='manual'&&onlyIssueIds?.length===1,'RECOVERY_SINGLE_ISSUE_REQUIRED');
  if(production){check(config.openaiMode==='live','PRODUCTION_REQUIRES_LIVE_API');check(approval?.authorized===true&&approval.workspaceId===config.linear.workspaceId&&(recoveryOnly||Date.parse(approval.expiresAt)>Date.now()),'PAID_API_NOT_AUTHORIZED');if(approval.scope!=='production')check(onlyIssueIds?.length===1&&onlyIssueIds[0]===approval.issueId,'SINGLE_TEST_ISSUE_REQUIRED');}
  if(nativeDocument&&!recoveryOnly)check(approval.workflowVersion===workflowVersion&&approval.maxCalls>=2,'NATIVE_TWO_STAGE_APPROVAL_REQUIRED');
  if(!linear)linear=new LinearClient({...config.linear,apiKey:await secretReader(config.linearSecretFile),transport});
  const store=jobStore(stateRoot);
  return store.lock(async()=>{
    const workspace=await linear.assertScope();check(workspace.id===config.linear.workspaceId,'LINEAR_WORKSPACE_FORBIDDEN');
    const ready=onlyIssueIds||await linear.readyIssues({labelId:config.linear.requestLabelId,readyStateIds:config.linear.readyStateIds,maxIssues:config.maxIssuesPerBatch});
    check(ready.length<=config.maxIssuesPerBatch&&ready.every(uuid),'LINEAR_BATCH_LIMIT');
    const pending=onlyIssueIds?[]:await store.pendingIds();
    const ids=[...new Set([...ready,...pending])].slice(0,config.maxIssuesPerBatch);const results=[];
    for(const issueId of ids){
      let state=null,issue;
      try{
        state=await store.load(issueId);
        if(production){
          const status=await linear.issueStatus(issueId);
          if(!config.linear.readyStateIds.includes(status.state.id)||!status.labels.nodes.some(l=>l.id===config.linear.requestLabelId)){
            if(state?.deliveryConfirmed&&state.phase!=='complete'&&status.state.id===config.linear.reviewStateId){state.phase='complete';await store.save(issueId,state);results.push({issueId,status:'recovered-complete'});}
            else results.push({issueId,status:'skipped-not-eligible'});
            continue;
          }
        }
        issue=await linear.issue(issueId);
        const eligible=issue.labels.nodes.some(l=>l.id===config.linear.requestLabelId);
        if(!eligible||!config.linear.readyStateIds.includes(issue.state.id)){
          if(state?.deliveryConfirmed&&state.phase!=='complete'&&[config.linear.reviewStateId].includes(issue.state.id)){
            const attachment=issue.attachments.nodes.find(a=>a.metadata?.cartorioJobId===state.jobId);check(attachment&&sha256(await linear.download(attachment.url))===state.outputSha256,'DELIVERY_NOT_CONFIRMED');
            state.phase='complete';await store.save(issueId,state);results.push({issueId,status:'recovered-complete'});
          }else results.push({issueId,status:'skipped-not-eligible'});
          continue;
        }
        const order=parseOrder(issue,production?bundle.catalog:null),pdfList=production?inputPdfs(issue):[inputPdf(issue)],pdf=pdfList[0],snapshot=snapshotOf(issue);
        const requireStillEligible=async code=>{if(production){const status=await linear.issueStatus(issueId);check(config.linear.readyStateIds.includes(status.state.id)&&status.labels.nodes.some(l=>l.id===config.linear.requestLabelId),code);}const current=await linear.issue(issueId);check(snapshotOf(current)===snapshot&&config.linear.readyStateIds.includes(current.state.id)&&current.labels.nodes.some(l=>l.id===config.linear.requestLabelId),code);return current;};
        const model=selectModel(bundle.catalog,order);
        const apiModel=model.origin==='approved-fictional-derivative'&&model.status==='ready-approved-fictional-api';
        const localModel=model.origin==='user-supplied-original'&&['ready-local-mock','ready-production'].includes(model.status)&&(!production||model.format!=='docx')||apiModel;
        if(model.origin==='user-supplied-original'&&!production)check(config.openaiMode==='mock'&&config.paidApiEnabled===false,'REGISTERED_TEMPLATE_LIVE_DISABLED');
        if(production)check(model.userApprovedUse===true&&['ready-production','ready-approved-fictional-api'].includes(model.status),'TEMPLATE_NOT_READY');
        const originalBytes=await downloadModel(linear,model);
        let templateBytes=originalBytes,fieldMetadata;
        if(templateFields){
          const assets=(bundle.fieldTemplates||[]).filter(a=>a.modelId===model.id&&a.version===model.version&&a.originalDocSha256===model.sha256);check(assets.length===1,'FIELD_TEMPLATE_ASSET_MISSING');
          const asset=assets[0];templateBytes=asset.path?await fs.readFile(asset.path):Buffer.from(asset.docxBase64,'base64');fieldMetadata=asset.metadata;
          check(fieldMetadata?.preparedSha256===sha256(templateBytes)&&fieldMetadata.originalDocSha256===sha256(originalBytes),'FIELD_TEMPLATE_SOURCE_MISMATCH');
        }else if(localModel){
          const assets=(bundle.localTemplates||[]).filter(a=>a.modelId===model.id&&a.version===model.version&&a.originalSha256===model.sha256);check(assets.length===1,'LOCAL_TEMPLATE_ASSET_MISSING');
          const asset=assets[0];templateBytes=production&&asset.path?await fs.readFile(asset.path):Buffer.from(asset.docxBase64,'base64');
          check(sha256(templateBytes)===asset.sha256&&asset.sha256===model.convertedCandidate?.sha256,'LOCAL_TEMPLATE_INTEGRITY_FAILED');
          check(asset.originalSha256===sha256(originalBytes)&&model.inventoryMode==='dynamic-per-request','LOCAL_TEMPLATE_SOURCE_MISMATCH');
        }else if(!production)requireOperationalModel(model,templateBytes);
        const pdfs=[];if(production)check(pdfList.length<=(nativeDocument?49:bundle.openaiConfig.limits.maxPdfs||4),'INVALID_INPUT_DOCUMENTS');
        for(let n=0;n<pdfList.length;n++)pdfs.push({id:'document-'+(n+1),bytes:await linear.download(pdfList[n].url,bundle.openaiConfig?.limits.maxPdfBytes)});
        const pdfBytes=pdfs[0].bytes;
        check(pdfBytes.subarray(0,5).toString()==='%PDF-','INVALID_PDF_INPUT');
        const fixture=production?null:bundle.fixtures.find(f=>f.pdfSha256===sha256(pdfBytes)&&f.templateSha256===sha256(templateBytes));if(!production)check(fixture,'FICTIONAL_INPUT_NOT_APPROVED');
        const apiTask=production?{workspaceId:workspace.id,issueId,title:issue.title||'',description:issue.description.replace(/```cartorio-json[\s\S]*?```/g,'').trim(),templateId:order.templateId,templateVersion:order.templateVersion}:apiModel?{description:issue.description.split('```cartorio-json')[0].trim(),templateId:order.templateId,templateVersion:order.templateVersion}:null;
        let preparedApi;
        const buildApi=()=>{if(!preparedApi)preparedApi=production?(templateFields?prepareTemplateFill:nativeDocument?prepareNativeDocument:prepareProduction)({templateBytes,fieldMetadata,pdfs,model,task:apiTask,config:bundle.openaiConfig,instructions:bundle.instructions,correction:bundle.analysisCorrection?.issueId===issueId?bundle.analysisCorrection:null}):prepareDynamicApi({templateBytes,pdfBytes,model,task:apiTask,fixture,config:bundle.openaiConfig});return preparedApi;};
        const payload=production?null:{...fixture.payload,mode:config.openaiMode,templateBase64:templateBytes.toString('base64'),pdfs:[{id:'pacote-ficticio',base64:pdfBytes.toString('base64')}]};
        const requestSha256=production?(await buildApi()).requestSha256:null;
        const jobId=hash({workspace:workspace.id,issueId,order,snapshot,model:model.sha256,...(localModel?{workingTemplate:sha256(templateBytes)}:{}),pdf:production?pdfs.map(p=>sha256(p.bytes)):sha256(pdfBytes),mode:config.openaiMode,...(production?{apiRequestSha256:requestSha256}:{})});
        if(recoveryOnly)check(state?.jobId===jobId&&state.apiAttempted===true&&state.usageAccounted===true&&state.reservation?.requestSha256===requestSha256,'RECOVERY_RECEIPT_REQUIRED');
        if(state&&state.jobId!==jobId){
          if(state.requestId===order.requestId){
            // Uma correcao tecnica nao concede nova chamada. Exige outra aprovacao
            // individual, valida e vinculada ao novo hash; tentativas sem recibo
            // contabilizado continuam bloqueadas, mesmo com outra aprovacao.
            if(production)check(state.usageAccounted===true&&(!state.editing?.apiAttempted||state.editing.usageAccounted===true)&&approval?.authorized===true&&approval.maxCalls===(nativeDocument?2:1)&&approval.maxBudgetUsd>0&&(nativeDocument||approval.maxBudgetUsd<=0.5||approval.scope==='production'&&approval.dataScope==='institutional'&&approval.budgetPolicy==='user-unlimited')&&approval.issueId===issueId&&approval.workspaceId===workspace.id&&approval.requestSha256===requestSha256&&requestSha256!==state.reservation?.requestSha256&&approval.approvalId!==state.reservation?.approvalId&&Date.parse(approval.expiresAt)>Date.now(),'NEW_REQUEST_AUTHORIZATION_REQUIRED');
            else check(false,'REQUEST_CHANGED_REQUIRES_NEW_ID');
          }
          if(state.jobId){check(/^[a-f0-9]{64}$/.test(state.jobId),'JOB_STATE_INVALID');await atomicJson(path.join(stateRoot,'history',state.jobId+'.json'),state);}else check(!state.apiAttempted,'JOB_STATE_INVALID');
        }
        if(state?.jobId===jobId&&state.phase==='complete'){results.push({issueId,status:'already-complete'});continue;}
        if(state?.jobId!==jobId)state={issueId,jobId,requestId:order.requestId,requestSnapshot:snapshot,inputSha256:sha256(pdfBytes),...(production?{inputDocuments:pdfs.map(p=>({id:p.id,sha256:sha256(p.bytes)})),modelId:model.id}:{}),templateSha256:sha256(templateBytes),...(localModel?{originalTemplateSha256:model.sha256,modelId:model.id}:{}),mode:config.openaiMode,phase:'downloaded',createdAt:new Date().toISOString()};
        await store.save(issueId,state);
        const dir=path.join(artifactRoot,jobId);await fs.mkdir(dir,{recursive:true});const outputFile=path.join(dir,'result.docx'),responseFile=path.join(dir,'response.json');
        if(bundle.analysisCorrection?.issueId===issueId){await atomicJson(path.join(dir,'correction.json'),bundle.analysisCorrection);state.analysisCorrectionSha256=hash(bundle.analysisCorrection);await store.save(issueId,state);}
        if(!state.outputSha256){
          let prepared;
          if(config.openaiMode==='live'){
            check((production&&approval?.scope==='production'||approval?.issueId===issueId)&&approval.workspaceId===workspace.id,'APPROVAL_ISSUE_MISMATCH');
            const saved=await readJson(responseFile);
            if(saved){
              // Reconstroi apenas contexto, sem reservar ou transmitir novamente.
              prepared=production||apiModel?await buildApi():await prepare({...payload,mode:'mock'},{manifest:bundle.manifest,truth:fixture.truth,config:bundle.openaiConfig});
              check(saved.requestSha256===prepared.requestSha256&&state.reservation,'RECOVERY_RESPONSE_MISMATCH');prepared={...prepared,mode:'live',config:state.liveConfig,reservation:state.reservation,ledgerPath:path.join(stateRoot,'openai-budget.json')};
            }else{
              check(!recoveryOnly,'RECOVERY_SAVED_RESPONSE_REQUIRED');
              check(!state.apiAttempted,'PAID_CALL_RESULT_UNKNOWN_REVIEW_REQUIRED');
              await requireStillEligible('REQUEST_CHANGED_BEFORE_OPENAI');
              const apiKey=await secretReader(config.openaiSecretFile);check(apiKey,'OPENAI_KEY_REQUIRED');
              prepared=production?await authorizeProduction(await buildApi(),(await buildApi()).config,approval,path.join(stateRoot,'openai-budget.json'),config.developmentOnly):apiModel?await authorizeDynamicApi(await buildApi(),bundle.openaiConfig,approval,path.join(stateRoot,'openai-budget.json')):await prepare(payload,{manifest:bundle.manifest,truth:fixture.truth,config:bundle.openaiConfig,approval,ledgerPath:path.join(stateRoot,'openai-budget.json')});
              state.apiAttempted=true;state.reservation=prepared.reservation;state.liveConfig=prepared.config;state.phase='api-reserved';await store.save(issueId,state);
              await checkpoint('before-openai',state);
              await requireStillEligible('REQUEST_CHANGED_BEFORE_OPENAI');
              require('./openai.cjs').assertMandatoryInstructions(prepared.request);
              const r=await transport('https://api.openai.com/v1/responses',{method:'POST',headers:{Authorization:'Bearer '+apiKey,'Content-Type':'application/json'},body:Buffer.from(JSON.stringify(prepared.request)),maxBytes:bundle.openaiConfig.limits.maxResponseBytes,timeoutMs:bundle.openaiConfig.timeoutMs});
              let raw;try{raw=JSON.parse(r.bytes.toString());}catch{check(false,'OPENAI_RESPONSE_INVALID');}
              await atomicJson(responseFile,{raw,httpStatus:r.status,requestSha256:prepared.requestSha256});
              await checkpoint('response-saved',state);
            }
          }else prepared=apiModel?await buildApi():localModel?await prepareDynamicMock({templateBytes,pdfBytes,model,task:{issueId:issue.id,description:issue.description,order},fixture,config}):await prepare(payload,{manifest:bundle.manifest,truth:fixture.truth,config:bundle.openaiConfig});
          const saved=config.openaiMode==='live'?await readJson(responseFile):null;
          // Persistir o consumo antes de validar a interpretacao: resposta
          // invalida nao significa que a chamada deixou de acontecer.
          if(config.openaiMode==='live'&&!state.usageAccounted&&saved?.raw?.usage){state.usage=await accountResponse(saved.raw,prepared.config,createLedger(path.join(stateRoot,'openai-budget.json')),prepared.reservation);state.usageAccounted=true;await store.save(issueId,state);}
          // Uma resposta recuperada nao gera outra chamada. O ledger pode registrar a mesma
          // resposta novamente apenas se a interrupcao precedeu a gravacao do recibo local.
          let result;
          if(nativeDocument){
            check(saved?.httpStatus===200,'OPENAI_HTTP_FAILED');
            const reading=validateReading(prepared,saved.raw);await atomicJson(path.join(dir,'reading.json'),reading);
            let editing=prepareNativeEditing(prepared,reading);const editingFile=path.join(dir,'editing-response.json');let edited=await readJson(editingFile);
            if(edited){
              check(state.editing?.reservation?.requestSha256===editing.requestSha256&&edited.requestSha256===editing.requestSha256,'RECOVERY_EDITING_RESPONSE_MISMATCH');
              editing={...editing,config:state.editing.liveConfig,reservation:state.editing.reservation};
            }else{
              check(!recoveryOnly,'RECOVERY_EDITING_RESPONSE_REQUIRED');
              check(!state.editing?.apiAttempted,'PAID_CALL_RESULT_UNKNOWN_REVIEW_REQUIRED');
              await requireStillEligible('REQUEST_CHANGED_BEFORE_OPENAI');
              editing=await authorizeProduction(editing,editing.config,approval,path.join(stateRoot,'openai-budget.json'),config.developmentOnly);
              state.editing={apiAttempted:true,reservation:editing.reservation,liveConfig:editing.config};state.phase='editing-api-reserved';await store.save(issueId,state);
              await checkpoint('before-editing-openai',state);await requireStillEligible('REQUEST_CHANGED_BEFORE_OPENAI');
              require('./openai.cjs').assertMandatoryInstructions(editing.request);
              const r=await transport('https://api.openai.com/v1/responses',{method:'POST',headers:{Authorization:'Bearer '+await secretReader(config.openaiSecretFile),'Content-Type':'application/json'},body:Buffer.from(JSON.stringify(editing.request)),maxBytes:bundle.openaiConfig.limits.maxResponseBytes,timeoutMs:bundle.openaiConfig.timeoutMs});
              let raw;try{raw=JSON.parse(r.bytes.toString());}catch{check(false,'OPENAI_RESPONSE_INVALID');}
              edited={raw,httpStatus:r.status,requestSha256:editing.requestSha256};await atomicJson(editingFile,edited);await checkpoint('editing-response-saved',state);
            }
            if(!state.editing.usageAccounted&&edited.raw?.usage){state.editing.usage=await accountResponse(edited.raw,editing.config,createLedger(path.join(stateRoot,'openai-budget.json')),editing.reservation);state.editing.usageAccounted=true;await store.save(issueId,state);}
            check(edited.httpStatus===200,'OPENAI_HTTP_FAILED');
            const nativeFile=path.join(dir,'native.docx');let bytes;
            try{bytes=await fs.readFile(nativeFile);}catch(e){check(e.code==='ENOENT','NATIVE_LOCAL_FILE_UNREADABLE');
              check(!recoveryOnly,'RECOVERY_NATIVE_FILE_REQUIRED');
              bytes=await downloadNativeDocument(edited.raw,{apiKey:await secretReader(config.openaiSecretFile),transport});
              await fs.writeFile(nativeFile,bytes);
            }
            result=await auditNativeDocument(templateBytes,bytes,reading);
            state.readingUsage||=state.usage;
            if(state.editing.usage)state.usage={...state.editing.usage,costEstimateUsd:state.readingUsage.costEstimateUsd+state.editing.usage.costEstimateUsd,inputTokens:state.readingUsage.inputTokens+state.editing.usage.inputTokens,outputTokens:state.readingUsage.outputTokens+state.editing.usage.outputTokens,stages:2};
          }else if(templateFields){check(saved?.httpStatus===200,'OPENAI_HTTP_FAILED');result=await evaluateTemplateFill(prepared,saved.raw);}
          else result=production||apiModel?await evaluateDynamicApi({...prepared,httpStatus:saved?.httpStatus||200,accountUsage:!state.usageAccounted},saved?.raw||prepared.mockResponse):localModel?await evaluateLocal(prepared,prepared.mockResponse):await evaluate(prepared,saved?.raw||prepared.mockResponse,{httpStatus:saved?.httpStatus||200,account:!state.usageAccounted});
          if(config.openaiMode==='live'){state.usageAccounted=true;state.usage=result.report.usage||state.usage;await store.save(issueId,state);}
          if(!result.bytes){
            state.pending=result.report.evaluation.expectedPending;state.pendingDetails=result.report.pendingDetails;
            state.documentPendingNotes=result.report.pendingNotes||[];
            await atomicJson(path.join(dir,'audit.json'),result.report);
            check(false,result.report.status==='documentary-review-blocked'?'DOCUMENTARY_REVIEW_REQUIRED':'PENDING_INFORMATION');
          }
          state.documentHasPending=result.report.hasPending===true;
          if(templateFields)state.documentaryValidation=result.report.documentaryValidation;
          state.documentPendingNotes=result.report.pendingNotes||[];
          await fs.writeFile(outputFile,result.bytes,{mode:0o600});await atomicJson(path.join(dir,'audit.json'),result.report);
          state.outputSha256=sha256(result.bytes);state.phase='generated';await store.save(issueId,state);await checkpoint('generated',state);
        }
        let output;try{output=await fs.readFile(outputFile);}catch{check(false,'LOCAL_RESULT_EXPIRED_REVIEW_REQUIRED');}
        check(sha256(output)===state.outputSha256,'LOCAL_RESULT_INTEGRITY_FAILED');
        if(templateFields){
          // Inclui retomadas: um DOCX antigo com pendencias nao herda validacao.
          const audit=await readJson(path.join(dir,'audit.json'));
          check(audit?.workflowVersion===fieldWorkflowVersion&&audit.technicalValidation==='passed'&&audit.documentGenerated===true&&audit.legalApproval===false&&['pending-human-review','requires-human-review'].includes(audit.documentaryValidation)&&audit.audit?.outputSha256===state.outputSha256,'DOCUMENTARY_DELIVERY_NOT_VALIDATED');
        }
        await requireStillEligible('REQUEST_CHANGED_DURING_PROCESSING');
        const mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        if(!state.assetUrl){const allocated=await linear.allocate('minuta-'+jobId.slice(0,16)+'.docx',output.length,mime);state.assetUrl=allocated.assetUrl;state.uploadTicket=allocated;state.phase='upload-allocated';await store.save(issueId,state);await checkpoint('upload-allocated',state);await linear.put(allocated,output,mime);delete state.uploadTicket;state.phase='uploaded';await store.save(issueId,state);}
        else{
          let accessible=false;try{accessible=sha256(await linear.download(state.assetUrl))===state.outputSha256;}catch{}
          if(!accessible){check(state.uploadTicket,'UPLOAD_RECOVERY_REVIEW_REQUIRED');await linear.put(state.uploadTicket,output,mime);delete state.uploadTicket;state.phase='uploaded';await store.save(issueId,state);}
        }
        await checkpoint('uploaded',state);
        await linear.attach(issueId,state.assetUrl,state.documentHasPending?(templateFields?'Minuta com pendências — revisão obrigatória':'Minuta para revisao — pendencias em vermelho'):production?'Minuta — revisao humana obrigatoria':'Minuta ficticia — revisao humana obrigatoria',{cartorioJobId:jobId,issueId,sha256:state.outputSha256,templateSha256:state.templateSha256,inputSha256:state.inputSha256,hasPending:state.documentHasPending===true,...(templateFields?{documentaryValidation:state.documentaryValidation,legalApproval:false}:{}),simulation:config.openaiMode==='mock'});
        await checkpoint('attached',state);
        const delivered=await linear.issue(issueId),attachment=delivered.attachments.nodes.find(a=>a.metadata?.cartorioJobId===jobId&&a.issue.id===issueId&&canonicalUrl(a.url)===canonicalUrl(state.assetUrl));
        check(attachment&&sha256(await linear.download(attachment.url))===state.outputSha256,'DELIVERY_NOT_CONFIRMED');
        check(snapshotOf(delivered)===snapshot&&config.linear.readyStateIds.includes(delivered.state.id)&&delivered.labels.nodes.some(l=>l.id===config.linear.requestLabelId),'REQUEST_CHANGED_DURING_PROCESSING');
        state.deliveryConfirmed=true;state.attachmentId=attachment.id;state.phase='delivered';await store.save(issueId,state);await checkpoint('delivery-confirmed',state);
        const marker='[cartorio-run:'+jobId+']';
        const reviewNotes=templateFields&&state.documentPendingNotes.length?'\n'+(state.documentHasPending?'Pendências em vermelho no documento.\n':'Informações da leitura:\n')+state.documentPendingNotes.join('\n'):state.documentHasPending?(nativeDocument?' Campos pendentes identificados por marcadores breves em vermelho. Pendencias: '+state.documentPendingNotes.join('; ')+'.':' Pendencias assinaladas em vermelho.'):'';
        await linear.comment(issueId,`${marker}\nMinuta ${config.openaiMode==='mock'?'com IA simulada':'com IA real'} entregue e integridade conferida por SHA-256. ${templateFields?'Minuta preliminar: exige conferência, saneamento e aprovação humana antes de qualquer utilização. In Review significa aguardando essa revisão; não é aprovação documental ou jurídica.':'Revisao juridica humana obrigatoria.'}${reviewNotes}${production?'':' Este teste utiliza exclusivamente dados ficticios.'}`,marker);
        await requireStillEligible('REQUEST_CHANGED_BEFORE_REVIEW');
        state.phase='review-intent';await store.save(issueId,state);
        await linear.review(issueId,config.linear.reviewStateId);await checkpoint('review-updated',state);
        const verified=production?await linear.issueStatus(issueId):await linear.issue(issueId);check(verified.state.id===config.linear.reviewStateId,'LINEAR_REVIEW_NOT_CONFIRMED');
        state.phase='complete';state.lastError=null;await store.save(issueId,state);results.push({issueId,scenario:production?'production':fixture.scenario,status:'delivered-in-review',simulation:config.openaiMode==='mock',sha256:state.outputSha256,attachmentId:state.attachmentId});
      }catch(e){
        const code=/^[A-Z0-9_]+$/.test(e.code||'')?e.code:'ISSUE_PROCESSING_FAILED';
        if(!state&&issue&&issue.team.id===config.linear.teamId)state={issueId,phase:'pending',jobId:null,requestId:null,createdAt:new Date().toISOString()};
        if(state){state.lastError=code;await store.save(issueId,state);}
        let pendingCommentRecorded=false;
        let mayComment=!!issue&&issue.team.id===config.linear.teamId;
        if(mayComment&&production){try{const status=await linear.issueStatus(issueId);mayComment=config.linear.readyStateIds.includes(status.state.id)&&status.labels.nodes.some(l=>l.id===config.linear.requestLabelId);}catch{mayComment=false;}}
        if(mayComment){
          const marker='[cartorio-pending:'+hash({issueId,job:state?.jobId||null,description:hash(issue.description),code})+']';
          const details=(state?.pendingDetails||[]).map(p=>p.code+(p.documentId?' ('+p.documentId+(p.page?', pagina '+p.page:'')+')':'')).join('; ');
          const documentary=code==='DOCUMENTARY_REVIEW_REQUIRED'?(state?.documentPendingNotes||[]).join('\n'):'';
          try{await linear.comment(issueId,`${marker}\nPendente de revisao humana. Codigo: ${code}. ${state?.pending?.length?'Campos: '+state.pending.filter(id=>/^[A-Za-z0-9_.-]{1,120}$/.test(id)).join(', ')+'. ':''}${details?'Pendencias: '+details+'. ':''}A automacao nao concluiu esta solicitacao.${documentary?'\n'+documentary:''}`,marker);pendingCommentRecorded=true;}catch{}
        }
        results.push({issueId,status:'pending',code,pendingCommentRecorded});
      }
    }
    const activeJobIds=[];let retentionSafe=true;
    for(const id of await store.pendingIds()){try{const s=await store.load(id);if(s?.jobId)activeJobIds.push(s.jobId);}catch{retentionSafe=false;}}
    const retention=retentionSafe?await cleanArtifacts(artifactRoot,config.retention,Date.now(),activeJobIds):{skipped:true,reason:'JOB_STATE_REVIEW_REQUIRED',receiptsPreserved:true};
    return {status:'batch-finished',trigger,simulation:config.openaiMode==='mock',results,retention,realVisualInterpretationVerified:config.openaiMode==='live'&&results.some(r=>r.scenario==='digitalizada'&&r.status==='delivered-in-review')};
  });
}
module.exports={runBatch,parseOrder,inputPdf,inputPdfs,requestSnapshot,cleanArtifacts};
