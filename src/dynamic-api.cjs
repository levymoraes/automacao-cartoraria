'use strict';
// API aceita apenas o contexto ficticio aprovado. DOCX e dados institucionais
// nunca entram no request. Chamada, reserva e retomada usam o fluxo existente.
const {sha256,openDocx,paragraphsOf}=require('./docx.cjs');
const {assertApproved}=require('./openai.cjs'),{compileInventory,digest}=require('./dynamic-template.cjs');
const {validateExtraction}=require('./extraction.cjs'),{applyLocal}=require('./local-template.cjs');
const {collectStories}=require('./operation.cjs'),{requireCondition:check}=require('./errors.cjs');
const clone=value=>JSON.parse(JSON.stringify(value));
const schema=clone(require('../config/extraction.schema.json'));
delete schema.properties.observations;schema.required=schema.required.filter(k=>k!=='observations');
schema.$defs.replacement.properties.category={type:'string',enum:['person-name','identity-document','personal-date','parentage','nationality','birthplace','profession','civil-status','capacity','marriage','address','property-identification','property-characteristic','transaction-value','transaction-date','transaction-payment','tax','certificate','record','other-variable']};schema.$defs.replacement.required.push('category');
schema.properties.templateCoverage={type:'array',items:{type:'object',additionalProperties:false,required:['part','paragraph','textSha256','reviewed'],properties:{part:{type:'string'},paragraph:{type:'integer'},textSha256:{type:'string'},reviewed:{type:'boolean'}}}};
schema.properties.inventoryAudit={type:'object',additionalProperties:false,required:['status','unmappedVariables'],properties:{status:{type:'string',enum:['passed','pending']},unmappedVariables:{type:'array',items:{type:'string'}}}};
schema.required.push('templateCoverage','inventoryAudit');
async function prepareDynamicApi({templateBytes,pdfBytes,model,task,fixture,config}){
 const material=fixture.apiMaterial;
 check(material?.approvedFictionalContext===true&&material.externalDocxSharingAllowed===false&&material.oldPersonalDataResidues===0&&material.institutionAndStaffMasked===true&&material.logoSentToApi===false,'FICTIONAL_API_CONTEXT_NOT_APPROVED');
 check(material.contextSha256===digest(material.context)&&material.templateSha256===sha256(templateBytes)&&material.context.templateSha256===sha256(templateBytes)&&model.origin==='approved-fictional-derivative','API_CONTEXT_INTEGRITY_FAILED');
 check(fixture.localOnly===false&&fixture.generatedFictionalPdf===true&&fixture.pdfSha256===sha256(pdfBytes)&&fixture.templateSha256===sha256(templateBytes),'FICTIONAL_INPUT_NOT_APPROVED');
 require('./openai.cjs').assertConfiguration(config);check(config.store===false,'OPENAI_CONFIGURATION_MISMATCH');
 check(pdfBytes.length<=config.limits.maxPdfBytes&&pdfBytes.subarray(0,5).toString()==='%PDF-','INVALID_PDF_INPUT');
 const pdfApproval=assertApproved(pdfBytes,'pdf',fixture.approvals);check(Number.isInteger(pdfApproval.pages)&&pdfApproval.pages>0&&pdfApproval.pages<=config.limits.maxPdfPages,'PDF_PAGE_LIMIT');
 const documents=[{id:'pacote-ficticio',sha256:sha256(pdfBytes),pages:pdfApproval.pages}];
 // Apenas a tarefa ficticia cadastrada pode compor a mensagem externa.
 check(digest(task)===fixture.taskSha256,'API_TASK_NOT_APPROVED');
 const context={...material.context,task,taskSha256:fixture.taskSha256,documents};
 const text=JSON.stringify(context);check(text.length<=config.limits.maxTemplateCharacters,'TEMPLATE_CONTEXT_LIMIT');
 const instructions=fixture.instructions;check(typeof instructions==='string'&&sha256(Buffer.from(instructions))===fixture.instructionsSha256,'API_INSTRUCTIONS_CHANGED');
 const request={model:config.model,service_tier:config.service_tier,reasoning:config.reasoning,store:false,prompt_cache_options:{mode:'explicit'},max_output_tokens:config.max_output_tokens,instructions:require('./openai.cjs').systemInstructions(instructions),text:{format:{type:'json_schema',name:'dynamic_minuta_analysis',strict:true,schema}},input:[{role:'user',content:[{type:'input_text',text}, {type:'input_file',filename:'pacote-ficticio.pdf',detail:'high',file_data:'data:application/pdf;base64,'+pdfBytes.toString('base64')}]}]};
 return {request,requestSha256:digest(request),templateBase64:templateBytes.toString('base64'),model,task,documents,fixture,config,scenario:fixture.scenario,mode:'mock',mockResponse:fixture.mockResponse};
}
async function evaluateDynamicApi(prepared,raw){
 let usage=null;
 if(prepared.mode==='live'&&prepared.accountUsage!==false)usage=await require('./openai.cjs').accountResponse(raw,prepared.config,require('./budget.cjs').createLedger(prepared.ledgerPath),prepared.reservation);
 check(!prepared.httpStatus||prepared.httpStatus>=200&&prepared.httpStatus<300,'OPENAI_HTTP_FAILED');
 check(raw?.status==='completed','OPENAI_RESPONSE_INCOMPLETE');
 const output=(raw.output||[]).flatMap(o=>o.type==='message'?(o.content||[]):[]);check(!output.some(c=>c.type==='refusal'),'OPENAI_REFUSAL');
 const texts=output.filter(c=>c.type==='output_text');check(texts.length===1,'OPENAI_OUTPUT_MISSING');let value;try{value=JSON.parse(texts[0].text);}catch{check(false,'OPENAI_OUTPUT_NOT_JSON');}
 if(prepared.reviewTargets){
  check(require('./generated/production-review-validator.cjs')(value),'PRODUCTION_REVIEW_SCHEMA_INVALID');
  check(value.templateSha256===prepared.reviewBase.templateSha256&&digest(value.documents)===digest(prepared.documents),'REVIEW_CONTEXT_MISMATCH');
  const patch=value,targets=prepared.reviewTargets;check(new Set(patch.fields.map(f=>f.fieldId)).size===patch.fields.length&&patch.fields.every(f=>targets.some(t=>t.fieldId===f.fieldId)),'REVIEW_FIELD_SCOPE_MISMATCH');
  value=clone(prepared.reviewBase);const ids=new Set(targets.flatMap(t=>t.occurrences));value.pending=value.pending.filter(p=>!ids.has(p.fieldId));
  for(const t of targets){const answer=patch.fields.find(f=>f.fieldId===t.fieldId),mismatch=answer?.sources.some(s=>s.entityId!==t.entityId||s.property!==t.property);
   for(const id of t.occurrences){const r=value.replacements.find(r=>r.fieldId===id);check(r&&r.entityId===t.entityId&&r.property===t.property,'REVIEW_FIELD_SCOPE_MISMATCH');
    Object.assign(r,answer&&!mismatch?{status:answer.status,value:answer.value,confirmedUnchanged:answer.confirmedUnchanged,sources:answer.sources}:{status:answer?'conflict':'missing',value:null,confirmedUnchanged:false,sources:[]});
    if(r.status!=='verified')value.pending.push({code:answer?'REVIEW_REQUIRED':'CORRECTION_FIELD_OMITTED',fieldId:id,entityId:r.entityId,documentId:null,page:null});
   }
  }
  value.documentSegments=require('./production.cjs').indexEvidencePages(value,prepared.documents);
 }else if(prepared.production){
  // Compatibilidade com respostas antigas arquivadas: o indice redundante
  // nao e evidencia e nao determina se uma informacao foi comprovada.
  const {documentSegments,...data}=value;check(require('./generated/production-validator.cjs')(data),'PRODUCTION_RESPONSE_SCHEMA_INVALID');
  value={...data,documentSegments:require('./production.cjs').indexEvidencePages(data,prepared.documents)};
 }
 else check(require('./generated/dynamic-validator.cjs')(value),'DYNAMIC_RESPONSE_SCHEMA_INVALID');
 if(prepared.production){
  if(!prepared.reviewTargets)value=await require('./production.cjs').locateReplacements(value,Buffer.from(prepared.templateBase64,'base64'),{allowPending:prepared.config.deliverPendingDrafts===true});
  const {executionRole}=require('./operation.cjs');value.templateEntities=value.templateEntities.map(e=>({...e,role:executionRole(e)}));
  value.entities=value.entities.map(e=>{
   let role=executionRole(e);const slot=value.templateEntities.find(s=>s.id===e.slotId),partner=value.entities.find(p=>p.id===e.spouseId);
   // Co-vendedor/comprador pode ser descrito como conjuge no modelo. O nome
   // do papel nao muda a composicao quando ambos os vinculos reciprocos e
   // o lado da transacao sao os mesmos; a identidade continua sendo validada.
   if(e.kind==='person'&&slot?.kind==='person'&&/^(seller|buyer)(_spouse)?$/.test(role)&&/^(seller|buyer)(_spouse)?$/.test(slot.role)&&role.replace(/_spouse$/,'')===slot.role.replace(/_spouse$/,'')&&partner?.spouseId===e.id&&partner.slotId===slot.spouseId)role=slot.role;
   return {...e,role};
  });
 }
 const entities=prepared.production?clone(value.templateEntities):value.entities.map(e=>({id:e.slotId,kind:e.kind,role:e.role,spouseId:e.spouseId?value.entities.find(s=>s.id===e.spouseId)?.slotId:null}));entities.push({id:'institution',kind:'institution',role:'staff',spouseId:null});
 const analysis={schemaVersion:1,templateSha256:value.templateSha256,taskSha256:digest(prepared.task),documentsSha256:digest(prepared.documents),entities,coverage:value.templateCoverage,audit:value.inventoryAudit,pending:value.pending,fields:value.replacements.map(r=>({id:r.fieldId,entityId:value.entities.find(e=>e.id===r.entityId)?.slotId,property:r.property,category:r.category,part:r.part,paragraph:r.paragraph,start:r.start,end:r.end,oldText:r.oldValue}))};
 const allowPending=prepared.production&&prepared.config.deliverPendingDrafts===true&&value.replacements.length>0&&value.inventoryAudit.unmappedVariables.length===0;
 if((value.pending.length||value.inventoryAudit.status!=='passed')&&!allowPending)return {report:{status:'pending-dynamic-analysis',documentGenerated:false,usage,evaluation:{expectedPending:value.pending.map(p=>p.fieldId).filter(id=>/^[A-Za-z0-9_.-]{1,120}$/.test(id||''))},pendingDetails:value.pending.slice(0,30).map(p=>({code:/^[A-Z0-9_]{1,128}$/.test(p.code.toUpperCase())?p.code.toUpperCase():'REVIEW_REQUIRED',entityId:/^[A-Za-z0-9_.-]{1,120}$/.test(p.entityId||'')?p.entityId:null,documentId:prepared.documents.some(d=>d.id===p.documentId)?p.documentId:null,page:Number.isInteger(p.page)&&p.page>0&&p.page<=Math.max(...prepared.documents.map(d=>d.pages))?p.page:null})),simulation:prepared.mode==='mock'}};
 const profile=await compileInventory({templateBytes:Buffer.from(prepared.templateBase64,'base64'),model:prepared.model,task:prepared.task,documents:prepared.documents,analysis,allowDataPending:allowPending});
 const extraction=clone(value);delete extraction.templateCoverage;delete extraction.inventoryAudit;delete extraction.templateEntities;extraction.observations=[];
 if(allowPending&&value.inventoryAudit.status==='pending'&&!extraction.pending.length)extraction.pending.push({code:'REVISAR_DOCUMENTOS',fieldId:null,entityId:null,documentId:null,page:null});
 for(const r of extraction.replacements){delete r.category;const key=r.entityId+':'+r.property,existing=extraction.observations.find(o=>o.entityId+':'+o.property===key);if(existing){check(existing.status===r.status&&existing.value===r.value,'REPEATED_VALUE_CONTRADICTION');for(const s of r.sources)if(!existing.sources.some(p=>digest(p)===digest(s)))existing.sources.push(s);}else extraction.observations.push({entityId:r.entityId,property:r.property,status:r.status,value:r.value,sources:[...r.sources]});}
 const plan=validateExtraction(extraction,profile,prepared.documents,{allowPending});
 for(const c of profile.operation.constraints){const e=extraction.entities.find(e=>e.slotId===c.entityId),o=e&&extraction.observations.find(o=>o.entityId===e.id&&o.property===c.property);if(allowPending&&o&&o.status!=='verified')continue;check(o?.status==='verified'&&o.value===c.value,'PARTICIPANT_CONSTRAINT_MISMATCH');}
 const result=await applyLocal(Buffer.from(prepared.templateBase64,'base64'),profile,plan),doc=await openDocx(result.bytes),stories=collectStories(doc,paragraphsOf);
 if(!prepared.production){const expected=prepared.fixture.expectedStories;check(Array.isArray(expected)&&expected.length===stories.length&&expected.every(e=>stories.some(s=>s.part===e.part&&s.paragraph===e.paragraph&&sha256(Buffer.from(s.text))===e.textSha256)),'FICTIONAL_DOCUMENT_GROUND_TRUTH_MISMATCH');}
 const hasPending=!!(plan.pendingAnnotations?.length||plan.pendingNotes?.length);
 return {bytes:result.bytes,report:{status:hasPending?'validated-draft-with-pending':'validated-dynamic-api-docx',documentGenerated:true,hasPending,simulation:prepared.mode==='mock',usage,evaluation:{variableOccurrences:profile.fields.length,expectedPending:(plan.pendingAnnotations||[]).map(n=>n.fieldId)},audit:result.audit,request:{model:prepared.request.model,serviceTier:prepared.request.service_tier,reasoning:prepared.request.reasoning.effort,contextOnly:true,originalDocxTransmitted:false}}};
}
async function authorizeDynamicApi(prepared,config,approval,ledgerPath){
 check(approval?.authorized===true&&typeof approval.evidence==='string'&&approval.evidence.trim().length>=10&&Date.parse(approval.expiresAt)>Date.now(),'PAID_API_NOT_AUTHORIZED');
 check(approval.scenario===prepared.scenario&&approval.requestSha256===prepared.requestSha256&&approval.maxCalls===1&&approval.maxBudgetUsd>0&&approval.maxBudgetUsd<=0.5,'APPROVAL_REQUEST_MISMATCH');
 config=clone(config);config.live={enabled:true,budgetApprovalId:approval.approvalId,maxCalls:1,maxBudgetUsd:approval.maxBudgetUsd};
 prepared.config=config;prepared.mode='live';prepared.ledgerPath=ledgerPath;
 prepared.reservation=await require('./openai.cjs').authorizeRequest(prepared,config,require('./budget.cjs').createLedger(ledgerPath));return prepared;
}
module.exports={schema,prepareDynamicApi,evaluateDynamicApi,authorizeDynamicApi};
