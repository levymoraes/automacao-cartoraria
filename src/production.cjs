'use strict';
// Caminho operacional: nenhum fixture, resposta simulada ou gabarito e consultado.
const {PDFDocument}=require('pdf-lib');
const {openDocx,paragraphsOf,sha256}=require('./docx.cjs');
const {collectStories,coverageOf}=require('./operation.cjs');
const {protectedLocations,fixedCoincidences,criticalCandidates,digest}=require('./dynamic-template.cjs');
const {requireCondition:check}=require('./errors.cjs');
const schema=JSON.parse(JSON.stringify(require('./dynamic-api.cjs').schema));
schema.properties.templateEntities={type:'array',minItems:1,items:{type:'object',additionalProperties:false,required:['id','kind','role','spouseId'],properties:{id:{type:'string'},kind:{type:'string',enum:['person','organization','property','transaction','marriage']},role:{type:'string'},spouseId:{type:['string','null']}}}};
schema.required.push('templateEntities');
// A IA informa trechos literais. Posicoes e repeticoes pertencem ao motor.
for(const key of ['part','paragraph','start','end']){delete schema.$defs.replacement.properties[key];schema.$defs.replacement.required=schema.$defs.replacement.required.filter(k=>k!==key);}
for(const key of ['contextBefore','contextAfter']){schema.$defs.replacement.properties[key]={type:'string'};schema.$defs.replacement.required.push(key);}
delete schema.properties.templateCoverage;schema.required=schema.required.filter(k=>k!=='templateCoverage');
// A pagina/entidade de cada evidencia e a unica fonte desse indice auxiliar.
// Pedir a mesma classificacao duas vezes a IA criava contradicoes artificiais.
delete schema.properties.documentSegments;schema.required=schema.required.filter(k=>k!=='documentSegments');
const reviewField=JSON.parse(JSON.stringify(schema.$defs.replacement));
for(const key of Object.keys(reviewField.properties))if(!['fieldId','value','status','confirmedUnchanged','sources'].includes(key)){delete reviewField.properties[key];reviewField.required=reviewField.required.filter(k=>k!==key);}
const reviewSchema={type:'object',additionalProperties:false,required:['templateSha256','documents','fields'],properties:{templateSha256:schema.properties.templateSha256,documents:schema.properties.documents,fields:{type:'array',items:{$ref:'#/$defs/reviewField'}}},$defs:{...schema.$defs,reviewField}};
function indexEvidencePages(value,documents){
 return documents.flatMap(d=>Array.from({length:d.pages},(_,i)=>({documentId:d.id,type:'other',startPage:i+1,endPage:i+1,entityIds:[...new Set(value.replacements.flatMap(r=>r.sources.filter(s=>s.documentId===d.id&&s.documentSha256===d.sha256&&s.page===i+1&&s.entityId===r.entityId&&s.property===r.property).map(s=>s.entityId)))]})));
}
function textForApi(p,fixed){let text=p.text;for(const f of fixed.filter(f=>f.part===p.part&&f.paragraph===p.paragraph).sort((a,b)=>b.start-a.start))text=text.slice(0,f.start)+' '.repeat(f.end-f.start)+text.slice(f.end);return text;}
const contextText=text=>text.replace(/\s+/gu,' ');
function uniqueAnchor(stories,text){const hits=[];if(!text)return null;for(const p of stories)for(let at=p.folded.indexOf(text);at>=0;at=p.folded.indexOf(text,at+1))hits.push({p,at});return hits.length===1?hits[0]:null;}
async function locateReplacements(value,templateBytes,{allowPending=false}={}){
 const doc=await openDocx(templateBytes),stories=collectStories(doc,paragraphsOf),fixed=protectedLocations(stories),found=[];
 const evidenceMismatch=r=>r.sources.some(s=>s.entityId!==r.entityId||s.property!==r.property);
 const people=new Set(value.entities.filter(e=>['person','organization'].includes(e.kind)).map(e=>e.id));
 for(const p of stories){
  p.preview=textForApi(p,fixed);p.folded=contextText(p.preview);const anchors=[];
  for(const r of value.replacements.filter(r=>r.category==='person-name'&&r.property==='nome'&&people.has(r.entityId)))for(let start=p.text.indexOf(r.oldValue);start>=0;start=p.text.indexOf(r.oldValue,start+1)){
   const end=start+r.oldValue.length;if(/[\p{L}\p{N}]/u.test(p.text[start-1]||'')||/[\p{L}\p{N}]/u.test(p.text[end]||'')||fixed.some(f=>f.part===p.part&&f.paragraph===p.paragraph&&f.start<end&&f.end>start))continue;
   anchors.push({start,end,entityId:r.entityId});
  }
  p.people=anchors.filter(a=>!anchors.some(b=>a!==b&&b.start<=a.start&&b.end>=a.end&&b.end-b.start>a.end-a.start)).sort((a,b)=>a.start-b.start);
 }
 for(const source of value.replacements){
  // Evidencia de outra entidade/propriedade nao comprova o valor. Nao
  // transferir fatos entre pessoas: a minuta parcial recebe uma pendencia.
  const mismatch=evidenceMismatch(source);
  const r=allowPending&&mismatch?{...source,status:'conflict',value:null,sources:[],confirmedUnchanged:false}:source;
  check(r.oldValue.trim().length>0,'REPLACEMENT_TEXT_NOT_FOUND');
  const eligible=[],matched=[];
  for(const p of stories){
  const preview=p.preview;
  for(let start=p.text.indexOf(r.oldValue);start>=0;start=p.text.indexOf(r.oldValue,start+1)){
   const end=start+r.oldValue.length,before=p.text.slice(0,start),after=p.text.slice(end);
   if(/[\p{L}\p{N}]$/u.test(r.oldValue)&&/^[\p{L}\p{N}]/u.test(after)||/^[\p{L}\p{N}]/u.test(r.oldValue)&&/[\p{L}\p{N}]$/u.test(before))continue;
   // Recusar um identificador abreviado dentro de outro maior: ABC-12 nao
   // corresponde a ABC-12-XYZ. Isto restringe a busca; nao suprime conflitos.
   const identifier=/^[A-Za-z0-9_.:/-]+$/.test(r.oldValue)&&/[A-Za-z]/.test(r.oldValue)&&/[-_/:]/.test(r.oldValue);
   if(identifier&&(/^[A-Za-z0-9_/-]/.test(after)||/[A-Za-z0-9_/-]$/.test(before)))continue;
   // Contextos sao literais: comparar tambem com a mesma vista mascarada
   // enviada a IA. Nao tratar espacos como curingas nem liberar areas fixas.
   const originalContext=before.endsWith(r.contextBefore)&&after.startsWith(r.contextAfter);
   const apiContext=preview.slice(0,start).endsWith(r.contextBefore)&&preview.slice(end).startsWith(r.contextAfter);
   if(fixed.some(f=>f.part===p.part&&f.paragraph===p.paragraph&&f.start<end&&f.end>start))continue;
   // Desambiguar qualificacoes pela pessoa explicitamente nomeada no modelo,
   // a mesma correspondencia exigida na validacao final. Nunca atribuir a data
   // ou documento de um participante ao bloco de outra pessoa.
   if(people.has(r.entityId)&&r.category!=='person-name'&&value.replacements.some(o=>o!==source&&o.oldValue===r.oldValue&&o.entityId!==r.entityId&&people.has(o.entityId))){
    const owner=[...p.people].reverse().find(a=>a.end<=start),person=value.entities.find(e=>e.id===r.entityId),partner=value.entities.find(e=>e.id===owner?.entityId);
    const shared=r.category==='civil-status'&&/^casad[oa]s$/iu.test(r.oldValue)&&person?.spouseId===partner?.id&&partner?.spouseId===person?.id;
    if(owner?.entityId!==r.entityId&&!shared)continue;
   }
   const {contextBefore,contextAfter,...replacement}=r,candidate={...replacement,part:p.part,paragraph:p.paragraph,start,end};eligible.push(candidate);
   const whitespaceContext=contextText(preview.slice(0,start)).endsWith(contextText(r.contextBefore))&&contextText(preview.slice(end)).startsWith(contextText(r.contextAfter));
   if(originalContext||apiContext||whitespaceContext)matched.push(candidate);
  }
  }
  // Uma referencia literal unica pode delimitar a qualificacao sem estar
  // colada ao dado. Aceitar somente UMA ocorrencia no intervalo delimitado;
  // outras referencias do mesmo dado encerram o intervalo. Nao escolher
  // primeira ocorrencia, nao inventar texto e nao modificar as areas fixas.
  if(!matched.length){
   const leftText=contextText(r.contextBefore),left=uniqueAnchor(stories,leftText);
   if(left){const from=left.at+leftText.length;let to=left.p.folded.length;
    for(const peer of value.replacements.filter(x=>x!==r&&x.oldValue===r.oldValue)){const anchor=contextText(peer.contextBefore);if(anchor===leftText)continue;const other=uniqueAnchor(stories,anchor);if(other&&other.p===left.p&&other.at>=from)to=Math.min(to,other.at);}
    const scope=left.p.folded.slice(from,to),rightText=contextText(r.contextAfter);
    if(!rightText||scope.includes(rightText)){const scoped=eligible.filter(c=>c.part===left.p.part&&c.paragraph===left.p.paragraph&&contextText(left.p.preview.slice(0,c.start)).length>=from&&contextText(left.p.preview.slice(0,c.end)).length<=to);if(scoped.length===1)matched.push(scoped[0]);}
   }
  }
  if(!matched.length){const error=new(require('./errors.cjs').Blocked)('REPLACEMENT_TEXT_NOT_FOUND');error.details={fieldId:r.fieldId,entityId:r.entityId,property:r.property};throw error;}found.push(...matched);
 }
 // Um nome completo pode conter outro nome. O trecho mais especifico prevalece;
 // qualquer outra disputa pelo mesmo trecho exige contexto, nunca um palpite.
 const candidates=found.filter(r=>!found.some(o=>o!==r&&o.part===r.part&&o.paragraph===r.paragraph&&o.start<=r.start&&o.end>=r.end&&o.end-o.start>r.end-r.start&&o.category===r.category&&['person-name','parentage'].includes(r.category)));
 const replacements=[],ids=new Map(),fieldIds=new Map(),reviewPending=[];
 for(const r of candidates){
  const overlap=replacements.find(o=>o.part===r.part&&o.paragraph===r.paragraph&&o.start<r.end&&o.end>r.start);
  if(overlap){
   const sameTarget=overlap.start===r.start&&overlap.end===r.end&&overlap.entityId===r.entityId&&overlap.category===r.category;
   if(allowPending&&sameTarget&&(overlap.property===r.property&&(overlap.status!==r.status||overlap.value!==r.value)||overlap.property!==r.property&&['certificate','record','transaction-date','tax'].includes(r.category))){
    // Propriedades distintas disputam um unico campo: nao reconciliar fatos.
    // Entregar APENAS um marcador de pendencia nesse local, sem valor antigo
    // nem qualquer valor novo nao resolvido. As outras substituicoes seguem.
    overlap.status='conflict';overlap.value=null;overlap.sources=[];overlap.confirmedUnchanged=false;
    if(!reviewPending.some(p=>p.fieldId===overlap.fieldId))reviewPending.push({code:'AMBIGUOUS_FIELD_MAPPING',fieldId:overlap.fieldId,entityId:overlap.entityId,documentId:null,page:null});continue;
   }
   if(!(sameTarget&&overlap.property===r.property&&overlap.value===r.value&&overlap.status===r.status)){const error=new(require('./errors.cjs').Blocked)('AMBIGUOUS_TEXT_REPLACEMENT');error.details=[overlap,r].map(c=>({fieldId:c.fieldId,entityId:c.entityId,property:c.property,category:c.category,start:c.start,end:c.end}));throw error;}continue;
  }
  const n=(ids.get(r.fieldId)||0)+1;ids.set(r.fieldId,n);const fieldId=r.fieldId.slice(0,100)+'.'+n;
  if(!fieldIds.has(r.fieldId))fieldIds.set(r.fieldId,[]);fieldIds.get(r.fieldId).push(fieldId);replacements.push({...r,fieldId});
 }
 const evidencePending=allowPending?replacements.filter(r=>r.status==='conflict'&&value.replacements.some(s=>r.fieldId.startsWith(s.fieldId+'.')&&evidenceMismatch(s))).map(r=>({code:'EVIDENCE_ENTITY_MISMATCH',fieldId:r.fieldId,entityId:r.entityId,documentId:null,page:null})):[];
 const pending=[...value.pending.flatMap(p=>fieldIds.has(p.fieldId)?fieldIds.get(p.fieldId).map(fieldId=>({...p,fieldId})):[p]),...reviewPending,...evidencePending];
 let entities=value.entities,templateEntities=value.templateEntities;
 if(allowPending){
  const fields=replacements.map(r=>({...r,id:r.fieldId,oldText:r.oldValue})),coincidence=fixedCoincidences(doc,stories,fields,fixed);
  const uncovered=criticalCandidates(stories,fields,fixed).filter(c=>['date','cpf','cnpj','identifier','value','parentage'].includes(c.type)&&![...fields,...fixed,...coincidence.fixedOldValueExceptions].some(f=>f.part===c.part&&f.paragraph===c.paragraph&&f.start<=c.start&&f.end>=c.end));
  if(uncovered.length){
   // Entidade tecnica local, sem pessoa ou dado novo. Serve exclusivamente
   // para retirar valores antigos nao cobertos e exibir a falta no rascunho.
   let id='local-unresolved';while(entities.some(e=>e.id===id)||templateEntities.some(e=>e.id===id))id+='-review';
   const needsUnknown=uncovered.some(c=>c.type!=='parentage'||!stories.find(p=>p.part===c.part&&p.paragraph===c.paragraph)?.people.some(a=>a.end<=c.start));
   if(needsUnknown){const entity={id,kind:'transaction',role:'revisao local de campo nao identificado',spouseId:null};entities=[...entities,{...entity,slotId:id}];templateEntities=[...templateEntities,entity];}
   for(const c of uncovered){
    if(replacements.some(r=>r.part===c.part&&r.paragraph===c.paragraph&&r.start<c.end&&r.end>c.start))continue;
    const p=stories.find(p=>p.part===c.part&&p.paragraph===c.paragraph),fieldId=id+'.'+replacements.length;
    const owner=c.type==='parentage'?[...p.people].reverse().find(a=>a.end<=c.start):null,entityId=owner?.entityId||id;
    replacements.push({fieldId,entityId,property:c.property||(c.type==='date'?'data':c.type==='value'?'preco':c.type),category:c.type==='date'?'transaction-date':c.type==='value'?'transaction-value':c.type==='parentage'?'parentage':'identity-document',oldValue:p.text.slice(c.start,c.end),value:null,status:'missing',confirmedUnchanged:false,sources:[],part:c.part,paragraph:c.paragraph,start:c.start,end:c.end});
    pending.push({code:'EXTRACTION_FIELD_OMITTED',fieldId,entityId,documentId:null,page:null});
   }
  }
 }
 return {...value,entities,templateEntities,replacements,pending,templateCoverage:coverageOf(doc,paragraphsOf).map(p=>({...p,reviewed:true}))};
}
async function prepareProduction({templateBytes,pdfs,model,task,config,instructions,correction=null}){
 check(model.userApprovedUse===true&&model.inventoryMode==='dynamic-per-request','TEMPLATE_NOT_APPROVED');
 require('./openai.cjs').assertConfiguration(config);check(config.store===false,'OPENAI_CONFIGURATION_MISMATCH');
 check(typeof task.description==='string'&&task.description.trim().length>0,'TASK_INSTRUCTIONS_REQUIRED');
 check(typeof instructions==='string'&&instructions.length>100,'DYNAMIC_INSTRUCTIONS_REQUIRED');
 check(pdfs.length>0&&pdfs.length<=(config.limits.maxPdfs||4)&&new Set(pdfs.map(p=>p.id)).size===pdfs.length,'INVALID_INPUT_DOCUMENTS');
 check(pdfs.reduce((sum,p)=>sum+p.bytes.length,0)<=config.limits.maxPdfBytes,'PDF_TOTAL_SIZE_LIMIT');
 const documents=[];
 for(const p of pdfs){
  check(/^[A-Za-z0-9_-]{1,80}$/.test(p.id)&&p.bytes.subarray(0,5).toString()==='%PDF-','INVALID_PDF_INPUT');
  // String evita diferencas de Uint8Array entre os contextos isolados.
  let pdf;try{pdf=await PDFDocument.load(p.bytes.toString('base64'),{updateMetadata:false});}catch{check(false,'PDF_UNREADABLE_OR_ENCRYPTED');}
  const pages=pdf.getPageCount();check(pages>0&&pages<=config.limits.maxPdfPages,'PDF_PAGE_LIMIT');
  documents.push({id:p.id,sha256:sha256(p.bytes),pages});
 }
 check(documents.reduce((n,d)=>n+d.pages,0)<=config.limits.maxPdfPages,'PDF_PAGE_LIMIT');
 const doc=await openDocx(templateBytes),stories=collectStories(doc,paragraphsOf),fixed=protectedLocations(stories);
 // Elementos institucionais ficam locais; a IA recebe somente o texto literal.
 const template=stories.map(p=>({...p,text:textForApi(p,fixed)}));
 const context={templateSha256:doc.hash,task,documents,template};
 const text=JSON.stringify(context);check(text.length<=config.limits.maxTemplateCharacters,'TEMPLATE_CONTEXT_LIMIT');
 const request={model:config.model,service_tier:config.service_tier,reasoning:config.reasoning,store:false,prompt_cache_options:{mode:'explicit'},max_output_tokens:config.max_output_tokens,instructions:require('./openai.cjs').systemInstructions(instructions),text:{format:{type:'json_schema',name:'production_minuta_analysis',strict:true,schema}},input:[{role:'user',content:[{type:'input_text',text},...pdfs.map(p=>({type:'input_file',filename:p.id+'.pdf',detail:'high',file_data:'data:application/pdf;base64,'+p.bytes.toString('base64')}))]}]};
 let reviewTargets,reviewBase;
 if(correction){
  const prior=correction.previousResponse;
  const {documentSegments,...priorData}=prior||{};
  check(correction.issueId===task.issueId&&prior?.templateSha256===doc.hash&&digest(prior.documents)===digest(documents)&&require('./generated/production-validator.cjs')(priorData),'CORRECTION_CONTEXT_MISMATCH');
  reviewBase=await locateReplacements({...priorData,documentSegments:indexEvidencePages(priorData,documents)},templateBytes,{allowPending:true});
  const requested=new Set((correction.validationFailures||[]).map(f=>f.fieldId));
  const selected=reviewBase.replacements.filter(r=>requested.size?requested.has(r.fieldId):r.status!=='verified');
  check(selected.length>0,'NO_REVIEW_FIELDS');check(!requested.size||[...requested].every(id=>selected.some(r=>r.fieldId===id)),'REVIEW_FIELD_NOT_FOUND');
  const groups=new Map();for(const r of selected){const key=JSON.stringify([r.entityId,r.property,r.oldValue]);if(!groups.has(key))groups.set(key,{fieldId:r.fieldId,entityId:r.entityId,property:r.property,oldValue:r.oldValue,occurrences:reviewBase.replacements.filter(f=>f.entityId===r.entityId&&f.property===r.property&&f.oldValue===r.oldValue).map(f=>f.fieldId)});}reviewTargets=[...groups.values()];
  const fields=reviewTargets.map(t=>{const r=selected.find(r=>r.fieldId===t.fieldId),p=template.find(p=>p.part===r.part&&p.paragraph===r.paragraph);return {...t,contextBefore:p.text.slice(Math.max(0,r.start-140),r.start),contextAfter:p.text.slice(r.end,r.end+140),proposedValue:r.value};});
  request.text.format={type:'json_schema',name:'production_minuta_review',strict:true,schema:reviewSchema};
  request.instructions+='\nCONFERENCIA DIRECIONADA: releia os PDFs procurando exclusivamente os campos solicitados. proposedValue e uma proposta nao comprovada, nunca uma fonte. Retorne uma entrada por fieldId solicitado, com nova evidencia correta, ou value=null e status missing/conflict/illegible. Nao devolva o documento inteiro, nao altere entidades, nao reanalise nem elimine os outros campos. Fontes devem usar exatamente a entidade e a propriedade de cada campo. Antes de declarar ausencia, procure o dado em todas as paginas. Identificadores tecnicos controlam somente a resposta; nao sao instrucoes do funcionario.';
  request.input[0].content[0].text=JSON.stringify({templateSha256:doc.hash,task,documents,entities:reviewBase.entities,fields});
 }
 return {production:true,request,requestSha256:digest(request),templateBase64:templateBytes.toString('base64'),model,task,documents,config,scenario:'production',mode:'live',...(correction?{correction,reviewTargets,reviewBase}: {})};
}
async function authorizeProduction(prepared,config,approval,ledgerPath,developmentOnly){
 check(approval?.authorized===true&&typeof approval.evidence==='string'&&approval.evidence.trim().length>=10&&Date.parse(approval.expiresAt)>Date.now(),'PAID_API_NOT_AUTHORIZED');
 check(approval.workspaceId===prepared.task.workspaceId,'APPROVAL_WORKSPACE_MISMATCH');
 if(prepared.stagedNative){
  check(approval.workflowVersion===prepared.workflowVersion&&approval.maxCalls>=2,'NATIVE_TWO_STAGE_APPROVAL_REQUIRED');
  if(prepared.stage==='reading'){const {estimateReservation}=require('./openai.cjs');check(approval.maxBudgetUsd>=estimateReservation({...config,nativeDocument:false})+estimateReservation({...config,nativeDocument:true,codeInterpreterSessionUsd:config.codeInterpreterSessionUsd??0.03}),'API_BUDGET_LIMIT');}
 }
 const single=approval.issueId===prepared.task.issueId&&approval.requestSha256===(prepared.parentRequestSha256||prepared.requestSha256)&&approval.maxCalls===(prepared.stagedNative?2:1)&&approval.maxBudgetUsd>0&&(prepared.stagedNative||approval.maxBudgetUsd<=0.5);
 const ongoing=!developmentOnly&&approval.scope==='production'&&approval.dataScope==='institutional'&&Number.isInteger(approval.maxCalls)&&approval.maxCalls>0&&approval.maxBudgetUsd>0;
 check(single||ongoing,'APPROVAL_REQUEST_MISMATCH');
 if(developmentOnly){check(single&&approval.dataScope==='fictional-only'&&['generated-synthetic','approved-fictional-derivative'].includes(prepared.model.origin),'REAL_DATA_EXTERNAL_SHARING_DISABLED');}
 else check(approval.dataScope==='institutional','REAL_DATA_EXTERNAL_SHARING_DISABLED');
 config=JSON.parse(JSON.stringify(config));config.live={enabled:true,budgetApprovalId:approval.approvalId,maxCalls:approval.maxCalls,maxBudgetUsd:approval.maxBudgetUsd,...(approval.budgetPolicy==='user-unlimited'?{budgetPolicy:'user-unlimited'}:{})};
 prepared={...prepared,config,ledgerPath};prepared.reservation=await require('./openai.cjs').authorizeRequest(prepared,config,require('./budget.cjs').createLedger(ledgerPath));return prepared;
}
module.exports={schema,reviewSchema,prepareProduction,authorizeProduction,locateReplacements,indexEvidencePages};
