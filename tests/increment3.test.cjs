'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
const {semanticFixture}=require('./semantic-fixture.cjs');
const {makeGroundTruth,compareGroundTruth}=require('../src/ground-truth.cjs');
const {sha256}=require('../src/docx.cjs');
const {prepare,evaluate,assertTestInputs}=require('../src/increment3.cjs');
const {buildRequest,transmitRequest,requestLive}=require('../src/openai.cjs');
const {createLedger}=require('../src/budget.cjs');
const defaults=require('../config/openai.json');
const blocked=(fn,code)=>assert.rejects(fn,e=>e.code===code);
async function fixture(){
  const f=await semanticFixture();f.profile.origin='generated-synthetic';
  const truth=makeGroundTruth(f.plan,'completo'),instructions='Instrucoes ficticias locais';
  const payload={mode:'mock',scenario:'completo',templateBase64:f.bytes.toString('base64'),profile:f.profile,pdfs:[{id:'pacote-ficticio',base64:f.pdf.toString('base64')}],approvals:f.approvals,instructions};
  for(const d of f.documents)d.id='pacote-ficticio';
  for(const d of f.extraction.documents)d.id='pacote-ficticio';
  for(const d of f.extraction.documentSegments)d.documentId='pacote-ficticio';
  for(const r of [...f.extraction.observations,...f.extraction.replacements])for(const s of r.sources)s.documentId='pacote-ficticio';
  payload.mockResponse={id:'ficticio',status:'completed',model:'gpt-6-luna',service_tier:'priority',usage:{input_tokens:200,output_tokens:100,output_tokens_details:{reasoning_tokens:30}},output:[{type:'message',content:[{type:'output_text',text:JSON.stringify(f.extraction)}]}]};
  const manifest={status:'FICTIONAL_INPUTS_VERIFIED',officialTemplateIncluded:false,assets:[{kind:'template',sha256:sha256(f.bytes)},{kind:'pdf',scenario:'completo',sha256:sha256(f.pdf)}],profileSha256:sha256(Buffer.from(JSON.stringify(f.profile))),instructionsSha256:sha256(Buffer.from(instructions)),truthSha256:{completo:sha256(Buffer.from(JSON.stringify(truth)))}};
  return {f,truth,payload,manifest,config:{...structuredClone(defaults),service_tier:'fast',reasoning:{effort:'high'},max_output_tokens:16000,pricing:{...defaults.pricing,inputPerMillionUsd:0.4,outputPerMillionUsd:1.5}}};
}
test('oraculo compara fatos, papeis e fontes sem consultar uma IA',async()=>{const {f,truth}=await fixture();const r=compareGroundTruth(f.extraction,truth,f.documents);assert.equal(r.passed,true);assert.equal(r.facts,5);});
test('oraculo aceita identificadores novos diferentes sem perder os vinculos',async()=>{
  const {f,truth}=await fixture();for(const e of f.extraction.entities)e.id='alternate-'+e.id;
  for(const r of [...f.extraction.observations,...f.extraction.replacements]){r.entityId='alternate-'+r.entityId;for(const s of r.sources)s.entityId=r.entityId;}
  assert.equal(compareGroundTruth(f.extraction,truth,f.documents).passed,true);
});
test('oraculo rejeita nome inventado ainda que observacao e substituicao concordem',async()=>{
  const {f,truth}=await fixture();f.extraction.observations[0].value=f.extraction.replacements[0].value='Nome Inventado Indevido';
  assert.throws(()=>compareGroundTruth(f.extraction,truth,f.documents),e=>e.code==='GROUND_TRUTH_VALUE_MISMATCH');
});
test('oraculo rejeita fonte na pagina errada mesmo dentro do intervalo permitido',async()=>{
  const {f,truth}=await fixture();f.extraction.observations[0].sources[0].page=2;
  assert.throws(()=>compareGroundTruth(f.extraction,truth,f.documents),e=>e.code==='GROUND_TRUTH_EVIDENCE_MISMATCH');
});
test('oraculo rejeita troca de papeis e campo atribuido a outra pessoa',async()=>{
  for(const edit of [v=>v.entities[0].role='buyer',v=>v.replacements[0].entityId='new-buyer']){
    const {f,truth}=await fixture();edit(f.extraction);assert.throws(()=>compareGroundTruth(f.extraction,truth,f.documents),e=>/GROUND_TRUTH_(ENTITIES|FIELD_OWNER)_MISMATCH/.test(e.code));
  }
});
test('oraculo nao tolera fatos omitidos, extras ou repetidos',async()=>{
  for(const edit of [v=>v.observations.pop(),v=>v.observations.push({...v.observations[0],property:'dado-inventado'}),v=>v.observations.push(v.observations[0])]){
    const {f,truth}=await fixture();edit(f.extraction);assert.throws(()=>compareGroundTruth(f.extraction,truth,f.documents),e=>e.code==='GROUND_TRUTH_FACT_COVERAGE');
  }
});
test('entrada completa ficticia gera DOCX somente apos comparacao independente',async()=>{const v=await fixture();const p=await prepare(v.payload,v);const r=await evaluate(p,v.payload.mockResponse);assert.equal(r.report.evaluation.passed,true);assert.ok(r.bytes);assert.equal(r.report.apiCalls,0);});
test('gabarito, perfil, instrucoes e PDF alterados bloqueiam o preflight',async()=>{
  for(const [edit,code] of [[v=>v.truth.facts[0].value='alterado','GROUND_TRUTH_PROVENANCE_MISMATCH'],[v=>v.payload.profile.fields[0].oldText='alterado','PREFLIGHT_TEMPLATE_MISMATCH'],[v=>v.payload.instructions+=' indevido','PREFLIGHT_INSTRUCTIONS_MISMATCH'],[v=>v.payload.pdfs[0].base64=Buffer.from('%PDF-OUTRO').toString('base64'),'PREFLIGHT_INPUT_MISMATCH']]){
    const v=await fixture();edit(v);assert.throws(()=>assertTestInputs(v.payload,v.manifest,v.truth),e=>e.code===code);
  }
});
test('modelo oficial ou candidato anonimizado nao entra neste teste',async()=>{
  const v=await fixture();v.payload.profile.origin='anonymized-reviewed';assert.throws(()=>assertTestInputs(v.payload,v.manifest,v.truth),e=>e.code==='OFFICIAL_TEMPLATE_NOT_ALLOWED_IN_TEST');
});
test('modo live sem autorizacao nao reserva nem envia',async()=>{const v=await fixture();v.payload.mode='live';await blocked(()=>prepare(v.payload,v),'PAID_API_NOT_AUTHORIZED');});
test('autorizacao vencida, de outro request ou sem comprovante nao libera o envio',async()=>{
  for(const variant of ['expired','request','evidence']){
    const v=await fixture();v.payload.mode='live';v.approval={authorized:true,approvalId:'ficticia',evidence:'Autorizacao ficticia de teste',scenario:'completo',expiresAt:new Date(Date.now()+60000).toISOString(),requestSha256:'0'.repeat(64),maxCalls:1,maxBudgetUsd:0.5};
    if(variant==='expired')v.approval.expiresAt='2020-01-01T00:00:00Z';
    if(variant==='evidence')v.approval.evidence=null;
    await blocked(()=>prepare(v.payload,v),variant==='request'?'APPROVAL_REQUEST_MISMATCH':'PAID_API_NOT_AUTHORIZED');
  }
});
test('gabarito e resposta simulada nunca entram no request enviado ao modelo',async()=>{
  const v=await fixture(),p=await prepare(v.payload,v),request=JSON.stringify(p.request);
  for(const r of v.truth.facts)assert.equal(request.includes(r.value),false);
  assert.equal(request.includes('mockResponse'),false);assert.equal(request.includes('expectedPending'),false);
  assert.deepEqual(p.request.prompt_cache_options,{mode:'explicit'});assert.equal(request.includes('prompt_cache_breakpoint'),false);
});
test('configuracao rejeita modelo, raciocinio e tier fora dos valores autorizados',async()=>{
  for(const edit of [c=>c.model='outro',c=>c.reasoning.effort='low',c=>c.service_tier='auto']){
    const v=await fixture();edit(v.config);await blocked(()=>prepare(v.payload,v),'OPENAI_CONFIGURATION_MISMATCH');
  }
});
test('ledger bloqueia mesmo request sob outra autorizacao, inclusive apos reinicio',async()=>{
  const base=path.join(__dirname,'..','storage','temp');await fs.mkdir(base,{recursive:true});const dir=await fs.mkdtemp(path.join(base,'dedup3-')),file=path.join(dir,'ledger.json');
  const reserve={approvalId:'ficticia-a',usd:0.444,maxCalls:2,budgetUsd:1,requestSha256:'1'.repeat(64)};
  await createLedger(file).reserve(reserve);
  await blocked(()=>createLedger(file).reserve({...reserve,approvalId:'ficticia-b'}),'DUPLICATE_PAID_REQUEST');
  await createLedger(file).reserve({...reserve,requestSha256:'2'.repeat(64)});
  const state=JSON.parse(await fs.readFile(file,'utf8'));assert.equal(state.approvals['ficticia-a'].calls,2);assert.equal(Object.keys(state.requests).length,2);
});
test('transporte recusa request sem reserva ou alterado depois da reserva',async()=>{
  const v=await fixture(),p=await prepare(v.payload,v);let calls=0;
  await blocked(()=>transmitRequest(p.request,v.config,{apiKey:'FICTICIA',transport:async()=>{calls++;}}),'PAID_API_NOT_AUTHORIZED');assert.equal(calls,0);
});
test('resposta real negativa registra tokens e tier antes de bloquear dados incompletos',async()=>{
  const v=await fixture();const args={config:{...v.config,live:{enabled:true,budgetApprovalId:'ficticia',maxCalls:1,maxBudgetUsd:0.5}},templateBytes:v.f.bytes,profile:v.f.profile,pdfs:[{id:'pacote-ficticio',bytes:v.f.pdf}],approvals:v.f.approvals,instructions:v.payload.instructions};
  let recorded,saved;const raw=structuredClone(v.payload.mockResponse);raw.status='incomplete';raw.output=[];
  const r=await requestLive(args,{apiKey:'FICTICIA',ledger:{reserve:async()=>{},recordUsage:async u=>{recorded=u;}},saveResponse:async s=>{saved=s;},transport:async()=>new Response(JSON.stringify(raw),{status:200})});
  assert.equal(recorded.actualTier,'priority');assert.equal(recorded.reasoningTokens,30);assert.equal(saved.httpStatus,200);assert.equal(r.raw.status,'incomplete');
});
test('workflow n8n permanece manual, sem chave, sem retry e com HTTP fixo protegido pelo desvio',async()=>{
  const w=JSON.parse(await fs.readFile(path.join(__dirname,'..','workflows','incremento3-manual.template.json'),'utf8'));
  assert.equal(w.active,false);assert.equal(w.nodes.filter(n=>/Trigger/.test(n.type)).length,1);
  const http=w.nodes.find(n=>n.type==='n8n-nodes-base.httpRequest');assert.equal(http.parameters.url,'https://api.openai.com/v1/responses');assert.equal(http.parameters.nodeCredentialType,'openAiApi');assert.equal(http.retryOnFail,false);assert.equal(http.parameters.options.redirect.redirect.followRedirects,false);assert.equal(http.credentials,undefined);
  assert.equal(w.connections['Preparar e reservar chamada autorizada'].main[0][0].node,'Chamada real autorizada');assert.equal(w.connections['Chamada real autorizada'].main[0][0].node,http.name);assert.equal(w.settings.saveManualExecutions,false);
});
async function fullNegative(scenario){
  const {createDocx}=require('../scripts/make-fixtures.cjs');
  const known=await createDocx({write:false});
  const documents=[{id:'pacote-ficticio',sha256:'a'.repeat(64),pages:10}];
  const replacements=known.plan.replacements.map(r=>{const f=known.profile.fields.find(f=>f.id===r.fieldId);return {fieldId:r.fieldId,entityId:r.entityId,property:r.property,status:'verified',value:r.value,sources:[{...r.source,documentSha256:documents[0].sha256,quote:r.value}]};});
  const pending=[];
  for(const fieldId of scenario==='incompleto'?['imovel.matricula']:['imovel.matricula','ato.preco']){
    const r=replacements.find(r=>r.fieldId===fieldId);r.value=null;r.status=scenario==='incompleto'?'missing':'conflict';
    r.sources=scenario==='incompleto'?[]:[...r.sources,{...r.sources[0],page:10,quote:fieldId==='imovel.matricula'?'MAT-FICTICIA-DIVERGENTE':'R$ 999.000,00'}];
    pending.push({fieldId,entityId:r.entityId,code:scenario==='incompleto'?'MISSING_REQUIRED':'CONTRADICTION'});
  }
  return {truth:makeGroundTruth(known.plan,scenario),documents,value:{templateSha256:known.plan.templateSha256,entities:known.plan.entities,documents,replacements,observations:replacements.map(r=>({...r})),pending}};
}
test('documentacao incompleta exige ausencia explicita e impede herdar matricula antiga',async()=>{
  const {truth,documents,value}=await fullNegative('incompleto');const r=compareGroundTruth(value,truth,documents);assert.equal(r.documentExpected,false);assert.deepEqual(r.expectedPending,['imovel.matricula']);
  value.observations.find(o=>o.property==='matricula').value='MAT-TESTE-ANTIGA';assert.throws(()=>compareGroundTruth(value,truth,documents),e=>e.code==='GROUND_TRUTH_VALUE_MISMATCH');
});
test('documentacao contraditoria exige as duas fontes e nao admite escolher uma',async()=>{
  const v=await fullNegative('contraditorio');assert.equal(compareGroundTruth(v.value,v.truth,v.documents).expectedPending.length,2);
  const o=v.value.observations.find(o=>o.property==='matricula');o.sources=o.sources.slice(0,1);assert.throws(()=>compareGroundTruth(v.value,v.truth,v.documents),e=>e.code==='GROUND_TRUTH_EVIDENCE_MISMATCH');
});
test('oraculo exige pendencias especificas e detecta vinculo conjugal trocado',async()=>{
  let v=await fullNegative('contraditorio');v.value.pending.pop();assert.throws(()=>compareGroundTruth(v.value,v.truth,v.documents),e=>e.code==='GROUND_TRUTH_PENDING_MISSING');
  v=await fullNegative('incompleto');v.value.entities[0].spouseId='new-buyer-spouse';assert.throws(()=>compareGroundTruth(v.value,v.truth,v.documents),e=>e.code==='GROUND_TRUTH_ENTITIES_MISMATCH');
});
test('timeout com ledger persistente nao autoriza repetir automaticamente a chamada',async()=>{
  const v=await fixture();const args={config:{...v.config,live:{enabled:true,budgetApprovalId:'ficticia',maxCalls:2,maxBudgetUsd:1}},templateBytes:v.f.bytes,profile:v.f.profile,pdfs:[{id:'pacote-ficticio',bytes:v.f.pdf}],approvals:v.f.approvals,instructions:v.payload.instructions};
  const base=path.join(__dirname,'..','storage','temp');await fs.mkdir(base,{recursive:true});const dir=await fs.mkdtemp(path.join(base,'timeout3-')),filename=path.join(dir,'ledger.json');let calls=0;
  const options={apiKey:'FICTICIA',ledger:createLedger(filename),transport:async()=>{calls++;throw Error('timeout-ficticio');}};
  await blocked(()=>requestLive(args,options),'OPENAI_TRANSPORT_FAILED');
  await blocked(()=>requestLive(args,{...options,ledger:createLedger(filename)}),'DUPLICATE_PAID_REQUEST');assert.equal(calls,1);
});
test('autorizacao exata libera uma reserva e contabiliza resposta ficticia sem rede',async()=>{
  const v=await fixture(),mock=await prepare(v.payload,v);v.payload.mode='live';
  const base=path.join(__dirname,'..','storage','temp');await fs.mkdir(base,{recursive:true});const dir=await fs.mkdtemp(path.join(base,'authorized3-'));v.ledgerPath=path.join(dir,'ledger.json');
  v.approval={authorized:true,approvalId:'autorizacao-FICTICIA',evidence:'Autorizacao ficticia apenas para teste unitario',scenario:'completo',expiresAt:new Date(Date.now()+60000).toISOString(),requestSha256:mock.requestSha256,maxCalls:1,maxBudgetUsd:0.5};
  const live=await prepare(v.payload,v);assert.equal(live.reservation.reservedUsd,0.444);
  const r=await evaluate(live,v.payload.mockResponse);assert.equal(r.report.simulation,false);assert.equal(r.report.usage.actualTier,'priority');assert.equal(r.report.evaluation.passed,true);
  const state=JSON.parse(await fs.readFile(v.ledgerPath,'utf8'));assert.equal(state.requests[mock.requestSha256].status,'response-accounted');
  await blocked(()=>prepare(v.payload,v),'DUPLICATE_PAID_REQUEST');
});
test('tier divergente e escrita de cache inesperada sao contabilizados e bloqueados',async()=>{
  const {accountResponse}=require('../src/openai.cjs');
  for(const [edit,code] of [[r=>r.service_tier='default','OPENAI_SERVICE_TIER_MISMATCH'],[r=>r.usage.input_tokens_details={cache_write_tokens:100},'UNEXPECTED_CACHE_WRITE']]){
    const v=await fixture();const raw=structuredClone(v.payload.mockResponse);edit(raw);let recorded;
    await blocked(()=>accountResponse(raw,v.config,{recordUsage:async r=>{recorded=r;}},{reservedUsd:0.444}),code);assert.ok(recorded);assert.equal(recorded.actualTier,raw.service_tier);
  }
});
