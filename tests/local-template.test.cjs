'use strict';
const {test}=require('node:test'),assert=require('node:assert/strict');
const {semanticFixture}=require('./semantic-fixture.cjs');
const {applyLocal,prepareLocal}=require('../src/local-template.cjs');
const {openDocx,sha256,paragraphsOf}=require('../src/docx.cjs');
const {buildRequest}=require('../src/openai.cjs');
const {assertTestInputs}=require('../src/increment3.cjs');
async function fixture(options){const f=await semanticFixture(options),doc=await openDocx(f.bytes);Object.assign(f.profile,{scope:'local-authorized-template',origin:'user-supplied-local',anonymizationApproved:false,externalSharingAllowed:false,localUse:{approved:true,originalSha256:'a'.repeat(64),convertedSha256:doc.hash,preservedParts:[...doc.parts].map(([part,b])=>({part,sha256:sha256(b)}))}});return f;}
test('modelo local preenche os campos, conserva funcionario e nao libera envio a OpenAI',async()=>{
 const f=await fixture({repeat:true}),result=await applyLocal(f.bytes,f.profile,f.plan),after=await openDocx(result.bytes);
 const text=paragraphsOf(after.parts.get('word/document.xml').toString()).map(p=>p.text).join('\n');assert.ok(text.includes('Novo Vendedor Inventado'));assert.ok(!text.includes('Antigo Vendedor Teste'));assert.ok(text.includes('Funcionário Fixo Teste'));assert.equal(result.audit.externalApiCalls,0);
 await assert.rejects(()=>buildRequest({templateBytes:f.bytes,profile:f.profile,pdfs:[],approvals:[],instructions:'FICTICIO',config:require('../config/openai.json')}),e=>e.code==='EXTERNAL_TEMPLATE_NOT_APPROVED');
 assert.throws(()=>assertTestInputs({profile:f.profile},{status:'FICTIONAL_INPUTS_VERIFIED',officialTemplateIncluded:true},{}),e=>e.code==='FICTIONAL_PREFLIGHT_REQUIRED');
});
test('caminho local recusa live antes de preparar documentos ou ler chave',()=>{
 for(const config of [{openaiMode:'live',paidApiEnabled:false},{openaiMode:'mock',paidApiEnabled:true}])assert.throws(()=>prepareLocal({},null,config),e=>e.code==='REGISTERED_TEMPLATE_LIVE_DISABLED');
});
test('inventario local exige aprovacao, hash exato e preservacao de todas as partes',async()=>{
 for(const change of [p=>p.localUse.approved=false,p=>p.localUse.convertedSha256='b'.repeat(64),p=>p.localUse.preservedParts.pop(),p=>p.operation.inventoryReviewed=false,p=>p.externalSharingAllowed=true]){
  const f=await fixture();change(f.profile);await assert.rejects(()=>applyLocal(f.bytes,f.profile,f.plan));
 }
});
test('campos esquecidos, trocas de proprietario e falta de fontes bloqueiam o resultado local',async()=>{
 for(const [mutate,code] of [[p=>p.replacements.pop(),'INCOMPLETE_VARIABLE_COVERAGE'],[p=>p.replacements[0].entityId='new-buyer','FIELD_ENTITY_MISMATCH'],[p=>delete p.replacements[0].source,'SOURCE_REQUIRED'],[p=>p.replacements[0].value='\ud800','INVALID_REPLACEMENT']]){const f=await fixture();mutate(f.plan);await assert.rejects(()=>applyLocal(f.bytes,f.profile,f.plan),e=>e.code===code);}
});
test('dado antigo em atributo auxiliar continua bloqueado no caminho local',async()=>{
 const f=await fixture({auxiliary:true});await assert.rejects(()=>applyLocal(f.bytes,f.profile,f.plan),e=>e.code==='OLD_VALUE_IN_AUXILIARY_PART');
});
