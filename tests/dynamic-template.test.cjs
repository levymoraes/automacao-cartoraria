'use strict';
const {test}=require('node:test'),assert=require('node:assert/strict');
const {semanticFixture}=require('./semantic-fixture.cjs');const {compileInventory,digest,criticalCandidates}=require('../src/dynamic-template.cjs');
async function setup(){const f=await semanticFixture(),task={description:'Vendedor e comprador conforme documentos ficticios anexados.'},model={id:'MODELO-NOVO',sha256:'a'.repeat(64),userApprovedUse:true},analysis={schemaVersion:1,templateSha256:f.profile.templateSha256,taskSha256:digest(task),documentsSha256:digest(f.documents),entities:f.profile.operation.entities,coverage:f.profile.operation.coverage.map(c=>({...c,reviewed:true})),audit:{status:'passed',unmappedVariables:[]},pending:[],fields:f.profile.fields.map((field,i)=>({...field,id:'ai-discovered-'+i,category:field.property==='nome'?'person-name':field.property==='cpf'?'identity-document':'profession'}))};return {...f,task,model,analysis,args:{templateBytes:f.bytes,task,model,documents:f.documents,analysis}};}
test('aceita campos descobertos por solicitacao sem consultar mapa de modelo cadastrado',async()=>{
 const f=await setup(),profile=await compileInventory(f.args);assert.equal(profile.modelId,'MODELO-NOVO');assert.equal(profile.fields.length,5);assert.equal(profile.operation.reviewType,'request-specific-ai-analysis-with-local-validation');assert.ok(profile.fields.every(f=>f.id.startsWith('ai-discovered-')));assert.equal(profile.externalSharingAllowed,false);
});
test('analise desatualizada, cobertura incompleta e pendencias nunca geram perfil',async()=>{
 for(const mutate of [a=>a.taskSha256='b'.repeat(64),a=>a.templateSha256='b'.repeat(64),a=>a.coverage.pop(),a=>a.coverage[0].reviewed=false,a=>a.audit.status='review',a=>a.audit.unmappedVariables.push('fato-nao-mapeado')]){const f=await setup();mutate(f.analysis);await assert.rejects(()=>compileInventory(f.args));}
});
test('proposta da IA nao pode converter escrevente em comprador nem ampliar trecho autorizado',async()=>{
 const f=await setup();const {openDocx,paragraphsOf}=require('../src/docx.cjs');const doc=await openDocx(f.bytes),text=paragraphsOf(doc.parts.get('word/document.xml').toString())[2].text,oldText='Funcionário Fixo Teste',start=text.indexOf(oldText);f.analysis.fields.push({id:'staff-as-buyer',entityId:'buyer',property:'nome',category:'person-name',part:'word/document.xml',paragraph:2,start,end:start+oldText.length,oldText});await assert.rejects(()=>compileInventory(f.args),e=>e.code==='INSTITUTIONAL_FIELD_PROTECTED');
 const g=await setup();g.analysis.fields[0].end++;await assert.rejects(()=>compileInventory(g.args),e=>e.code==='DYNAMIC_FIELD_ANCHOR_INVALID');
});
test('scanner identifica dados criticos sem confundir a virgula apos valor monetario',()=>{
 const text='CPF 111.111.111-11; data 01/01/2026; preco R$ 500,00, quitado.';const found=criticalCandidates([{part:'word/document.xml',paragraph:0,text}]);assert.ok(found.some(c=>c.type==='cpf'));assert.ok(found.some(c=>c.type==='date'));const money=found.find(c=>c.type==='value');assert.equal(text.slice(money.start,money.end),'500,00');
});
test('novo identificador encontrado no modelo nao pode ser omitido pela analise',async()=>{
 const f=await setup();const {openDocx,sha256,paragraphsOf}=require('../src/docx.cjs'),{coverageOf}=require('../src/operation.cjs');
 const doc=await openDocx(f.bytes),xml=doc.parts.get('word/document.xml').toString().replace('</w:body>','<w:p><w:r><w:t>Outro CPF antigo: 111.111.111-11.</w:t></w:r></w:p></w:body>');doc.zip.file('word/document.xml',xml);
 const bytes=await doc.zip.generateAsync({type:'nodebuffer'}),updated=await openDocx(bytes);f.args.templateBytes=bytes;f.analysis.templateSha256=sha256(bytes);f.analysis.coverage=coverageOf(updated,paragraphsOf).map(c=>({...c,reviewed:true}));
 await assert.rejects(()=>compileInventory(f.args),e=>e.code==='UNREGISTERED_CRITICAL_VARIABLE');
});
