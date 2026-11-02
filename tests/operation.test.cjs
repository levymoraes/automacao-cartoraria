'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const {semanticFixture}=require('./semantic-fixture.cjs');
const {applyAuthorized,openDocx,paragraphsOf}=require('../src/docx.cjs');
const blocked=(fn,code)=>assert.rejects(fn,e=>e.code===code);
test('exige novos dados em todas as ocorrências e preserva funcionário e cláusula',async()=>{
  const f=await semanticFixture({repeat:true}),out=await applyAuthorized(f.bytes,f.profile,f.plan),doc=await openDocx(out.bytes);
  const text=paragraphsOf(doc.parts.get('word/document.xml').toString()).map(p=>p.text).join('\n');
  for(const r of f.plan.replacements)assert.ok(text.includes(r.value));
  for(const field of f.profile.fields)assert.ok(!text.includes(field.oldText));
  assert.ok(text.includes('Funcionário Fixo Teste. Cláusula jurídica fixa.'));assert.equal(out.audit.operation.occurrences,6);
});
test('bloqueia perfil legado e inventário sem revisão',async()=>{
  for(const edit of [f=>delete f.profile.schemaVersion,f=>f.profile.operation.inventoryReviewed=false]){const f=await semanticFixture();edit(f);await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'VARIABLE_INVENTORY_REVIEW_REQUIRED');}
});
test('bloqueia inventário desatualizado',async()=>{
  const f=await semanticFixture();f.profile.operation.coverage[0].textSha256='0'.repeat(64);await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'VARIABLE_INVENTORY_STALE');
});
test('nem campo marcado opcional pode conservar dado antigo por omissão',async()=>{
  const f=await semanticFixture();f.profile.fields[0].required=false;f.plan.replacements.shift();await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'INCOMPLETE_VARIABLE_COVERAGE');
});
test('bloqueia profissão sem fonte da nova solicitação',async()=>{
  const f=await semanticFixture();delete f.plan.replacements.find(r=>r.property==='profissao').source;await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'SOURCE_REQUIRED');
});
test('bloqueia troca de documentos entre vendedor e comprador',async()=>{
  const f=await semanticFixture(),r=f.plan.replacements.find(r=>r.fieldId==='seller.cpf');r.entityId=r.source.entityId='new-buyer';await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'FIELD_ENTITY_MISMATCH');
});
test('bloqueia troca de papel das partes',async()=>{
  const f=await semanticFixture();f.plan.entities[0].role='buyer';await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'PARTICIPANT_COMPOSITION_MISMATCH');
});
test('bloqueia o mesmo CPF atribuído a pessoas distintas',async()=>{
  const f=await semanticFixture();f.plan.replacements.find(r=>r.fieldId==='buyer.cpf').value=f.plan.replacements.find(r=>r.fieldId==='seller.cpf').value;await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'PARTICIPANT_IDENTIFIER_CONFLICT');
});
test('bloqueia composição de seis participantes substituída por quatro',async()=>{
  const f=await semanticFixture();for(let i=0;i<4;i++)f.profile.operation.entities.push({id:'extra'+i,kind:'person',role:'buyer_spouse',spouseId:null});
  f.plan.entities.push(...f.profile.operation.entities.filter(e=>e.id.startsWith('extra')).slice(0,2).map(e=>({...e,id:'new-'+e.id,slotId:e.id})));
  await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'PARTICIPANT_COMPOSITION_MISMATCH');
});
test('bloqueia cônjuge ligado à pessoa incorreta',async()=>{
  const f=await semanticFixture();f.profile.operation.entities[0].spouseId='buyer';f.profile.operation.entities[1].spouseId='seller';f.plan.entities[0].spouseId='new-seller';f.plan.entities[1].spouseId='new-seller';
  await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'SPOUSE_CORRESPONDENCE_MISMATCH');
});
test('funcionário da serventia não pode receber dados de parte',async()=>{
  const f=await semanticFixture();f.profile.fields[0].entityId='office';await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'INSTITUTIONAL_FIELD_PROTECTED');
});
test('intervalo institucional protegido bloqueia campo mesmo com entidade falsamente atribuída',async()=>{
  const f=await semanticFixture();f.profile.operation.protectedFields=[{...f.profile.fields[0]}];await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'INSTITUTIONAL_FIELD_PROTECTED');
});
test('bloqueia nomes diferentes para a mesma pessoa em ocorrências repetidas',async()=>{
  const f=await semanticFixture({repeat:true});f.plan.replacements.at(-1).value='Outro Nome Inventado';await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'REPEATED_VALUE_CONTRADICTION');
});
test('detecta ocorrência antiga esquecida fora dos campos',async()=>{
  const f=await semanticFixture({unmapped:true});await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'UNMAPPED_OLD_VALUE_OCCURRENCE');
});
test('detecta dado antigo residual em atributos de partes auxiliares',async()=>{
  const f=await semanticFixture({auxiliary:true});await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'OLD_VALUE_IN_AUXILIARY_PART');
});
test('valor igual ao antigo exige confirmação explícita e fonte nova',async()=>{
  const f=await semanticFixture(),r=f.plan.replacements[0];r.value=f.profile.fields[0].oldText;
  await blocked(()=>applyAuthorized(f.bytes,f.profile,f.plan),'OLD_VALUE_REQUIRES_NEW_EVIDENCE');
  r.confirmedUnchanged=true;const out=await applyAuthorized(f.bytes,f.profile,f.plan);assert.ok(out.bytes.length);
});
