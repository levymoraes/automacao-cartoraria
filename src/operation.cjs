'use strict';
const {createHash}=require('node:crypto');
const {requireCondition:check}=require('./errors.cjs');
const hash=text=>createHash('sha256').update(text).digest('hex');
const story=/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/;
const normalize=value=>value.normalize('NFKC').replace(/\s+/gu,' ').trim().toLocaleLowerCase('pt-BR');
// Rotulos descritivos da IA nao sao regras de negocio. Normaliza apenas a
// terminologia para as verificacoes de correspondencia e localizacao do motor.
function executionRole(entity){
 if(['property','marriage','transaction'].includes(entity.kind))return entity.kind;
 const role=normalize(entity.role).normalize('NFD').replace(/\p{M}/gu,'');
 if(/^(seller|buyer)(_[a-z_]+)?$/.test(role))return role;
 const seller=/\b(?:seller|vendedor(?:a|es|as)?)\b/.test(role),buyer=/\b(?:buyer|comprador(?:a|es|as)?)\b/.test(role);
 if(seller===buyer)return role;
 const side=seller?'seller':'buyer';
 if(/\b(?:spouse|conjuge|espos[oa])\b/.test(role))return side+'_spouse';
 if(seller&&/\badministrador(?:a|es|as)?\b/.test(role))return 'seller_administrator';
 if(seller&&/\bsoci[oa]s?\b/.test(role))return /\bmenor\b/.test(role)?'seller_minor_shareholder':'seller_shareholder';
 return side;
}
function matches(text,value){
  const escaped=value.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
  return [...text.matchAll(new RegExp(`(?<![\\p{L}\\p{N}])${escaped}(?![\\p{L}\\p{N}])`,'giu'))].map(m=>({start:m.index,end:m.index+m[0].length}));
}
function collectStories(doc,paragraphsOf){
  return [...doc.parts].filter(([name])=>story.test(name)).flatMap(([part,b])=>paragraphsOf(b.toString()).map(p=>({part,paragraph:p.index,text:p.text})));
}
function coverageOf(doc,paragraphsOf){
  // Isto registra a versao do texto revisado, nao aprova seu inventario.
  return collectStories(doc,paragraphsOf).map(p=>({part:p.part,paragraph:p.paragraph,textSha256:hash(p.text)}));
}
function validateOperation(doc,profile,plan,paragraphsOf){
  check(profile.schemaVersion===2&&profile.operation?.inventoryReviewed===true,'VARIABLE_INVENTORY_REVIEW_REQUIRED');
  const op=profile.operation;
  const stories=collectStories(doc,paragraphsOf);
  check(Array.isArray(op.coverage)&&op.coverage.length===stories.length&&new Set(op.coverage.map(p=>p.part+':'+p.paragraph)).size===stories.length&&stories.every(p=>op.coverage.some(c=>c.part===p.part&&c.paragraph===p.paragraph&&c.textSha256===hash(p.text))),'VARIABLE_INVENTORY_STALE');
  const identifier=id=>typeof id==='string'&&/^[A-Za-z0-9_.-]{1,120}$/.test(id);
  check(Array.isArray(op.entities)&&op.entities.length>0&&op.entities.every(e=>identifier(e.id)&&['person','organization','property','transaction','marriage','institution'].includes(e.kind)&&typeof e.role==='string'&&e.role.length>0)&&new Set(op.entities.map(e=>e.id)).size===op.entities.length,'INVALID_ENTITY_INVENTORY');
  check(Array.isArray(plan.entities)&&plan.entities.every(e=>identifier(e.id)&&identifier(e.slotId))&&new Set(plan.entities.map(e=>e.id)).size===plan.entities.length&&new Set(plan.entities.map(e=>e.slotId)).size===plan.entities.length,'INVALID_ENTITY_BINDINGS');
  const slots=new Map(op.entities.map(e=>[e.id,e]));
  for(const e of op.entities)if(e.spouseId)check(e.spouseId!==e.id&&slots.get(e.spouseId)?.spouseId===e.id,'SPOUSE_CORRESPONDENCE_MISMATCH');
  const mutable=op.entities.filter(e=>e.kind!=='institution'&&e.role!=='staff');
  check(plan.entities.length===mutable.length&&plan.entities.every(e=>slots.has(e.slotId)&&slots.get(e.slotId).kind!=='institution'&&slots.get(e.slotId).role!=='staff'),'PARTICIPANT_COMPOSITION_MISMATCH');
  const binding=new Map(plan.entities.map(e=>[e.slotId,e]));
  for(const slot of mutable){
    const entity=binding.get(slot.id);
    check(entity&&entity.kind===slot.kind&&entity.role===slot.role,'PARTICIPANT_COMPOSITION_MISMATCH');
    check((entity.spouseId||null)===(slot.spouseId?binding.get(slot.spouseId)?.id:null),'SPOUSE_CORRESPONDENCE_MISMATCH');
  }
  const byId=new Map(plan.replacements.map(r=>[r.fieldId,r]));
  const variableValues=new Map();
  const cpfOwners=new Map();
  for(const f of profile.fields){
    check(!(op.protectedFields||[]).some(p=>p.part===f.part&&p.paragraph===f.paragraph&&p.start<f.end&&p.end>f.start),'INSTITUTIONAL_FIELD_PROTECTED');
    const slot=slots.get(f.entityId);
    check(slot&&typeof f.property==='string'&&f.property.length>0,'FIELD_ENTITY_REQUIRED');
    check(slot.kind!=='institution'&&slot.role!=='staff','INSTITUTIONAL_FIELD_PROTECTED');
    const r=byId.get(f.id);
    check(r,'INCOMPLETE_VARIABLE_COVERAGE');
    check(r.entityId===binding.get(f.entityId)?.id&&r.property===f.property&&r.source.entityId===r.entityId&&r.source.property===f.property,'FIELD_ENTITY_MISMATCH');
    check(normalize(r.value)!==normalize(f.oldText)||r.confirmedUnchanged===true,'OLD_VALUE_REQUIRES_NEW_EVIDENCE');
    const key=f.entityId+':'+f.property;
    if(variableValues.has(key))check(variableValues.get(key)===normalize(r.value),'REPEATED_VALUE_CONTRADICTION');
    variableValues.set(key,normalize(r.value));
    const pendingMarker=(plan.pendingAnnotations||[]).some(p=>p.fieldId===f.id&&p.text===r.value&&r.source.kind==='pending'&&r.value.startsWith('pendente: '));
    if(slot.kind==='person'&&/(^|\.)cpf$/i.test(f.property)&&!pendingMarker){
      const cpf=normalize(r.value).replace(/[.\-\s]/g,'');
      check(!cpfOwners.has(cpf)||cpfOwners.get(cpf)===f.entityId,'PARTICIPANT_IDENTIFIER_CONFLICT');cpfOwners.set(cpf,f.entityId);
    }
  }
  // Toda ocorrencia antiga deve estar dentro de um campo variavel mapeado
  // ou de uma excecao fixa explicita, revisada e vinculada ao hash do modelo.
  for(const f of profile.fields)for(const p of stories)for(const match of matches(p.text,f.oldText)){
    const covered=profile.fields.some(c=>c.part===p.part&&c.paragraph===p.paragraph&&c.start<=match.start&&c.end>=match.end);
    const fixed=(op.fixedOldValueExceptions||[]).some(e=>e.part===p.part&&e.paragraph===p.paragraph&&e.start===match.start&&e.end===match.end&&e.value===p.text.slice(match.start,match.end)&&typeof e.reason==='string'&&e.reason.trim().length>=8);
    check(covered||fixed,'UNMAPPED_OLD_VALUE_OCCURRENCE');
  }
  return {entities:mutable.length,variables:variableValues.size,occurrences:profile.fields.length,coverageReviewed:true};
}
function validateOperationResult(before,after,profile,plan,paragraphsOf,parseXml){
  const replacements=new Map(plan.replacements.map(r=>[r.fieldId,r]));
  const stories=collectStories(after,paragraphsOf);
  for(const p of collectStories(before,paragraphsOf)){
    const fields=profile.fields.filter(f=>f.part===p.part&&f.paragraph===p.paragraph).sort((a,b)=>a.start-b.start);
    let expected=p.text;
    for(const f of [...fields].reverse())expected=expected.slice(0,f.start)+replacements.get(f.id).value+expected.slice(f.end);
    const result=stories.find(a=>a.part===p.part&&a.paragraph===p.paragraph);
    check(result?.text===expected,'VARIABLE_RESULT_MISMATCH');
  }
  // A comparacao integral de cada paragrafo verifica presenca por localizacao,
  // inclusive quando outra pessoa tem um valor igual ao antigo, documentado.
  // Metadados, comentarios e atributos nao podem conservar dados antigos.
  const auxiliary=[...after.parts].filter(([part])=>/\.(xml|rels)$/.test(part)).flatMap(([part,b])=>parseXml(b.toString()).nodes.map((n,nodeIndex)=>({part,nodeIndex,element:n.local,value:(story.test(part)&&n.uri==='http://schemas.openxmlformats.org/wordprocessingml/2006/main'&&n.local==='t'?'':n.text)+' '+Object.values(n.attrs).map(a=>a.value).join(' ')})));
  const joined=records=>{const parts=new Map();for(const r of records){if(!parts.has(r.part))parts.set(r.part,[]);parts.get(r.part).push(r.value);}return [...parts.values()].map(v=>v.join(' ')).join('\n');};
  const allAuxiliary=joined(auxiliary);
  for(const f of profile.fields){
    const escaped=f.oldText.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),pattern=new RegExp(`(?<![\\p{L}\\p{N}])${escaped}(?![\\p{L}\\p{N}])`,'iu');
    const exceptions=(profile.operation.fixedAuxiliaryValueExceptions||[]).filter(e=>e.fieldId===f.id);
    // Excecao pontual para coincidencia numerica em formato/metadado tecnico:
    // ex.: ato 12 e fonte de tamanho 12. O valor do no e o hash da parte
    // original devem coincidir. A busca conjunta mantem a deteccao entre nos.
    const values=exceptions.length?joined(auxiliary.map(r=>exceptions.some(e=>e.part===r.part&&e.nodeIndex===r.nodeIndex&&e.element===r.element&&e.valueSha256===hash(r.value)&&e.partSha256===hash(before.parts.get(r.part))&&typeof e.reason==='string'&&e.reason.length>=20)?{...r,value:''}:r)):allAuxiliary;
    check(!pattern.test(values),'OLD_VALUE_IN_AUXILIARY_PART');
  }
  for(const [part,bytes] of before.parts)if(part.startsWith('word/media/'))check(after.parts.get(part)?.equals(bytes),'ORIGINAL_LOGO_CHANGED');
}
module.exports={coverageOf,collectStories,normalize,validateOperation,validateOperationResult,executionRole};
