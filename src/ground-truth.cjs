'use strict';
const {requireCondition:check}=require('./errors.cjs');
const {normalize}=require('./operation.cjs');
// Este oraculo e local e deterministico. Nunca integra as mensagens da API.
function makeGroundTruth(plan,scenario){
  check(['completo','digitalizada','incompleto','contraditorio'].includes(scenario),'INVALID_SCENARIO');
  const entities=plan.entities.map(e=>({...e}));
  const facts=new Map();
  for(const r of plan.replacements){
    const slotId=entities.find(e=>e.id===r.entityId).slotId;
    const key=slotId+':'+r.property;
    const status=scenario==='incompleto'&&r.fieldId==='imovel.matricula'?'missing':scenario==='contraditorio'&&['imovel.matricula','ato.preco'].includes(r.fieldId)?'conflict':'verified';
    const evidence=status==='missing'?[]:[{page:r.source.page,value:r.value},...(status==='conflict'?[{page:10,value:r.fieldId==='imovel.matricula'?'MAT-FICTICIA-DIVERGENTE':'R$ 999.000,00'}]:[])];
    if(!facts.has(key))facts.set(key,{slotId,property:r.property,status,value:status==='verified'?r.value:null,evidence,fieldIds:[]});
    facts.get(key).fieldIds.push(r.fieldId);
  }
  return {scenario,templateSha256:plan.templateSha256,entities,facts:[...facts.values()]};
}
function compareGroundTruth(value,truth,documents){
  check(value.templateSha256===truth.templateSha256,'GROUND_TRUTH_TEMPLATE_MISMATCH');
  check(value.documents.length===documents.length&&documents.every(d=>value.documents.some(v=>v.id===d.id&&v.sha256===d.sha256&&v.pages===d.pages)),'GROUND_TRUTH_DOCUMENT_MISMATCH');
  check(value.entities.length===truth.entities.length&&new Set(value.entities.map(e=>e.id)).size===value.entities.length&&new Set(value.entities.map(e=>e.slotId)).size===value.entities.length,'GROUND_TRUTH_ENTITIES_MISMATCH');
  const actual=new Map(value.entities.map(e=>[e.slotId,e]));
  for(const expected of truth.entities){
    const entity=actual.get(expected.slotId);
    const spouse=expected.spouseId?truth.entities.find(e=>e.id===expected.spouseId):null;
    check(entity&&entity.kind===expected.kind&&entity.role===expected.role&&(entity.spouseId||null)===(spouse?actual.get(spouse.slotId)?.id:null),'GROUND_TRUTH_ENTITIES_MISMATCH');
  }
  const byId=new Map(value.entities.map(e=>[e.id,e.slotId]));
  const observations=new Map(value.observations.map(o=>[byId.get(o.entityId)+':'+o.property,o]));
  check(observations.size===value.observations.length&&observations.size===truth.facts.length,'GROUND_TRUTH_FACT_COVERAGE');
  check(value.replacements.length===truth.facts.reduce((n,f)=>n+f.fieldIds.length,0)&&new Set(value.replacements.map(r=>r.fieldId)).size===value.replacements.length,'GROUND_TRUTH_REPLACEMENT_COVERAGE');
  const expectedPending=new Set();
  for(const fact of truth.facts){
    const o=observations.get(fact.slotId+':'+fact.property);
    const compare=(item)=>{
      check(item&&item.status===fact.status&&(fact.value===null?item.value===null:typeof item.value==='string'&&normalize(item.value)===normalize(fact.value)),'GROUND_TRUTH_VALUE_MISMATCH');
      check(item.sources.length>=fact.evidence.length&&(fact.status!=='missing'||item.sources.length===0),'GROUND_TRUTH_EVIDENCE_MISMATCH');
      for(const s of item.sources){
        const doc=documents.find(d=>d.id===s.documentId);
        check(doc&&s.documentSha256===doc.sha256&&s.page<=doc.pages&&s.entityId===item.entityId&&s.property===fact.property&&fact.evidence.some(e=>e.page===s.page&&normalize(s.quote).includes(normalize(e.value))),'GROUND_TRUTH_EVIDENCE_MISMATCH');
      }
      for(const e of fact.evidence)check(item.sources.some(s=>s.page===e.page&&normalize(s.quote).includes(normalize(e.value))),'GROUND_TRUTH_EVIDENCE_MISMATCH');
    };
    compare(o);
    for(const fieldId of fact.fieldIds){
      const r=value.replacements.find(r=>r.fieldId===fieldId);
      check(r&&byId.get(r.entityId)===fact.slotId&&r.property===fact.property,'GROUND_TRUTH_FIELD_OWNER_MISMATCH');compare(r);
      if(fact.status!=='verified'){
        expectedPending.add(fieldId);
        check(value.pending.some(p=>p.fieldId===fieldId&&p.entityId===r.entityId&&p.code===(fact.status==='missing'?'MISSING_REQUIRED':'CONTRADICTION')),'GROUND_TRUTH_PENDING_MISSING');
      }
    }
  }
  check(value.pending.every(p=>expectedPending.has(p.fieldId))&&new Set(value.pending.map(p=>p.fieldId)).size===expectedPending.size,'GROUND_TRUTH_UNEXPECTED_PENDING');
  return {passed:true,method:'deterministic-fictional-ground-truth',facts:truth.facts.length,entities:truth.entities.length,variableOccurrences:value.replacements.length,expectedPending:[...expectedPending],documentExpected:expectedPending.size===0};
}
module.exports={makeGroundTruth,compareGroundTruth};
