'use strict';
const schema=require('../config/extraction.schema.json');
const {requireCondition:check}=require('./errors.cjs');
const validate=require('./generated/extraction-validator.cjs');
function validateExtraction(value,profile,documents,{allowPending=false}={}){
  check(validate(value),'EXTRACTION_SCHEMA_INVALID');
  check(value.templateSha256===profile.templateSha256,'EXTRACTION_TEMPLATE_MISMATCH');
  check(value.documents.length===documents.length&&new Set(value.documents.map(d=>d.id)).size===documents.length&&documents.every(d=>value.documents.some(v=>v.id===d.id&&v.sha256===d.sha256&&v.pages===d.pages)),'EXTRACTION_DOCUMENT_MISMATCH');
  for(const segment of value.documentSegments){
    const doc=documents.find(d=>d.id===segment.documentId);
    check(doc&&segment.startPage<=segment.endPage&&segment.endPage<=doc.pages&&segment.entityIds.every(id=>value.entities.some(e=>e.id===id)),'DOCUMENT_SEGMENT_INVALID');
  }
  for(const doc of documents)for(let page=1;page<=doc.pages;page++)check(value.documentSegments.some(s=>s.documentId===doc.id&&s.startPage<=page&&s.endPage>=page),'DOCUMENT_SEGMENT_MISSING');
  const sourceValid=(s,r)=>{
    const d=documents.find(d=>d.id===s.documentId);
    check(d&&d.sha256===s.documentSha256&&s.page<=d.pages,'EVIDENCE_DOCUMENT_MISMATCH');
    check(s.entityId===r.entityId&&s.property===r.property&&s.quote.trim().length>0,'EVIDENCE_ENTITY_MISMATCH');
    check(value.documentSegments.some(segment=>segment.documentId===s.documentId&&segment.startPage<=s.page&&segment.endPage>=s.page&&segment.entityIds.includes(s.entityId)),'EVIDENCE_SEGMENT_MISMATCH');
  };
  const observations=new Map();
  for(const o of value.observations){
    const key=o.entityId+':'+o.property;check(!observations.has(key),'DUPLICATE_OBSERVATION');observations.set(key,o);
    for(const s of o.sources)sourceValid(s,o);
    if(o.status==='verified')check(typeof o.value==='string'&&o.value.trim()&&o.sources.length>0,'VERIFIED_VALUE_WITHOUT_EVIDENCE');
  }
  // Mesmo uma resposta com pendencias nao pode alterar os intervalos autorizados.
  for(const r of value.replacements){
    const f=profile.fields.find(f=>f.id===r.fieldId);
    check(f&&r.oldValue===f.oldText&&r.part===f.part&&r.paragraph===f.paragraph&&r.start===f.start&&r.end===f.end,'EXTRACTION_LOCATION_MISMATCH');
    for(const s of r.sources)sourceValid(s,r);
  }
  if(!allowPending){check(value.pending.length===0,'PENDING_INFORMATION');check(value.observations.every(o=>o.status==='verified'),'PENDING_INFORMATION');}
  const replacements=[],pendingAnnotations=[];
  for(const r of value.replacements){
    const f=profile.fields.find(f=>f.id===r.fieldId);
    check(f&&r.oldValue===f.oldText&&r.part===f.part&&r.paragraph===f.paragraph&&r.start===f.start&&r.end===f.end,'EXTRACTION_LOCATION_MISMATCH');
    const o=observations.get(r.entityId+':'+r.property);
    if(allowPending&&r.status!=='verified'){
      check(o&&o.status===r.status&&['missing','conflict','illegible'].includes(r.status)&&r.value===null&&o.value===null,'UNVERIFIED_VALUE_FORBIDDEN');
      const labels={cpf:'CPF',rg:'RG',cin:'CIN',cnh:'CNH',nome:'nome',nascimento:'data de nascimento',pai:'filiacao',mae:'filiacao',profissao:'profissao',endereco:'endereco',matricula:'matricula',preco:'valor',pagamento:'pagamento',emissao:'data',data:'data',validade:'validade',estadoCivil:'estado civil'};
      const label=r.property.startsWith('filiacao.')?'filiacao':labels[r.property]||labels[r.property.split('.').at(-1)]||'informacao';
      const omitted=value.pending.some(p=>p.fieldId===r.fieldId&&['EXTRACTION_FIELD_OMITTED','CORRECTION_FIELD_OMITTED'].includes(p.code));
      const text=omitted?'pendente: conferir '+label:r.status==='missing'?'pendente: falta '+label:r.status==='conflict'?'pendente: conferir '+label:'pendente: '+label+' ilegivel';
      pendingAnnotations.push({fieldId:r.fieldId,entityId:r.entityId,property:r.property,status:r.status,text});
      replacements.push({fieldId:r.fieldId,entityId:r.entityId,property:r.property,value:text,confirmedUnchanged:false,source:{kind:'pending',entityId:r.entityId,property:r.property}});continue;
    }
    check(o&&r.status==='verified'&&typeof r.value==='string'&&o.value===r.value&&r.sources.length>0,'REPLACEMENT_NOT_VERIFIED');
    for(const s of r.sources){sourceValid(s,r);check(o.sources.some(a=>JSON.stringify(a)===JSON.stringify(s)),'REPLACEMENT_SOURCE_MISMATCH');}
    replacements.push({fieldId:r.fieldId,entityId:r.entityId,property:r.property,value:r.value,confirmedUnchanged:r.confirmedUnchanged,source:r.sources[0]});
  }
  const pendingNotes=allowPending?[...new Set(value.pending.filter(p=>!p.fieldId||!pendingAnnotations.some(a=>a.fieldId===p.fieldId)).map(p=>{
    const text=p.code.replace(/_/g,' ').toLowerCase().replace(/^missing(?: required)?\b/,'falta').replace(/^contradiction\b/,'conferir').replace(/[^\p{L}\p{N} ]/gu,'').slice(0,65).trim();
    return 'pendente: '+(text==='falta'?'falta informacao':/^(falta|conferir|ilegivel|divergencia)\b/.test(text)?text:'revisar documentos');
  }))].slice(0,3):[];
  return {schemaVersion:2,scope:'synthetic-only',templateSha256:profile.templateSha256,entities:value.entities,pending:[],replacements,...(pendingAnnotations.length?{pendingAnnotations}: {}),...(pendingNotes.length?{pendingNotes}: {})};
}
function extractResponse(response){
  check(response?.status==='completed','OPENAI_RESPONSE_INCOMPLETE');
  const outputs=(response.output||[]).flatMap(o=>o.type==='message'?(o.content||[]):[]);
  check(!outputs.some(c=>c.type==='refusal'),'OPENAI_REFUSAL');
  const texts=outputs.filter(c=>c.type==='output_text');check(texts.length===1,'OPENAI_OUTPUT_MISSING');
  let value;try{value=JSON.parse(texts[0].text);}catch{check(false,'OPENAI_OUTPUT_NOT_JSON');}
  check(validate(value),'EXTRACTION_SCHEMA_INVALID');
  return value;
}
function parseResponse(response,profile,documents){
  const value=extractResponse(response);
  return {extraction:value,plan:validateExtraction(value,profile,documents)};
}
module.exports={schema,validateExtraction,parseResponse,extractResponse};
