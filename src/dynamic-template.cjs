'use strict';
// O inventario e uma resposta da IA para a solicitacao atual, nao um cadastro
// permanente de campos. Este modulo so valida e aplica dados; nao usa rede.
const {openDocx,paragraphsOf,parseXml,sha256}=require('./docx.cjs');
const {collectStories,coverageOf}=require('./operation.cjs');
const {prepareLocal,evaluateLocal}=require('./local-template.cjs');
const {requireCondition:check}=require('./errors.cjs');
const digest=value=>sha256(Buffer.from(JSON.stringify(value)));
const escape=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');
const categories=new Set(['person-name','identity-document','personal-date','parentage','nationality','birthplace','profession','civil-status','capacity','marriage','address','property-identification','property-characteristic','transaction-value','transaction-date','transaction-payment','tax','certificate','record','other-variable']);
function protectedLocations(stories){
 const result=[];
 const add=(p,start,end,reason)=>{if(end>start)result.push({id:'protected-'+result.length,part:p.part,paragraph:p.paragraph,start,end,reason,valueSha256:sha256(Buffer.from(p.text.slice(start,end)))});};
 const institutionalHeaders=new Set(stories.filter(p=>/^word\/(header|footer)\d+\.xml$/.test(p.part)&&/CART[ÓO]RIO|TABELI|SERVENTIA|REGISTRO|COMARCA|ESTADO D[AO]/iu.test(p.text)).map(p=>p.part));
 for(const p of stories){
  if(institutionalHeaders.has(p.part)){add(p,0,p.text.length,'institutional-story');continue;}
  const following=stories.find(s=>s.part===p.part&&s.paragraph===p.paragraph+1);
  if(/^(?:ESCREVENTE|TABELI|SUBSTITUT)/i.test(p.text.trim())||following&&/^(?:ESCREVENTE\s+AUTORIZAD[OA]|TABELI[ÃA]O(?:\s+SUBSTITUTO)?|SUBSTITUTO)[\s.]*$/i.test(following.text.trim()))add(p,0,p.text.length,'staff-signature');
  for(const m of p.text.matchAll(/(?:perante mim,\s*|Eu,\s*[_\s]*,?\s*)([^,;]{3,160}),\s*(?:Escrevente\s+Autorizad[oa]|Tabeli[aã]o[^,.]*)/giu)){
   const name=m[1].trim(),at=m.index+m[0].indexOf(name);add(p,at,at+name.length,'staff-identification');
  }
  for(const m of p.text.matchAll(/nesta\s+Serventia[\s\S]{0,450}?(?=perante mim,)/giu))add(p,m.index,m.index+m[0].length,'institution-identification');
  for(const m of p.text.matchAll(/[“"]Para os bens imóveis[\s\S]*?[”"]/giu))add(p,m.index,m.index+m[0].length,'normative-quotation');
  for(const m of p.text.matchAll(/(?:Lei|Provimento|Artigo|art\.)\s*(?:n[º°o.]?\s*)?\d[\d.,/\sº°-]{0,35}/giu))add(p,m.index,m.index+m[0].length,'normative-reference');
  for(const m of p.text.matchAll(/(?:Código (?:Civil|Penal)|Cadastro Imobiliário) brasileiro/giu))add(p,m.index,m.index+m[0].length,'normative-name');
 }
 return result;
}
function fixedCoincidences(doc,stories,fields,protectedFields){
 const fixedOldValueExceptions=[],fixedAuxiliaryValueExceptions=[];
 for(const f of fields)for(const p of stories)for(const m of p.text.matchAll(new RegExp(`(?<![\\p{L}\\p{N}])${escape(f.oldText)}(?![\\p{L}\\p{N}])`,'giu'))){
  const fixed=protectedFields.some(c=>c.part===p.part&&c.paragraph===p.paragraph&&c.start<=m.index&&c.end>=m.index+m[0].length);
  const system=/^brasileiro$/i.test(m[0])&&/(?:Código (?:Civil|Penal)|Cadastro Imobiliário) $/i.test(p.text.slice(0,m.index));
  if(fixed||system)fixedOldValueExceptions.push({part:p.part,paragraph:p.paragraph,start:m.index,end:m.index+m[0].length,value:m[0],reason:'Coincidencia em elemento institucional, norma ou nome oficial de sistema; nao pertence aos dados da operacao'});
 }
 const short=fields.filter(f=>/^\d{1,4}$/.test(f.oldText));
 for(const [part,data] of doc.parts){if(!/\.(xml|rels)$/.test(part))continue;
  parseXml(data.toString()).nodes.forEach((n,nodeIndex)=>{
   const size=n.uri==='http://schemas.openxmlformats.org/wordprocessingml/2006/main'&&['sz','szCs','spacing','pgMar','cols','pgSz','ind','tab','pos','zoom','defaultTabStop','tblInd','compatSetting'].includes(n.local);
   const app=part==='docProps/app.xml'&&['AppVersion','TotalTime','Pages','Words','Characters','CharactersWithSpaces','Lines','Paragraphs','DocSecurity','vector','i4'].includes(n.local);
   const namespace=nodeIndex===0&&Object.keys(n.attrs).every(name=>name==='xmlns'||name.startsWith('xmlns:')||name==='mc:Ignorable');
   const shape=n.uri==='urn:schemas-microsoft-com:vml'&&n.local==='shapetype';
   const timestamp=part==='docProps/core.xml'&&['created','modified','lastPrinted'].includes(n.local);
   if(!size&&!app&&!namespace&&!shape&&!timestamp)return;
   const value=n.text+' '+Object.values(n.attrs).map(a=>a.value).join(' ');
   for(const f of short)if(new RegExp(`(?<![\\p{L}\\p{N}])${escape(f.oldText)}(?![\\p{L}\\p{N}])`,'iu').test(value))fixedAuxiliaryValueExceptions.push({fieldId:f.id,part,nodeIndex,element:n.local,valueSha256:sha256(Buffer.from(value)),partSha256:sha256(data),reason:'Coincidencia numerica em estrutura, configuracao ou metadado tecnico original, preservado por hash e posicao'});
  });
 }
 return {fixedOldValueExceptions,fixedAuxiliaryValueExceptions};
}
function criticalCandidates(stories,fields=[],protectedFields=[]){
 const result=[];
 for(const p of stories){
  const add=(start,end,type)=>result.push({part:p.part,paragraph:p.paragraph,start,end,type});
  for(const [pattern,type] of [[/\b\d{3}\.\d{3}\.\d{3}-\d{2}\b/gu,'cpf'],[/\b\d{2}\.\d{3}\.\d{3}\/\d{4}-\d{2}\b/gu,'cnpj'],[/\b\d{2}\/\d{2}\/\d{4}\b/gu,'date'],[/\b\d{9,}\b/gu,'identifier'],[/\b(?:solteir[oa]|divorciad[oa]|casad[oa]s?)\b/giu,'civil-status'],[/\bbrasileir[oa](?=,\s*(?:maior|menor))/giu,'nationality']])for(const m of p.text.matchAll(pattern))add(m.index,m.index+m[0].length,type);
  for(const m of p.text.matchAll(/R\$\s*(\d+(?:\.\d{3})*(?:,\d{2})?)/gu)){const at=m.index+m[0].lastIndexOf(m[1]);add(at,at+m[1].length,'value');}
  // Defesa generica para filiacao em qualificacoes; nao depende de nomes,
  // identificadores de participantes ou inventario cadastrado do modelo.
  for(const m of p.text.matchAll(/filh[oa]\s+de\s+([^,;]+?)\s+e\s+([^,;]+)(?=[,;])/giu)){
   const owner=fields.filter(f=>f.part===p.part&&f.paragraph===p.paragraph&&f.category==='person-name'&&f.property==='nome'&&f.end<=m.index).sort((a,b)=>b.end-a.end)[0];
   if(!owner||protectedFields.some(f=>f.part===p.part&&f.paragraph===p.paragraph&&f.start>=owner.end&&f.start<m.index))continue;
   if(/\s+e\s+/iu.test(m[1])||/\s+e\s+/iu.test(m[2]))continue;
   for(const group of [1,2]){const name=m[group].trim(),start=m.index+m[0].indexOf(name,group===1?0:m[0].indexOf(m[1])+m[1].length);result.push({part:p.part,paragraph:p.paragraph,start,end:start+name.length,type:'parentage',property:'filiacao.'+group});}
  }
  for(const pattern of [/OUTORGANTES?\s+VENDEDOR(?:A|ES|AS)?\s*:\s*([^,;]{2,180})(?=,\s*(?:brasileir|inscrit))/giu,/OUTORGAD[OA]S?\s+COMPRADOR(?:A|ES|AS)?\s*:\s*([^,;]{2,180})(?=,\s*brasileir)/giu,/(?:sua\s+esposa|seu\s+esposo|casad[oa]\s+com)\s+([^,;]{2,180})(?=,\s*brasileir)/giu])for(const m of p.text.matchAll(pattern)){const at=m.index+m[0].lastIndexOf(m[1]);add(at,at+m[1].length,'person-name');}
 }
 return result;
}
function assertRoleLocations(stories,fields,entities){
 const people=new Map(entities.filter(e=>e.kind==='person'||e.kind==='organization').map(e=>[e.id,e]));
 for(const p of stories.filter(p=>p.part==='word/document.xml')){
  const seller=/OUTORGANTES?\s+VENDEDOR(?:A|ES|AS)?\s*:/iu.exec(p.text),buyer=/OUTORGAD[OA]S?\s+COMPRADOR(?:A|ES|AS)?\s*:/iu.exec(p.text);
  if(!seller||!buyer||buyer.index<=seller.index)continue;
  const end=p.text.search(/DA IDENTIFICAÇÃO|VERIFICAÇÃO DA IDENTIDADE|DO IMÓVEL/iu);const limit=end>buyer.index?end:p.text.length;
  const names=fields.filter(f=>f.part===p.part&&f.paragraph===p.paragraph&&f.property==='nome'&&people.has(f.entityId)&&f.start>=seller.index&&f.start<limit).sort((a,b)=>a.start-b.start);
  for(const f of fields.filter(f=>f.part===p.part&&f.paragraph===p.paragraph&&people.has(f.entityId)&&f.start>=seller.index&&f.start<limit)){
   const role=people.get(f.entityId).role;check(f.start<buyer.index?/^seller/.test(role):/^buyer/.test(role),'PARTICIPANT_ROLE_LOCATION_MISMATCH');
   const owner=[...names].reverse().find(n=>n.start<=f.start);
   // O estado civil plural depois da qualificacao dos dois conjuges pertence
   // ao casal. Aceitar apenas vinculo reciproco, nunca uma terceira pessoa.
   const person=people.get(f.entityId),partner=people.get(owner?.entityId);
   const sharedCivilStatus=f.category==='civil-status'&&/^casad[oa]s$/iu.test(f.oldText)&&person.spouseId===partner?.id&&partner.spouseId===person.id;
   check(owner?.entityId===f.entityId||sharedCivilStatus,'PARTICIPANT_QUALIFICATION_MISMATCH');
  }
 }
}
async function compileInventory({templateBytes,model,task,documents,analysis,allowDataPending=false}){
 const doc=await openDocx(templateBytes),stories=collectStories(doc,paragraphsOf);
 check(typeof task.description==='string'&&task.description.trim().length>0,'TASK_INSTRUCTIONS_REQUIRED');
 check(analysis?.schemaVersion===1&&analysis.templateSha256===doc.hash&&analysis.taskSha256===digest(task)&&analysis.documentsSha256===digest(documents),'DYNAMIC_ANALYSIS_CONTEXT_MISMATCH');
 check((analysis.audit?.status==='passed'||allowDataPending&&analysis.audit?.status==='pending')&&(analysis.pending?.length===0||allowDataPending)&&analysis.audit.unmappedVariables?.length===0,'DYNAMIC_INVENTORY_REVIEW_REQUIRED');
 check(Array.isArray(analysis.entities)&&analysis.entities.length>0&&Array.isArray(analysis.fields)&&analysis.fields.length>0&&analysis.fields.length<=1000,'DYNAMIC_INVENTORY_INVALID');
 const expected=coverageOf(doc,paragraphsOf);
 check(Array.isArray(analysis.coverage)&&analysis.coverage.length===expected.length&&new Set(analysis.coverage.map(c=>c.part+':'+c.paragraph)).size===expected.length&&expected.every(c=>analysis.coverage.some(a=>a.part===c.part&&a.paragraph===c.paragraph&&a.textSha256===c.textSha256&&a.reviewed===true)),'DYNAMIC_COVERAGE_INCOMPLETE');
 const protectedFields=protectedLocations(stories),fields=[];
 for(const f of analysis.fields){
  const p=stories.find(p=>p.part===f.part&&p.paragraph===f.paragraph);
  check(/^[A-Za-z0-9_.-]{1,120}$/.test(f.id)&&categories.has(f.category)&&typeof f.entityId==='string'&&typeof f.property==='string'&&/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(f.part)&&p&&Number.isInteger(f.start)&&Number.isInteger(f.end)&&f.start>=0&&f.end>f.start&&f.end<=p.text.length&&p.text.slice(f.start,f.end)===f.oldText,'DYNAMIC_FIELD_ANCHOR_INVALID');
  check(!protectedFields.some(c=>c.part===f.part&&c.paragraph===f.paragraph&&c.start<f.end&&c.end>f.start),'INSTITUTIONAL_FIELD_PROTECTED');
  check(!/(?:\b(?:outorgantes?|outorgados?|quitação|Provimento|Escrevente|Tabelião)\b|Código\s+(?:Civil|Penal|de Normas)|perante mim|sob pena|Lei\s+n[º°])/iu.test(f.oldText),'FIXED_CLAUSE_CHANGE_FORBIDDEN');
  check(!fields.some(c=>c.id===f.id||c.part===f.part&&c.paragraph===f.paragraph&&c.start<f.end&&c.end>f.start),'DYNAMIC_FIELD_OVERLAP');
  fields.push({...f,required:true,anchorContext:{before:p.text.slice(Math.max(0,f.start-32),f.start),after:p.text.slice(f.end,f.end+32)}});
 }
 const constraints=[];
 assertRoleLocations(stories,fields,analysis.entities);
 const coincidence=fixedCoincidences(doc,stories,fields,protectedFields);
 for(const candidate of criticalCandidates(stories,fields,protectedFields)){
  const covered=fields.some(f=>f.part===candidate.part&&f.paragraph===candidate.paragraph&&f.start<=candidate.start&&f.end>=candidate.end)||protectedFields.some(f=>f.part===candidate.part&&f.paragraph===candidate.paragraph&&f.start<=candidate.start&&f.end>=candidate.end)||coincidence.fixedOldValueExceptions.some(f=>f.part===candidate.part&&f.paragraph===candidate.paragraph&&f.start<=candidate.start&&f.end>=candidate.end);
  check(covered,'UNREGISTERED_CRITICAL_VARIABLE');
 }
 return {schemaVersion:2,modelId:model.id,scope:'local-authorized-template',origin:'user-supplied-local',templateSha256:doc.hash,anonymizationApproved:false,externalSharingAllowed:false,localUse:{approved:model.userApprovedUse===true,originalSha256:model.sha256,convertedSha256:doc.hash,preservedParts:[...doc.parts].map(([part,b])=>({part,sha256:sha256(b)}))},operation:{inventoryReviewed:true,reviewType:'request-specific-ai-analysis-with-local-validation',requiresHumanLegalReview:true,entities:analysis.entities,couples:analysis.couples||[],constraints,coverage:expected,protectedFields,...coincidence},fields};
}
async function prepareDynamicMock({templateBytes,pdfBytes,model,task,fixture,config}){
 check(config.openaiMode==='mock'&&config.paidApiEnabled===false,'REGISTERED_TEMPLATE_LIVE_DISABLED');
 check(fixture.localOnly===true&&fixture.generatedFictionalPdf===true&&fixture.analysisSha256===digest(fixture.payload.mockAnalysis),'DYNAMIC_MOCK_NOT_APPROVED');
 // O simulador recebe o contexto da tarefa; nao afirma ter executado IA real.
 const analysis={...fixture.payload.mockAnalysis,taskSha256:digest(task),documentsSha256:digest(fixture.documents)};
 const profile=await compileInventory({templateBytes,model,task,documents:fixture.documents,analysis});
 const payload={...fixture.payload,profile,templateBase64:templateBytes.toString('base64'),pdfs:[{id:'pacote-ficticio',base64:pdfBytes.toString('base64')}]};
 const prepared=prepareLocal(payload,{...fixture,profileSha256:digest(profile)},config);prepared.dynamicAudit={taskSha256:analysis.taskSha256,inventorySha256:digest(analysis),fields:profile.fields.length};return prepared;
}
module.exports={compileInventory,prepareDynamicMock,evaluateLocal,protectedLocations,fixedCoincidences,digest,criticalCandidates};
