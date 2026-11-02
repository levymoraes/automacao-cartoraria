'use strict';
// Uma leitura pela IA; tags, avisos e ajustes locais fechados da qualificação.
const Docxtemplater=require('docxtemplater'),PizZip=require('pizzip');
const {PDFDocument}=require('pdf-lib');
const {openDocx,paragraphsOf,parseXml,sha256}=require('./docx.cjs');
const {pendingFormatting}=require('./local-template.cjs');
const {patchText}=require('./anonymize.cjs');
const {requireCondition:check}=require('./errors.cjs');
const {sourceSchema,auditSchema,validSource,auditReading,uniqueFindings,formatFindings,normalize,fold}=require('./documentary-audit.cjs');
const {documentFamily,notePosition,qualificationEdits,fieldLabel,fieldValue}=require('./template-presentation.cjs');
const workflowVersion='template-fields-v3-review-draft',W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const story=/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/;
const digest=value=>sha256(Buffer.from(JSON.stringify(value)));
const object=properties=>({type:'object',additionalProperties:false,required:Object.keys(properties),properties});
const string={type:'string'},nullable={type:['string','null']};
const schema=object({templateSha256:string,reviewedPages:{type:'array',items:object({documentId:string,pages:{type:'array',items:{type:'integer',minimum:1}}})},entities:{type:'array',items:object({slotId:string,participantId:string,name:nullable,spouseSlotId:nullable})},fields:{type:'array',items:object({tag:string,value:nullable,status:{type:'string',enum:['verified','missing','conflict','illegible']},sources:{type:'array',items:sourceSchema}})},pending:{type:'array',items:{type:'string',maxLength:240}},audit:auditSchema});
const tagPattern=()=>/\{\{([A-Z][A-Z0-9_]*)\}\}/g;
function identificationPlan(prepared,answers,invalidEntities,fieldOrigins){
 const metadata=prepared.fieldMetadata,aliases=new Map(),mappings=[],omitted=new Set(),invalidTags=new Set();
 const verified=(a,entity,property)=>a?.status==='verified'&&typeof a.value==='string'&&a.value.trim()&&a.sources.length&&a.sources.every(s=>s.purpose==='identity'&&validSource(prepared,s,entity,property));
 for(const slot of metadata.entities.filter(e=>e.kind==='person'&&!invalidEntities.has(e.id))){
  const fields=metadata.fields.filter(f=>f.entityId===slot.id),type=fields.find(f=>f.property==='documento.tipo'),sigla=fields.find(f=>f.property==='documento.sigla');
  if(!type||!verified(answers.get(type.tag),slot.id,type.property))continue;
  const family=documentFamily(answers.get(type.tag).value);if(!family)continue;
  if(sigla&&(!verified(answers.get(sigla.tag),slot.id,sigla.property)||documentFamily(answers.get(sigla.tag).value)!==family)){invalidTags.add(type.tag);invalidTags.add(sigla.tag);continue;}
  for(const f of fields.filter(f=>/^(?:rg|cin|cnh)\.(?:numero|orgao|uf|emissao|validade)$/u.test(f.property))){
   const [original,component]=f.property.split('.');if(original===family)continue;
   // Uma referencia secundaria expressa a RG continua sendo RG, mesmo se o
   // documento principal for CNH. Somente a qualificacao junto ao tipo varia.
   const primary=prepared.templateParagraphs.some(p=>{const at=p.text.indexOf('{{'+f.tag+'}}'),typeAt=p.text.lastIndexOf('{{'+type.tag+'}}',at);return at>=0&&typeAt>=0&&!/na qual consta o RG/iu.test(p.text.slice(typeAt,at));});
   if(!primary)continue;
   const property=family+'.'+component;aliases.set(f.tag,property);
   let answer=answers.get(f.tag),origin='field-source';
   if(!answer||answer.status==='missing'){
    const candidates=fields.filter(other=>other.property===property).map(other=>answers.get(other.tag)).filter(a=>verified(a,slot.id,property));
    if(new Set(candidates.map(a=>normalize(a.value))).size===1){answer={...candidates[0],tag:f.tag};origin='equivalent-document-field';}
    else if(component==='numero'&&family==='rg'&&sigla){
     const sources=answers.get(sigla.tag).sources.map(s=>({s,match:/\b(?:RG|REGISTRO(?: GERAL)?)\s*(?:n[º°.o]\s*)?:?\s*([0-9][\d./-]{3,25})\b/iu.exec(s.quote)})).filter(x=>x.match);
     if(new Set(sources.map(x=>x.match[1])).size===1){answer={tag:f.tag,status:'verified',value:sources[0].match[1],sources:sources.map(({s})=>({...s,property}))};origin='explicit-document-reference';}
    }
    if(answer?.status==='verified')answers.set(f.tag,answer);
   }
   if(verified(answer,slot.id,property)){mappings.push({tag:f.tag,entityId:slot.id,from:f.property,to:property,origin});if(origin!=='field-source')fieldOrigins.push({tag:f.tag,origin});}
  }
  if(family==='rg')for(const f of fields.filter(f=>/^(?:cin|cnh)\.validade$/u.test(f.property))){
   const a=answers.get(f.tag),number=fields.find(n=>aliases.get(n.tag)==='rg.numero'&&verified(answers.get(n.tag),slot.id,'rg.numero'));
   if((!a||a.status==='missing')&&number&&prepared.templateParagraphs.some(p=>new RegExp(',\\s*data de validade\\s*\\{\\{'+f.tag+'\\}\\}','iu').test(p.text)))omitted.add(f.tag);
  }
 }
 return {aliases,mappings,omitted,invalidTags};
}
function renderer(bytes){return new Docxtemplater(new PizZip(bytes),{delimiters:{start:'{{',end:'}}'},paragraphLoop:false,linebreaks:false,errorLogging:false,nullGetter(){check(false,'TEMPLATE_FIELD_VALUE_MISSING');}});}
function canonical(node,withoutColor=false){
 if(!node||withoutColor&&node.uri===W&&node.local==='color')return null;
 const children=node.children.map(n=>canonical(n,withoutColor)).filter(Boolean),attrs=Object.values(node.attrs).filter(a=>a.uri!=='http://www.w3.org/2000/xmlns/').map(a=>[a.uri+'|'+a.local,a.value]).sort();
 if(node.local==='rPr'&&!children.length&&!attrs.length)return null;
 return [node.uri+'|'+node.local,attrs,node.text,children];
}
function characters(paragraph){return paragraph.texts.flatMap(t=>{let run=t.parent;while(run&&!(run.uri===W&&run.local==='r'))run=run.parent;const rp=run?.children.find(n=>n.uri===W&&n.local==='rPr');return Array.from({length:t.text.length},()=>({full:JSON.stringify(canonical(rp)),base:JSON.stringify(canonical(rp,true)),red:rp?.children.some(n=>n.local==='color'&&Object.values(n.attrs).some(a=>a.local==='val'&&a.value.toUpperCase()==='FF0000'))===true}));});}
function expectedParagraph(p,values,pending,notes=[],metadata={fields:[]}){
 const styles=characters(p),expected=[],ranges=[];let text='',at=0;
 for(const match of p.text.matchAll(tagPattern())){const value=values[match[1]];check(typeof value==='string','TEMPLATE_FIELD_VALUE_MISSING');text+=p.text.slice(at,match.index);expected.push(...styles.slice(at,match.index));const start=text.length;text+=value;expected.push(...Array.from({length:value.length},()=>({...styles[match.index],pending:pending.has(match[1])})));if(pending.has(match[1]))ranges.push({paragraph:p.index,start,end:text.length,oldText:value});at=match.index+match[0].length;}
 text+=p.text.slice(at);expected.push(...styles.slice(at));
 const offset=at=>at+[...p.text.matchAll(tagPattern())].filter(m=>m.index+m[0].length<=at).reduce((sum,m)=>sum+values[m[1]].length-m[0].length,0);
 const edits=qualificationEdits(p,metadata,values,pending).map(e=>({...e,start:offset(e.start),end:offset(e.end)}));
 for(const e of [...edits].reverse()){const replacement=Array.from({length:e.value.length},(_,i)=>({...expected[Math.min(e.start+i,e.end-1)]}));text=text.slice(0,e.start)+e.value+text.slice(e.end);expected.splice(e.start,e.end-e.start,...replacement);}
 const groups=new Map();for(const n of notes){let at=offset(notePosition(p.text,n));for(const e of edits){if(e.end<=offset(notePosition(p.text,n)))at+=e.value.length-(e.end-e.start);}
  const topics=n.text.replace(/^pendente: /u,'').replace(/\. Ver tarefa\.$/u,'').split('; ').map(s=>s.trim()).filter(Boolean);if(!groups.has(at))groups.set(at,new Set());for(const topic of topics)groups.get(at).add(topic);
 }
 const insertions=[];for(const [at,topics]of [...groups].sort(([a],[b])=>b-a)){
  const note='pendente: '+[...topics].join('; '),value=' ['+note+']',style=expected[Math.max(0,at-1)];check(style,'FIELD_NOTE_ANCHOR_INVALID');
  const start=at>0?at-1:0,end=at>0?at:1,patchValue=at>0?text.slice(start,end)+value:value+text.slice(0,1);insertions.push({paragraph:p.index,start,end,value:patchValue});
  text=text.slice(0,at)+value+text.slice(at);expected.splice(at,0,...Array.from({length:value.length},(_,i)=>({...style,pending:i>=2&&i<value.length-1})));
 }
 ranges.length=0;for(let start=0;start<expected.length;){if(!expected[start].pending){start++;continue;}let end=start+1;while(end<expected.length&&expected[end].pending)end++;ranges.push({paragraph:p.index,start,end,oldText:text.slice(start,end)});start=end;}
 return {text,styles:expected,ranges,edits:edits.map(e=>({...e,paragraph:p.index})),insertions};
}
function structure(xml){
 const collapse=items=>items.filter((n,i)=>n!=='TEXT'||items[i-1]!=='TEXT');
 const visit=n=>{
  if(n.uri===W&&n.local==='rPr'&&n.parent?.local==='r')return [];
  if(n.uri===W&&n.local==='t')return ['TEXT'];
  const children=collapse(n.children.flatMap(visit));if(n.uri===W&&n.local==='r')return children;
  const attrs=Object.values(n.attrs).filter(a=>a.uri!=='http://www.w3.org/2000/xmlns/').map(a=>[a.uri+'|'+a.local,a.value]).sort();
  return [[n.uri+'|'+n.local,attrs,n.text.trim(),children]];
 };
 return visit(parseXml(xml).nodes[0]);
}
async function fillTemplate(templateBytes,metadata,values,pendingTags=new Set(),paragraphNotes=[]){
 const source=await openDocx(templateBytes);check(source.hash===metadata.preparedSha256,'FIELD_TEMPLATE_INTEGRITY_FAILED');
 const fields=new Map(metadata.fields.map(f=>[f.tag,f]));check(fields.size===metadata.fields.length&&Object.keys(values).length===fields.size&&Object.keys(values).every(k=>fields.has(k)),'FIELD_VALUES_INVALID');
 const changed=new Set(),counts=new Map();
 const notes=new Map();for(const n of paragraphNotes){const key=n.part+':'+n.paragraph;check(story.test(n.part)&&typeof n.text==='string'&&n.text.startsWith('pendente: ')&&n.text.length<=600&&!/[\u0000-\u001f\u007f]|\{\{|\}\}/u.test(n.text),'FIELD_NOTE_INVALID');check(paragraphsOf(source.parts.get(n.part)?.toString()||'').some(p=>p.index===n.paragraph&&p.text.length),'FIELD_NOTE_ANCHOR_INVALID');if(!notes.has(key))notes.set(key,[]);notes.get(key).push(n);changed.add(n.part);}
 for(const [part,b]of source.parts)if(story.test(part))for(const p of paragraphsOf(b.toString()))for(const m of p.text.matchAll(tagPattern())){check(fields.has(m[1]),'UNKNOWN_TEMPLATE_FIELD');counts.set(m[1],(counts.get(m[1])||0)+1);changed.add(part);}
 check(metadata.fields.every(f=>counts.get(f.tag)===f.occurrences),'TEMPLATE_FIELD_OCCURRENCES_CHANGED');
 let rendered;try{rendered=renderer(templateBytes);rendered.render(values);}catch(e){check(e.code==='TEMPLATE_FIELD_VALUE_MISSING','TEMPLATE_LIBRARY_RENDER_FAILED');throw e;}
 // A biblioteca nao recebe permissao para modificar estilos, imagens ou metadados.
 for(const [part,b]of source.parts)if(!changed.has(part))rendered.getZip().file(part,b);
 let bytes=rendered.getZip().generate({type:'nodebuffer',compression:'DEFLATE'}),output=await openDocx(bytes);
 const controlledEdits=[];
 for(const part of changed){const original=source.parts.get(part).toString(),paragraphs=paragraphsOf(original);let xml=output.parts.get(part).toString();const plans=paragraphs.map(p=>expectedParagraph(p,values,pendingTags,notes.get(part+':'+p.index),metadata));
  const edits=plans.flatMap(p=>p.edits);if(edits.length){xml=patchText(xml,edits).xml;controlledEdits.push(...edits.map(({paragraph,reason})=>({part,paragraph,reason})));}
  // Offsets pertencem ao texto apos os ajustes locais; aplicar do fim para o
  // inicio evita que uma anotacao desloque a ancora da seguinte.
  for(const insertion of plans.flatMap(p=>p.insertions))xml=patchText(xml,[insertion]).xml;
  const actual=paragraphsOf(xml);check(actual.length===paragraphs.length,'FIELD_PARAGRAPHS_CHANGED');const ranges=[];
  for(let i=0;i<paragraphs.length;i++){const expected=plans[i];check(actual[i].text===expected.text,'FIELD_FIXED_TEXT_CHANGED');ranges.push(...expected.ranges);}
  output.zip.file(part,ranges.length?pendingFormatting(xml,ranges):xml);
 }
 bytes=await output.zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});output=await openDocx(bytes);
 check(source.parts.size===output.parts.size,'FIELD_PACKAGE_CHANGED');
 for(const [part,b]of source.parts){const after=output.parts.get(part);check(after,'FIELD_PACKAGE_CHANGED');if(!changed.has(part)){check(after.equals(b),'FIELD_FIXED_PART_CHANGED');continue;}
  check(digest(structure(b.toString()))===digest(structure(after.toString())),'FIELD_LAYOUT_CHANGED');const original=paragraphsOf(b.toString()),actual=paragraphsOf(after.toString());
  for(let i=0;i<original.length;i++){const expected=expectedParagraph(original[i],values,pendingTags,notes.get(part+':'+i),metadata),chars=characters(actual[i]);check(actual[i].text===expected.text&&chars.length===expected.styles.length,'FIELD_FIXED_TEXT_CHANGED');check(chars.every((c,j)=>expected.styles[j].pending?c.red&&c.base===expected.styles[j].base:c.full===expected.styles[j].full),'FIELD_FORMAT_CHANGED');}
 }
 return {bytes,audit:{templateSha256:source.hash,outputSha256:output.hash,changedParts:[...changed],fixedTextVerified:true,controlledQualificationEdits:controlledEdits,characterFormattingVerified:true,originalMediaPreserved:true,pendingFields:pendingTags.size,redParagraphNotes:paragraphNotes.length,library:'docxtemplater',libraryVersion:require('../package.json').dependencies.docxtemplater}};
}
async function prepareTemplateFill({templateBytes,fieldMetadata,pdfs,model,task,config,instructions}){
 require('./openai.cjs').assertConfiguration(config);check(config.store===false&&model.userApprovedUse===true,'FIELD_TEMPLATE_NOT_AUTHORIZED');
 check(fieldMetadata?.modelId===model.id&&fieldMetadata.version===model.version&&fieldMetadata.preparedSha256===sha256(templateBytes),'FIELD_TEMPLATE_METADATA_MISMATCH');
 check(typeof task.description==='string'&&task.description.trim()&&typeof instructions==='string'&&instructions.length>100,'TASK_INSTRUCTIONS_REQUIRED');
 check(instructions.includes('audit.operation')&&instructions.includes('paragraphReviews'),'DOCUMENTARY_INSTRUCTIONS_REQUIRED');
 check(pdfs.length>0&&pdfs.length<=(config.limits.maxPdfs||20)&&new Set(pdfs.map(p=>p.id)).size===pdfs.length,'INVALID_INPUT_DOCUMENTS');check(pdfs.reduce((n,p)=>n+p.bytes.length,0)<=config.limits.maxPdfBytes,'PDF_TOTAL_SIZE_LIMIT');
 const documents=[];for(const p of pdfs){check(/^[A-Za-z0-9_-]{1,80}$/.test(p.id)&&p.bytes.subarray(0,5).toString()==='%PDF-','INVALID_PDF_INPUT');let pdf;try{pdf=await PDFDocument.load(p.bytes.toString('base64'),{updateMetadata:false});}catch{check(false,'PDF_UNREADABLE_OR_ENCRYPTED');}documents.push({id:p.id,sha256:sha256(p.bytes),pages:pdf.getPageCount()});}
 check(documents.every(d=>d.pages>0)&&documents.reduce((n,d)=>n+d.pages,0)<=config.limits.maxPdfPages,'PDF_PAGE_LIMIT');
 const fields=fieldMetadata.fields.map(({tag,entityId,property,case:letterCase,contextBefore,contextAfter,label})=>({tag,entityId,property,case:letterCase,label,contextBefore,contextAfter}));
 // Somente contexto dos campos; nao envia DOCX, valores antigos ou arquivos privados.
 const template=await openDocx(templateBytes);
 const templateParagraphs=[...template.parts].filter(([part])=>story.test(part)).flatMap(([part,b])=>paragraphsOf(b.toString()).map(p=>({part,paragraph:p.index,text:p.text})));
 const referenceDate=new Intl.DateTimeFormat('en-CA',{timeZone:'America/Sao_Paulo',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
 const context={workflowVersion,referenceDate,templateSha256:sha256(templateBytes),selectedModel:{id:model.id,name:model.name||model.id,actType:model.actType||task.actType||'minuta'},task,documents,entities:fieldMetadata.entities,couples:fieldMetadata.couples,fields,templateParagraphs};
 check(JSON.stringify(context).length<=config.limits.maxTemplateCharacters,'TEMPLATE_CONTEXT_LIMIT');
 const request={model:config.model,service_tier:config.service_tier,reasoning:config.reasoning,store:false,prompt_cache_options:{mode:'explicit'},max_output_tokens:config.max_output_tokens,instructions:require('./openai.cjs').systemInstructions(instructions),text:{format:{type:'json_schema',name:'minuta_template_fields',strict:true,schema}},input:[{role:'user',content:[{type:'input_text',text:JSON.stringify(context)},...pdfs.map(p=>({type:'input_file',filename:p.id+'.pdf',detail:'high',file_data:'data:application/pdf;base64,'+p.bytes.toString('base64')}))]}]};
 return {production:true,templateFields:true,workflowVersion,referenceDate,templateParagraphs,request,requestSha256:digest(request),templateBase64:templateBytes.toString('base64'),fieldMetadata,documents,model,task,config:{...config,nativeDocument:false},mode:'live'};
}
async function evaluateTemplateFill(prepared,raw){
 check(raw?.status==='completed','OPENAI_RESPONSE_INCOMPLETE');const content=(raw.output||[]).filter(o=>o.type==='message').flatMap(o=>o.content||[]);check(!content.some(c=>c.type==='refusal'),'OPENAI_REFUSAL');const texts=content.filter(c=>c.type==='output_text');check(texts.length===1,'OPENAI_OUTPUT_MISSING');let value;try{value=JSON.parse(texts[0].text);}catch{check(false,'OPENAI_OUTPUT_NOT_JSON');}
 check(require('./generated/template-fields-validator.cjs')(value),'FIELD_RESPONSE_SCHEMA_INVALID');check(value.templateSha256===prepared.fieldMetadata.preparedSha256,'FIELD_RESPONSE_TEMPLATE_MISMATCH');
 const reviewed=new Set();for(const r of value.reviewedPages){const doc=prepared.documents.find(d=>d.id===r.documentId);check(doc,'FIELD_REVIEW_DOCUMENT_INVALID');for(const page of r.pages){const key=r.documentId+':'+page;check(page<=doc.pages&&!reviewed.has(key),'FIELD_REVIEW_PAGE_INVALID');reviewed.add(key);}}
 check(prepared.documents.every(d=>Array.from({length:d.pages},(_,i)=>reviewed.has(d.id+':'+(i+1))).every(Boolean)),'FIELD_DOCUMENT_NOT_FULLY_REVIEWED');
 const review=auditReading(prepared,value),findings=review.findings;
 const annotate=(p,topic,anchor='end')=>{let n=review.paragraphNotes.find(n=>n.part===p.part&&n.paragraph===p.paragraph&&(n.anchor||'end')===anchor);if(!n){n={part:p.part,paragraph:p.paragraph,anchor,text:'pendente: '+topic};review.paragraphNotes.push(n);}else if(!n.text.includes(topic))n.text+='; '+topic;};
 const metadata=prepared.fieldMetadata,answers=new Map(),assignments=new Map(),invalidEntities=new Set(review.invalidSlots),fieldOrigins=[];
 for(const e of value.entities){check(!assignments.has(e.slotId)&&metadata.entities.some(s=>s.id===e.slotId&&['person','organization'].includes(s.kind)),'FIELD_ENTITY_INVALID');assignments.set(e.slotId,e);}
 for(const slot of metadata.entities.filter(s=>['person','organization'].includes(s.kind))){const e=assignments.get(slot.id);if(!e?.name?.trim())invalidEntities.add(slot.id);}
 const identities=new Map();for(const e of assignments.values())if(e.name){const key=normalize(e.name);if(identities.has(key)){invalidEntities.add(e.slotId);invalidEntities.add(identities.get(key));}identities.set(key,e.slotId);}
 for(const f of value.fields){check(metadata.fields.some(m=>m.tag===f.tag)&&!answers.has(f.tag),'FIELD_RESPONSE_UNKNOWN_OR_DUPLICATE');answers.set(f.tag,f);}
 // Aproveita nomes explicitamente extraidos no inventario de participantes,
 // sem inferir identidades nem substituir uma resposta de campo existente.
 for(const f of metadata.fields.filter(f=>f.property==='nome'&&!answers.has(f.tag)&&!invalidEntities.has(f.entityId))){const p=review.participants.get(assignments.get(f.entityId)?.participantId);if(!p)continue;const sources=p.sources.filter(s=>normalize(s.quote).includes(normalize(p.name))).map(s=>({...s,entityId:f.entityId,property:f.property,purpose:'identity'}));if(sources.length){answers.set(f.tag,{tag:f.tag,value:p.name,status:'verified',sources});fieldOrigins.push({tag:f.tag,origin:'audit.operation.participants'});}}
 const identification=identificationPlan(prepared,answers,invalidEntities,fieldOrigins);
 const identifiers=new Map(),conflictingTags=new Set(identification.invalidTags);
 for(const f of metadata.fields.filter(f=>['cpf','cnpj'].includes(f.property))){const a=answers.get(f.tag);if(a?.status!=='verified'||typeof a.value!=='string')continue;const key=f.property+':'+a.value.replace(/\W/gu,'').toUpperCase(),prior=identifiers.get(key);if(prior&&prior.entityId!==f.entityId){conflictingTags.add(prior.tag);conflictingTags.add(f.tag);}else identifiers.set(key,f);}
 const properties=new Map();for(const f of metadata.fields){const a=answers.get(f.tag);if(a?.status!=='verified'||typeof a.value!=='string')continue;const key=f.entityId+':'+(identification.aliases.get(f.tag)||f.property);if(!properties.has(key))properties.set(key,[]);properties.get(key).push({tag:f.tag,value:normalize(a.value)});}for(const group of properties.values())if(new Set(group.map(f=>f.value)).size>1)for(const f of group)conflictingTags.add(f.tag);
 const values={},pendingTags=new Set(),grammarCorrections=[];
 for(const message of value.pending)findings.push({category:'B',subject:'Confirmação documental',message,critical:true,sources:[]});
 for(const f of metadata.fields){const answer=answers.get(f.tag),entity=assignments.get(f.entityId);let reason=answer?.status||'omitted';
  if(identification.omitted.has(f.tag)){values[f.tag]='';continue;}
  let verified=answer?.status==='verified'&&typeof answer.value==='string'&&answer.value.trim().length>0&&answer.value.length<=2000&&!/[\u0000-\u001f\u007f]|\{\{|\}\}|pendente\s*:/iu.test(answer.value)&&!invalidEntities.has(f.entityId)&&!conflictingTags.has(f.tag);
  if(answer?.status==='verified'&&!verified)reason='conflict';
  if(verified&&f.property==='nome'&&entity?.name&&normalize(answer.value)!==normalize(entity.name)){verified=false;reason='conflict';}
  if(verified&&f.property==='estadoCivil'){
   const participant=review.participants.get(entity?.participantId),state=fold(answer.value),prefix={married:'casad',single:'solteir',divorced:'divorciad',widowed:'viuv'}[participant?.maritalStatus];
   if(prefix&&!state.startsWith(prefix)){verified=false;reason='conflict';}
  }
  if(verified)verified=answer.sources.length>0&&answer.sources.every(s=>validSource(prepared,s,f.entityId,identification.aliases.get(f.tag)||f.property));
  const property=fold(f.property);
  if(verified&&/avaliacao.*valor/.test(property)&&!answer.sources.some(s=>s.purpose==='municipal_valuation')){verified=false;reason='conflict';}
  if(verified&&/itbi\.base/.test(property)&&!answer.sources.some(s=>s.purpose==='itbi_base')){verified=false;reason='conflict';}
  if(verified&&/preco|price/.test(property)&&!answer.sources.some(s=>['contract_confirmation','declared_transaction_value'].includes(s.purpose))){verified=false;reason='conflict';}
  if(verified&&/preco|price/.test(property)&&answer.sources.some(s=>s.purpose==='declared_transaction_value')&&!answer.sources.some(s=>s.purpose==='contract_confirmation')){findings.push({category:'B',subject:'Preço contratual',message:'Valor declarado encontrado: '+answer.value+'. Conferir a confirmação contratual; não comprova pagamento ou quitação.',critical:true,sources:answer.sources});for(const p of prepared.templateParagraphs.filter(p=>p.text.includes('{{'+f.tag+'}}')))annotate(p,'confirmar preço contratual; valor declarado aproveitado','preco');}
  if(verified&&!/^itbi\./.test(property)&&/pagamento|payment|quitacao/.test(property)){
   const purposes=/meio|forma|method/.test(property)?['payment_method','contract_confirmation']:/data|date|momento/.test(property)?['payment_date','contract_confirmation']:/valor|amount/.test(property)?['paid_amount','contract_confirmation']:['contract_confirmation'];
   if(!answer.sources.some(s=>purposes.includes(s.purpose))){verified=false;reason='confirmation';}
  }
  if(answer?.status==='verified'&&!verified)reason='conflict';
  if(verified){let text=fieldValue(f,answer.value,prepared.templateParagraphs);if(text!==answer.value)grammarCorrections.push({tag:f.tag,entityId:f.entityId,before:answer.value,after:text});const participant=review.participants.get(entity?.participantId);
   // Flexao de um valor atomico comprovado; nao altera frases ou fatos juridicos.
   if(['nacionalidade','estadoCivil'].includes(f.property)&&['masculine','feminine'].includes(participant?.gender)&&/^(?:brasileir|solteir|casad|divorciad|viúv)[oa]$/iu.test(text)){
    const suffix=participant.gender==='feminine'?'a':'o',inflected=text.slice(0,-1)+(text.at(-1)===text.at(-1).toUpperCase()?suffix.toUpperCase():suffix);
    if(inflected!==text){grammarCorrections.push({tag:f.tag,entityId:f.entityId,before:text,after:inflected});text=inflected;}
   }
   values[f.tag]=f.case==='upper'?text.toLocaleUpperCase('pt-BR'):text;
  }else{
   const administrative=/^(?:ato\.)?(?:livro|folha|folhas|termo)(?:\.|$)/iu.test(f.property)||metadata.entities.some(e=>e.id===f.entityId&&e.kind==='transaction')&&/^(?:ato\.)?numero$/iu.test(f.property);
   const label=fieldLabel(f).replace(/[\u0000-\u001f\[\]{}]/gu,' ').slice(0,70),unassigned=invalidEntities.has(f.entityId);
   values[f.tag]=unassigned?'pendente: revisar posição do modelo':'pendente: '+(reason==='missing'?'informar ':reason==='illegible'?'legibilidade de ':'conferir ')+label;pendingTags.add(f.tag);
   if(!unassigned)findings.push({category:reason==='missing'?'A':'B',subject:administrative?'Dados administrativos':f.entityId+' / '+f.property.split('.')[0],message:label+': '+(answer?.value?'Valor extraído não aceito para este campo: '+answer.value+'. Conferir as fontes e a finalidade.':reason==='missing'?'Informação não localizada após a leitura integral.':reason==='illegible'?'Informação ilegível; conferir documento.':'Conferir a correspondência ou a extração deste dado.'),critical:!administrative,sources:answer?.sources||[]});
  }
 }
 // A IA confere o texto com as substituicoes na mesma leitura. Aqui se
 // verifica a composicao efetiva, sem permitir reescrita de texto fixo.
 for(const p of prepared.templateParagraphs){
  const text=p.text.replace(tagPattern(),(_,tag)=>values[tag]||'');
  if(/\b(n[º°o.]?\s*\d+)\s*,?\s*\1\b/iu.test(text)){findings.push({category:'B',subject:'Endereço',message:'Número de endereço duplicado após o preenchimento.',critical:true,sources:[]});annotate(p,'conferir duplicação do endereço');}
  for(const [role,expression]of [['seller',/\boutorgantes vendedores\b/iu],['buyer',/\boutorgados compradores\b/iu]]){
   const count=metadata.entities.filter(e=>e.role===role||e.role===role+'_spouse').length;
   if(count===1&&expression.test(text)){findings.push({category:'C',subject:'Concordância de participantes',message:'Referência plural incompatível com uma única parte no papel '+role+'.',critical:true,sources:[]});annotate(p,'conferir concordância dos participantes no texto fixo');}
  }
  const tags=[...p.text.matchAll(tagPattern())];
  for(const [i,m]of tags.entries()){
   const field=metadata.fields.find(f=>f.tag===m[1]);if(!field||field.property!=='nome')continue;
   const participant=review.participants.get(assignments.get(field.entityId)?.participantId);if(!participant)continue;
   const end=tags.slice(i+1).find(n=>metadata.fields.some(f=>f.tag===n[1]&&f.property==='nome'))?.index||p.text.length,tail=p.text.slice(m.index+m[0].length,end),first=tags[i+1];
   const ownsQualification=first&&first.index<end&&metadata.fields.some(f=>f.tag===first[1]&&f.entityId===field.entityId);
   if(!ownsQualification&&!/^[,;\s]*(?:portador[a]?|filh[oa]|brasileir[oa])\b/iu.test(tail))continue;
   const qualification=fold(tail.replace(tagPattern(),(_,tag)=>['nacionalidade','estadoCivil'].includes(metadata.fields.find(f=>f.tag===tag)?.property)?values[tag]||'':''));
   const wrong=participant.gender==='feminine'?/\b(?:portador|filho|brasileiro|solteiro|divorciado|viuvo)\b/u:participant.gender==='masculine'?/\b(?:portadora|filha|brasileira|solteira|divorciada|viuva)\b/u:null;
   if(wrong?.test(qualification)){findings.push({category:'C',subject:'Concordância / '+field.entityId,message:'A qualificação fixa e os dados preenchidos não concordam. O texto foi preservado para revisão do funcionário.',critical:true,sources:[]});annotate(p,'conferir gênero da qualificação fixa');}
  }
 }
 const hasPending=findings.length>0||pendingTags.size>0,notes=review.paragraphNotes.filter(n=>prepared.templateParagraphs.some(p=>p.part===n.part&&p.paragraph===n.paragraph&&p.text.length));
 if(hasPending){const first=prepared.templateParagraphs.find(p=>p.part==='word/document.xml'&&p.text.length);check(first,'FIELD_DRAFT_LABEL_ANCHOR_REQUIRED');const label='pendente: MINUTA PRELIMINAR — exige conferência e aprovação humana antes de qualquer utilização';notes.push({part:first.part,paragraph:first.paragraph,anchor:'end',text:label});}
 const filled=await fillTemplate(Buffer.from(prepared.templateBase64,'base64'),metadata,values,pendingTags,notes);
 const facts=review.financialFacts.map(f=>f.nature+' — '+({declared_transaction_value:'valor declarado',municipal_valuation:'avaliação municipal',itbi_base:'base de cálculo',paid_amount:'valor recolhido/pagamento informado',payment_method:'meio informado',payment_date:'data informada'}[f.purpose])+': '+f.value+' ['+[...new Set(f.sources.map(s=>s.kind==='task'?'solicitação':s.documentId+', página '+s.page))].join('; ')+']');
 const pendingNotes=[formatFindings(findings,review.requiredModelCharacteristics),...(facts.length?['Informações documentais aproveitáveis (não confirmam quitação):\n'+facts.map(f=>'- '+f).join('\n')]:[])].filter(Boolean);
 return {bytes:filled.bytes,report:{status:hasPending?'template-fields-draft-with-pending':'template-fields-docx',workflowVersion,documentGenerated:true,technicalValidation:'passed',documentaryValidation:hasPending?'pending-human-review':'requires-human-review',legalApproval:false,hasPending,pendingNotes,findings:uniqueFindings(findings),financialFacts:review.financialFacts,documentRecords:review.records,evaluation:{expectedPending:[...pendingTags]},audit:{...filled.audit,quarantinedRecords:review.quarantinedRecords,quarantinedFinancialFacts:review.quarantinedFinancialFacts,grammarCorrections,fieldOrigins,identificationMappings:identification.mappings.filter(m=>!pendingTags.has(m.tag)),notApplicableFields:[...identification.omitted],filledFields:metadata.fields.length-pendingTags.size-identification.omitted.size,documentaryReviewVerified:review.unreviewedParagraphs.length===0,unreviewedParagraphs:review.unreviewedParagraphs,pendingLocation:'red-markers-and-linear-issue',clausesRewritten:false},request:{model:prepared.config.model,serviceTier:prepared.config.service_tier,reasoning:prepared.config.reasoning.effort,originalDocxTransmitted:false,calls:1,codeInterpreter:false},semanticContentRequiresHumanReview:true}};
}
module.exports={workflowVersion,schema,prepareTemplateFill,evaluateTemplateFill,fillTemplate};
