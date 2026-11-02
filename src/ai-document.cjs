'use strict';
// Integracao da ferramenta nativa e conferencia somente leitura. Nenhuma
// funcao deste modulo localiza dados variaveis ou edita o documento.
const {PDFDocument}=require('pdf-lib');
const {openDocx,parseXml,paragraphsOf,sha256}=require('./docx.cjs');
const {requireCondition:check}=require('./errors.cjs');
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const story=/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/;
const digest=x=>sha256(Buffer.from(JSON.stringify(x)));
const id=x=>typeof x==='string'&&/^[A-Za-z0-9_-]{1,200}$/.test(x);
const referenceId=x=>typeof x==='string'&&/^[A-Za-z0-9_.-]{1,200}$/.test(x);
const workflowVersion='native-reviewed-v3';
const readingSchema=require('../config/native-reading.schema.json');
async function prepareNativeDocument({templateBytes,pdfs,model,task,config,instructions}){
 require('./openai.cjs').assertConfiguration(config);
 check(config.store===false&&model.userApprovedUse===true,'NATIVE_DOCUMENT_NOT_AUTHORIZED');
 check(typeof instructions==='string'&&instructions.length>100&&typeof task.description==='string'&&task.description.trim(),'TASK_INSTRUCTIONS_REQUIRED');
 const template=await openDocx(templateBytes);
 check(pdfs.length>0&&pdfs.length<=49,'INVALID_INPUT_DOCUMENTS');
 check(templateBytes.length+pdfs.reduce((n,p)=>n+p.bytes.length,0)<50000000,'PROVIDER_FILE_SIZE_LIMIT');
 const documents=[];
 for(const p of pdfs){let pdf;try{pdf=await PDFDocument.load(p.bytes.toString('base64'),{updateMetadata:false});}catch{check(false,'PDF_UNREADABLE_OR_ENCRYPTED');}
  check(pdf.getPageCount()>0&&id(p.id),'INVALID_PDF_INPUT');documents.push({id:p.id,sha256:sha256(p.bytes),pages:pdf.getPageCount()});
 }
 const nativeConfig={...config,nativeDocument:false};
 const request={model:config.model,service_tier:config.service_tier,reasoning:config.reasoning,store:false,prompt_cache_options:{mode:'explicit'},max_output_tokens:config.max_output_tokens,
  instructions:require('./openai.cjs').systemInstructions(require('../config/native-reading-instructions.json').join('\n')),
  text:{format:{type:'json_schema',name:'native_document_reading',strict:true,schema:readingSchema}},
  input:[{role:'user',content:[{type:'input_text',text:JSON.stringify({workflowVersion,editingContractSha256:digest(instructions),task,documents,template:{filename:'modelo-original.docx',sha256:template.hash},templateText:[...template.parts].filter(([p])=>story.test(p)).map(([part,b])=>({part,paragraphs:paragraphsOf(b.toString()).map(p=>p.text)}))})},
   {type:'input_file',filename:'modelo-original.docx',file_data:'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,'+templateBytes.toString('base64')},
   ...pdfs.map(p=>({type:'input_file',filename:p.id+'.pdf',detail:'high',file_data:'data:application/pdf;base64,'+p.bytes.toString('base64')}))]}]};
 return {production:true,nativeDocument:true,stagedNative:true,workflowVersion,stage:'reading',editingInstructions:instructions,templateTexts:[...template.parts].filter(([p])=>story.test(p)).flatMap(([,b])=>paragraphsOf(b.toString()).map(p=>p.text)),request,requestSha256:digest(request),templateBase64:templateBytes.toString('base64'),model,task,documents,config:nativeConfig,mode:'live'};
}
function validateReading(prepared,raw){
 check(raw?.status==='completed','OPENAI_RESPONSE_INCOMPLETE');
 const content=(raw.output||[]).filter(o=>o.type==='message').flatMap(o=>o.content||[]);check(!content.some(c=>c.type==='refusal'),'OPENAI_REFUSAL');
 let value;try{value=JSON.parse(content.filter(c=>c.type==='output_text').map(c=>c.text).join(''));}catch{check(false,'NATIVE_READING_INVALID');}
 check(require('./generated/native-reading-validator.cjs')(value),'NATIVE_READING_SCHEMA_INVALID');
 check(value.templateSha256===sha256(Buffer.from(prepared.templateBase64,'base64')),'NATIVE_READING_TEMPLATE_MISMATCH');
 const pagesValid=items=>{const keys=new Set();for(const item of items){const doc=prepared.documents.find(d=>d.id===item.documentId);check(doc&&item.pages.length>0,'NATIVE_READING_PAGE_INVALID');for(const page of item.pages){const key=item.documentId+':'+page;check(page<=doc.pages&&!keys.has(key),'NATIVE_READING_PAGE_INVALID');keys.add(key);}}return keys;};
 const allPages=items=>{const keys=pagesValid(items);return prepared.documents.every(d=>Array.from({length:d.pages},(_,i)=>keys.has(d.id+':'+(i+1))).every(Boolean));};
 check(allPages(value.reviewedPages),'NATIVE_READING_PAGES_NOT_REVIEWED');
 const entities=new Set(value.entities.map(e=>e.id));check(entities.size===value.entities.length&&entities.size>0&&value.entities.every(e=>referenceId(e.id)&&e.relatedEntityIds.every(r=>entities.has(r))),'NATIVE_READING_ENTITIES_INVALID');
 const fields=new Set();check(value.fields.length>0,'NATIVE_READING_FIELDS_REQUIRED');
 for(const field of value.fields){
  check(referenceId(field.id)&&!fields.has(field.id)&&entities.has(field.entityId)&&field.label.trim(),'NATIVE_READING_FIELD_INVALID');fields.add(field.id);pagesValid(field.checkedPages);
  check(prepared.templateTexts.some(t=>t.includes(field.oldValue)),'NATIVE_READING_OLD_VALUE_NOT_IN_TEMPLATE');
  for(const source of field.sources){const doc=prepared.documents.find(d=>d.id===source.documentId);check(doc&&source.page<=doc.pages&&source.quote.trim(),'NATIVE_READING_SOURCE_INVALID');}
  if(field.status==='verified')check(typeof field.value==='string'&&field.value.trim()&&field.sources.length>0,'NATIVE_VALUE_WITHOUT_SOURCE');
  else {check(field.value===null&&field.rechecked&&field.checkedPages.length>0,'NATIVE_PENDING_NOT_RECHECKED');if(field.status==='absent')check(allPages(field.checkedPages),'NATIVE_ABSENCE_NOT_CONFIRMED');}
 }
 return value;
}
function prepareNativeEditing(prepared,reading){
 const config={...prepared.config,nativeDocument:true,codeInterpreterMemory:prepared.config.codeInterpreterMemory||'1g',codeInterpreterSessionUsd:prepared.config.codeInterpreterSessionUsd??0.03};
 const {model,service_tier,reasoning,store,prompt_cache_options,max_output_tokens}=prepared.request;
 const request={model,service_tier,reasoning,store,prompt_cache_options,max_output_tokens,instructions:require('./openai.cjs').systemInstructions(prepared.editingInstructions),
  tools:[{type:'code_interpreter',container:{type:'auto',memory_limit:config.codeInterpreterMemory,network_policy:{type:'disabled'}}}],tool_choice:'required',
  input:[{role:'user',content:[{type:'input_text',text:JSON.stringify({workflowVersion,task:{issueId:prepared.task.issueId,templateId:prepared.model.id},documents:prepared.documents,readingFile:'leitura-validada.json',outputFilename:'minuta-final.docx'})},
   {type:'input_file',filename:'modelo-original.docx',file_data:'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,'+prepared.templateBase64},
   {type:'input_file',filename:'leitura-validada.json',file_data:'data:application/json;base64,'+Buffer.from(JSON.stringify({...reading,fields:reading.fields.map(f=>({...f,editValue:editingValue(f),markRed:f.status!=='verified'}))})).toString('base64')}]}]};
 return {...prepared,stage:'editing',parentRequestSha256:prepared.requestSha256,request,requestSha256:digest(request),config};
}
function pendingMarker(field){return '[PENDENTE: '+field.label.replace(/[\[\]\r\n\t]/g,' ').trim().slice(0,40).toUpperCase()+']';}
function editingValue(field){return field.status==='verified'?field.value:pendingMarker(field);}
function pendingNotes(reading){
 const reasons={not_located:'nao localizado apos reconferencia; revisar',absent:'ausente apos reconferencia documental',conflict:'documentos divergentes',illegible:'leitura ilegivel'};
 return [...new Set(reading.fields.filter(f=>f.status!=='verified').map(f=>{const entity=reading.entities.find(e=>e.id===f.entityId);return (entity?.label||entity?.role||'Parte')+' — '+f.label.replace(/[\r\n\t]/g,' ')+': '+reasons[f.status];}))];
}
async function downloadNativeDocument(raw,{apiKey,transport}){
 check(raw?.status==='completed','OPENAI_RESPONSE_INCOMPLETE');
 const messages=(raw.output||[]).filter(o=>o.type==='message').flatMap(o=>o.content||[]);
 check(!messages.some(c=>c.type==='refusal'),'OPENAI_REFUSAL');
 const containers=[...new Set((raw.output||[]).filter(o=>o.type==='code_interpreter_call'&&o.status==='completed').map(o=>o.container_id))];
 check(containers.length>0&&containers.every(id),'NATIVE_TOOL_NOT_EXECUTED');
 let files=messages.flatMap(c=>c.annotations||[]).filter(a=>a.type==='container_file_citation'&&/(?:^|\/)minuta-final\.docx$/i.test(a.filename||'')&&containers.includes(a.container_id)&&id(a.file_id));
 // O arquivo e recuperado por IDs oficiais, nunca por uma URL escrita pela IA.
 if(!files.length)for(const container of containers){
  const r=await transport('https://api.openai.com/v1/containers/'+container+'/files?limit=100',{headers:{Authorization:'Bearer '+apiKey},maxBytes:2097152});check(r.status===200,'NATIVE_CONTAINER_FILES_UNAVAILABLE');
  let list;try{list=JSON.parse(r.bytes.toString());}catch{check(false,'NATIVE_CONTAINER_FILES_INVALID');}
  files.push(...(list.data||[]).filter(f=>f.source==='assistant'&&/(?:^|\/)minuta-final\.docx$/i.test(f.path||'')&&id(f.id)).map(f=>({container_id:container,file_id:f.id})));
 }
 files=[...new Map(files.map(f=>[f.container_id+':'+f.file_id,f])).values()];check(files.length===1,'NATIVE_FINAL_DOCX_NOT_UNIQUE');
 const f=files[0],r=await transport('https://api.openai.com/v1/containers/'+f.container_id+'/files/'+f.file_id+'/content',{headers:{Authorization:'Bearer '+apiKey},maxBytes:20971520,timeoutMs:60000});
 check(r.status===200,'NATIVE_FINAL_DOCX_DOWNLOAD_FAILED');return r.bytes;
}
function canonical(n,withoutColor=false){
 if(!n)return null;if(withoutColor&&n.uri===W&&n.local==='color')return null;
 const attrs=Object.values(n.attrs).filter(a=>a.uri!=='http://www.w3.org/2000/xmlns/').map(a=>[a.uri+'|'+a.local,a.value]).sort((a,b)=>a[0].localeCompare(b[0]));
 return [n.uri+'|'+n.local,attrs,n.text.trim(),n.children.map(c=>canonical(c,withoutColor)).filter(Boolean)];
}
function collapse(items){return items.filter((x,i)=>x!=='TEXT'||items[i-1]!=='TEXT');}
function skeleton(n){
 if(n.uri===W&&n.local==='rPr'&&n.parent?.local==='r')return [];
 if(n.uri===W&&n.local==='t')return ['TEXT'];
 const children=collapse(n.children.flatMap(skeleton));
 if(n.uri===W&&n.local==='r')return children;
 const own=canonical(n);return [[own[0],own[1],own[2],children]];
}
function paragraphOf(n){for(let p=n.parent;p;p=p.parent)if(p.uri===W&&p.local==='p')return p;return null;}
function markerRuns(parsed,p,markers){
 const runs=parsed.nodes.filter(n=>n.uri===W&&n.local==='r'&&paragraphOf(n)===p),allowed=new Set(),counts=new Map();let group=[];
 const finish=()=>{if(!group.length)return;const text=group.map(r=>r.children.filter(n=>n.uri===W&&n.local==='t').map(n=>n.text).join('')).join(''),found=text.match(/\[PENDENTE: [^\[\]\r\n]+\]/gu)||[];
  if(found.length&&found.join('')===text&&found.every(m=>markers.includes(m))){for(const r of group)allowed.add(r);for(const marker of found)counts.set(marker,(counts.get(marker)||0)+1);}group=[];
 };
 for(const run of runs){const red=run.children.find(n=>n.uri===W&&n.local==='rPr')?.children.some(n=>n.uri===W&&n.local==='color'&&Object.values(n.attrs).some(a=>a.local==='val'&&a.value.toUpperCase()==='FF0000'));if(red)group.push(run);else finish();}finish();return {allowed,counts};
}
function runStyles(parsed,p,markers=[]){
 const runs=parsed.nodes.filter(n=>n.uri===W&&n.local==='r'&&paragraphOf(n)===p),values=[];
 const {allowed}=markerRuns(parsed,p,markers);
 for(let i=0;i<runs.length;i++){
  const run=runs[i],rp=run.children.find(n=>n.uri===W&&n.local==='rPr');
  const key=node=>node&&node[1].length===0&&node[2]===''&&node[3].length===0?'null':JSON.stringify(node);
  const style={full:key(rp?canonical(rp):null),base:key(rp?canonical(rp,true):null),pending:allowed.has(run)};
  if(values.at(-1)?.full!==style.full||values.at(-1)?.pending!==style.pending)values.push(style);
 }
 return values;
}
// Aceita somente diferencas de texto compostas dos valores autorizados. Esta
// verificacao nao localiza campos para edicao e nunca modifica o arquivo.
function authorizedTextChange(before,after,pairs){
 const byFirst=new Map();for(const [oldValue,value] of pairs){if(oldValue===value)continue;const key=oldValue[0];if(!byFirst.has(key))byFirst.set(key,[]);byFirst.get(key).push([oldValue,value]);}
 const stack=[[0,0,new Map()]],seen=new Set();
 while(stack.length){let [i,j,used]=stack.pop();const key=i+':'+j;if(seen.has(key))continue;seen.add(key);check(seen.size<=250000,'NATIVE_TEXT_CHANGE_AMBIGUOUS');
  for(;;){if(i===before.length&&j===after.length)return used;
   const candidates=(byFirst.get(before[i])||[]).filter(([oldValue,value])=>before.startsWith(oldValue,i)&&after.startsWith(value,j));
   for(const [oldValue,value] of candidates){const next=new Map(used),pair=JSON.stringify([oldValue,value]);next.set(pair,(next.get(pair)||0)+1);stack.push([i+oldValue.length,j+value.length,next]);}
   if(i>=before.length||j>=after.length||before[i]!==after[j])break;i++;j++;
  }
 }return false;
}
async function auditNativeDocument(templateBytes,bytes,reading){
 const before=await openDocx(templateBytes),after=await openDocx(bytes),changed=[],applied=new Map();
 const markers=reading?reading.fields.filter(f=>f.status!=='verified').map(pendingMarker):[],coloredMarkers=new Map();
 check(before.parts.size===after.parts.size&&[...before.parts.keys()].every(p=>after.parts.has(p)),'NATIVE_PACKAGE_PARTS_CHANGED');
 for(const [part,source] of before.parts){const output=after.parts.get(part);if(source.equals(output))continue;
  check(story.test(part),'NATIVE_FIXED_PART_CHANGED');
  const a=parseXml(source.toString()),b=parseXml(output.toString());
  check(digest(skeleton(a.nodes[0]))===digest(skeleton(b.nodes[0])),'NATIVE_DOCUMENT_STRUCTURE_CHANGED');
  check(a.paragraphs.length===b.paragraphs.length,'NATIVE_PARAGRAPHS_CHANGED');
  for(let i=0;i<a.paragraphs.length;i++){
   const expected=runStyles(a,a.paragraphs[i]),actual=runStyles(b,b.paragraphs[i],markers);let states=new Set([-1]);
   for(const style of actual){const next=new Set();for(const at of states)for(const index of [at,at+1])if(index>=0&&index<expected.length&&(style.pending?style.base===expected[index].base:style.full===expected[index].full))next.add(index);states=next;}
   check(states.has(expected.length-1),'NATIVE_TEXT_FORMAT_CHANGED');
   for(const [marker,count] of markerRuns(b,b.paragraphs[i],markers).counts)coloredMarkers.set(marker,(coloredMarkers.get(marker)||0)+count);
  }
  if(reading){const sourceText=paragraphsOf(source.toString()),outputText=paragraphsOf(output.toString()),pairs=reading.fields.map(f=>[f.oldValue,editingValue(f)]);
   for(let i=0;i<sourceText.length;i++){const used=authorizedTextChange(sourceText[i].text,outputText[i].text,pairs);check(used,'NATIVE_UNAUTHORIZED_TEXT_CHANGE');for(const [key,count] of used)applied.set(key,(applied.get(key)||0)+count);}
  }
  changed.push(part);
 }
 check(changed.length>0||reading&&reading.fields.every(f=>f.status==='verified'&&f.oldValue===f.value),'NATIVE_DOCUMENT_UNCHANGED');
 if(reading){const outputTexts=[...after.parts].filter(([p])=>story.test(p)).flatMap(([,b])=>paragraphsOf(b.toString()).map(p=>p.text));
  for(const field of reading.fields.filter(f=>f.status==='verified'))check(outputTexts.some(t=>t.includes(field.value)),'NATIVE_VERIFIED_VALUE_NOT_APPLIED');
  const required=new Map(),requiredMarkers=new Map();for(const field of reading.fields){const value=editingValue(field);if(field.status!=='verified'){if(!requiredMarkers.has(value))requiredMarkers.set(value,new Set());requiredMarkers.get(value).add(field.entityId+':'+field.label);}if(value===field.oldValue)continue;const pair=JSON.stringify([field.oldValue,value]);if(!required.has(pair))required.set(pair,new Set());required.get(pair).add(field.entityId+':'+field.label);}
  for(const [pair,fields] of required)check((applied.get(pair)||0)>=fields.size,'NATIVE_VALUE_NOT_APPLIED');
  for(const [marker,fields] of requiredMarkers)check((coloredMarkers.get(marker)||0)>=fields.size,'NATIVE_PENDING_MARKER_NOT_RED');
 }
 const notes=reading?pendingNotes(reading):[];
 return {bytes,report:{status:notes.length?'native-ai-draft-with-pending':'native-ai-docx',documentGenerated:true,hasPending:notes.length>0,pendingNotes:notes,simulation:false,evaluation:{expectedPending:[]},audit:{templateSha256:before.hash,outputSha256:after.hash,changedParts:changed,unchangedParts:before.parts.size-changed.length,originalMediaPreserved:true,structureVerified:true,runFormattingVerified:true,authorizedValuesVerified:!!reading,pendingLocation:'red-markers-and-linear-issue',localTextEdits:0,semanticContentRequiresHumanReview:true}}};
}
module.exports={workflowVersion,prepareNativeDocument,validateReading,prepareNativeEditing,pendingMarker,pendingNotes,downloadNativeDocument,auditNativeDocument};
