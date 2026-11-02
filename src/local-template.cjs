'use strict';
// Caminho estritamente local: nao importa cliente OpenAI nem constroi requisicoes.
// Modelos reais continuam recusados por src/openai.cjs e src/increment3.cjs.
const {openDocx,paragraphsOf,parseXml,sha256,privacyFindings,textRangeSafe}=require('./docx.cjs');
const {patchText}=require('./anonymize.cjs');
const {validateOperation,validateOperationResult}=require('./operation.cjs');
const {extractResponse,validateExtraction}=require('./extraction.cjs');
const {compareGroundTruth}=require('./ground-truth.cjs');
const {requireCondition:check}=require('./errors.cjs');
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xmlText=s=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
function pendingFormatting(xml,fields,notes=[]){
 const paragraphs=paragraphsOf(xml),runs=new Map();
 for(const f of fields){const p=paragraphs[f.paragraph];check(p&&p.text.slice(f.start,f.end)===f.oldText&&f.oldText.startsWith('pendente: '),'PENDING_MARKER_LOCATION_INVALID');
  for(const t of p.texts.filter(t=>t.from<f.end&&t.to>f.start)){let run=t.parent;while(run&&!(run.uri===W&&run.local==='r'))run=run.parent;check(run,'PENDING_RUN_REQUIRED');if(!runs.has(run.start))runs.set(run.start,{run,p,ranges:[]});runs.get(run.start).ranges.push(f);}
 }
 const patches=[];
 for(const {run,p,ranges} of runs.values()){
  const prefix=run.name.split(':')[0]+':';check(run.name.includes(':'),'PENDING_STYLE_NAMESPACE_REQUIRED');
  const rp=run.children.find(n=>n.uri===W&&n.local==='rPr'),original=rp?xml.slice(rp.start,rp.end):'',color='<'+prefix+'color '+prefix+'val="FF0000"/>';
  let red=original;
  if(rp){for(const n of rp.children.filter(n=>n.uri===W&&n.local==='color').sort((a,b)=>b.start-a.start))red=red.slice(0,n.start-rp.start)+red.slice(n.end-rp.start);
   red=rp.selfClosing?red.replace(/\/>$/, '>'+color+'</'+rp.name+'>'):red.slice(0,red.lastIndexOf('</'))+color+red.slice(red.lastIndexOf('</'));
  }else red='<'+prefix+'rPr>'+color+'</'+prefix+'rPr>';
  const open=xml.slice(run.start,run.openEnd),close=xml.slice(run.contentEnd,run.end);let replacement='';
  for(const child of run.children){if(child===rp)continue;
   const t=p.texts.find(t=>t.start===child.start);
   if(!t){replacement+=open+original+xml.slice(child.start,child.end)+close;continue;}
   const cuts=[...new Set([t.from,t.to,...ranges.flatMap(f=>[Math.max(t.from,Math.min(t.to,f.start)),Math.max(t.from,Math.min(t.to,f.end))])])].sort((a,b)=>a-b);
   if(cuts.length===1){replacement+=open+original+xml.slice(t.start,t.end)+close;continue;}
   for(let i=0;i<cuts.length-1;i++){const from=cuts[i],to=cuts[i+1];if(from===to)continue;const isPending=ranges.some(f=>f.start<=from&&f.end>=to),text=t.text.slice(from-t.from,to-t.from);
    let tag=xml.slice(t.start,t.openEnd);if(!/xml:space=/.test(tag))tag=tag.replace(/>$/,' xml:space="preserve">');
    replacement+=open+(isPending?red:original)+tag+xmlText(text)+'</'+t.name+'>'+close;
   }
  }
  patches.push({start:run.start,end:run.end,value:replacement});
 }
 for(const patch of patches.sort((a,b)=>b.start-a.start))xml=xml.slice(0,patch.start)+patch.value+xml.slice(patch.end);
 const styled=paragraphsOf(xml);check(styled.length===paragraphs.length&&styled.every((p,i)=>p.text===paragraphs[i].text),'PENDING_FORMAT_CHANGED_TEXT');
 if(notes.length){const body=parseXml(xml).nodes.find(n=>n.uri===W&&n.local==='body');check(body,'PENDING_NOTE_BODY_REQUIRED');const at=body.children.find(n=>n.uri===W&&n.local==='sectPr')?.start||body.contentEnd;
  const paragraph=paragraphsOf(xml).find(p=>p.texts.length),run=paragraph?.texts[0]?.parent,rp=run?.children.find(n=>n.uri===W&&n.local==='rPr');
  const font=rp?rp.children.filter(n=>n.uri===W&&['rFonts','sz','szCs'].includes(n.local)).map(n=>xml.slice(n.start,n.end)).join(''):'';
  const noteXml=notes.map(text=>'<w:p><w:r><w:rPr>'+font+'<w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">'+xmlText(text)+'</w:t></w:r></w:p>').join('');xml=xml.slice(0,at)+noteXml+xml.slice(at);
 }
 return xml;
}
function assertLocalProfile(profile,hash){
 check(profile?.scope==='local-authorized-template'&&profile.origin==='user-supplied-local'&&profile.localUse?.approved===true&&profile.localUse.convertedSha256===hash&&/^[a-f0-9]{64}$/.test(profile.localUse.originalSha256||'')&&profile.templateSha256===hash&&profile.operation?.inventoryReviewed===true&&profile.externalSharingAllowed===false&&profile.anonymizationApproved===false,'LOCAL_TEMPLATE_USE_NOT_APPROVED');
}
async function applyLocal(bytes,profile,plan){
 const doc=await openDocx(bytes);assertLocalProfile(profile,doc.hash);
 check(plan.scope==='synthetic-only'&&plan.templateSha256===doc.hash&&plan.pending?.length===0,'LOCAL_PLAN_INVALID');
 const byId=new Map(plan.replacements.map(r=>[r.fieldId,r]));
 check(byId.size===profile.fields.length&&byId.size===plan.replacements.length,'INCOMPLETE_VARIABLE_COVERAGE');
 for(const f of profile.fields){
  const r=byId.get(f.id),p=doc.parts.has(f.part)&&paragraphsOf(doc.parts.get(f.part).toString())[f.paragraph];
  check(/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(f.part)&&textRangeSafe(p,f.start,f.end)&&p.text.slice(f.start,f.end)===f.oldText,'ANCHOR_MISMATCH');
  check(r&&typeof r.value==='string'&&r.value.trim().length>0&&r.value.length<=2000&&!/[\u0000-\u001f\u007f]/u.test(r.value)&&!/[\uD800-\uDFFF]/u.test(r.value.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g,'')),'INVALID_REPLACEMENT');
  const note=(plan.pendingAnnotations||[]).find(n=>n.fieldId===f.id);
  if(note)check(r.source?.kind==='pending'&&r.value===note.text&&r.value.startsWith('pendente: ')&&note.entityId===r.entityId&&note.property===r.property,'INVALID_PENDING_MARKER');
  else check(r.source&&typeof r.source.documentId==='string'&&/^[A-Za-z0-9_-]{1,80}$/.test(r.source.documentId)&&Number.isInteger(r.source.page)&&r.source.page>0,'SOURCE_REQUIRED');
 }
 for(const [part,data] of doc.parts){
  check(!privacyFindings({parts:new Map([[part,data]])}).some(f=>['ACTIVE_OR_EMBEDDED_CONTENT','TRACKED_CHANGES','HIDDEN_TEXT'].includes(f)),'DOCUMENT_REQUIRES_PRIVACY_REVIEW');
  check(profile.localUse.preservedParts.some(p=>p.part===part&&p.sha256===sha256(data)),'LOCAL_TEMPLATE_PART_NOT_REVIEWED');
 }
 const operation=validateOperation(doc,profile,plan,paragraphsOf),expected=new Map(doc.parts),changed=[];
 for(const [part,data] of doc.parts){
  const fields=profile.fields.filter(f=>f.part===part);if(!fields.length)continue;
  const patched=patchText(data.toString(),fields.map(f=>({...f,value:byId.get(f.id).value}))).xml;
  const output=Buffer.from(patched);expected.set(part,output);doc.zip.file(part,output);changed.push(part);
 }
 let output=await doc.zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'}),after=await openDocx(output);
 check(after.parts.size===doc.parts.size&&[...expected].every(([p,b])=>after.parts.get(p)?.equals(b)),'UNAUTHORIZED_PACKAGE_CHANGE');
 validateOperationResult(doc,after,profile,plan,paragraphsOf,parseXml);
 const markers=plan.pendingAnnotations||[],notes=plan.pendingNotes||[];
 if(markers.length||notes.length){const styledExpected=new Map(after.parts);
  for(const [part,data] of after.parts){const ids=new Set(markers.map(n=>n.fieldId)),partFields=profile.fields.filter(f=>f.part===part);const final=partFields.length?patchText(doc.parts.get(part).toString(),partFields.map(f=>({...f,value:byId.get(f.id).value}))).fields.filter(f=>ids.has(f.id)):[];
   if(!final.length&&!(notes.length&&part==='word/document.xml'))continue;
   const styled=Buffer.from(pendingFormatting(data.toString(),final,part==='word/document.xml'?notes:[]));styledExpected.set(part,styled);after.zip.file(part,styled);
  }
  output=await after.zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});after=await openDocx(output);
  check(after.parts.size===doc.parts.size&&[...styledExpected].every(([part,b])=>after.parts.get(part)?.equals(b)),'UNAUTHORIZED_PENDING_FORMAT_CHANGE');
 }
 return {bytes:output,audit:{scope:'local-authorized-template',templateSha256:doc.hash,originalSha256:profile.localUse.originalSha256,outputSha256:after.hash,changedParts:changed,unchangedParts:doc.parts.size-changed.length,operation,originalMediaPreserved:true,pendingMarkers:markers.length,pendingNotes:notes.length,externalApiCalls:0,requiresHumanLegalReview:true}};
}
function prepareLocal(payload,fixture,config){
 check(config.openaiMode==='mock'&&config.paidApiEnabled===false,'REGISTERED_TEMPLATE_LIVE_DISABLED');
 const bytes=Buffer.from(payload.templateBase64,'base64'),pdf=Buffer.from(payload.pdfs[0].base64,'base64');
 assertLocalProfile(payload.profile,sha256(bytes));
 check(fixture.localOnly===true&&fixture.templateSha256===sha256(bytes)&&fixture.pdfSha256===sha256(pdf)&&fixture.profileSha256===sha256(Buffer.from(JSON.stringify(payload.profile)))&&fixture.mockResponseSha256===sha256(Buffer.from(JSON.stringify(payload.mockResponse)))&&fixture.truthSha256===sha256(Buffer.from(JSON.stringify(fixture.truth))),'LOCAL_FIXTURE_INTEGRITY_FAILED');
 check(fixture.documents.length===1&&fixture.documents[0].sha256===sha256(pdf)&&fixture.generatedFictionalPdf===true,'FICTIONAL_INPUT_NOT_APPROVED');
 return {profile:payload.profile,templateBase64:payload.templateBase64,truth:fixture.truth,documents:fixture.documents,mockResponse:payload.mockResponse,mode:'mock',scenario:fixture.scenario};
}
async function evaluateLocal(prepared,raw){
 check(prepared.mode==='mock','REGISTERED_TEMPLATE_LIVE_DISABLED');
 const value=extractResponse(raw),evaluation=compareGroundTruth(value,prepared.truth,prepared.documents);
 let plan;try{plan=validateExtraction(value,prepared.profile,prepared.documents);}catch(e){if(e.code==='PENDING_INFORMATION'&&!evaluation.documentExpected)return {report:{status:'expected-block',evaluation,documentGenerated:false,apiCalls:0,simulation:true}};throw e;}
 for(const c of prepared.profile.operation.constraints||[]){
  const e=value.entities.find(e=>e.slotId===c.entityId),o=e&&value.observations.find(o=>o.entityId===e.id&&o.property===c.property);
  check(o?.status==='verified'&&o.value===c.value,'PARTICIPANT_CONSTRAINT_MISMATCH');
 }
 const result=await applyLocal(Buffer.from(prepared.templateBase64,'base64'),prepared.profile,plan);
 return {bytes:result.bytes,report:{status:'validated-local-template-docx',evaluation,audit:{...result.audit,...(prepared.dynamicAudit?{dynamicAnalysis:prepared.dynamicAudit}:{})},documentGenerated:true,apiCalls:0,simulation:true}};
}
module.exports={assertLocalProfile,applyLocal,prepareLocal,evaluateLocal,pendingFormatting};
