'use strict';
const {sha256}=require('./docx.cjs');
const {requireCondition:check}=require('./errors.cjs');
const uuid=value=>typeof value==='string'&&/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(value);
function documentFiles(content){
 check(typeof content==='string'&&content.length<=1048576,'TEMPLATE_DOCUMENT_CONTENT_INVALID');
 // Somente links de arquivos. O texto do recurso nunca e executado como instrucao.
 return [...content.matchAll(/\[([^\]\r\n]+\.docx?)\]\((https:\/\/[^\s)]+)\)/gi)].map(m=>({filename:m[1],url:m[2]}));
}
function modelIdentity(filename){
 const name=filename.replace(/\.docx?$/i,'').trim(),number=/^MODELO[\s_-]*(\d+)\b/i.exec(name);
 const slug=name.normalize('NFD').replace(/\p{M}/gu,'').toUpperCase().replace(/[^A-Z0-9]+/g,'-').replace(/^-|-$/g,'');
 return {id:number?'MODELO-'+number[1].padStart(2,'0'):'MINUTA-'+slug.slice(0,60)+'-'+sha256(Buffer.from(filename)).slice(0,8),name:name.replace(/^(?:MODELO|MINUTA)\b[\s_-]*(?:\d+)?[\s_-]*/i,''),format:/\.docx$/i.test(filename)?'docx':'doc'};
}
function selectModel(catalog,order){
 check(Array.isArray(catalog),'TEMPLATE_CATALOG_INVALID');
 const matches=catalog.filter(m=>m.id===order.templateId&&m.version===order.templateVersion&&m.actType===order.actType);
 check(matches.length===1,'EXACT_TEMPLATE_NOT_FOUND');return matches[0];
}
async function downloadModel(linear,model){
 let url;
 if(model.source?.kind==='linear-document'){
  check(model.userApprovedUse===true&&uuid(model.source.documentId),'TEMPLATE_NOT_APPROVED');
  const document=await linear.document(model.source.documentId);
  check(document.title===model.source.documentTitle,'TEMPLATE_DOCUMENT_TITLE_CHANGED');
  const files=documentFiles(document.content).filter(f=>f.filename===model.source.filename);
  check(files.length===1,'MODEL_DOCUMENT_FILE_NOT_UNIQUE');url=files[0].url;
 }else{
  check(model.origin==='generated-synthetic'&&uuid(model.issueId),'TEMPLATE_NOT_APPROVED');
  const issue=await linear.issue(model.issueId);
  const files=issue.attachments.nodes.filter(a=>a.id===model.attachmentId);
  check(files.length===1,'MODEL_ATTACHMENT_NOT_FOUND');url=files[0].url;
 }
 const bytes=await linear.download(url);check(sha256(bytes)===model.sha256,'MODEL_INTEGRITY_FAILED');
 return bytes;
}
function requireOperationalModel(model,bytes){
 if(model.source?.kind==='linear-document'){
  check(model.format==='docx'&&bytes.subarray(0,2).toString()==='PK','LEGACY_TEMPLATE_CONVERSION_REQUIRED');
  check(model.status==='ready'&&model.inventoryReviewed===true&&model.externalSharingAllowed===true,'TEMPLATE_FIELD_REVIEW_REQUIRED');
 }
}
module.exports={documentFiles,modelIdentity,selectModel,downloadModel,requireOperationalModel};
