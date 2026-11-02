'use strict';
const JSZip=require('jszip');
const {sha256,openDocx,paragraphsOf}=require('../src/docx.cjs');
const {coverageOf}=require('../src/operation.cjs');
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
async function semanticFixture({repeat=false,unmapped=false,auxiliary=false}={}){
  const paragraphs=['Vendedor: Antigo Vendedor Teste; CPF: CPF-ANTIGO-V01; profissão: profissão antiga.','Comprador: Antiga Compradora Teste; CPF: CPF-ANTIGO-C01.','Escrevente: Funcionário Fixo Teste. Cláusula jurídica fixa.'];
  if(repeat||unmapped)paragraphs.push('Assinatura: Antigo Vendedor Teste.');
  const z=new JSZip();
  z.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',{createFolders:false});
  z.file('_rels/.rels','<Relationships/>',{createFolders:false});
  z.file('word/styles.xml',`<w:styles xmlns:w="${W}"/>`,{createFolders:false});
  z.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${paragraphs.map(p=>`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${p}</w:t></w:r></w:p>`).join('')}</w:body></w:document>`,{createFolders:false});
  if(auxiliary)z.file('word/settings.xml',`<w:settings xmlns:w="${W}"><w:docVars><w:docVar w:name="old" w:val="Antigo Vendedor Teste"/></w:docVars></w:settings>`,{createFolders:false});
  const bytes=await z.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
  const defs=[['seller.name','seller','nome',0,'Antigo Vendedor Teste','Novo Vendedor Inventado'],['seller.cpf','seller','cpf',0,'CPF-ANTIGO-V01','CPF-NOVO-V11'],['seller.profession','seller','profissao',0,'profissão antiga','profissão nova'],['buyer.name','buyer','nome',1,'Antiga Compradora Teste','Nova Compradora Inventada'],['buyer.cpf','buyer','cpf',1,'CPF-ANTIGO-C01','CPF-NOVO-C11']];
  if(repeat)defs.push(['seller.signature','seller','nome',3,'Antigo Vendedor Teste','Novo Vendedor Inventado']);
  const fields=defs.map(([id,entityId,property,paragraph,oldText])=>{const text=paragraphs[paragraph],start=text.indexOf(oldText);return {id,entityId,property,part:'word/document.xml',paragraph,start,end:start+oldText.length,oldText,required:true};});
  const entities=[{id:'seller',kind:'person',role:'seller',spouseId:null},{id:'buyer',kind:'person',role:'buyer',spouseId:null},{id:'office',kind:'institution',role:'staff',spouseId:null}];
  const profile={schemaVersion:2,scope:'synthetic-only',templateSha256:sha256(bytes),anonymizationApproved:true,externalSharingAllowed:true,operation:{inventoryReviewed:true,coverage:coverageOf(await openDocx(bytes),paragraphsOf),entities},fields};
  const pdf=Buffer.from('%PDF-1.4\n% fixture for request assembly only\n%%EOF');
  const documents=[{id:'ficticio',sha256:sha256(pdf),pages:2}];
  const plan={scope:'synthetic-only',templateSha256:profile.templateSha256,pending:[],entities:entities.filter(e=>e.kind!=='institution').map(e=>({...e,slotId:e.id,id:'new-'+e.id})),replacements:defs.map(([fieldId,entityId,property,paragraph,oldText,value])=>({fieldId,entityId:'new-'+entityId,property,value,confirmedUnchanged:false,source:{documentId:'ficticio',documentSha256:documents[0].sha256,page:1,entityId:'new-'+entityId,property,quote:value}}))};
  const approvals=[{kind:'template',sha256:profile.templateSha256,approved:true,externalSharingAllowed:true,origin:'generated-synthetic'},{kind:'pdf',sha256:documents[0].sha256,pages:2,approved:true,externalSharingAllowed:true,origin:'generated-synthetic'}];
  const extraction={schemaVersion:2,templateSha256:profile.templateSha256,documents,documentSegments:[{documentId:'ficticio',type:'identity',startPage:1,endPage:2,entityIds:plan.entities.map(e=>e.id)}],entities:plan.entities,observations:plan.replacements.filter((r,i,a)=>a.findIndex(x=>x.entityId===r.entityId&&x.property===r.property)===i).map(r=>({entityId:r.entityId,property:r.property,value:r.value,status:'verified',sources:[r.source]})),replacements:plan.replacements.map(r=>{const f=fields.find(f=>f.id===r.fieldId);return {fieldId:r.fieldId,entityId:r.entityId,property:r.property,oldValue:f.oldText,value:r.value,part:f.part,paragraph:f.paragraph,start:f.start,end:f.end,status:'verified',confirmedUnchanged:false,sources:[r.source]};}),pending:[]};
  return {bytes,profile,plan,pdf,documents,approvals,extraction};
}
module.exports={semanticFixture};
