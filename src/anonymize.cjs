'use strict';
// Preparacao local de uma copia de desenvolvimento. Nao e certificado de
// anonimizacao nem libera compartilhamento. Nao contem clientes de rede.
const {openDocx,parseXml,paragraphsOf,sha256}=require('./docx.cjs');
const {requireCondition:check}=require('./errors.cjs');
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const storyPart=/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/;
const escape=s=>s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

function patchText(xml,changes){
  const paragraphs=paragraphsOf(xml), updates=new Map();
  const byParagraph=new Map();
  for(const change of changes){
    const p=paragraphs[change.paragraph];
    check(p&&Number.isInteger(change.start)&&Number.isInteger(change.end)&&change.start>=0&&change.end>change.start&&change.end<=p.text.length&&typeof change.value==='string','ANONYMIZATION_ANCHOR');
    check(!/[\u0000-\u001f]/u.test(change.value),'ANONYMIZATION_VALUE');
    if(!byParagraph.has(p.index))byParagraph.set(p.index,[]);
    const list=byParagraph.get(p.index);
    check(!list.some(c=>change.start<c.end&&change.end>c.start),'ANONYMIZATION_OVERLAP');list.push(change);
    const affected=p.texts.filter(t=>t.from<change.end&&t.to>change.start);
    check(affected.length&&!affected.some(t=>t.selfClosing),'ANONYMIZATION_TEXT_LOCATION');
    let remaining=change.value;
    affected.forEach((node,i)=>{
      const start=Math.max(0,change.start-node.from),end=Math.min(node.text.length,change.end-node.from);
      const value=i===affected.length-1?remaining:remaining.slice(0,end-start);remaining=remaining.slice(value.length);
      if(!updates.has(node.start))updates.set(node.start,{node,edits:[]});
      updates.get(node.start).edits.push({start,end,value});
    });
  }
  const patches=[];
  for(const {node,edits} of updates.values()){
    let value=node.text;
    for(const e of edits.sort((a,b)=>b.start-a.start))value=value.slice(0,e.start)+e.value+value.slice(e.end);
    patches.push({start:node.openEnd,end:node.contentEnd,value:escape(value)});
    if(value.trim()!==value&&!Object.values(node.attrs).some(a=>a.name==='xml:space'&&a.value==='preserve')){
      patches.push({start:node.openEnd-1,end:node.openEnd-1,value:' xml:space="preserve"'});
    }
  }
  let result=xml;
  for(const patch of patches.sort((a,b)=>b.start-a.start))result=result.slice(0,patch.start)+patch.value+result.slice(patch.end);
  const after=paragraphsOf(result);
  const finalFields=[];
  for(const p of paragraphs){
    const ordered=(byParagraph.get(p.index)||[]).sort((a,b)=>a.start-b.start);
    let expected=p.text;
    for(const c of [...ordered].reverse())expected=expected.slice(0,c.start)+c.value+expected.slice(c.end);
    check(after[p.index].text===expected,'ANONYMIZATION_TEXT_VALIDATION');
    let delta=0;
    for(const c of ordered){finalFields.push({id:c.id,paragraph:p.index,start:c.start+delta,end:c.start+delta+c.value.length,oldText:c.value});delta+=c.value.length-(c.end-c.start);}
  }
  // Nenhum texto fora dos intervalos e recriado.
  return {xml:result,fields:finalFields};
}

function removeNodes(xml,predicate){
  const nodes=parseXml(xml).nodes.filter(predicate);
  const chosen=nodes.filter(n=>!nodes.some(p=>p!==n&&p.start<=n.start&&p.end>=n.end));
  for(const n of chosen.sort((a,b)=>b.start-a.start))xml=xml.slice(0,n.start)+xml.slice(n.end);
  return xml;
}
function stripHyperlinks(xml){
  const patches=[];
  for(const n of parseXml(xml).nodes.filter(n=>n.uri===W&&n.local==='hyperlink')){
    if(n.selfClosing)patches.push({start:n.start,end:n.end});
    else patches.push({start:n.start,end:n.openEnd},{start:n.contentEnd,end:n.end});
  }
  for(const p of patches.sort((a,b)=>b.start-a.start))xml=xml.slice(0,p.start)+xml.slice(p.end);
  return xml;
}

async function sanitizePackage(doc,textChanges){
  const zip=doc.zip,removed=[],replaced=[],fields=[];
  const expected=new Map(doc.parts);
  for(const [name,bytes] of doc.parts){
    if(!/\.(xml|rels)$/.test(name))continue;
    let xml=bytes.toString();
    if(storyPart.test(name)){
      const changes=textChanges.filter(c=>c.part===name);
      const patched=patchText(xml,changes);xml=patched.xml;
      fields.push(...patched.fields.map(f=>({...f,part:name})));
    }
    // Somente valores pessoais de propriedades sao esvaziados. Nao remover
    // imagens, vinculos, custom XML, propriedades estruturais ou elementos.
    if(name==='docProps/core.xml'||name==='docProps/app.xml'){
      if(name==='docProps/core.xml')xml=removeNodes(xml,n=>['created','modified'].includes(n.local));
      const sensitive=name==='docProps/core.xml'?['creator','lastModifiedBy','title','subject','description','keywords']:['Manager'];
      const nodes=parseXml(xml).nodes.filter(n=>sensitive.includes(n.local)&&!n.selfClosing&&n.text);
      for(const n of nodes.sort((a,b)=>b.openEnd-a.openEnd))xml=xml.slice(0,n.openEnd)+xml.slice(n.contentEnd);
    }
    parseXml(xml);
    if(xml!==bytes.toString()){zip.file(name,xml,{createFolders:false});expected.set(name,Buffer.from(xml));replaced.push(name);}
  }
  const bytes=await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'}),result=await openDocx(bytes);
  check(result.parts.size===expected.size,'TEMPLATE_PARTS_CHANGED');
  for(const [part,data] of expected)check(result.parts.get(part)?.equals(data),'TEMPLATE_UNAUTHORIZED_CHANGE');
  return {bytes,doc:result,fields,removed,replaced,media:[...result.parts].filter(([n])=>n.startsWith('word/media/')).map(([part,b])=>({part,sha256:sha256(b)}))};
}
module.exports={patchText,removeNodes,stripHyperlinks,sanitizePackage};
