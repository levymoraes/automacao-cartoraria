'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const JSZip=require('jszip');
const {sha256,openDocx,paragraphsOf,privacyFindings,applyAuthorized,checkZip}=require('../src/docx.cjs');
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const {reviewSynthetic}=require('./helpers.cjs');
const r=text=>`<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
async function sample({body,extra={}}={}){
  const zip=new JSZip();
  zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',{createFolders:false});
  zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',{createFolders:false});
  zip.file('word/styles.xml',`<w:styles xmlns:w="${W}"/>`,{createFolders:false});
  zip.file('word/document.xml',`<w:document xmlns:w="${W}"><w:body>${body||`<w:p>${r('Vendedor: Ana ')}${r('Teste')}${r('; cláusula fixa intacta.')}</w:p>`}</w:body></w:document>`,{createFolders:false});
  for(const [name,text] of Object.entries(extra))zip.file(name,text,{createFolders:false});
  const bytes=await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
  const profile={scope:'synthetic-only',templateSha256:sha256(bytes),fields:[{id:'nome',part:'word/document.xml',paragraph:0,start:10,end:19,oldText:'Ana Teste',required:true}]};
  const plan={scope:'synthetic-only',templateSha256:sha256(bytes),pending:[],replacements:[{fieldId:'nome',value:'Bruna Inventada',source:{documentId:'pacote-ficticio',page:1}}]};
  return reviewSynthetic(bytes,profile,plan);
}
const blocked=(fn,code)=>assert.rejects(fn,e=>e.code===code);

test('edita nome dividido entre runs e preserva cláusula e estilos byte a byte',async()=>{
  const a=await sample(),original=await openDocx(a.bytes);
  const out=await applyAuthorized(a.bytes,a.profile,a.plan),result=await openDocx(out.bytes);
  assert.equal(paragraphsOf(result.parts.get('word/document.xml').toString())[0].text,'Vendedor: Bruna Inventada; cláusula fixa intacta.');
  const mask=x=>x.replace(/(<w:t[^>]*>)[\s\S]*?(<\/w:t>)/g,'$1$2');
  assert.equal(mask(original.parts.get('word/document.xml').toString()),mask(result.parts.get('word/document.xml').toString()));
  for(const [name,bytes] of original.parts)if(name!=='word/document.xml')assert.deepEqual(result.parts.get(name),bytes);
  assert.equal(sha256(a.bytes),a.profile.templateSha256);
  assert.equal(out.audit.changes.length,1);assert.equal(out.audit.validation,'exact-package-comparison-passed');
  assert.ok(!JSON.stringify(out.audit).includes('Bruna'));assert.ok(!JSON.stringify(out.audit).includes('Ana'));
});

test('substituições menores preservam os elementos de todos os runs',async()=>{
  const a=await sample();a.plan.replacements[0].value='Bia';
  const out=await applyAuthorized(a.bytes,a.profile,a.plan),doc=await openDocx(out.bytes);
  assert.equal(paragraphsOf(doc.parts.get('word/document.xml').toString())[0].text,'Vendedor: Bia; cláusula fixa intacta.');
  assert.equal((doc.parts.get('word/document.xml').toString().match(/<w:r>/g)||[]).length,3);
});

test('escapa XML sem perder acentos ou conteúdo',async()=>{
  const a=await sample();a.plan.replacements[0].value='João & Léa <Teste>';
  const out=await applyAuthorized(a.bytes,a.profile,a.plan),doc=await openDocx(out.bytes);
  assert.ok(paragraphsOf(doc.parts.get('word/document.xml').toString())[0].text.includes('João & Léa <Teste>'));
});

test('bloqueia versão diferente do modelo',async()=>{
  const a=await sample();a.profile.templateSha256='0'.repeat(64);await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'TEMPLATE_HASH_MISMATCH');
});
test('bloqueia hash diferente no plano',async()=>{
  const a=await sample();a.plan.templateSha256='0'.repeat(64);await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'TEMPLATE_HASH_MISMATCH');
});
test('bloqueia uso fora do escopo fictício',async()=>{
  const a=await sample();a.plan.scope='production';await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'SYNTHETIC_ONLY');
});
for(const code of ['MISSING_REQUIRED','ILLEGIBLE','CONTRADICTION'])test(`bloqueia pendência ${code}`,async()=>{
  const a=await sample();a.plan.pending=[{code,fieldId:'nome'}];await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'PENDING_INFORMATION');
});
test('bloqueia lista de pendências não informada',async()=>{
  const a=await sample();delete a.plan.pending;await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'PENDING_INFORMATION');
});
test('bloqueia campo obrigatório ausente',async()=>{
  const a=await sample();a.plan.replacements=[];await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'REQUIRED_FIELD_MISSING');
});
test('bloqueia tentativa de alteração de cláusula não cadastrada',async()=>{
  const a=await sample();a.plan.replacements.push({...a.plan.replacements[0],fieldId:'clausula'});await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'UNKNOWN_OR_DUPLICATE_FIELD');
});
test('bloqueia campo duplicado no plano',async()=>{
  const a=await sample();a.plan.replacements.push(a.plan.replacements[0]);await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'UNKNOWN_OR_DUPLICATE_FIELD');
});
test('bloqueia perfil com identificadores duplicados',async()=>{
  const a=await sample();a.profile.fields.push(a.profile.fields[0]);await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'INVALID_PROFILE');
});
test('bloqueia valor vazio, controle XML e falta de fonte',async()=>{
  for(const value of ['', '\u0000', '\ud800']){
    const a=await sample();a.plan.replacements[0].value=value;await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'INVALID_REPLACEMENT');
  }
  const a=await sample();delete a.plan.replacements[0].source;await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'SOURCE_REQUIRED');
});
test('bloqueia página inválida',async()=>{
  const a=await sample();a.plan.replacements[0].source.page=0;await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'SOURCE_REQUIRED');
});
test('bloqueia âncora divergente',async()=>{
  const a=await sample();a.profile.fields[0].start++;await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'ANCHOR_MISMATCH');
});
test('bloqueia ocorrência repetida no mesmo parágrafo',async()=>{
  const a=await sample({body:`<w:p>${r('Vendedor: Ana Teste e Ana Teste')}</w:p>`});await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'AMBIGUOUS_ANCHOR');
});
test('bloqueia intervalos sobrepostos',async()=>{
  const a=await sample();a.profile.fields.push({...a.profile.fields[0],id:'outro'});a.plan.replacements.push({...a.plan.replacements[0],fieldId:'outro'});await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'OVERLAPPING_REPLACEMENTS');
});
test('bloqueia parte não autorizada',async()=>{
  const a=await sample();a.profile.fields[0].part='word/styles.xml';await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'PART_NOT_ALLOWED');
});
for(const markup of ['<w:tab/>','<w:br/>','<w:fldChar w:fldCharType="begin"/>','<w:hyperlink/>'])test(`bloqueia parágrafo com ${markup}`,async()=>{
  const a=await sample({body:`<w:p>${r('Vendedor: Ana Teste')}${markup}</w:p>`});await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'UNSAFE_PARAGRAPH');
});
test('bloqueia comentários, histórico, metadados pessoais, mídia, vínculo externo e texto oculto',async()=>{
  const cases=[
    {extra:{'word/comments.xml':`<w:comments xmlns:w="${W}"/>`}},
    {body:`<w:p><w:ins>${r('Vendedor: Ana Teste')}</w:ins></w:p>`},
    {extra:{'docProps/core.xml':'<cp:coreProperties xmlns:cp="urn:cp" xmlns:dc="urn:dc"><dc:creator>Pessoa Fictícia</dc:creator></cp:coreProperties>'}},
    {extra:{'word/media/image.png':Buffer.from([1,2,3])}},
    {extra:{'word/_rels/document.xml.rels':'<Relationships><Relationship TargetMode="External" Target="https://example.invalid"/></Relationships>'}},
    {body:`<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>Vendedor: Ana Teste</w:t></w:r></w:p>`}
  ];
  for(const input of cases){const a=await sample(input);assert.ok(privacyFindings(await openDocx(a.bytes)).length);await blocked(()=>applyAuthorized(a.bytes,a.profile,a.plan),'DOCUMENT_REQUIRES_PRIVACY_REVIEW');}
});
test('recusa arquivo que não é DOCX',async()=>{
  await blocked(()=>openDocx(Buffer.from('não é zip')),'INVALID_ZIP');
  const zip=new JSZip();zip.file('a.txt','x');await blocked(()=>openDocxBuffer(zip),'NOT_DOCX');
});
async function openDocxBuffer(zip){return openDocx(await zip.generateAsync({type:'nodebuffer'}));}
test('impõe limite de tamanho antes de descompactar',async()=>{
  const a=await sample();assert.throws(()=>checkZip(a.bytes,{maxFileBytes:1}),e=>e.code==='FILE_SIZE_LIMIT');
});
test('impõe limite de expansão antes de abrir o pacote',async()=>{
  const a=await sample();await blocked(()=>openDocx(a.bytes,{maxFileBytes:1e6,maxEntries:100,maxExpandedBytes:1,maxPartBytes:1e6}),'ZIP_EXPANSION_LIMIT');
});
test('detecta tamanho descompactado adulterado antes do JSZip',async()=>{
  const a=await sample();const bytes=Buffer.from(a.bytes);const c=bytes.indexOf(Buffer.from('504b0102','hex'));bytes.writeUInt32LE(1,c+24);
  await blocked(()=>openDocx(bytes),'ZIP_EXPANSION_OR_DATA_ERROR');
});
test('recusa XML com DTD ou entidades declaradas',async()=>{
  await blocked(()=>sample({extra:{'word/danger.xml':'<!DOCTYPE x [<!ENTITY x "qualquer">]><x>&x;</x>'}}),'XML_UNSUPPORTED_DECLARATION');
});
test('recusa XML malformado',async()=>{
  await blocked(()=>sample({body:'<w:p><w:r></w:p>'}),'INVALID_XML');
});
test('preserva isolamento de requisições simultâneas',async()=>{
  const a=await sample(),b=await sample();b.plan.replacements[0].value='Outra Pessoa Fictícia';
  const results=await Promise.all([applyAuthorized(a.bytes,a.profile,a.plan),applyAuthorized(b.bytes,b.profile,b.plan)]);
  const texts=await Promise.all(results.map(async r=>paragraphsOf((await openDocx(r.bytes)).parts.get('word/document.xml').toString())[0].text));
  assert.ok(texts[0].includes('Bruna Inventada'));assert.ok(!texts[0].includes('Outra Pessoa'));assert.ok(texts[1].includes('Outra Pessoa'));
});
