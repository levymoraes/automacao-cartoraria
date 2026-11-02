'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const JSZip=require('jszip');
const {patchText,sanitizePackage}=require('../src/anonymize.cjs');
const logo=Buffer.from('LOGO-FICTICIA-INALTERAVEL');
const {reviewSynthetic}=require('./helpers.cjs');
const {openDocx,paragraphsOf,parseXml,privacyFindings,sha256,applyAuthorized}=require('../src/docx.cjs');
const W='http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const xml=text=>`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Nome: ${text}</w:t></w:r></w:p></w:body></w:document>`;
async function fixture(body=xml('Pessoa Ficticia')){
  const zip=new JSZip();
  zip.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpeg" ContentType="image/jpeg"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/></Types>',{createFolders:false});
  zip.file('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',{createFolders:false});
  zip.file('word/document.xml',body,{createFolders:false});
  zip.file('word/styles.xml',`<w:styles xmlns:w="${W}"/>`,{createFolders:false});
  return zip;
}
async function buffer(zip){return zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});}

test('anonimização altera apenas os spans e preserva rPr e texto fixo',()=>{
  const before=xml('Pessoa Ficticia; cláusula fixa');
  const out=patchText(before,[{id:'nome',paragraph:0,start:6,end:21,value:'Outra Pessoa Teste'}]);
  assert.equal(paragraphsOf(out.xml)[0].text,'Nome: Outra Pessoa Teste; cláusula fixa');
  assert.ok(out.xml.includes('<w:rPr><w:b/></w:rPr>'));
  assert.deepEqual(out.fields[0],{id:'nome',paragraph:0,start:6,end:24,oldText:'Outra Pessoa Teste'});
});
test('anonimização em runs acrescenta xml:space quando necessário',()=>{
  const before=`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t>ABCDE</w:t></w:r><w:r><w:t>FGHIJ</w:t></w:r></w:p></w:body></w:document>`;
  const out=patchText(before,[{id:'campo',paragraph:0,start:1,end:9,value:' X Y Z '}]);
  assert.equal(paragraphsOf(out.xml)[0].text,'A X Y Z J');assert.ok(out.xml.includes('xml:space="preserve"'));
});
test('anonimização bloqueia sobreposição e offsets inválidos',()=>{
  assert.throws(()=>patchText(xml('Pessoa Ficticia'),[{id:'a',paragraph:0,start:1,end:8,value:'x'},{id:'b',paragraph:0,start:4,end:9,value:'y'}]),e=>e.code==='ANONYMIZATION_OVERLAP');
  assert.throws(()=>patchText(xml('Pessoa Ficticia'),[{id:'a',paragraph:99,start:1,end:8,value:'x'}]),e=>e.code==='ANONYMIZATION_ANCHOR');
});
test('anonimização remove somente valores pessoais e preserva partes auxiliares e vínculos',async()=>{
  const z=await fixture(xml('Pessoa Ficticia').replace('<w:p>','<w:p w:rsidR="12345678">'));
  z.file('docProps/core.xml','<x xmlns:dc="urn:dc" xmlns:dcterms="urn:dates" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Autor Ficticio</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">2026-01-01T00:00:00Z</dcterms:created></x>');
  z.file('docProps/app.xml','<Properties><Company>Empresa Ficticia</Company></Properties>');
  z.file('customXml/item1.xml','<private>VALOR-FICTICIO-ANTIGO</private>');
  z.file('customXml/itemProps1.xml','<props/>');
  z.file('word/_rels/document.xml.rels','<Relationships><Relationship Id="custom" Target="../customXml/item1.xml"/><Relationship Id="ext" Target="mailto:teste@example.invalid" TargetMode="External"/></Relationships>');
  z.file('word/settings.xml',`<w:settings xmlns:w="${W}"><w:rsids><w:rsid w:val="12345678"/></w:rsids></w:settings>`);
  const doc=await openDocx(await buffer(z)),out=await sanitizePackage(doc,[]);
  assert.deepEqual(privacyFindings(out.doc),['COMMENTS_OR_CUSTOM_DATA','EXTERNAL_RELATIONSHIP']);
  for(const [name,bytes] of doc.parts)if(name!=='docProps/core.xml')assert.deepEqual(out.doc.parts.get(name),bytes);
  const all=[...out.doc.parts.values()].map(b=>b.toString()).join('');
  assert.ok(!all.includes('Autor Ficticio'));
  assert.ok(!all.includes('dcterms:created'));
  for(const text of ['Empresa Ficticia','VALOR-FICTICIO-ANTIGO','mailto:','rsid','../customXml'])assert.ok(all.includes(text));
});
test('preserva literalmente a logo, referências e tipos de conteúdo sem criar imagens',async()=>{
  const z=await fixture();z.file('word/media/image1.jpeg',Buffer.from('imagem antiga ficticia'));
  z.file('word/_rels/document.xml.rels','<Relationships><Relationship Id="im" Type="image" Target="media/image1.jpeg"/></Relationships>');
  const out=await sanitizePackage(await openDocx(await buffer(z)),[]);
  assert.deepEqual(out.doc.parts.get('word/media/image1.jpeg'),Buffer.from('imagem antiga ficticia'));
  assert.ok(!out.doc.parts.has('word/media/development-mark.png'));
  assert.equal(out.doc.parts.size,4+2);
  assert.ok(out.doc.parts.get('word/_rels/document.xml.rels').toString().includes('media/image1.jpeg'));
  assert.deepEqual(out.removed,[]);assert.deepEqual(out.replaced,[]);
  assert.deepEqual(privacyFindings(out.doc),['MEDIA_REQUIRES_LOCAL_REVIEW']);
});
test('preserva literalmente o hiperlink, seu texto e formato',async()=>{
  const z=await fixture(`<w:document xmlns:w="${W}" xmlns:r="urn:r"><w:body><w:p><w:hyperlink r:id="h"><w:r><w:rPr><w:b/></w:rPr><w:t>Contato ficticio</w:t></w:r></w:hyperlink></w:p></w:body></w:document>`);
  const doc=await openDocx(await buffer(z)),out=await sanitizePackage(doc,[]),body=out.doc.parts.get('word/document.xml').toString();
  assert.equal(paragraphsOf(body)[0].text,'Contato ficticio');assert.ok(body.includes('<w:b/>'));assert.ok(body.includes('hyperlink'));assert.deepEqual(out.doc.parts.get('word/document.xml'),doc.parts.get('word/document.xml'));parseXml(body);
});
test('motor bloqueia expressamente uma cópia que aguarda validação',async()=>{
  const bytes=await buffer(await fixture());
  await assert.rejects(()=>applyAuthorized(bytes,{scope:'synthetic-only',templateSha256:sha256(bytes),anonymizationApproved:false},{scope:'synthetic-only',templateSha256:sha256(bytes)}),e=>e.code==='ANONYMIZATION_APPROVAL_REQUIRED');
});
test('motor permite imagem fictícia somente com aprovação e hash exatos',async()=>{
  const z=await fixture();z.file('word/media/original.png',logo);const bytes=await buffer(z);
  const profile={scope:'synthetic-only',templateSha256:sha256(bytes),anonymizationApproved:true,approvedMedia:[{part:'word/media/original.png',sha256:sha256(logo)}],fields:[{id:'nome',part:'word/document.xml',paragraph:0,start:6,end:21,oldText:'Pessoa Ficticia',required:true}]};
  const plan={scope:'synthetic-only',templateSha256:sha256(bytes),pending:[],replacements:[{fieldId:'nome',value:'Outra Pessoa',source:{documentId:'ficticio',page:1}}]};
  await reviewSynthetic(bytes,profile,plan);
  const out=await applyAuthorized(bytes,profile,plan);assert.deepEqual((await openDocx(out.bytes)).parts.get('word/media/original.png'),logo);
  profile.approvedMedia[0].sha256='0'.repeat(64);
  await assert.rejects(()=>applyAuthorized(bytes,profile,plan),e=>e.code==='DOCUMENT_REQUIRES_PRIVACY_REVIEW');
});
test('contextos distinguem valores iguais no mesmo parágrafo',async()=>{
  const text='Vendedor: Pessoa Ficticia; comprador: Pessoa Ficticia; cláusula fixa.';
  const z=await fixture(`<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p></w:body></w:document>`),bytes=await buffer(z);
  const start=text.lastIndexOf('Pessoa Ficticia'),end=start+'Pessoa Ficticia'.length;
  const field={id:'comprador',part:'word/document.xml',paragraph:0,start,end,oldText:'Pessoa Ficticia',required:true,anchorContext:{before:text.slice(start-16,start),after:text.slice(end,end+16)}};
  const profile={scope:'synthetic-only',templateSha256:sha256(bytes),fields:[field]};
  const plan={scope:'synthetic-only',templateSha256:sha256(bytes),pending:[],replacements:[{fieldId:'comprador',value:'Outra Pessoa',source:{documentId:'ficticio',page:1}}]};
  await reviewSynthetic(bytes,profile,plan);
  profile.operation.fixedOldValueExceptions=[{part:'word/document.xml',paragraph:0,start:text.indexOf('Pessoa Ficticia'),end:text.indexOf('Pessoa Ficticia')+15,value:'Pessoa Ficticia',reason:'Outra pessoa fixa neste teste de contexto.'}];
  const out=await applyAuthorized(bytes,profile,plan);
  assert.equal(paragraphsOf((await openDocx(out.bytes)).parts.get('word/document.xml').toString())[0].text,'Vendedor: Pessoa Ficticia; comprador: Outra Pessoa; cláusula fixa.');
  field.anchorContext.before='contexto divergente';await assert.rejects(()=>applyAuthorized(bytes,profile,plan),e=>e.code==='AMBIGUOUS_ANCHOR');
});
test('dados auxiliares fixos revisados por hash são preservados sem desativar os alertas gerais',async()=>{
  const z=await fixture();const data=Buffer.from('<fixed>Informacao institucional ficticia</fixed>');z.file('customXml/item1.xml',data,{createFolders:false});const bytes=await buffer(z);
  const profile={scope:'synthetic-only',templateSha256:sha256(bytes),anonymizationApproved:true,fields:[{id:'nome',part:'word/document.xml',paragraph:0,start:6,end:21,oldText:'Pessoa Ficticia',required:true}],reviewedPreservedParts:[{part:'customXml/item1.xml',sha256:sha256(data),reviewed:true}]};
  const plan={scope:'synthetic-only',templateSha256:sha256(bytes),pending:[],replacements:[{fieldId:'nome',value:'Outra Pessoa',source:{documentId:'ficticio',page:1}}]};await reviewSynthetic(bytes,profile,plan);
  const out=await applyAuthorized(bytes,profile,plan);assert.deepEqual((await openDocx(out.bytes)).parts.get('customXml/item1.xml'),data);
  profile.reviewedPreservedParts[0].sha256='0'.repeat(64);await assert.rejects(()=>applyAuthorized(bytes,profile,plan),e=>e.code==='DOCUMENT_REQUIRES_PRIVACY_REVIEW');
});
