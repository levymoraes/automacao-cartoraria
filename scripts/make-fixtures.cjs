'use strict';
// Gerador de dados ficticios. Nao le modelos nem documentos de pessoas reais.
const fs = require('node:fs/promises');
const path = require('node:path');
const { spawnSync } = require('node:child_process');
const JSZip = require('jszip');
const { PDFDocument, StandardFonts, rgb } = require('pdf-lib');
const { sha256, paragraphsOf, openDocx } = require('../src/docx.cjs');
const { coverageOf } = require('../src/operation.cjs');
const root = path.resolve(__dirname, '..');
const output = path.join(root, 'storage', 'development');
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const esc = s => s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const run = (s, bold = false) => `<w:r><w:rPr>${bold ? '<w:b/>' : ''}<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr><w:t xml:space="preserve">${esc(s)}</w:t></w:r>`;
const p = (s, title = false) => `<w:p><w:pPr>${title ? '<w:pStyle w:val="Title"/>' : ''}<w:spacing w:after="140"/></w:pPr>${Array.isArray(s) ? s.map((v,i) => run(v, i === 1)).join('') : run(s,title)}</w:p>`;
const para = [
  'Compra e venda rural para teste de software',
  'DOCUMENTO INTEIRAMENTE FICTÍCIO SEM VALIDADE JURÍDICA. Serve apenas à verificação de edição restrita e não é um modelo oficial de escritura.',
  ['Vendedor: ', 'Arnaldo Exemplo', ' de Teste, brasileiro, agricultor, identificação CPF-TESTE-V01, residente na Rua Inventada 10, Município Fictício.'],
  'Cônjuge do vendedor: Beatriz Exemplo de Teste, brasileira, professora, identificação CPF-TESTE-V02.',
  'Casamento dos vendedores: regime comunhão parcial; certidão CAS-TESTE-ANTIGA-V; data 01/01/2000.',
  'Comprador: Caio Exemplo de Teste, brasileiro, agrônomo, identificação CPF-TESTE-C01, residente na Rua Inventada 20, Município Fictício.',
  'Cônjuge do comprador: Diana Exemplo de Teste, brasileira, veterinária, identificação CPF-TESTE-C02.',
  'Casamento dos compradores: regime separação convencional; certidão CAS-TESTE-ANTIGA-C; data 02/02/2001.',
  'Imóvel rural: Fazenda Exemplo Antiga; área 100,00 ha; município Município Fictício; matrícula MAT-TESTE-ANTIGA.',
  'Cadastro rural: CCIR-TESTE-ANTIGO; cadastro fiscal CIB-TESTE-ANTIGO; cadastro ambiental CAR-TESTE-ANTIGO.',
  'Preço: R$ 100.000,00; pagamento: transferência fictícia REC-TESTE-ANTIGO; data do ato: 01/01/2020.',
  'Documentos fiscais apresentados: certidão fiscal CERT-TESTE-ANTIGA; imposto GUIA-TESTE-ANTIGA; exercício 2020.',
  'Texto fixo para teste: as alterações autorizadas limitam-se aos dados expressamente cadastrados no perfil local. Esta frase deve permanecer intacta.',
  'Texto fixo para teste: este documento não atesta regularidade jurídica, pagamento de tributos nem suficiência da documentação. A revisão humana é obrigatória.'
];
const mappings = [
  [2,'brasileiro','brasileiro','vendedor.nacionalidade',1],
  [2,'agricultor','engenheiro','vendedor.profissao',1],
  [2,'Rua Inventada 10, Município Fictício','Rua Simulada 101, Cidade Nova Fictícia','vendedor.endereco',1],
  [3,'brasileira','brasileira','vendedor.conjuge.nacionalidade',2],
  [3,'professora','pesquisadora','vendedor.conjuge.profissao',2],
  [4,'comunhão parcial','comunhão parcial','vendedor.casamento.regime',5],
  [5,'brasileiro','brasileiro','comprador.nacionalidade',3],
  [5,'agrônomo','economista','comprador.profissao',3],
  [5,'Rua Inventada 20, Município Fictício','Rua Simulada 202, Cidade Nova Fictícia','comprador.endereco',3],
  [6,'brasileira','brasileira','comprador.conjuge.nacionalidade',4],
  [6,'veterinária','analista','comprador.conjuge.profissao',4],
  [7,'separação convencional','separação convencional','comprador.casamento.regime',6],
  [8,'Município Fictício','Cidade Nova Fictícia','imovel.municipio',7],
  [2,'Arnaldo Exemplo de Teste','Eduardo Pessoa Inventada','vendedor.nome',1],
  [2,'CPF-TESTE-V01','ID-FICTICIO-V11','vendedor.identificacao',1],
  [3,'Beatriz Exemplo de Teste','Fernanda Pessoa Inventada','vendedor.conjuge.nome',2],
  [3,'CPF-TESTE-V02','ID-FICTICIO-V12','vendedor.conjuge.identificacao',2],
  [4,'CAS-TESTE-ANTIGA-V','CAS-FICTICIO-NOVA-V','vendedor.casamento.certidao',5],
  [4,'01/01/2000','03/03/2010','vendedor.casamento.data',5],
  [5,'Caio Exemplo de Teste','Gustavo Pessoa Inventada','comprador.nome',3],
  [5,'CPF-TESTE-C01','ID-FICTICIO-C11','comprador.identificacao',3],
  [6,'Diana Exemplo de Teste','Helena Pessoa Inventada','comprador.conjuge.nome',4],
  [6,'CPF-TESTE-C02','ID-FICTICIO-C12','comprador.conjuge.identificacao',4],
  [7,'CAS-TESTE-ANTIGA-C','CAS-FICTICIO-NOVA-C','comprador.casamento.certidao',6],
  [7,'02/02/2001','04/04/2011','comprador.casamento.data',6],
  [8,'Fazenda Exemplo Antiga','Fazenda Horizonte Inventado','imovel.nome',7],
  [8,'100,00 ha','125,50 ha','imovel.area',7],
  [8,'MAT-TESTE-ANTIGA','MAT-FICTICIA-NOVA','imovel.matricula',7],
  [9,'CCIR-TESTE-ANTIGO','CCIR-FICTICIO-NOVO','imovel.ccir',8],
  [9,'CIB-TESTE-ANTIGO','CIB-FICTICIO-NOVO','imovel.cib',8],
  [9,'CAR-TESTE-ANTIGO','CAR-FICTICIO-NOVO','imovel.car',8],
  [10,'R$ 100.000,00','R$ 275.000,00','ato.preco',9],
  [10,'REC-TESTE-ANTIGO','REC-FICTICIO-NOVO','ato.pagamento',9],
  [10,'01/01/2020','30/10/2026','ato.data',9],
  [11,'CERT-TESTE-ANTIGA','CERT-FICTICIA-NOVA','fiscal.certidao',10],
  [11,'GUIA-TESTE-ANTIGA','GUIA-FICTICIA-NOVA','fiscal.guia',10],
  [11,'2020','2026','fiscal.exercicio',10],
  [14,'PROTOCOLO-ANTIGO','PROTOCOLO-FICTICIO-NOVO','ato.protocolo',9]
];

async function createDocx({write=true}={}) {
  const zip = new JSZip();
  const entries = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>',
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>',
    'docProps/core.xml': '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"/>',
    'word/document.xml': `<w:document xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${para.map((s,i)=>p(s,i===0)).join('')}<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders>${['top','left','bottom','right','insideH','insideV'].map(s=>`<w:${s} w:val="single" w:sz="4" w:color="D9D9D9"/>`).join('')}</w:tblBorders><w:tblCellMar><w:top w:w="120" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:bottom w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9000" w:type="dxa"/></w:tcPr>${p('Protocolo do ato: PROTOCOLO-ANTIGO')}</w:tc></w:tr></w:tbl><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:footerReference w:type="default" r:id="rIdFooter"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567"/></w:sectPr></w:body></w:document>`,
    'word/_rels/document.xml.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdFooter" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>',
    'word/styles.xml': `<w:styles xmlns:w="${W}"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="22"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:color w:val="000000"/><w:sz w:val="28"/></w:rPr></w:style></w:styles>`,
    'word/header1.xml': `<w:hdr xmlns:w="${W}">${p('SERVENTIA FICTÍCIA - LIVRO-TESTE-ANTIGO')}</w:hdr>`,
    'word/footer1.xml': `<w:ftr xmlns:w="${W}">${p('SEM VALIDADE JURÍDICA - SELO-TESTE-ANTIGO')}</w:ftr>`
  };
  for (const [name, text] of Object.entries(entries)) zip.file(name,text,{date:new Date('2026-01-01T00:00:00Z'),createFolders:false});
  const bytes = await zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
  const fields = mappings.map(([paragraph,oldText,value,id,page]) => {
    const text = paragraphsOf(entries['word/document.xml'])[paragraph].text;
    const start = text.indexOf(oldText);
    if (start < 0) throw new Error('FIXTURE_MAPPING_ERROR');
    return {id,part:'word/document.xml',paragraph,start,end:start+oldText.length,oldText,required:true,value,page};
  });
  for (const [part,oldText,value,id] of [
    ['word/header1.xml','LIVRO-TESTE-ANTIGO','LIVRO-FICTICIO-NOVO','ato.livro'],
    ['word/footer1.xml','SELO-TESTE-ANTIGO','SELO-FICTICIO-NOVO','ato.selo']]) {
    const start = paragraphsOf(entries[part])[0].text.indexOf(oldText);
    fields.push({id,part,paragraph:0,start,end:start+oldText.length,oldText,required:true,value,page:9});
  }
  const hash = sha256(bytes);
  const entities=[
    {id:'seller',kind:'person',role:'seller',spouseId:'seller-spouse'},
    {id:'seller-spouse',kind:'person',role:'seller_spouse',spouseId:'seller'},
    {id:'buyer',kind:'person',role:'buyer',spouseId:'buyer-spouse'},
    {id:'buyer-spouse',kind:'person',role:'buyer_spouse',spouseId:'buyer'},
    {id:'seller-marriage',kind:'marriage',role:'marriage',spouseId:null},
    {id:'buyer-marriage',kind:'marriage',role:'marriage',spouseId:null},
    {id:'property',kind:'property',role:'property',spouseId:null},
    {id:'transaction',kind:'transaction',role:'transaction',spouseId:null}
  ];
  const owner=id=>id.startsWith('vendedor.conjuge.')?'seller-spouse':id.startsWith('comprador.conjuge.')?'buyer-spouse':id.startsWith('vendedor.casamento.')?'seller-marriage':id.startsWith('comprador.casamento.')?'buyer-marriage':id.startsWith('vendedor.')?'seller':id.startsWith('comprador.')?'buyer':id.startsWith('imovel.')?'property':'transaction';
  for(const f of fields){f.entityId=owner(f.id);f.property=f.id.split('.').at(-1);}
  const profile={schemaVersion:2,scope:'synthetic-only',templateSha256:hash,anonymizationApproved:true,externalSharingAllowed:true,origin:'generated-synthetic',operation:{inventoryReviewed:true,entities,coverage:coverageOf(await openDocx(bytes),paragraphsOf)},fields:fields.map(({value,page,...f})=>f)};
  const plan={schemaVersion:2,scope:'synthetic-only',templateSha256:hash,pending:[],entities:entities.map(e=>({...e,slotId:e.id,id:'new-'+e.id,spouseId:e.spouseId?'new-'+e.spouseId:null})),replacements:fields.map(f=>({fieldId:f.id,entityId:'new-'+f.entityId,property:f.property,value:f.value,confirmedUnchanged:f.value===f.oldText,source:{documentId:'pacote-ficticio',page:f.page,entityId:'new-'+f.entityId,property:f.property}}))};
  if(!write)return {fields,bytes,profile,plan};
  await fs.writeFile(path.join(output,'modelo-ficticio.docx'),bytes);
  await fs.writeFile(path.join(output,'perfil-ficticio.json'),JSON.stringify(profile,null,2));
  await fs.writeFile(path.join(output,'plano-ficticio.json'),JSON.stringify(plan,null,2));
  return fields;
}

async function createPdf(fields, mode, dates={}) {
  const pdf = await PDFDocument.create();
  pdf.setTitle('Documentacao inteiramente ficticia para teste'); pdf.setAuthor(''); pdf.setCreator('Gerador local de testes');
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const bold = await pdf.embedFont(StandardFonts.HelveticaBold);
  const titles = ['Identificação do vendedor','Identificação do cônjuge do vendedor','Identificação do comprador','Identificação do cônjuge do comprador','Casamento dos vendedores','Casamento dos compradores','Registro do imóvel rural','Cadastros rurais','Operação e pagamento','Documentos fiscais'];
  for (let number=1;number<=10;number++) {
    const page = pdf.addPage([595.28,841.89]);
    page.drawText('DOCUMENTO FICTÍCIO SEM VALIDADE',{x:42,y:790,size:15,font:bold,color:rgb(0,0,0)});
    page.drawText(`${number}. ${titles[number-1]}`,{x:42,y:752,size:14,font:bold});
    page.drawText('Todos os nomes e identificadores foram inventados para testes.',{x:42,y:723,size:10,font});
    let y=674;
    for (const field of fields.filter(f=>f.page===number)) {
      page.drawText(field.id,{x:42,y,size:11,font:bold}); y-=21;
      let value = field.value;
      if (mode==='incompleto' && field.id==='imovel.matricula') value='INFORMAÇÃO NÃO APRESENTADA';
      page.drawText(value,{x:42,y,size:12,font}); y-=42;
    }
    if (mode==='contraditorio' && number===10) {
      page.drawText('INCONSISTÊNCIA PROPOSITAL NO PACOTE DE TESTE',{x:42,y:y-10,size:11,font:bold});
      page.drawText('Matrícula informada aqui: MAT-FICTICIA-DIVERGENTE',{x:42,y:y-37,size:11,font});
      page.drawText('Preço informado aqui: R$ 999.000,00',{x:42,y:y-60,size:11,font});
    }
    page.drawText(`Pacote de desenvolvimento local | página ${number} de 10`,{x:42,y:40,size:10,font});
  }
  if(dates.creationDate)pdf.setCreationDate(dates.creationDate);
  if(dates.modificationDate)pdf.setModificationDate(dates.modificationDate);
  return pdf.save();
}

async function main() {
  await fs.mkdir(output,{recursive:true});
  const fields = await createDocx();
  for (const mode of ['completo','incompleto','contraditorio']) await fs.writeFile(path.join(output,`entrada-${mode}.pdf`),await createPdf(fields,mode));
  const rasterizer = process.env.PDFTOPPM;
  let scanned = false;
  if (rasterizer) {
    const prefix = path.join(root,'storage','temp','scan');
    await fs.mkdir(path.dirname(prefix),{recursive:true});
    const r = spawnSync(rasterizer,['-r','110','-png',path.join(output,'entrada-completo.pdf'),prefix],{windowsHide:true,timeout:120000,stdio:'pipe'});
    if (r.status!==0) throw new Error('RASTERIZER_FAILED');
    const pdf = await PDFDocument.create();
    for (const name of (await fs.readdir(path.dirname(prefix))).filter(n=>/^scan-\d+\.png$/.test(n)).sort((a,b)=>a.localeCompare(b,undefined,{numeric:true}))) {
      const img=await pdf.embedPng(await fs.readFile(path.join(path.dirname(prefix),name)));
      const page=pdf.addPage([595.28,841.89]); page.drawImage(img,{x:0,y:0,width:595.28,height:841.89});
    }
    if (pdf.getPageCount()!==10) throw new Error('SCAN_PAGE_COUNT');
    await fs.writeFile(path.join(output,'entrada-digitalizada.pdf'),await pdf.save()); scanned=true;
  }
  const complete=JSON.parse(await fs.readFile(path.join(output,'plano-ficticio.json'),'utf8'));
  for (const [mode,pending] of [['incompleto',[{code:'MISSING_REQUIRED',fieldId:'imovel.matricula'}]],['contraditorio',[{code:'CONTRADICTION',fieldId:'imovel.matricula',pages:[7,10]},{code:'CONTRADICTION',fieldId:'ato.preco',pages:[9,10]}]]]) {
    const plan=structuredClone(complete); plan.pending=pending;
    if(mode==='incompleto') plan.replacements=plan.replacements.filter(r=>r.fieldId!=='imovel.matricula');
    await fs.writeFile(path.join(output,`plano-${mode}.json`),JSON.stringify(plan,null,2));
  }
  await fs.writeFile(path.join(output,'origem-dos-dados.json'),JSON.stringify({scope:'synthetic-only',source:'Constantes inventadas em scripts/make-fixtures.cjs; nenhum documento original lido.',pdfPages:10,scannedPdfCreated:scanned,extractionPerformed:false,plans:'Planos determinísticos escritos pelo gerador; não extraídos dos PDFs por IA.'},null,2));
  const approvals=[{kind:'template',sha256:sha256(await fs.readFile(path.join(output,'modelo-ficticio.docx'))),approved:true,externalSharingAllowed:true,origin:'generated-synthetic'}];
  for(const name of ['entrada-completo.pdf','entrada-incompleto.pdf','entrada-contraditorio.pdf',...(scanned?['entrada-digitalizada.pdf']:[])])approvals.push({kind:'pdf',sha256:sha256(await fs.readFile(path.join(output,name))),pages:10,approved:true,externalSharingAllowed:true,origin:'generated-synthetic'});
  await fs.writeFile(path.join(output,'fixture-approvals.json'),JSON.stringify(approvals,null,2));
  console.log(JSON.stringify({fixturesCreated:true,scannedPdfCreated:scanned,fields:fields.length,externalCalls:0}));
}
if(require.main===module)main().catch(()=>{console.error('FIXTURE_GENERATION_FAILED');process.exitCode=1;});
module.exports={createDocx,createPdf};
