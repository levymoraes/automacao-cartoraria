'use strict';
// Demonstração inteiramente local. A resposta da IA é simulada de propósito.
const fs=require('node:fs/promises'),path=require('node:path'),assert=require('node:assert/strict');
const {PDFDocument,StandardFonts}=require('pdf-lib');
const {semanticFixture}=require('../tests/semantic-fixture.cjs');
const {automationFixture}=require('../tests/automation-fixture.cjs');
const {createSimulator}=require('../tests/linear-simulator.cjs');
const {openDocx,paragraphsOf,sha256}=require('../src/docx.cjs');
const {patchText}=require('../src/anonymize.cjs');
const {prepareTemplateFill}=require('../src/template-fill.cjs');
const {runBatch}=require('../src/automation.cjs');
const {request}=require('../src/http.cjs');
const root=path.resolve(__dirname,'..');
async function main(){
 const source=await semanticFixture(),doc=await openDocx(source.bytes);
 const fields=source.profile.fields.map((f,i)=>({...f,tag:'CAMPO_'+i}));
 for(const [part,bytes]of doc.parts){const edits=fields.filter(f=>f.part===part).map(f=>({...f,value:'{{'+f.tag+'}}'}));if(edits.length)doc.zip.file(part,patchText(bytes.toString(),edits).xml);}
 const templateBytes=await doc.zip.generateAsync({type:'nodebuffer'});
 const metadata={modelId:'DEMO-DOCUMENTAL',version:'1',originalDocSha256:sha256(source.bytes),preparedSha256:sha256(templateBytes),entities:source.profile.operation.entities.filter(e=>e.kind==='person'),couples:[],fields:fields.map(f=>({tag:f.tag,entityId:f.entityId,property:f.property,occurrences:1,label:f.property,case:'source'}))};
 const pdf=await PDFDocument.create(),page=pdf.addPage(),font=await pdf.embedFont(StandardFonts.Helvetica);
 const lines=['DOCUMENTOS INTEIRAMENTE FICTICIOS - SEM VALIDADE JURIDICA','Vendedor: Joao Exemplo Ficticio','Profissao do vendedor: analista de teste','CPF do vendedor: nao apresentado','Compradora: Maria Exemplo Ficticio','Identificador da compradora: ID-FICTICIO-C02'];
 lines.forEach((line,i)=>page.drawText(line,{x:35,y:780-i*24,size:12,font}));pdf.setTitle('Dossie inteiramente ficticio');
 const pdfBytes=Buffer.from(await pdf.save()),frame=await automationFixture();
 const bundle={config:{...frame.bundle.config,pipeline:'production',documentEditingMode:'template-fields',developmentOnly:false,openaiMode:'live',paidApiEnabled:true},catalog:[{...frame.bundle.catalog[0],id:metadata.modelId,format:'docx',status:'ready-production',userApprovedUse:true,sha256:sha256(source.bytes)}],openaiConfig:frame.bundle.openaiConfig,instructions:await fs.readFile(path.join(root,'config/template-fields-instructions.txt'),'utf8'),fieldTemplates:[{modelId:metadata.modelId,version:'1',originalDocSha256:metadata.originalDocSha256,docxBase64:templateBytes.toString('base64'),metadata}]};
 // "live" exercita o caminho operacional; todo transporte externo é interceptado.
 const sim=createSimulator({bundle,assets:{...frame.assets,template:source.bytes,digitalizada:pdfBytes},cases:['digitalizada']});
 const origin=await sim.start();Object.assign(bundle.config.linear,{apiUrl:origin+'/graphql',simulation:true});
 const issue=sim.issues.get(sim.stats.issues[0].id);issue.title='Solicitacao inteiramente ficticia';issue.description='Usar '+metadata.modelId+' e registrar o CPF ausente.';
 const prepared=await prepareTemplateFill({templateBytes,fieldMetadata:metadata,pdfs:[{id:'document-1',bytes:pdfBytes}],model:bundle.catalog[0],task:{description:issue.description},config:bundle.openaiConfig,instructions:bundle.instructions});
 const values={'seller.nome':'Joao Exemplo Ficticio','seller.cpf':null,'seller.profissao':'analista de teste','buyer.nome':'Maria Exemplo Ficticio','buyer.cpf':'ID-FICTICIO-C02'};
 const evidence=(entityId,property)=>({kind:'document',documentId:'document-1',page:1,entityId,property,purpose:'identity',quote:lines.join('; ')});
 const reading={templateSha256:metadata.preparedSha256,reviewedPages:[{documentId:'document-1',pages:[1]}],entities:metadata.entities.map(e=>({slotId:e.id,participantId:'participant-'+e.id,name:values[e.id+'.nome'],spouseSlotId:null})),fields:metadata.fields.map(f=>({tag:f.tag,value:values[f.entityId+'.'+f.property],status:values[f.entityId+'.'+f.property]===null?'missing':'verified',sources:values[f.entityId+'.'+f.property]===null?[]:[evidence(f.entityId,f.property)]})),pending:[],audit:{operation:{actType:'compra-venda-rural',objectNature:'rural',participants:metadata.entities.map(e=>({id:'participant-'+e.id,name:values[e.id+'.nome'],role:e.role,maritalStatus:'single',gender:'unknown',spouseId:null,representsId:null,sources:[{...evidence('participant-'+e.id,'role'),purpose:'role'}]}))},compatibility:{status:'compatible',reasons:[],requiredModelCharacteristics:[]},financialFacts:[],records:[],paragraphReviews:prepared.templateParagraphs.map(p=>({part:p.part,paragraph:p.paragraph,status:'not_applicable',reason:'Texto de demonstracao ficticio.',sources:[]})),checks:['participants','source_ownership','old_data','values_dates','grammar','duplicates_fragments','fixed_assertions'].map(name=>({name,passed:true,message:''})),findings:[]}};
 const demoRoot=path.join(root,'storage/demo');await fs.mkdir(demoRoot,{recursive:true});const out=await fs.mkdtemp(path.join(demoRoot,'execucao-'));
 let simulatedCalls=0;
 try{
  const options={bundle,stateRoot:path.join(out,'state'),artifactRoot:path.join(out,'artifacts'),approval:{authorized:true,scope:'production',dataScope:'institutional',approvalId:'demonstracao-ficticia',evidence:'Simulacao local; nenhuma API externa',expiresAt:new Date(Date.now()+60000).toISOString(),workspaceId:bundle.config.linear.workspaceId,maxCalls:1,maxBudgetUsd:1},secretReader:async()=> 'FICTICIA-SEM-ACESSO',transport:async(url,opts)=>{
   if(url==='https://api.openai.com/v1/responses'){simulatedCalls++;return {status:200,bytes:Buffer.from(JSON.stringify({id:'RESPOSTA-FICTICIA',status:'completed',model:bundle.openaiConfig.model,service_tier:'default',usage:{input_tokens:100,output_tokens:50},output:[{type:'message',content:[{type:'output_text',text:JSON.stringify(reading)}]}]}))};}
   assert.ok(url.startsWith(origin+'/'),'Destino externo bloqueado na demonstracao');return request(url,opts);
  }};
  const first=await runBatch(options),second=await runBatch(options);
  assert.equal(first.results[0]?.status,'delivered-in-review',JSON.stringify(first));assert.equal(sim.stats.attachmentsCreated,1);assert.equal(sim.stats.reviewUpdates,1);assert.equal(simulatedCalls,1);assert.equal(second.results.length,0,'A tarefa em revisao sai da fila elegivel');
  const state=JSON.parse(await fs.readFile(path.join(out,'state/jobs',issue.id+'.json'),'utf8')),file=path.join(out,'artifacts',state.jobId,'result.docx'),result=await openDocx(await fs.readFile(file));
  const xml=result.parts.get('word/document.xml').toString(),text=paragraphsOf(xml).map(p=>p.text).join('\n');assert.match(xml,/FF0000/);assert.match(text,/pendente/iu);assert.match(text,/Joao Exemplo Ficticio/);
  const summary={simulation:true,externalApiCalls:0,credentialsRead:false,generatedDraft:true,missingDataMarkedRed:true,confirmedAttachment:true,reviewState:'In Review',attachmentsCreated:1,duplicateAttachmentsOnRerun:0,humanApprovalRequired:true};
  await fs.writeFile(path.join(out,'demo-summary.json'),JSON.stringify(summary,null,2));console.log(JSON.stringify({...summary,outputDirectory:path.relative(root,out)},null,2));
 }finally{await sim.close();}
}
main().catch(e=>{console.error(e.code||e.message);process.exitCode=1;});
