'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
const {randomUUID}=require('node:crypto');
const vm=require('node:vm');
const esbuild=require('esbuild');
const {automationFixture}=require('./automation-fixture.cjs');
const {createSimulator}=require('./linear-simulator.cjs');
const {runBatch,parseOrder,inputPdf}=require('../src/automation.cjs');
const {LinearClient}=require('../src/linear.cjs');
const {withLock,jobStore,cleanArtifacts}=require('../src/job-store.cjs');
const {request}=require('../src/http.cjs');
const root=path.resolve(__dirname,'..');
async function setup(t,cases=['completo'],controls={}){
 const fixture=await automationFixture();const sim=createSimulator({...fixture,cases,controls});const origin=await sim.start();t.after(()=>sim.close());
 Object.assign(fixture.bundle.config.linear,{apiUrl:origin+'/graphql',simulation:true});
 await fs.mkdir(path.join(root,'storage','temp'),{recursive:true});const dir=await fs.mkdtemp(path.join(root,'storage','temp','automation-regression-'));
 const options={bundle:fixture.bundle,stateRoot:path.join(dir,'state'),artifactRoot:path.join(dir,'artifacts'),secretReader:async()=> 'FICTICIA-SEM-ACESSO'};
 return {...fixture,sim,origin,dir,options};
}
test('lote completo entrega so apos download de confirmacao e nao repete anexos',async t=>{
 const {sim,options}=await setup(t,['completo','digitalizada','incompleto','contraditorio','modelo-inexistente','sem-pdf','pdf-invalido','falha-download','completo']);
 const first=await runBatch(options),second=await runBatch(options);
 assert.equal(first.results.filter(r=>r.status==='delivered-in-review').length,3);assert.equal(first.results.filter(r=>r.status==='pending').length,6);
 assert.equal(sim.stats.attachmentsCreated,3);assert.equal(sim.stats.reviewUpdates,3);assert.equal(sim.stats.commentsCreated,9);assert.equal(sim.stats.openaiCalls,0);assert.equal(second.results.some(r=>r.status==='delivered-in-review'),false);
});
test('auth invalida ou workspace incorreto interrompem antes de baixar ou escrever',async t=>{
 for(const [controls,code] of [[{auth:true},'LINEAR_AUTH_FAILED'],[{wrongWorkspace:true},'LINEAR_WORKSPACE_FORBIDDEN']]){
  const {sim,options}=await setup(t,['completo'],controls);await assert.rejects(()=>runBatch(options),e=>e.code===code);assert.equal(sim.stats.downloads,0);assert.equal(sim.stats.uploadAllocations,0);
 }
});
test('agendamento e chamada paga bloqueados antes de ler credenciais',async t=>{
 const {options}=await setup(t);options.secretReader=()=>{throw Error('MUST_NOT_READ_SECRET');};
 await assert.rejects(()=>runBatch({...options,trigger:'automatic'}),e=>e.code==='AUTOMATIC_EXECUTION_DISABLED');
 options.bundle.config.openaiMode='live';await assert.rejects(()=>runBatch(options),e=>e.code==='PAID_API_NOT_AUTHORIZED');
});
test('falha de upload nao altera status e retomada reutiliza mesmo asset',async t=>{
 const {sim,options}=await setup(t,['completo'],{put:1});const a=await runBatch(options);assert.equal(a.results[0].code,'LINEAR_UPLOAD_FAILED');assert.equal(sim.stats.reviewUpdates,0);
 const b=await runBatch(options);assert.equal(b.results[0].status,'delivered-in-review');assert.equal(sim.stats.uploadAllocations,1);assert.equal(sim.stats.attachmentsCreated,1);
});
test('falha de comentario ou status retoma sem anexar documento duplicado',async t=>{
 for(const failure of ['CartorioComment','CartorioReview']){
  const {sim,options}=await setup(t,['completo'],{[failure]:1});await runBatch(options);assert.equal(sim.stats.reviewUpdates,0);const retry=await runBatch(options);assert.equal(retry.results[0].status,'delivered-in-review');assert.equal(sim.stats.attachmentsCreated,1);assert.equal(sim.stats.uploadAllocations,1);
 }
});
test('recuperacao nas fronteiras de persistencia nao repete geracao ou entrega',async t=>{
 for(const phase of ['generated','upload-allocated','uploaded','attached','delivery-confirmed','review-updated']){
  const {sim,options}=await setup(t);let fired=false;
  await runBatch({...options,checkpoint:async p=>{if(p===phase&&!fired){fired=true;throw Object.assign(Error(),{code:'INTERRUPTED_FICTITIOUS'});}}});
  const r=await runBatch(options);assert.ok(['delivered-in-review','recovered-complete'].includes(r.results[0].status),phase);assert.equal(sim.stats.uploadAllocations,1,phase);assert.equal(sim.stats.attachmentsCreated,1,phase);assert.equal(sim.stats.reviewUpdates,1,phase);
 }
});
test('alteracao da solicitacao durante o processamento impede entrega',async t=>{
 const {sim,options}=await setup(t);const id=sim.stats.issues[0].id;
 const r=await runBatch({...options,checkpoint:async phase=>{if(phase==='generated')sim.issues.get(id).description+='\nAlterado durante execucao';}});
 assert.equal(r.results[0].code,'REQUEST_CHANGED_DURING_PROCESSING');assert.equal(sim.stats.attachmentsCreated,0);assert.equal(sim.stats.reviewUpdates,0);
});
test('mudanca do PDF ou modo exige requestId novo e nao entrega na issue errada',async t=>{
 const {sim,options}=await setup(t);let once=false;
 await runBatch({...options,checkpoint:async phase=>{if(phase==='generated'&&!once){once=true;throw Object.assign(Error(),{code:'INTERRUPTED_FICTITIOUS'});}}});
 const issue=sim.issues.get(sim.stats.issues[0].id);issue.description+='\nPedido editado';const r=await runBatch(options);assert.equal(r.results[0].code,'REQUEST_CHANGED_REQUIRES_NEW_ID');assert.equal(sim.stats.attachmentsCreated,0);
});
test('limite de arquivos, endpoints e redirects nao encaminham credenciais a terceiros',async t=>{
 const {options}=await setup(t);const c=new LinearClient({...options.bundle.config.linear,apiKey:'FICTICIA'});
 assert.throws(()=>c.checkAsset('http://outra-maquina/arquivo'),e=>e.code==='ASSET_URL_FORBIDDEN');
 assert.throws(()=>new LinearClient({apiKey:'FICTICIA',workspaceId:'f',teamId:'f',apiUrl:'https://malicioso.invalid/graphql'}),e=>e.code==='LINEAR_ENDPOINT_FORBIDDEN');
 assert.throws(()=>request('file:///arquivo'),e=>e.code==='HTTPS_REQUIRED');
});
test('lock ativo impede execucoes simultaneas e libera apos erro',async t=>{
 const {dir}=await setup(t);const file=path.join(dir,'mutex.lock');let release;const gate=new Promise(r=>{release=r;});let started;const begun=new Promise(r=>{started=r;});
 const first=withLock(file,async()=>{started();await gate;});await begun;
 await assert.rejects(()=>withLock(file,async()=>{}),e=>e.code==='BATCH_ALREADY_RUNNING');release();await first;await withLock(file,async()=>{});
});
test('retencao apaga apenas artefatos permitidos e preserva recibos',async t=>{
 const {dir}=await setup(t),artifacts=path.join(dir,'retention'),name='a'.repeat(64);await fs.mkdir(path.join(artifacts,name),{recursive:true});
 const file=path.join(artifacts,name,'response.json');await fs.writeFile(file,'FICTICIO');await fs.utimes(file,new Date(0),new Date(0));
 const receipts=path.join(dir,'state','jobs');await fs.mkdir(receipts,{recursive:true});await fs.writeFile(path.join(receipts,'receipt.json'),'FICTICIO');
 const r=await cleanArtifacts(artifacts,{artifactsDays:7});assert.equal(r.removed,1);assert.equal(await fs.readFile(path.join(receipts,'receipt.json'),'utf8'),'FICTICIO');
});
test('bundle funciona em VM sem globais URL ou fetch, como runner isolado',async t=>{
 const {sim,options}=await setup(t);const built=await esbuild.build({entryPoints:[path.join(root,'src','automation.cjs')],bundle:true,write:false,platform:'node',format:'iife',globalName:'Cartorio',external:['node:crypto','node:zlib','node:fs/promises','node:path','node:https','node:http','node:os','node:process','node:url']});
 const context=vm.createContext({require,Buffer,setTimeout,clearTimeout,setImmediate,clearImmediate,options});vm.runInContext(built.outputFiles[0].text,context);
 const r=await vm.runInContext('Cartorio.runBatch(options)',context);assert.equal(r.results[0].status,'delivered-in-review');assert.equal(sim.stats.reviewUpdates,1);
});
test('uma chamada com transporte ficticio percorre o fluxo real e retoma resposta salva sem nova chamada',async t=>{
 const {buildRequest}=require('../src/openai.cjs');const {sha256}=require('../src/docx.cjs');
 for(const interruption of [false,true]){
  const {sim,bundle,options,assets}=await setup(t,['digitalizada']);bundle.config.openaiMode='live';bundle.config.paidApiEnabled=true;
  const f=bundle.fixtures.find(f=>f.scenario==='digitalizada');const built=await buildRequest({templateBytes:assets.template,profile:f.payload.profile,pdfs:[{id:'pacote-ficticio',bytes:assets.digitalizada}],approvals:f.payload.approvals,instructions:f.payload.instructions,config:bundle.openaiConfig});
  options.approval={authorized:true,approvalId:'TESTE-FICTICIO-'+randomUUID(),evidence:'Autorizacao ficticia somente para teste automatizado sem API',expiresAt:new Date(Date.now()+60000).toISOString(),workspaceId:bundle.config.linear.workspaceId,issueId:sim.stats.issues[0].id,scenario:'digitalizada',requestSha256:sha256(Buffer.from(JSON.stringify(built.request))),maxCalls:1,maxBudgetUsd:0.5};
  let calls=0;options.transport=async(url,opts)=>{if(url==='https://api.openai.com/v1/responses'){calls++;return {status:200,bytes:Buffer.from(JSON.stringify(f.payload.mockResponse))};}return request(url,opts);};
  if(interruption){const first=await runBatch({...options,checkpoint:async phase=>{if(phase==='response-saved')throw Object.assign(Error(),{code:'INTERRUPTED_FICTITIOUS'});}});assert.equal(first.results[0].status,'pending');}
  const result=await runBatch(options);assert.equal(result.results[0].status,'delivered-in-review');assert.equal(calls,1);assert.equal(sim.stats.attachmentsCreated,1);assert.equal(result.realVisualInterpretationVerified,true);
  const ledger=JSON.parse(await fs.readFile(path.join(options.stateRoot,'openai-budget.json'),'utf8'));assert.equal(ledger.usage.length,1);
 }
});
test('timeout apos reserva impede segunda chamada mesmo com nova execucao',async t=>{
 const {buildRequest}=require('../src/openai.cjs');const {sha256}=require('../src/docx.cjs');const {sim,bundle,options,assets}=await setup(t,['digitalizada']);bundle.config.openaiMode='live';bundle.config.paidApiEnabled=true;
 const f=bundle.fixtures.find(f=>f.scenario==='digitalizada');const built=await buildRequest({templateBytes:assets.template,profile:f.payload.profile,pdfs:[{id:'pacote-ficticio',bytes:assets.digitalizada}],approvals:f.payload.approvals,instructions:f.payload.instructions,config:bundle.openaiConfig});
 options.approval={authorized:true,approvalId:'FICTICIA',evidence:'Autorizacao ficticia sem API real',expiresAt:new Date(Date.now()+60000).toISOString(),workspaceId:bundle.config.linear.workspaceId,issueId:sim.stats.issues[0].id,scenario:'digitalizada',requestSha256:sha256(Buffer.from(JSON.stringify(built.request))),maxCalls:1,maxBudgetUsd:0.5};
 let calls=0;options.transport=async(url,opts)=>{if(url==='https://api.openai.com/v1/responses'){calls++;throw Object.assign(Error(),{code:'HTTP_TIMEOUT'});}return request(url,opts);};
 await runBatch(options);const retry=await runBatch(options);assert.equal(retry.results[0].code,'PAID_CALL_RESULT_UNKNOWN_REVIEW_REQUIRED');assert.equal(calls,1);assert.equal(sim.stats.attachmentsCreated,0);
});
test('recibo de issue corrompido nao interrompe as demais',async t=>{
 const {sim,options}=await setup(t,['completo','completo']);await fs.mkdir(path.join(options.stateRoot,'jobs'),{recursive:true});await fs.writeFile(path.join(options.stateRoot,'jobs',sim.stats.issues[0].id+'.json'),'{INVALIDO');
 const r=await runBatch(options);assert.equal(r.results[0].code,'JOB_STATE_INVALID');assert.equal(r.results[1].status,'delivered-in-review');
});
