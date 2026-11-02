'use strict';
const fs=require('node:fs/promises');
const path=require('node:path');
const esbuild=require('esbuild');
const root=path.resolve(__dirname,'..');
async function main(){
  require('./build-schema.cjs');
  // A distribuicao UMD evita os getters de interoperabilidade bloqueados no n8n.
  const r=await esbuild.build({entryPoints:[path.join(root,'src','automation.cjs')],bundle:true,write:false,platform:'node',alias:{'pdf-lib':'pdf-lib/dist/pdf-lib.min.js'},format:'iife',globalName:'Cartorio',target:'node22',minify:true,external:['docxtemplater','docxtemplater/package.json','pizzip','node:crypto','node:zlib','node:fs/promises','node:path','node:https','node:http','node:os','node:process','node:url']});
  const code=r.outputFiles[0].text;
  const jsCode=code+`\ntry{const fs=require('fs');const production=fs.existsSync('/work/production-bundle.json');const bundle=JSON.parse(fs.readFileSync(production?'/work/production-bundle.json':'/work/bundle.json','utf8'));const control=JSON.parse(fs.readFileSync(production?'/work/production-control.json':'/work/control.json','utf8'));Object.assign(bundle.config,control);let approval=bundle.config.testApproval||null;if(!approval)try{approval=JSON.parse(fs.readFileSync('/state/live-approval.local.json','utf8'));}catch{}const onlyIssueIds=bundle.config.testOnlyIssueIds||(bundle.config.openaiMode==='live'?(approval?.scope==='production'?null:approval?.issueId?[approval.issueId]:[]):null);const result=await Cartorio.runBatch({bundle,stateRoot:'/state',artifactRoot:'/data/artifacts',trigger:$input.first().json.trigger||'manual',approval,onlyIssueIds,recoveryOnly:bundle.config.recoveryOnly===true});return [{json:result}];}catch(e){return [{json:{status:'blocked',code:/^[A-Z0-9_]+$/.test(e.code||'')?e.code:(e.name==='ReferenceError'?'RUNNER_GLOBAL_UNAVAILABLE':'AUTOMATION_FAILED'),realVisualInterpretationVerified:false}}];}`;
  let config;try{config=JSON.parse((await fs.readFile(path.join(root,'storage/automation/runtime/production-bundle.json'),'utf8')).replace(/^\uFEFF/,'')).config;}catch(e){if(e.code!=='ENOENT')throw e;config=JSON.parse(await fs.readFile(path.join(root,'config/automation.example.json'),'utf8'));}
  const schedule=config.schedule;const hours=schedule.hours; if(!Array.isArray(hours)||!hours.length||hours.some(h=>!Number.isInteger(h)||h<0||h>23))throw Object.assign(Error(),{code:'INVALID_SCHEDULE_HOURS'});
  new Intl.DateTimeFormat('pt-BR',{timeZone:schedule.timezone});
  const settings={timezone:schedule.timezone,executionTimeout:2400,saveExecutionProgress:false,saveDataSuccessExecution:'none',saveDataErrorExecution:'none',saveManualExecutions:false};
  const execute={id:'execute',name:'Processar lote com entrega confirmada',type:'n8n-nodes-base.code',typeVersion:2,position:[500,0],parameters:{jsCode}};
  const manual={id:'cartorio-completo-manual',name:'Cartorio — fluxo completo manual',active:false,nodes:[{id:'manual',name:'Execucao manual',type:'n8n-nodes-base.manualTrigger',typeVersion:1,position:[0,0],parameters:{}},execute],connections:{'Execucao manual':{main:[[{node:execute.name,type:'main',index:0}]]}},settings};
  const context={id:'context',name:'Execucao automatica protegida',type:'n8n-nodes-base.code',typeVersion:2,position:[240,0],parameters:{jsCode:"return [{json:{trigger:'automatic'}}];"}};
  const clockName=hours.map(h=>h+'h').join(' e ')+' '+schedule.timezone;
  const scheduled={id:'cartorio-completo-agendado',name:'Cartorio — '+hours.map(h=>h+'h').join(' ')+' e recuperacao ao iniciar — DESATIVADO',active:false,nodes:[{id:'clock',name:clockName,type:'n8n-nodes-base.scheduleTrigger',typeVersion:1.2,disabled:true,position:[0,0],parameters:{rule:{interval:[{field:'cronExpression',expression:'0 '+hours.join(',')+' * * *'}]}}},{id:'startup',name:'Recuperar pendentes ao iniciar',type:'n8n-nodes-base.n8nTrigger',typeVersion:1,disabled:true,position:[0,180],parameters:{events:['init']}},context,execute],connections:{[clockName]:{main:[[{node:context.name,type:'main',index:0}]]},'Recuperar pendentes ao iniciar':{main:[[{node:context.name,type:'main',index:0}]]},[context.name]:{main:[[{node:execute.name,type:'main',index:0}]]}},settings};
  for(const w of [manual,scheduled])for(const n of w.nodes)if(n.type==='n8n-nodes-base.code')new(require('node:vm').Script)('(async()=>{'+n.parameters.jsCode+'\n})()');
  const out=path.join(root,'workflows','exported');await fs.mkdir(out,{recursive:true});
  for(const [name,w] of [['cartorio-manual.json',manual],['cartorio-agendado-desativado.json',scheduled]])await fs.writeFile(path.join(out,name),JSON.stringify(w,null,2));
  console.log(JSON.stringify({exportedWorkflows:2,bundleBytes:Buffer.byteLength(code),paidApiEnabled:false,schedulesEnabled:false,productionServicesAdded:0}));
}
main().catch(e=>{console.error(e.code||'AUTOMATION_BUILD_FAILED');process.exitCode=1;});
