'use strict';
// Servidor de teste: nenhuma dependencia externa, apenas arquivos ficticios recebidos.
const http=require('node:http');
const {randomUUID,createHash}=require('node:crypto');
const hash=b=>createHash('sha256').update(b).digest('hex');
function createSimulator({bundle,assets,cases=['completo','digitalizada','incompleto','contraditorio'],controls={}}){
  const issues=new Map(),files=new Map(),stats={scopeChecks:0,downloads:0,uploadAllocations:0,puts:0,attachmentsCreated:0,commentsCreated:0,reviewUpdates:0,openaiCalls:0,issues:[],issueReads:[]};let origin;
  const config=bundle.config,model=bundle.catalog[0];
  const asset=(bytes,name)=>{const id=randomUUID();files.set(id,{bytes,name});return id;};
  const modelFile=asset(assets.template,'modelo-ficticio.docx');
  const plain=(id,state,description,attachments=[],labels=[])=>({id,identifier:'FIC-'+id.slice(0,6),url:'https://linear.invalid/fictional/'+id,description,team:{id:config.linear.teamId},state:{id:state,name:state===config.linear.reviewStateId?'In Review':'Todo',type:'unstarted'},labels:{nodes:labels.map(id=>({id}))},attachments:{nodes:attachments,pageInfo:{hasNextPage:false,endCursor:null}},comments:[]});
  const modelIssue=plain(model.issueId,config.linear.reviewStateId,'Modelo inteiramente ficticio');issues.set(model.issueId,modelIssue);
  const fail=(name)=>{if(controls[name]>0){controls[name]--;return true;}return false;};
  let seeded=false;
  function seed(){if(seeded)return;seeded=true;
    modelIssue.attachments.nodes.push({id:model.attachmentId,title:'modelo-ficticio.docx',url:origin+'/assets/'+modelFile,metadata:{},issue:{id:model.issueId}});
    for(const scenario of cases){
      const id=randomUUID(),pdfCase=['completo','digitalizada','incompleto','contraditorio'].includes(scenario)?scenario:'completo';
      const order={requestId:randomUUID(),actType:model.actType,templateId:scenario==='modelo-inexistente'?'MODELO-NAO-CADASTRADO':model.id,templateVersion:model.version};
      const inputId=asset(scenario==='pdf-invalido'?Buffer.from('NAO E PDF'):assets[pdfCase],'documentos-'+scenario+'.pdf');
      files.get(inputId).scenario=scenario;
      const a={id:randomUUID(),title:'documentos-'+scenario+'.pdf',url:origin+'/assets/'+inputId,metadata:{},issue:{id}};
      const issue=plain(id,config.linear.readyStateIds[0],'```cartorio-json\n'+JSON.stringify(order)+'\n```',scenario==='sem-pdf'?[]:[a],[config.linear.requestLabelId]);issue.scenario=scenario;issues.set(id,issue);stats.issues.push({id,scenario});
    }
  }
  const server=http.createServer(async(req,res)=>{
    const send=(status,value)=>{res.writeHead(status,{'content-type':'application/json'});res.end(JSON.stringify(value));};
    if(req.url==='/stats'){send(200,{...stats,results:[...issues.values()].filter(i=>i.scenario).map(i=>({issueId:i.id,scenario:i.scenario,stateId:i.state.id,outputs:i.attachments.nodes.filter(a=>a.metadata?.cartorioJobId).length,comments:i.comments.length}))});return;}
    if(req.url==='/v1/responses'){stats.openaiCalls++;send(500,{error:'OPENAI_MUST_NOT_RUN'});return;}
    if(req.url.startsWith('/assets/')){
      stats.downloads++;const file=files.get(req.url.split('/')[2]);
      if(!file?.bytes||file.scenario==='falha-download'||fail('download')){send(404,{error:'fictional-download-failure'});return;}
      res.writeHead(200,{'content-type':'application/octet-stream'});res.end(file.bytes);return;
    }
    if(req.url.startsWith('/put/')){
      stats.puts++;const id=req.url.split('/')[2],file=files.get(id);if(!file||req.headers['x-test-upload']!=='approved'){send(403,{});return;}
      if(fail('put')){send(503,{});return;}const chunks=[];for await(const c of req)chunks.push(c);file.bytes=Buffer.concat(chunks);send(200,{});return;
    }
    if(req.url!=='/graphql'){send(404,{});return;}
    if(req.headers.authorization!=='FICTICIA-SEM-ACESSO'||controls.auth){send(401,{});return;}
    const chunks=[];for await(const c of req)chunks.push(c);let body;try{body=JSON.parse(Buffer.concat(chunks));}catch{send(400,{});return;}
    const {operationName:op,variables:v}=body;
    if(fail(op)){send(503,{});return;}
    const connection=nodes=>({nodes,pageInfo:{hasNextPage:false,endCursor:null}});
    let data;
    switch(op){
      case 'CartorioScope':stats.scopeChecks++;data={organization:{id:controls.wrongWorkspace?'00000000-0000-4000-8000-000000000099':config.linear.workspaceId,name:'Workspace SIMULADO'}};break;
      case 'CartorioReady':data={issues:connection([...issues.values()].filter(i=>i.scenario&&config.linear.readyStateIds.includes(i.state.id)).map(i=>({id:i.id})))};break;
      case 'CartorioIssueStatus':{const i=issues.get(v.id);data={issue:i?{id:i.id,team:i.team,state:i.state,labels:i.labels}:null};break;}
      case 'CartorioIssue':{stats.issueReads.push(v.id);const i=issues.get(v.id);data={issue:i?{...i,comments:undefined}:null};break;}
      case 'CartorioUpload':{stats.uploadAllocations++;const id=asset(null,v.filename);data={fileUpload:{success:true,uploadFile:{assetUrl:origin+'/assets/'+id,uploadUrl:origin+'/put/'+id,headers:[{key:'x-test-upload',value:'approved'}]}}};break;}
      case 'CartorioAttach':{const i=issues.get(v.input.issueId);if(!i){data={attachmentCreate:{success:false}};break;}
        let a=i.attachments.nodes.find(a=>a.url===v.input.url);if(!a){a={id:randomUUID(),issue:{id:i.id}};i.attachments.nodes.push(a);stats.attachmentsCreated++;}Object.assign(a,v.input);data={attachmentCreate:{success:true,attachment:a}};break;}
      case 'CartorioComments':{const i=issues.get(v.id);data={issue:i?{id:i.id,comments:connection(i.comments)}:null};break;}
      case 'CartorioComment':{const i=issues.get(v.input.issueId);const c={id:randomUUID(),body:v.input.body};i.comments.push(c);stats.commentsCreated++;data={commentCreate:{success:true,comment:c}};break;}
      case 'CartorioReview':{const i=issues.get(v.id);if(!i.attachments.nodes.some(a=>a.metadata?.cartorioJobId)){send(400,{error:'NO_OUTPUT'});return;}i.state={id:v.input.stateId,name:'In Review',type:'started'};stats.reviewUpdates++;data={issueUpdate:{success:true,issue:{id:i.id,state:i.state}}};break;}
      default:send(400,{error:'UNKNOWN_OPERATION'});return;
    }
    send(200,{data});
  });
  return {server,issues,files,stats,controls,
    async start({port=0,host='127.0.0.1',publicOrigin}={}){await new Promise(r=>server.listen(port,host,r));origin=publicOrigin||'http://127.0.0.1:'+server.address().port;seed();return origin;},
    close:()=>new Promise(r=>{server.closeAllConnections();server.close(r);})};
}
module.exports={createSimulator};
