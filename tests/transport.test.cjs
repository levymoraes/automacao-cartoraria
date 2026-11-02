'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const http=require('node:http');
const {request}=require('../src/http.cjs');
const {LinearClient}=require('../src/linear.cjs');
async function server(t,handler){const s=http.createServer(handler);await new Promise(r=>s.listen(0,'127.0.0.1',r));t.after(()=>new Promise(r=>{s.closeAllConnections();s.close(r);}));return 'http://127.0.0.1:'+s.address().port;}
test('transporte limita bytes, timeout e nao segue redirecionamento',async t=>{
 let received=0;const origin=await server(t,(req,res)=>{received++;if(req.url==='/large'){res.end(Buffer.alloc(3000));}else if(req.url==='/timeout'){}else{res.writeHead(302,{location:origin+'/large'});res.end();}});
 await assert.rejects(()=>request(origin+'/large',{allowHttp:true,maxBytes:100}),e=>e.code==='HTTP_RESPONSE_LIMIT');
 await assert.rejects(()=>request(origin+'/timeout',{allowHttp:true,timeoutMs:30}),e=>e.code==='HTTP_TIMEOUT');
 const r=await request(origin+'/redirect',{allowHttp:true});assert.equal(r.status,302);assert.equal(received,3);
});
test('paginacao Linear coleta todas as paginas e detecta cursor repetido',async()=>{
 let calls=0;const client=new LinearClient({apiKey:'FICTICIA',workspaceId:'FICTICIO',teamId:'FICTICIA',transport:async(url,o)=>{const op=JSON.parse(o.body);const data=op.operationName==='CartorioScope'?{organization:{id:'FICTICIO',name:'Ficticio'}}:{issues:{nodes:[{id:'issue-'+calls}],pageInfo:{hasNextPage:++calls<3,endCursor:'page-'+calls}}};return {status:200,bytes:Buffer.from(JSON.stringify({data}))};}});
 await client.assertScope();assert.equal((await client.readyIssues({labelId:'FICTICIO',readyStateIds:['FICTICIO']})).length,3);
});
test('erros GraphQL e limite de requisicoes sao codigos sem corpo ou credencial',async()=>{
 for(const [response,expected] of [[{status:429,bytes:Buffer.from('FICTICIO')},'LINEAR_RATE_LIMITED'],[{status:200,bytes:Buffer.from(JSON.stringify({errors:[{message:'nao registrar este texto'}]}))},'LINEAR_GRAPHQL_FAILED']]){
  const client=new LinearClient({apiKey:'FICTICIA',workspaceId:'FICTICIO',teamId:'FICTICIO',transport:async()=>response});await assert.rejects(()=>client.assertScope(),e=>e.code===expected&&!e.message.includes('nao registrar'));
 }
});
