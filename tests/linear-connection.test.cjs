'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs/promises');
const path=require('node:path');
const root=path.resolve(__dirname,'..');
test('os dois nomes do verificador existem no diretorio scripts',async()=>{
 for(const name of ['Test-LinearConnection.ps1','Test-Linear-Connection.ps1','check-linear-connection.cjs'])await fs.access(path.join(root,'scripts',name));
});
test('verificador identifica workspace usando somente uma consulta autenticada',async()=>{
 const {verifyLinearConnection}=require('../src/linear-connection.cjs');let calls=0;
 const result=await verifyLinearConnection({apiKey:'CREDENCIAL-FICTICIA-NAO-IMPRIMIR',config:{workspaceId:'workspace-ficticio',teamId:'equipe-ficticia'},transport:async(url,options)=>{
  calls++;assert.equal(url,'https://api.linear.app/graphql');assert.equal(options.headers.Authorization,'CREDENCIAL-FICTICIA-NAO-IMPRIMIR');
  const body=JSON.parse(options.body);assert.equal(body.operationName,'CartorioScope');assert.equal(body.query.includes('mutation'),false);
  return {status:200,bytes:Buffer.from(JSON.stringify({data:{organization:{id:'workspace-ficticio',name:'Workspace Ficticio'}}}))};
 }});
 assert.equal(calls,1);assert.equal(result.workspaceName,'Workspace Ficticio');assert.equal(result.status,'LINEAR_CONNECTION_VERIFIED');assert.equal(result.openaiCalls,0);assert.equal(JSON.stringify(result).includes('CREDENCIAL-FICTICIA'),false);
});
test('verificador bloqueia credencial invalida e workspace diferente sem expor resposta',async()=>{
 const {verifyLinearConnection}=require('../src/linear-connection.cjs');
 for(const [status,data,code] of [[401,{},'LINEAR_AUTH_FAILED'],[200,{organization:{id:'outro-workspace',name:'Outro'}},'LINEAR_WORKSPACE_FORBIDDEN']]){
  await assert.rejects(()=>verifyLinearConnection({apiKey:'FICTICIA',config:{workspaceId:'esperado',teamId:'equipe'},transport:async()=>({status,bytes:Buffer.from(JSON.stringify({data}))})}),e=>e.code===code&&!e.message.includes('FICTICIA'));
 }
});
test('configuracao de conexao funciona sem bundle operacional preparado',async()=>{
 const {connectionConfig}=require('../scripts/check-linear-connection.cjs');
 const config=await connectionConfig(async name=>{
  if(name==='config/development-binding.json')return {workspaceId:'esperado',teamId:'equipe'};
  throw Object.assign(Error(),{code:'ENOENT'});
 });
 assert.deepEqual(config,{workspaceId:'esperado',teamId:'equipe'});
});
