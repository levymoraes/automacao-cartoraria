'use strict';
const fs=require('node:fs/promises');
const path=require('node:path');
const {verifyLinearConnection}=require('../src/linear-connection.cjs');
const {requireCondition:check}=require('../src/errors.cjs');
const {readCredential,inspectCredential}=require('../src/credentials.cjs');
const root=path.resolve(__dirname,'..');
async function connectionConfig(read=async p=>JSON.parse(await fs.readFile(path.join(root,p),'utf8'))){
 try{const bundle=await read('storage/automation/runtime/bundle.json');return {workspaceId:bundle.config.linear.workspaceId,teamId:bundle.config.linear.teamId};}
 catch(e){check(e.code==='ENOENT','LINEAR_CONFIGURATION_INVALID');}
 try{const binding=await read('config/development-binding.json');return {workspaceId:binding.workspaceId,teamId:binding.teamId};}
 catch(e){check(e.code==='ENOENT','LINEAR_CONFIGURATION_INVALID');check(false,'LINEAR_CONFIGURATION_REQUIRED');}
}
async function main(){
 const keyPath=path.join(root,'storage','secrets','linear-api-key.txt');
 if(process.argv.includes('--diagnose-credential')){
  const bytes=await fs.readFile(keyPath);let result;try{result=inspectCredential(bytes,'linear');}finally{bytes.fill(0);}
  console.log(JSON.stringify({status:'LINEAR_CREDENTIAL_DIAGNOSTIC',...result,networkRequests:0,openaiCalls:0},null,2));
  if(!result.valid)process.exitCode=1;return;
 }
 const config=await connectionConfig();
 const apiKey=await readCredential(keyPath,'linear');
 const result=await verifyLinearConnection({apiKey,config});
 await fs.mkdir(path.join(root,'reports'),{recursive:true});
 await fs.writeFile(path.join(root,'reports','linear-connection.json'),JSON.stringify(result,null,2));console.log(JSON.stringify(result,null,2));
}
if(require.main===module)main().catch(e=>{console.error(/^[A-Z0-9_]+$/.test(e.code||'')?e.code:'LINEAR_CONNECTION_FAILED');process.exitCode=1;});
module.exports={connectionConfig,main};
