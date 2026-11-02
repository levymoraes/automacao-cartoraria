'use strict';
const {LinearClient}=require('./linear.cjs');
const {requireCondition:check}=require('./errors.cjs');
async function verifyLinearConnection({apiKey,config,transport}){
 check(config?.workspaceId&&config?.teamId,'LINEAR_CONFIGURATION_REQUIRED');
 // Endpoint fixo: o teste so consulta a identidade do workspace, sem mutations.
 const client=new LinearClient({apiKey,workspaceId:config.workspaceId,teamId:config.teamId,...(transport?{transport}:{})});
 const workspace=await client.assertScope();
 return {status:'LINEAR_CONNECTION_VERIFIED',authenticated:true,workspaceId:workspace.id,workspaceName:workspace.name,workspaceMatchesConfiguration:true,readOnly:true,openaiCalls:0};
}
module.exports={verifyLinearConnection};
