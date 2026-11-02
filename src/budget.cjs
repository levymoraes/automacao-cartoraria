'use strict';
const fs=require('node:fs/promises');
const path=require('node:path');
const {requireCondition:check}=require('./errors.cjs');
function createLedger(filename){
  async function locked(update){
    await fs.mkdir(path.dirname(filename),{recursive:true});let lock;
    try{lock=await fs.open(filename+'.lock','wx');}catch{check(false,'BUDGET_LEDGER_LOCKED');}
    try{
      let state={approvals:{},usage:[]};
      try{state=JSON.parse(await fs.readFile(filename,'utf8'));}catch(e){check(e.code==='ENOENT','BUDGET_LEDGER_INVALID');}
      const result=update(state);
      const temporary=filename+'.next';
      await fs.writeFile(temporary,JSON.stringify(state,null,2),{mode:0o600});
      await fs.rename(temporary,filename);return result;
    }finally{await lock.close();await fs.unlink(filename+'.lock');}
  }
  return {
    reserve:({approvalId,usd,maxCalls,budgetUsd,requestSha256})=>locked(state=>{
      check(Number.isFinite(usd)&&usd>0&&Number.isInteger(maxCalls)&&maxCalls>0,'INVALID_BUDGET');
      state.requests||={};
      if(requestSha256){
        check(/^[a-f0-9]{64}$/.test(requestSha256),'INVALID_REQUEST_FINGERPRINT');
        check(!state.requests[requestSha256],'DUPLICATE_PAID_REQUEST');
      }
      const previous=state.approvals[approvalId]||{calls:0,reservedUsd:0};
      check(previous.calls<maxCalls&&previous.reservedUsd+usd<=budgetUsd,'API_BUDGET_LIMIT');
      state.approvals[approvalId]={calls:previous.calls+1,reservedUsd:previous.reservedUsd+usd};
      if(requestSha256)state.requests[requestSha256]={approvalId,reservedUsd:usd,status:'reserved-no-automatic-retry',at:new Date().toISOString()};
    }),
    recordUsage:usage=>locked(state=>{
      if(usage.responseId&&state.usage.some(u=>u.responseId===usage.responseId&&u.requestSha256===usage.requestSha256))return;
      state.usage.push({...usage,at:new Date().toISOString()});
      const request=state.requests?.[usage.requestSha256];
      if(request)Object.assign(request,{status:'response-accounted',actualTier:usage.actualTier,inputTokens:usage.inputTokens,outputTokens:usage.outputTokens,costEstimateUsd:usage.costEstimateUsd});
    })
  };
}
module.exports={createLedger};
