'use strict';
const {sha256,applyAuthorized}=require('./docx.cjs');
const {buildRequest,authorizeRequest,accountResponse}=require('./openai.cjs');
const {extractResponse,validateExtraction}=require('./extraction.cjs');
const {compareGroundTruth}=require('./ground-truth.cjs');
const {createLedger}=require('./budget.cjs');
const {requireCondition:check}=require('./errors.cjs');
function assertTestInputs(payload,manifest,truth){
  check(manifest?.status==='FICTIONAL_INPUTS_VERIFIED'&&manifest.officialTemplateIncluded===false,'FICTIONAL_PREFLIGHT_REQUIRED');
  check(payload.profile.origin==='generated-synthetic','OFFICIAL_TEMPLATE_NOT_ALLOWED_IN_TEST');
  const template=manifest.assets.find(a=>a.kind==='template');
  check(template?.sha256===sha256(Buffer.from(payload.templateBase64,'base64'))&&sha256(Buffer.from(JSON.stringify(payload.profile)))===manifest.profileSha256,'PREFLIGHT_TEMPLATE_MISMATCH');
  check(payload.pdfs.length===1&&payload.pdfs[0].id==='pacote-ficticio','PREFLIGHT_INPUT_MISMATCH');
  const pdf=manifest.assets.find(a=>a.kind==='pdf'&&a.scenario===payload.scenario);
  check(pdf?.sha256===sha256(Buffer.from(payload.pdfs[0].base64,'base64')),'PREFLIGHT_INPUT_MISMATCH');
  check(manifest.truthSha256?.[payload.scenario]===sha256(Buffer.from(JSON.stringify(truth)))&&truth.scenario===payload.scenario,'GROUND_TRUTH_PROVENANCE_MISMATCH');
  check(manifest.instructionsSha256===sha256(Buffer.from(payload.instructions)),'PREFLIGHT_INSTRUCTIONS_MISMATCH');
}
async function prepare(payload,{manifest,truth,config,approval=null,ledgerPath}={}){
  check(['mock','live'].includes(payload.mode),'INVALID_EXECUTION_MODE');
  assertTestInputs(payload,manifest,truth);
  const args={templateBytes:Buffer.from(payload.templateBase64,'base64'),profile:payload.profile,pdfs:payload.pdfs.map(p=>({id:p.id,bytes:Buffer.from(p.base64,'base64')})),approvals:payload.approvals,instructions:payload.instructions,config};
  const built=await buildRequest(args);
  const requestSha256=sha256(Buffer.from(JSON.stringify(built.request)));
  let reservation=null;
  if(payload.mode==='live'){
    check(approval?.authorized===true&&typeof approval.evidence==='string'&&approval.evidence.trim().length>=10&&Date.parse(approval.expiresAt)>Date.now(),'PAID_API_NOT_AUTHORIZED');
    check(approval.scenario===payload.scenario&&approval.requestSha256===requestSha256,'APPROVAL_REQUEST_MISMATCH');
    check(approval.maxCalls===1&&approval.maxBudgetUsd>0&&approval.maxBudgetUsd<=0.5,'APPROVAL_BUDGET_MISMATCH');
    config=structuredClone(config);
    config.live={enabled:true,budgetApprovalId:approval.approvalId,maxCalls:approval.maxCalls,maxBudgetUsd:approval.maxBudgetUsd};
    check(typeof ledgerPath==='string'&&ledgerPath.length>0,'DURABLE_BUDGET_LEDGER_REQUIRED');
    reservation=await authorizeRequest(built,config,createLedger(ledgerPath));
  }
  return {mode:payload.mode,scenario:payload.scenario,templateBase64:payload.templateBase64,profile:payload.profile,truth,documents:built.documents,request:built.request,config,reservation,ledgerPath,
    requestSha256,requestSummary:{model:built.request.model,serviceTier:built.request.service_tier,reasoning:built.request.reasoning.effort,store:false,pdfImageDetail:'high'},
    // Apenas o mock usa este campo; request foi construido por selecao explicita.
    ...(payload.mode==='mock'?{mockResponse:payload.mockResponse}:{})};
}
async function evaluate(prepared,raw,{httpStatus=200,account=true}={}){
  let usage=null;
  if(prepared.mode==='live'&&account){
    if(raw?.usage)usage=await accountResponse(raw,prepared.config,createLedger(prepared.ledgerPath),prepared.reservation);
    else check(false,'OPENAI_USAGE_MISSING');
  }
  check(httpStatus>=200&&httpStatus<300,'OPENAI_HTTP_FAILED');
  const value=extractResponse(raw);
  const evaluation=compareGroundTruth(value,prepared.truth,prepared.documents);
  let plan;
  try{plan=validateExtraction(value,prepared.profile,prepared.documents);}
  catch(e){
    if(e.code==='PENDING_INFORMATION'&&!evaluation.documentExpected)return {report:{status:'expected-block',code:e.code,scenario:prepared.scenario,simulation:prepared.mode==='mock',apiCalls:prepared.mode==='live'?1:0,documentGenerated:false,evaluation,usage}};
    throw e;
  }
  check(evaluation.documentExpected,'PENDING_INFORMATION');
  const result=await applyAuthorized(Buffer.from(prepared.templateBase64,'base64'),prepared.profile,plan);
  return {bytes:result.bytes,report:{status:'validated-synthetic-docx',scenario:prepared.scenario,simulation:prepared.mode==='mock',apiCalls:prepared.mode==='live'?1:0,documentGenerated:true,evaluation,usage,request:prepared.requestSummary,audit:result.audit}};
}
module.exports={assertTestInputs,prepare,evaluate};
