'use strict';
const {buildRequest}=require('./openai.cjs');
const {parseResponse}=require('./extraction.cjs');
const {applyAuthorized}=require('./docx.cjs');
const {requireCondition:check}=require('./errors.cjs');
async function interpretMock(payload){
  check(payload.mode==='mock','WORKFLOW_MOCK_ONLY');
  const templateBytes=Buffer.from(payload.templateBase64,'base64');
  const pdfs=payload.pdfs.map(p=>({id:p.id,bytes:Buffer.from(p.base64,'base64')}));
  const built=await buildRequest({templateBytes,profile:payload.profile,pdfs,approvals:payload.approvals,instructions:payload.instructions,config:payload.config});
  const {plan}=parseResponse(payload.mockResponse,payload.profile,built.documents);
  return {mode:'mock',templateBase64:payload.templateBase64,profile:payload.profile,plan,
    requestSummary:{model:built.request.model,serviceTier:built.request.service_tier,reasoning:built.request.reasoning.effort,store:built.request.store,pdfCount:pdfs.length,pdfInputType:'input_file',pdfImageDetail:'high'},
    simulation:true,apiCalls:0,costUsd:0};
}
async function applyMapping(prepared){
  const result=await applyAuthorized(Buffer.from(prepared.templateBase64,'base64'),prepared.profile,prepared.plan);
  return {bytes:result.bytes,report:{status:'validated-synthetic-docx',simulation:prepared.simulation,apiCalls:prepared.simulation?0:1,costUsd:prepared.usage?.costEstimateUsd||0,request:prepared.requestSummary,audit:result.audit}};
}
module.exports={interpretMock,applyMapping};
