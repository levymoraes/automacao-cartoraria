'use strict';
const {sha256,openDocx,paragraphsOf}=require('./docx.cjs');
const {schema,parseResponse}=require('./extraction.cjs');
const {requireCondition:check}=require('./errors.cjs');
const mandatoryRules=require('../config/mandatory-ai-rules.json').join('\n');
const modelCatalog=require('../config/openai-models.json');
function assertConfiguration(config){
  // Fast/high permanece reconhecido para recibos historicos; o padrao atual
  // e Standard/xhigh. Nao usar auto, que pode herdar Fast do projeto remoto.
  check(modelCatalog.models[config.model]?.reasoning.includes(config.reasoning?.effort)&&['default','fast','priority'].includes(config.service_tier),'OPENAI_CONFIGURATION_MISMATCH');
}
function configurationForModel(base,model,effort){
 const config=JSON.parse(JSON.stringify(base));config.model=model;config.reasoning={effort};assertConfiguration(config);
 const rate=modelCatalog.models[model],tier=['fast','priority'].includes(config.service_tier)?2:1;
 config.pricing={...config.pricing,inputPerMillionUsd:rate.inputPerMillionUsd*2*tier,outputPerMillionUsd:rate.outputPerMillionUsd*1.5*tier,source:modelCatalog.source,verifiedOn:modelCatalog.verifiedOn};config.compatibility={...config.compatibility,verifiedOn:modelCatalog.verifiedOn,modelReasoningStructuredOutputs:'https://developers.openai.com/api/docs/models/'+model};return config;
}
function systemInstructions(instructions){return mandatoryRules+'\n\nDETALHES INTERNOS DO CONTRATO DOCUMENTAL\n'+(instructions||'');}
function assertMandatoryInstructions(request){check(typeof request?.instructions==='string'&&request.instructions.startsWith(mandatoryRules+'\n\n'),'MANDATORY_AI_RULES_MISSING');}
function assertApproved(bytes,kind,approvals){
  const record=approvals.find(a=>a.kind===kind&&a.sha256===sha256(bytes)&&a.approved===true&&a.externalSharingAllowed===true&&['generated-synthetic','anonymized-reviewed'].includes(a.origin));
  check(record,'EXTERNAL_ASSET_NOT_APPROVED');return record;
}
async function buildRequest({templateBytes,profile,pdfs,approvals,instructions,config}){
  assertConfiguration(config);
  check(config.store===false,'STORE_MUST_BE_FALSE');
  check(profile.anonymizationApproved===true&&profile.externalSharingAllowed===true,'EXTERNAL_TEMPLATE_NOT_APPROVED');
  assertApproved(templateBytes,'template',approvals);
  const doc=await openDocx(templateBytes);check(doc.hash===profile.templateSha256,'TEMPLATE_HASH_MISMATCH');
  check(profile.schemaVersion===2&&profile.operation?.inventoryReviewed===true,'VARIABLE_INVENTORY_REVIEW_REQUIRED');
  check(pdfs.length>0&&pdfs.length<=4&&new Set(pdfs.map(p=>p.id)).size===pdfs.length&&pdfs.every(p=>/^[A-Za-z0-9_-]{1,80}$/.test(p.id)),'INVALID_INPUT_DOCUMENTS');
  check(pdfs.reduce((n,p)=>n+p.bytes.length,0)<=config.limits.maxPdfBytes,'PDF_TOTAL_SIZE_LIMIT');
  const documents=pdfs.map(pdf=>{
    check(pdf.bytes.length<=config.limits.maxPdfBytes&&pdf.bytes.subarray(0,5).toString()==='%PDF-','INVALID_PDF_INPUT');
    const approval=assertApproved(pdf.bytes,'pdf',approvals);
    check(Number.isInteger(approval.pages)&&approval.pages>0&&approval.pages<=config.limits.maxPdfPages,'PDF_PAGE_LIMIT');
    return {id:pdf.id,sha256:sha256(pdf.bytes),pages:approval.pages};
  });
  const text=[...doc.parts].filter(([n])=>/^word\/(document|header\d+|footer\d+|footnotes|endnotes)\.xml$/.test(n)).map(([part,b])=>({part,paragraphs:paragraphsOf(b.toString()).map(p=>({paragraph:p.index,text:p.text}))}));
  const context=JSON.stringify({templateSha256:doc.hash,entities:profile.operation.entities,fields:profile.fields,documents,template:text});
  check(context.length<=config.limits.maxTemplateCharacters,'TEMPLATE_CONTEXT_LIMIT');
  return {documents,request:{model:config.model,service_tier:config.service_tier,reasoning:config.reasoning,store:false,prompt_cache_options:{mode:'explicit'},max_output_tokens:config.max_output_tokens,
    instructions:systemInstructions(instructions),text:{format:{type:'json_schema',name:'minuta_extraction',strict:true,schema}},
    input:[{role:'user',content:[{type:'input_text',text:context},...pdfs.map(pdf=>({type:'input_file',filename:pdf.id+'.pdf',detail:'high',file_data:'data:application/pdf;base64,'+pdf.bytes.toString('base64')}))]}]}};
}
function estimateReservation(config){
  return (config.pricing.maxBillableInputTokens*config.pricing.inputPerMillionUsd+config.max_output_tokens*config.pricing.outputPerMillionUsd)/1e6+(config.nativeDocument?config.codeInterpreterSessionUsd:0);
}
async function authorizeRequest(built,config,ledger){
  check(config.live.enabled===true&&typeof config.live.budgetApprovalId==='string'&&config.live.budgetApprovalId.length>0,'PAID_API_NOT_AUTHORIZED');
  assertMandatoryInstructions(built.request);
  const reserve=estimateReservation(config);
  const fast=['fast','priority'].includes(config.service_tier);
  if(config.nativeDocument){const minimum={'1g':0.03,'4g':0.12}[config.codeInterpreterMemory];check(minimum&&config.codeInterpreterSessionUsd>=minimum,'UNSAFE_TOOL_COST_RESERVATION');}
  const rate=modelCatalog.models[config.model];check(rate,'OPENAI_CONFIGURATION_MISMATCH');
  check(config.pricing.inputPerMillionUsd>=rate.inputPerMillionUsd*2*(fast?2:1)&&config.pricing.outputPerMillionUsd>=rate.outputPerMillionUsd*1.5*(fast?2:1)&&config.pricing.maxBillableInputTokens>=1050000&&Number.isInteger(config.max_output_tokens)&&config.max_output_tokens>0&&config.max_output_tokens<=128000,'UNSAFE_COST_RESERVATION');
  check(Number.isFinite(reserve)&&reserve>0&&reserve<=config.live.maxBudgetUsd,'API_BUDGET_LIMIT');
  check(ledger&&typeof ledger.reserve==='function'&&typeof ledger.recordUsage==='function','DURABLE_BUDGET_LEDGER_REQUIRED');
  const requestSha256=sha256(Buffer.from(JSON.stringify(built.request)));
  await ledger.reserve({approvalId:config.live.budgetApprovalId,usd:reserve,maxCalls:config.live.maxCalls,budgetUsd:config.live.maxBudgetUsd,requestSha256});
  return {requestSha256,reservedUsd:reserve,approvalId:config.live.budgetApprovalId};
}
async function accountResponse(raw,config,ledger,reservation){
  const usage=raw?.usage;check(usage&&Number.isInteger(usage.input_tokens)&&usage.input_tokens>=0&&Number.isInteger(usage.output_tokens)&&usage.output_tokens>=0,'OPENAI_USAGE_MISSING');
  const details=usage.input_tokens_details||{};
  const writes=details.cache_write_tokens||0;
  const write12h=details.cache_write_tokens_details?.ttl_12h_tokens||0;
  // Ignorar descontos de leitura conserva uma estimativa superior. Escritas inesperadas
  // sao contabilizadas antes de bloquear, embora o request nao autorize cache writes.
  const rate=modelCatalog.models[config.model];check(rate,'OPENAI_CONFIGURATION_MISMATCH');
  const tier=['fast','priority'].includes(config.service_tier)?2:1,long=usage.input_tokens>modelCatalog.longContextThreshold,inputRate=rate.inputPerMillionUsd*tier*(long?2:1),outputRate=rate.outputPerMillionUsd*tier*(long?1.5:1);
  const tokenCost=((usage.input_tokens+Math.max(0,writes-write12h)*0.25+write12h*1.5)*inputRate+usage.output_tokens*outputRate)/1e6;
  const containers=config.nativeDocument?new Set((raw.output||[]).filter(o=>o.type==='code_interpreter_call'&&o.container_id).map(o=>o.container_id)).size:0,toolCost=containers*(config.codeInterpreterSessionUsd||0),cost=tokenCost+toolCost;
  const record={...reservation,responseId:raw.id,requestedTier:config.service_tier,actualTier:raw.service_tier,inputTokens:usage.input_tokens,outputTokens:usage.output_tokens,cachedInputTokens:details.cached_tokens||0,cacheWriteTokens:writes,reasoningTokens:usage.output_tokens_details?.reasoning_tokens??null,costEstimateUsd:cost,estimateOnly:true,pricingBasis:{model:config.model,inputPerMillionUsd:inputRate,outputPerMillionUsd:outputRate,longContext:long,source:modelCatalog.source,verifiedOn:modelCatalog.verifiedOn},...(config.nativeDocument?{tokenCostEstimateUsd:tokenCost,toolCostEstimateUsd:toolCost,toolContainers:containers}: {})};
  await ledger.recordUsage(record);
  check(raw.model===config.model||raw.model?.startsWith(config.model+'-'),'OPENAI_MODEL_MISMATCH');
  check(config.service_tier==='default'?raw.service_tier==='default':['fast','priority'].includes(raw.service_tier),'OPENAI_SERVICE_TIER_MISMATCH');
  check(writes===0,'UNEXPECTED_CACHE_WRITE');
  check(cost<=reservation.reservedUsd||config.nativeDocument&&config.live?.budgetPolicy==='user-unlimited','ACTUAL_USAGE_EXCEEDS_RESERVATION');
  return record;
}
async function requestLive(args,{apiKey,ledger,transport=fetch,saveResponse=async()=>{}}={}){
  const {config}=args;
  check(config.live.enabled===true,'PAID_API_NOT_AUTHORIZED');
  check(typeof apiKey==='string'&&apiKey.trim().length>0,'OPENAI_KEY_REQUIRED');
  const built=await buildRequest(args);
  const reservation=await authorizeRequest(built,config,ledger);
  const {raw,httpStatus}=await transmitRequest(built.request,config,{apiKey,transport,saveResponse,reservation});
  if(httpStatus<200||httpStatus>=300){if(raw.usage)await accountResponse(raw,config,ledger,reservation);check(false,'OPENAI_HTTP_FAILED');}
  const usage=await accountResponse(raw,config,ledger,reservation);
  return {raw,documents:built.documents,usage,reservation};
}
async function transmitRequest(request,config,{apiKey,transport=fetch,saveResponse=async()=>{},reservation}={}){
  check(config.live.enabled===true&&reservation?.approvalId===config.live.budgetApprovalId&&reservation.requestSha256===sha256(Buffer.from(JSON.stringify(request))),'PAID_API_NOT_AUTHORIZED');
  assertMandatoryInstructions(request);
  check(typeof apiKey==='string'&&apiKey.trim().length>0,'OPENAI_KEY_REQUIRED');
  // Sem retries automaticos: timeout ou falha pode ter sido cobrada.
  let response;
  try{response=await transport('https://api.openai.com/v1/responses',{method:'POST',redirect:'error',headers:{Authorization:'Bearer '+apiKey,'Content-Type':'application/json'},body:JSON.stringify(request),signal:AbortSignal.timeout(config.timeoutMs)});}catch{check(false,'OPENAI_TRANSPORT_FAILED');}
  let raw;try{raw=await readLimited(response,config.limits.maxResponseBytes);}catch{check(false,'OPENAI_RESPONSE_INVALID');}
  await saveResponse({raw,httpStatus:response.status,...reservation});
  return {raw,httpStatus:response.status};
}
async function interpretLive(args,options){
  const result=await requestLive(args,options);
  return {...parseResponse(result.raw,args.profile,result.documents),usage:result.usage};
}
async function readLimited(response,limit){
  const reader=response.body.getReader(),chunks=[];let length=0;
  try{while(true){const {done,value}=await reader.read();if(done)break;length+=value.length;check(length<=limit,'OPENAI_RESPONSE_LIMIT');chunks.push(Buffer.from(value));}}finally{await reader.cancel();}
  return JSON.parse(Buffer.concat(chunks).toString('utf8'));
}
module.exports={assertApproved,assertConfiguration,configurationForModel,buildRequest,interpretLive,requestLive,authorizeRequest,accountResponse,estimateReservation,transmitRequest,systemInstructions,assertMandatoryInstructions};
