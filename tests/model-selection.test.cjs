'use strict';
const {test}=require('node:test'),assert=require('node:assert/strict');
const {configurationForModel,assertConfiguration,accountResponse,estimateReservation}=require('../src/openai.cjs');
const base=require('../config/openai.json');
test('modelos: Terra low permitido explicitamente sem reduzir a reserva pelo custo do Luna',async()=>{
 const terra=configurationForModel(base,'gpt-6.1-terra','low'),luna=configurationForModel(base,'gpt-6-luna','xhigh');assertConfiguration(terra);assert.equal(terra.model,'gpt-6.1-terra');assert.equal(terra.reasoning.effort,'low');assert.equal(estimateReservation(terra),estimateReservation(luna)*10);
 assert.throws(()=>configurationForModel(base,'modelo-desconhecido','low'),e=>e.code==='OPENAI_CONFIGURATION_MISMATCH');
 const records=[],ledger={recordUsage:async r=>records.push(r)},raw={id:'FICTICIO',model:terra.model,service_tier:'default',usage:{input_tokens:150000,output_tokens:20000,input_tokens_details:{cached_tokens:0},output_tokens_details:{reasoning_tokens:1000}}};
 const r=await accountResponse(raw,terra,ledger,{reservedUsd:estimateReservation(terra)});assert.equal(r.costEstimateUsd,0.25);assert.equal(r.pricingBasis.longContext,false);assert.equal(r.pricingBasis.inputPerMillionUsd,1);assert.equal(records.length,1);
 const long=await accountResponse({...raw,usage:{...raw.usage,input_tokens:300000}},terra,ledger,{reservedUsd:estimateReservation(terra)});assert.equal(long.costEstimateUsd,0.75);assert.equal(long.pricingBasis.longContext,true);
});
