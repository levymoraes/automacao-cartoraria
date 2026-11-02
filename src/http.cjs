'use strict';
const https=require('node:https');
const http=require('node:http');
const {URL}=require('node:url');
const {requireCondition:check}=require('./errors.cjs');
// Sem redirects, proxies implicitos ou retries. Nunca registra URL assinada ou headers.
function request(url,{method='GET',headers={},body,timeoutMs=30000,maxBytes=20971520,allowHttp=false}={}){
  const target=new URL(url);check(target.protocol==='https:'||(allowHttp&&target.protocol==='http:'),'HTTPS_REQUIRED');
  check(!target.username&&!target.password,'URL_CREDENTIALS_FORBIDDEN');
  return new Promise((resolve,reject)=>{
    const fail=code=>{const e=new Error(code);e.code=code;reject(e);};
    const req=(target.protocol==='https:'?https:http).request(target,{method,headers},res=>{
      const chunks=[];let size=0;
      res.on('data',chunk=>{size+=chunk.length;if(size>maxBytes){req.destroy();fail('HTTP_RESPONSE_LIMIT');}else chunks.push(chunk);});
      res.on('error',()=>fail('HTTP_TRANSPORT_FAILED'));
      res.on('end',()=>resolve({status:res.statusCode,headers:res.headers,bytes:Buffer.concat(chunks)}));
    });
    req.setTimeout(timeoutMs,()=>{req.destroy();fail('HTTP_TIMEOUT');});
    req.on('error',()=>fail('HTTP_TRANSPORT_FAILED'));if(body)req.write(body);req.end();
  });
}
module.exports={request};
