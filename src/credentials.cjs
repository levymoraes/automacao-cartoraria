'use strict';
const fs=require('node:fs/promises');
const {requireCondition:check}=require('./errors.cjs');
const code=(service,suffix)=>(service==='openai'?'OPENAI':'LINEAR')+'_KEY_'+suffix;
const boundary=/^[\s\u200b\u200e\u200f\u2060\ufeff]+|[\s\u200b\u200e\u200f\u2060\ufeff]+$/gu;
function normalizeCredential(value,service='linear'){
 check(typeof value==='string',code(service,'INVALID_FORMAT'));
 const normalized=value.replace(boundary,'');
 check(normalized.length>0,code(service,'EMPTY'));
 check(normalized.length<=4096&&/^[A-Za-z0-9_-]+$/.test(normalized),code(service,'INVALID_CHARACTERS'));
 return normalized;
}
function decodeText(bytes,service){
 check(Buffer.isBuffer(bytes)&&bytes.length<=16384,code(service,'FILE_INVALID'));
 let encoding='utf8',data=bytes;
 if(bytes.length>=2&&bytes[0]===0xff&&bytes[1]===0xfe){encoding='utf16le';data=bytes.subarray(2);}
 else if(bytes.length>=2&&bytes[0]===0xfe&&bytes[1]===0xff){encoding='utf16be';data=bytes.subarray(2);}
 else if(bytes.length>=4&&bytes.length%2===0){
  const zeroOdd=bytes.every((b,i)=>i%2===0||b===0),zeroEven=bytes.every((b,i)=>i%2===1||b===0);
  if(zeroOdd&&!zeroEven)encoding='utf16le';else if(zeroEven&&!zeroOdd)encoding='utf16be';
 }
 let value;
 if(encoding.startsWith('utf16')){
  check(data.length%2===0,code(service,'ENCODING_INVALID'));
  if(encoding==='utf16be'){data=Buffer.from(data);data.swap16();}
  value=data.toString('utf16le');
 }else{value=data.toString('utf8');check(Buffer.from(value,'utf8').equals(data),code(service,'ENCODING_INVALID'));}
 return {value,encoding};
}
function decodeCredential(bytes,service='linear'){
 const decoded=decodeText(bytes,service),value=normalizeCredential(decoded.value,service);
 return {value,diagnostics:{valid:true,encoding:decoded.encoding,formattingAtBoundariesRemoved:value!==decoded.value,credentialFileModified:false}};
}
function inspectCredential(bytes,service='linear'){
 try{
  const decoded=decodeText(bytes,service),trimmed=decoded.value.replace(boundary,'');
  const diagnostics={encoding:decoded.encoding,formattingAtBoundariesRemoved:trimmed!==decoded.value,hasInternalNonAscii:/[^\x20-\x7e]/.test(trimmed),hasInternalWhitespace:/\s/.test(trimmed),credentialFileModified:false};
  try{normalizeCredential(decoded.value,service);return {...diagnostics,valid:true};}catch(e){return {...diagnostics,valid:false,code:e.code};}
 }catch(e){return {valid:false,code:e.code||code(service,'INVALID_FORMAT'),credentialFileModified:false};}
}
async function readCredential(filename,service='linear'){
 const bytes=await fs.readFile(filename);try{return decodeCredential(bytes,service).value;}finally{bytes.fill(0);}
}
module.exports={normalizeCredential,decodeCredential,inspectCredential,readCredential};
