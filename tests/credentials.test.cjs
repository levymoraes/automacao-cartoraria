'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const {validateHeaderValue}=require('node:http');
const {LinearClient}=require('../src/linear.cjs');
test('chave ficticia com Unicode invisivel nas bordas nao chega como ERR_INVALID_CHAR',async()=>{
 const c=new LinearClient({apiKey:'\u200bFICTICIA\u200b',workspaceId:'ficticio',teamId:'ficticio',transport:async(url,request)=>{
  validateHeaderValue('Authorization',request.headers.Authorization);assert.equal(request.headers.Authorization,'FICTICIA');return {status:200,bytes:Buffer.from(JSON.stringify({data:{organization:{id:'ficticio',name:'Ficticio'}}}))};
 }});await c.assertScope();
});
test('leitura de UTF8, UTF16LE e UTF16BE preserva arquivo e nao expoe o valor',()=>{
 const {decodeCredential}=require('../src/credentials.cjs');const expected='lin_api_FICTICIA_123';
 const be=Buffer.from(expected+'\r\n','utf16le');be.swap16();
 for(const bytes of [Buffer.from(expected+'\r\n'),Buffer.concat([Buffer.from([0xef,0xbb,0xbf]),Buffer.from(expected)]),Buffer.concat([Buffer.from([0xff,0xfe]),Buffer.from(expected,'utf16le')]),Buffer.from(expected,'utf16le'),Buffer.concat([Buffer.from([0xfe,0xff]),be])]){
  const before=Buffer.from(bytes),r=decodeCredential(bytes,'linear');assert.equal(r.value,expected);assert.deepEqual(bytes,before);assert.equal(JSON.stringify(r.diagnostics).includes(expected),false);
 }
});
test('caractere invisivel interno, quebra interna e formato invalido bloqueiam antes do transporte',()=>{
 const {normalizeCredential}=require('../src/credentials.cjs');
 for(const value of ['FICTICIA\u200bINTERNA','FICTICIA\nINTERNA','FICTICIA\0','CHAVE inválida','"FICTICIA"','Bearer FICTICIA'])assert.throws(()=>normalizeCredential(value,'linear'),e=>e.code==='LINEAR_KEY_INVALID_CHARACTERS'&&!e.message.includes(value));
});
test('diagnostico de credencial invalida devolve somente classificacoes seguras',()=>{
 const {inspectCredential}=require('../src/credentials.cjs');const secret='FICTICIA\u200bINTERNA';const d=inspectCredential(Buffer.from(secret),'linear');assert.equal(d.valid,false);assert.equal(d.hasInternalNonAscii,true);assert.equal(JSON.stringify(d).includes('FICTICIA'),false);
});
