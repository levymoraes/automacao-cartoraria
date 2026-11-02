'use strict';
const {test}=require('node:test');
const assert=require('node:assert/strict');
const JSZip=require('jszip');
const {encrypt,decrypt,inspectArchive}=require('../src/backup.cjs');
test('backup criptografado autenticado recupera dados ficticios e rejeita senha incorreta',()=>{
 const data=Buffer.from('DOCUMENTACAO E CHAVE EXCLUSIVAMENTE FICTICIAS'),password='senha-ficticia-para-teste';const bytes=encrypt(data,password);
 assert.equal(bytes.includes(data),false);assert.deepEqual(decrypt(bytes,password),data);assert.throws(()=>decrypt(bytes,'senha-errada-ficticia'),e=>e.code==='BACKUP_AUTHENTICATION_FAILED');
});
test('backup adulterado ou senha curta nao sao aceitos',()=>{
 assert.throws(()=>encrypt(Buffer.from('ficticio'),'curta'),e=>e.code==='BACKUP_PASSWORD_TOO_SHORT');const encrypted=encrypt(Buffer.from('ficticio'),'senha-ficticia-longa');encrypted[encrypted.length-1]^=1;assert.throws(()=>decrypt(encrypted,'senha-ficticia-longa'),e=>e.code==='BACKUP_AUTHENTICATION_FAILED');
});
test('restauracao inspeciona caminhos e estrutura antes de extrair qualquer arquivo',async()=>{
 const z=new JSZip();z.file('manifest.json','{}');z.file('n8n-volume.tar.gz','TAR-FICTICIO');z.file('automation/state/jobs/arquivo.json','{}');z.file('secrets/linear-api-key.txt','FICTICIA');const bytes=await z.generateAsync({type:'nodebuffer'});assert.equal((await inspectArchive(bytes)).files.length,4);
 z.file('../fora.txt','FICTICIO');await assert.rejects(()=>z.generateAsync({type:'nodebuffer'}).then(inspectArchive),e=>e.code==='BACKUP_UNSAFE_PATH');
});
test('backup inclui chaves separadas do cartorio sem aceitar outros arquivos de segredos',async()=>{
 const z=new JSZip();z.file('manifest.json','{}');z.file('n8n-volume.tar.gz','TAR-FICTICIO');
 for(const service of ['linear','openai'])z.file('secrets/cartorio/'+service+'-api-key.txt','CREDENCIAL-EXCLUSIVAMENTE-FICTICIA');
 assert.equal((await inspectArchive(await z.generateAsync({type:'nodebuffer'}))).files.length,4);
 z.file('secrets/cartorio/outro-arquivo.txt','FICTICIO');await assert.rejects(()=>z.generateAsync({type:'nodebuffer'}).then(inspectArchive),e=>e.code==='BACKUP_UNEXPECTED_FILE');
});
