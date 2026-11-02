'use strict';
const {randomBytes,scryptSync,createCipheriv,createDecipheriv}=require('node:crypto');
const JSZip=require('jszip');
const {requireCondition:check}=require('./errors.cjs');
const magic=Buffer.from('CARTORIO-BACKUP-1\n');
function encrypt(bytes,password){
 check(typeof password==='string'&&password.length>=12,'BACKUP_PASSWORD_TOO_SHORT');check(bytes.length<=536870912,'BACKUP_SIZE_LIMIT');
 const salt=randomBytes(16),iv=randomBytes(12),key=scryptSync(password,salt,32);
 try{const cipher=createCipheriv('aes-256-gcm',key,iv);cipher.setAAD(magic);const encrypted=Buffer.concat([cipher.update(bytes),cipher.final()]);return Buffer.concat([magic,salt,iv,cipher.getAuthTag(),encrypted]);}finally{key.fill(0);}
}
function decrypt(bytes,password){
 check(bytes.length<=536871000&&bytes.subarray(0,magic.length).equals(magic),'BACKUP_FORMAT_INVALID');const offset=magic.length;
 const salt=bytes.subarray(offset,offset+16),iv=bytes.subarray(offset+16,offset+28),tag=bytes.subarray(offset+28,offset+44),key=scryptSync(password,salt,32);
 try{const decipher=createDecipheriv('aes-256-gcm',key,iv);decipher.setAAD(magic);decipher.setAuthTag(tag);return Buffer.concat([decipher.update(bytes.subarray(offset+44)),decipher.final()]);}catch{check(false,'BACKUP_AUTHENTICATION_FAILED');}finally{key.fill(0);}
}
async function inspectArchive(bytes){
 check(bytes.length<=536870912,'BACKUP_SIZE_LIMIT');const z=await JSZip.loadAsync(bytes,{checkCRC32:true});const files=[];let expanded=0;
 for(const [name,entry] of Object.entries(z.files)){
  check(name===entry.unsafeOriginalName||entry.dir,'BACKUP_UNSAFE_PATH');
  check(!name.includes('\\')&&!name.startsWith('/')&&!name.includes(':')&&!name.split('/').some(p=>p==='..'),'BACKUP_UNSAFE_PATH');
  check(!((entry.unixPermissions||0)&0xf000)||((entry.unixPermissions||0)&0xf000)!==0xa000,'BACKUP_SYMLINK_FORBIDDEN');
  if(entry.dir)continue;
  check(name==='n8n-volume.tar.gz'||/^automation\/(runtime|state)\/[A-Za-z0-9_.\/-]+$/.test(name)||/^secrets\/(cartorio\/)?(linear|openai)-api-key\.txt$/.test(name)||name==='manifest.json','BACKUP_UNEXPECTED_FILE');
  const size=entry._data?.uncompressedSize;check(Number.isSafeInteger(size)&&size>=0&&size<=536870912,'BACKUP_SIZE_LIMIT');expanded+=size;check(expanded<=536870912&&files.length<10000,'BACKUP_SIZE_LIMIT');files.push(name);
 }
 check(files.includes('manifest.json')&&files.includes('n8n-volume.tar.gz'),'BACKUP_INCOMPLETE');return {zip:z,files};
}
module.exports={encrypt,decrypt,inspectArchive};
