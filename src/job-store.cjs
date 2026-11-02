'use strict';
const fs=require('node:fs/promises');
const path=require('node:path');
const {randomUUID}=require('node:crypto');
const os=require('node:os');
const proc=require('node:process');
const {requireCondition:check}=require('./errors.cjs');
const uuid=s=>typeof s==='string'&&/^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$/i.test(s);
async function atomicJson(file,value){
  await fs.mkdir(path.dirname(file),{recursive:true});const tmp=file+'.'+randomUUID()+'.next';
  await fs.writeFile(tmp,JSON.stringify(value,null,2),{mode:0o600});await fs.rename(tmp,file);
}
async function processIdentity(pid=proc.pid){
  try{const stat=await fs.readFile('/proc/'+pid+'/stat','utf8');return stat.slice(stat.lastIndexOf(')')+2).split(' ')[19];}catch(e){if(e.code==='ENOENT')return null;throw e;}
}
async function withLock(file,fn){
  await fs.mkdir(path.dirname(file),{recursive:true});const token=randomUUID();let fd;
  const owner={token,pid:proc.pid,host:os.hostname(),start:await processIdentity(),createdAt:new Date().toISOString()};
  for(let attempt=0;attempt<2;attempt++){
    try{fd=await fs.open(file,'wx');await fd.writeFile(JSON.stringify(owner));break;}catch(e){
      check(e.code==='EEXIST','JOB_LOCK_FAILED');
      let old;try{old=JSON.parse(await fs.readFile(file,'utf8'));}catch{check(false,'JOB_LOCK_REVIEW_REQUIRED');}
      // Outro host pode estar vivo: exige recuperacao explicita, sem roubar seu lock.
      check(old.host===owner.host,'JOB_LOCK_OTHER_HOST_REVIEW_REQUIRED');
      let dead=old.start&&await processIdentity(old.pid)!==old.start;
      if(!old.start&&old.host===owner.host){try{proc.kill(old.pid,0);}catch(error){dead=error.code==='ESRCH';}}
      check(dead,'BATCH_ALREADY_RUNNING');
      // Guarda o vestigio do processo encerrado; nunca rouba lock de processo vivo.
      await fs.rename(file,file+'.stale-'+randomUUID());
    }
  }
  check(fd,'BATCH_ALREADY_RUNNING');
  try{return await fn();}finally{await fd.close();const current=JSON.parse(await fs.readFile(file,'utf8'));if(current.token===token)await fs.unlink(file);}
}
function jobStore(root){
  const jobs=path.join(root,'jobs');
  const filename=id=>{check(uuid(id),'INVALID_ISSUE_ID');return path.join(jobs,id+'.json');};
  return {
    root,
    async load(id){try{const state=JSON.parse(await fs.readFile(filename(id),'utf8'));check(state?.issueId===id&&typeof state.phase==='string','JOB_STATE_INVALID');return state;}catch(e){check(e.code==='ENOENT','JOB_STATE_INVALID');return null;}},
    async save(id,state){check(state.issueId===id,'JOB_ISSUE_MISMATCH');await atomicJson(filename(id),{...state,updatedAt:new Date().toISOString()});},
    async pendingIds(){await fs.mkdir(jobs,{recursive:true});const ids=[];for(const name of await fs.readdir(jobs)){if(!/^[a-f0-9-]{36}\.json$/i.test(name))continue;try{const state=JSON.parse(await fs.readFile(path.join(jobs,name),'utf8'));if(state.phase!=='complete')ids.push(name.slice(0,-5));}catch{ids.push(name.slice(0,-5));}}return ids;},
    lock:fn=>withLock(path.join(root,'batch.lock'),fn)
  };
}
// Apaga apenas artefatos enumerados dentro da raiz gerenciada. Recibos de deduplicacao ficam.
async function cleanArtifacts(root,retention,now=Date.now(),activeJobIds=[]){
  check(Number.isFinite(retention.artifactsDays)&&retention.artifactsDays>=1,'INVALID_RETENTION');
  const resolved=path.resolve(root),removed=[];await fs.mkdir(resolved,{recursive:true});
  check(!(await fs.lstat(resolved)).isSymbolicLink(),'UNSAFE_ARTIFACT_PATH');
  for(const name of await fs.readdir(resolved)){
    if(!/^[a-f0-9]{64}$/.test(name)||activeJobIds.includes(name))continue;const directory=path.join(resolved,name),info=await fs.lstat(directory);
    check(info.isDirectory()&&!info.isSymbolicLink(),'UNSAFE_ARTIFACT_PATH');
    for(const file of await fs.readdir(directory)){
      check(['result.docx','native.docx','response.json','reading.json','editing-response.json','audit.json','correction.json'].includes(file),'UNEXPECTED_ARTIFACT_FILE');
      const target=path.join(directory,file),stat=await fs.lstat(target);check(stat.isFile()&&!stat.isSymbolicLink()&&target.startsWith(resolved+path.sep),'UNSAFE_ARTIFACT_PATH');
      const days=file==='audit.json'?(retention.auditDays||90):retention.artifactsDays;
      if(now-stat.mtimeMs>days*86400000){await fs.unlink(target);removed.push(name+'/'+file);}
    }
  }return {removed:removed.length,receiptsPreserved:true};
}
module.exports={uuid,atomicJson,withLock,jobStore,cleanArtifacts};
