'use strict';
const fs=require('node:fs/promises'),path=require('node:path'),{execFileSync}=require('node:child_process');
const root=path.resolve(__dirname,'..');
const ignored=new Set(['.git','node_modules','storage','reports','release','.cache']);
async function main(){
 let files;try{files=execFileSync('git',['ls-files','-z'],{cwd:root,encoding:'utf8',stdio:['ignore','pipe','ignore']}).split('\0').filter(Boolean);}catch{files=[];}
 if(!files.length){files=[];async function walk(rel=''){for(const e of await fs.readdir(path.join(root,rel),{withFileTypes:true})){if(ignored.has(e.name))continue;if(e.isSymbolicLink())throw Error('PUBLIC_SYMLINK_FORBIDDEN');const name=rel?rel+'/'+e.name:e.name;if(e.isDirectory())await walk(name);else files.push(name);}}await walk();}
 const findings=[];
 const blocked=/^(?:storage|reports|release|node_modules|\.cache)\/|(?:^|\/)(?:[^/]+\.local\.[^/]+|development-binding\.json|modelo06-anonymization\.json)$|\.(?:docx?|pdf|png|jpe?g|zip|db|sqlite\d*|pem|key|pfx|log)$/i;
 const patterns=[/\b(?:sk-(?:proj-|svcacct-)?[A-Za-z0-9_-]{40,}|gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}|lin_api_[A-Za-z0-9]{30,})\b/,/-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/,/https:\/\/[^\s"']+[?&](?:X-Amz-Signature|signature|token)=[A-Za-z0-9%_-]{15,}/i,/[A-Z]:[\\/]Users[\\/][^\s"']+/i];
 function validCpf(value){const d=value.replace(/\D/g,'');if(/^(\d)\1+$/.test(d))return false;return [9,10].every(n=>{const r=11-[...d.slice(0,n)].reduce((sum,c,i)=>sum+Number(c)*(n+1-i),0)%11;return Number(d[n])===(r>=10?0:r);});}
 for(const name of files){if(blocked.test(name)||/^\.env(?:\.|$)/.test(name)&&name!=='.env.example'){findings.push({file:name,reason:'private-file'});continue;}const info=await fs.lstat(path.join(root,name));if(info.isSymbolicLink()){findings.push({file:name,reason:'symlink'});continue;}const content=await fs.readFile(path.join(root,name),'utf8');if(patterns.some(re=>re.test(content))||[...content.matchAll(/\b\d{3}\.\d{3}\.\d{3}-\d{2}\b/g)].some(m=>validCpf(m[0])))findings.push({file:name,reason:'potential-sensitive-content'});}
 if(findings.length){console.error(JSON.stringify({status:'PUBLICATION_BLOCKED',findings},null,2));process.exitCode=1;return;}
 console.log(JSON.stringify({status:'PUBLIC_SOURCE_CHECK_PASSED',filesChecked:files.length,scope:'Public source only; excludes ignored local runtime. Heuristic check does not replace manual review.'}));
}
main().catch(()=>{console.error('PUBLIC_SOURCE_CHECK_FAILED');process.exitCode=1;});
