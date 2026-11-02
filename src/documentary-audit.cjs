'use strict';
// Conferencia da leitura existente. Nao chama IA, nao edita nem decide requisitos juridicos.
const {requireCondition:check}=require('./errors.cjs');
const {claims,visibleText}=require('./template-presentation.cjs');
const object=properties=>({type:'object',additionalProperties:false,required:Object.keys(properties),properties});
const string={type:'string'},nullable={type:['string','null']},list=items=>({type:'array',items});
const choice=values=>({type:'string',enum:values});
const purposes=['identity','role','civil_status','declared_transaction_value','municipal_valuation','itbi_base','paid_amount','payment_method','payment_date','contract_confirmation','property','certificate','administrative','other'];
const sourceSchema=object({kind:choice(['document','task']),documentId:nullable,page:{type:['integer','null'],minimum:1},entityId:string,property:string,purpose:choice(purposes),quote:string});
const findingSchema=object({category:choice(['A','B','C']),subject:string,message:string,critical:{type:'boolean'},sources:list(sourceSchema)});
const auditSchema=object({
 operation:object({actType:string,objectNature:string,participants:list(object({id:string,name:string,role:string,maritalStatus:choice(['married','single','divorced','widowed','unknown']),gender:choice(['masculine','feminine','unknown']),spouseId:nullable,representsId:nullable,sources:list(sourceSchema)}))}),
 compatibility:object({status:choice(['compatible','incompatible','undetermined']),reasons:list(string),requiredModelCharacteristics:list(string)}),
 financialFacts:list(object({entityId:string,nature:string,purpose:choice(['declared_transaction_value','municipal_valuation','itbi_base','paid_amount','payment_method','payment_date']),value:string,sources:list(sourceSchema)})),
 records:list(object({documentId:string,entityId:string,kind:choice(['certificate','property_record']),holderRole:choice(['owner','coowner','declarant','taxpayer','unknown']),issuedOn:nullable,expiresOn:nullable,holder:nullable,propertyId:nullable,registration:nullable,area:nullable,identifier:nullable,identifierType:nullable,sources:list(sourceSchema)})),
 paragraphReviews:list(object({part:string,paragraph:{type:'integer',minimum:0},status:choice(['verified','not_applicable','needs_confirmation','incompatible']),reason:string,sources:list(sourceSchema)})),
 checks:list(object({name:choice(['participants','source_ownership','old_data','values_dates','grammar','duplicates_fragments','fixed_assertions']),passed:{type:'boolean'},message:string})),
 findings:list(findingSchema)
});
const normalize=s=>s.normalize('NFKC').replace(/\s+/gu,' ').trim().toLocaleUpperCase('pt-BR');
const fold=s=>s.normalize('NFD').replace(/\p{M}/gu,'').toLowerCase();
function roleCode(role){
 const code=fold(role).split(';')[0].trim().replace(/[-\s]+/gu,'_');
 if(/^(?:seller|buyer)(?:_spouse|_administrator|_shareholder|_minor_shareholder)?$/u.test(code))return code;
 if(/^(?:spouse|conjuge)(?:_|$)/u.test(code)){
  if(/(?:^|_)(?:buyer|comprador[ae]?)(?:_|$)/u.test(code))return 'buyer_spouse';
  if(/(?:^|_)(?:seller|vendedor[ae]?)(?:_|$)/u.test(code))return 'seller_spouse';
  return 'spouse';
 }
 if(/^vendedor[ae]?(?:_|$)/u.test(code))return 'seller';
 if(/^comprador[ae]?(?:_|$)/u.test(code))return 'buyer';
 return code;
}
function validSource(prepared,s,entityId,property){
 return s.entityId===entityId&&(!property||s.property===property)&&s.quote.trim()&&(s.kind==='task'?s.documentId===null&&s.page===null&&((prepared.task.title||'')+'\n'+prepared.task.description).includes(s.quote):prepared.documents.some(d=>d.id===s.documentId&&Number.isInteger(s.page)&&s.page>0&&s.page<=d.pages));
}
function uniqueFindings(findings){
 const grouped=new Map();for(const f of findings){const key=normalize(f.category+' '+f.subject),prior=grouped.get(key);if(!prior){grouped.set(key,{...f,sources:[...f.sources]});continue;}prior.critical||=f.critical;if(!prior.message.includes(f.message))prior.message+='; '+f.message;const sources=new Map([...prior.sources,...f.sources].map(s=>[JSON.stringify(s),s]));prior.sources=[...sources.values()];}return [...grouped.values()];
}
function recordRelations(metadata,assignments,participants,invalidSlots){
 const aliases=new Map(metadata.entities.map(e=>[e.id,e.id]));
 for(const e of assignments)if(!invalidSlots.has(e.slotId)&&participants.has(e.participantId))aliases.set(e.participantId,e.slotId);
 const marriages=[];for(const c of metadata.couples||[]){
  const a=assignments.find(e=>e.slotId===c.a),b=assignments.find(e=>e.slotId===c.b),pa=participants.get(a?.participantId),pb=participants.get(b?.participantId);
  if(pa&&pb&&!invalidSlots.has(c.a)&&!invalidSlots.has(c.b)&&pa.spouseId===pb.id&&pb.spouseId===pa.id)marriages.push(c);
 }
 return (a,b)=>{const left=aliases.get(a),right=aliases.get(b);return !!left&&!!right&&(left===right||marriages.some(c=>left===c.id&&[c.a,c.b].includes(right)||right===c.id&&[c.a,c.b].includes(left)));};
}
function auditReading(prepared,value){
 const audit=value.audit,metadata=prepared.fieldMetadata,findings=[],invalidSlots=new Set(),paragraphNotes=new Map();
 const note=(p,topic,anchor='end',occurrence=0)=>{const key=p.part+':'+p.paragraph+':'+anchor+':'+occurrence;if(!paragraphNotes.has(key))paragraphNotes.set(key,{part:p.part,paragraph:p.paragraph,anchor,occurrence,topics:new Set()});paragraphNotes.get(key).topics.add(topic);};
 const add=(category,subject,message,critical=true,sources=[])=>findings.push({category,subject,message,critical,sources});
 const people=metadata.entities.filter(e=>['person','organization'].includes(e.kind)),participants=new Map(),assigned=new Map();
 for(const p of audit.operation.participants){
  check(p.id.trim()&&p.name.trim()&&!participants.has(p.id),'AUDIT_PARTICIPANT_INVALID');participants.set(p.id,p);
  const mapped=value.entities.filter(e=>e.participantId===p.id&&people.some(slot=>slot.id===e.slotId)),ownIds=new Set([p.id,...(mapped.length===1?[mapped[0].slotId]:[])]);
  if(!p.sources.some(s=>ownIds.has(s.entityId)&&validSource(prepared,s,s.entityId))){for(const e of value.entities.filter(e=>e.participantId===p.id))invalidSlots.add(e.slotId);add('B','Identidade sem fonte comprovada','A fonte do participante '+p.id+' não pôde ser validada. Seus campos ficam pendentes para conferir a identidade; os demais participantes são tratados separadamente.');}
 }
 const names=new Set();for(const p of participants.values()){
  if(names.has(normalize(p.name)))add('B','Participantes','Identidades repetidas na composição informada.');names.add(normalize(p.name));
  if(p.spouseId){const spouse=participants.get(p.spouseId);if(!spouse||spouse.spouseId!==p.id||p.maritalStatus!=='married'||spouse.maritalStatus!=='married')add('B','Participantes','Vínculo conjugal não comprovado ou inconsistente.');}
  if(p.representsId&&!participants.has(p.representsId)||['attorney','legal_representative'].includes(p.role)&&(!p.representsId||p.representsId===p.id))add('B','Representação','Identificar a parte representada e comprovar os poderes de representação.');
 }
 if(!audit.operation.actType.trim()||!audit.operation.objectNature.trim())add('B','Operação','Confirmar o ato solicitado e a natureza do objeto.');
 const expectedAct=prepared.model.actType;if(expectedAct&&expectedAct!=='minuta'&&fold(expectedAct)!==fold(audit.operation.actType))add('C','Ato solicitado','O ato identificado nos documentos difere do ato do modelo selecionado.');
 const modelName=fold(prepared.model.name||'');
 if(/\brural\b/u.test(modelName)&&!/^rural$/iu.test(audit.operation.objectNature)||/\burban[oa]\b/u.test(modelName)&&!/^urban[oa]$/iu.test(audit.operation.objectNature))add('C','Objeto','A natureza do objeto difere da natureza prevista no modelo.');
 if(audit.compatibility.status!=='compatible')add('C','Modelo solicitado',audit.compatibility.reasons.join('; ')||'Não foi possível comprovar a compatibilidade do modelo.');
 for(const e of value.entities){
  const slot=people.find(s=>s.id===e.slotId),p=participants.get(e.participantId);
  if(!slot||!p||assigned.has(p.id)){invalidSlots.add(e.slotId);if(p&&assigned.has(p.id))invalidSlots.add(assigned.get(p.id));add('C','Composição','O modelo exige participantes sem correspondência única na operação.');continue;}
  assigned.set(p.id,slot.id);
  const expected=fold(slot.role).replace(/[- ]/g,'_'),role=expected.includes('spouse')||expected.includes('conjuge')?'spouse':expected==='vendedor'?'seller':expected==='comprador'?'buyer':expected==='procurador'?'attorney':expected;
  const actualRole=roleCode(p.role),spouseOwner=expected.endsWith('_spouse')&&actualRole===expected.replace(/_spouse$/u,'')&&p.spouseId&&participants.get(p.spouseId)?.spouseId===p.id;
  if(role!==actualRole&&expected!==actualRole&&!spouseOwner||normalize(e.name||'')!==normalize(p.name)){invalidSlots.add(slot.id);add('B','Correspondência / '+slot.id,'Confirmar a identidade ou o papel associado a esta posição; seus campos não foram preenchidos.');}
  if((e.spouseSlotId||null)!==(slot.spouseId||null))add('C','Composição','Os vínculos do modelo não correspondem à operação; o texto fixo foi preservado para revisão.');
 }
 for(const slot of people){const e=value.entities.find(e=>e.slotId===slot.id),p=participants.get(e?.participantId),spouse=participants.get(p?.spouseId);
  if(!p){invalidSlots.add(slot.id);add('C','Posições excedentes','O modelo prevê posição sem participante comprovado: '+slot.id+' ('+slot.role+'). Não foram inventados participantes.');}
  if(p&&((slot.spouseId||null)!==(spouse?assigned.get(spouse.id)||null:null)))add('C','Composição','A relação conjugal prevista pelo modelo não corresponde à operação.');
 }
 if([...participants.keys()].some(id=>!assigned.has(id)))add('C','Composição','A operação contém participante ou representante sem posição correspondente no modelo.');
 const sourceEntityIds=new Set([...participants.keys(),...metadata.entities.map(e=>e.id)]);
 const sourceValid=s=>sourceEntityIds.has(s.entityId)&&validSource(prepared,s,s.entityId);
 for(const f of audit.findings){const sources=f.sources.filter(sourceValid);findings.push({...f,sources});if(sources.length!==f.sources.length)add('B','Referências da revisão','Uma observação da leitura contém referência de fonte inválida. Conferir a observação no documento; a referência não foi aceita.');}
 const relatedRecord=recordRelations(metadata,value.entities,participants,invalidSlots),records=[],financialFacts=[],quarantinedRecords=[],quarantinedFinancialFacts=[];
 for(const p of participants.values())if(p.sources.some(s=>!sourceValid(s)||!relatedRecord(p.id,s.entityId)))add('B','Referências do participante','O participante '+p.id+' possui referência auxiliar sem vínculo ou fonte comprovada. A referência não confirma o vínculo; os campos de identidade só usam suas próprias fontes válidas.');
 const date=s=>/^\d{4}-\d{2}-\d{2}$/.test(s||'')&&Number.isFinite(Date.parse(s))&&new Date(s).toISOString().slice(0,10)===s;
 for(const [index,r]of audit.records.entries()){
  const sourcesValid=sourceEntityIds.has(r.entityId)&&prepared.documents.some(d=>d.id===r.documentId)&&r.sources.length&&r.sources.every(s=>s.kind==='document'&&s.documentId===r.documentId&&sourceValid(s));
  const linked=sourcesValid&&r.sources.some(s=>relatedRecord(r.entityId,s.entityId));
  const datesValid=(r.issuedOn===null||date(r.issuedOn))&&(r.expiresOn===null||date(r.expiresOn));
  if(!linked||!datesValid){
   const reason=!linked?'vínculo da certidão e suas fontes não comprovado':'data documental inválida';quarantinedRecords.push({index,reason,record:r});
   add('B','Conferência de registro documental','Registro '+(index+1)+' ('+r.documentId+'): '+reason+'. Registro isolado das comparações automáticas; conferir na revisão humana.',true,r.sources.filter(sourceValid));continue;
  }
  records.push(r);
 }
 for(const [index,f]of audit.financialFacts.entries()){
  if(sourceEntityIds.has(f.entityId)&&f.value.trim()&&f.sources.length&&f.sources.every(s=>validSource(prepared,s,f.entityId)&&s.purpose===f.purpose)){financialFacts.push(f);continue;}
  quarantinedFinancialFacts.push({index,fact:f});add('B','Conferência financeira','Informação financeira '+(index+1)+' com fonte ou finalidade não comprovada. Informação isolada; conferir o documento antes de aproveitá-la.',true,f.sources.filter(sourceValid));
 }
 const expiredSource=s=>records.some(r=>r.expiresOn&&r.expiresOn<prepared.referenceDate&&r.documentId===s.documentId&&r.sources.some(origin=>origin.page===s.page&&relatedRecord(origin.entityId,s.entityId)));
 const checked=new Set();for(const original of audit.paragraphReviews){const r={...original,sources:original.sources.filter(sourceValid)};
  const key=r.part+':'+r.paragraph,p=prepared.templateParagraphs.find(p=>p.part===r.part&&p.paragraph===r.paragraph);
  if(!p||checked.has(key)){add('B','Referências da revisão','Revisão com posição inexistente ou repetida: '+r.part+', parágrafo '+(r.paragraph+1)+'. Conferir a leitura; a referência adicional não foi aproveitada.');if(p)note(p,'conferir referências repetidas da revisão');continue;}checked.add(key);
  if(r.sources.length!==original.sources.length){r.status='needs_confirmation';add('B','Referências da revisão','Revisão do parágrafo '+(r.paragraph+1)+' de '+r.part+' contém fonte inválida, retirada da conferência automática.');}
  const presentClaims=claims.filter(({pattern})=>pattern.test(visibleText(p.text)));
  if(r.status==='needs_confirmation'||r.status==='incompatible'){add(r.status==='incompatible'?'C':'B',r.part+' / parágrafo '+(r.paragraph+1),r.reason||'Confirmar a condição declarada no texto fixo.',true,r.sources);if(!presentClaims.length&&!p.text.includes('{{'))note(r,'revisar este trecho; ver tarefa');}
  const unconfirmed=[];for(const {key:claim,purposes,label,pattern}of presentClaims)if(r.status!=='verified'||!r.sources.some(s=>fold(s.property)===claim&&purposes.includes(s.purpose)&&!expiredSource(s))){unconfirmed.push(label||claim);const count=claim==='certificacoes'?[...visibleText(p.text).matchAll(new RegExp(pattern.source,pattern.flags+'g'))].length:1;for(let i=0;i<count;i++)note(r,'confirmar '+(label||claim),claim,i);}
  if(unconfirmed.length)add('B','Declarações não confirmadas',r.part+', parágrafo '+(r.paragraph+1)+': '+unconfirmed.join(', ')+'. Texto preservado para revisão; não significa confirmação ou aprovação do ato.',true,r.sources);
 }
 const unreviewedParagraphs=prepared.templateParagraphs.filter(p=>p.text.trim()&&!checked.has(p.part+':'+p.paragraph));
 for(const p of unreviewedParagraphs){
  add('B','Revisão não informada',p.part+', parágrafo '+(p.paragraph+1)+': a resposta não informou a revisão deste trecho. Texto preservado; conferir integralmente antes de concluir o ato.',true);
  // Nao aumentar cabecalhos/rodapes repetidos em todas as paginas. O aviso
  // vermelho fica no corpo; os trechos exatos permanecem listados na tarefa.
  const anchor=p.part==='word/document.xml'?p:prepared.templateParagraphs.find(p=>p.part==='word/document.xml'&&p.text.trim());
  check(anchor,'FIELD_DRAFT_LABEL_ANCHOR_REQUIRED');
  note(anchor,p.part==='word/document.xml'?'trecho não revisado pela leitura automática; conferir integralmente':'revisar cabeçalhos, rodapés e notas com trechos não revisados pela leitura automática');
 }
 const required=['participants','source_ownership','old_data','values_dates','grammar','duplicates_fragments','fixed_assertions'];
 if(audit.checks.length!==required.length||new Set(audit.checks.map(c=>c.name)).size!==required.length)add('B','Conferência incompleta','A resposta omitiu ou repetiu itens da lista de conferência. O funcionário deverá conferir integralmente o conteúdo antes da aprovação.');
 for(const c of audit.checks)if(!c.passed)add('B','Conferência do conteúdo',c.name+': '+(c.message||'Conferir o conteúdo antes de finalizar a minuta.'));
 for(const r of records){
  if(r.expiresOn&&r.expiresOn<prepared.referenceDate)add('B','Validade documental','Documento '+r.documentId+' com validade expressa encerrada em '+r.expiresOn+'. A serventia deve avaliar a exigência de atualização.',false,r.sources);
  if(r.issuedOn&&r.issuedOn>prepared.referenceDate||r.issuedOn&&r.expiresOn&&r.expiresOn<r.issuedOn)add('B','Datas documentais','Datas de emissão/validade inconsistentes em '+r.documentId+'.',true,r.sources);
 }
 for(let i=0;i<records.length;i++)for(const b of records.slice(i+1)){
  const a=records[i];if(!a.propertyId||a.propertyId!==b.propertyId)continue;
  const differences=['holder','registration','area','identifier'].filter(k=>(k!=='holder'||['owner','coowner'].includes(a.holderRole)&&a.holderRole===b.holderRole)&&(k!=='identifier'||a.identifierType&&b.identifierType&&normalize(a.identifierType)===normalize(b.identifierType))&&a[k]&&b[k]&&normalize(a[k])!==normalize(b[k]));
  if(differences.length)add('B','Identificação do imóvel','Diferenças em '+differences.map(k=>k+': '+a[k]+' / '+b[k]).join('; ')+' entre '+a.documentId+' e '+b.documentId+'. Conferir a vinculação documental; nenhuma causa jurídica foi presumida.',true,[...a.sources,...b.sources]);
 }
 return {findings:uniqueFindings(findings),records,financialFacts,quarantinedRecords,quarantinedFinancialFacts,requiredModelCharacteristics:audit.compatibility.requiredModelCharacteristics,participants,invalidSlots,unreviewedParagraphs:unreviewedParagraphs.map(({part,paragraph})=>({part,paragraph})),paragraphNotes:[...paragraphNotes.values()].map(p=>({...p,text:'pendente: '+[...p.topics].join('; ')+'. Ver tarefa.'}))};
}
function formatFindings(findings,characteristics=[]){
 const labels={A:'A — Informação ausente',B:'B — Informação disponível / confirmação',C:'C — Incompatibilidade do modelo'},lines=[];
 for(const category of ['A','B','C']){const rows=uniqueFindings(findings).filter(f=>f.category===category);if(!rows.length)continue;lines.push(labels[category]);for(const f of rows){const sources=[...new Set(f.sources.map(s=>s.kind==='task'?'solicitação':s.documentId+', página '+s.page))];lines.push('- '+f.subject+': '+f.message+(sources.length?' [Fontes: '+sources.join('; ')+']':''));}}
 if(characteristics.length)lines.push('Características necessárias no modelo: '+[...new Set(characteristics)].join('; ')+'. Manter o modelo solicitado até autorização do funcionário.');
 return lines.join('\n');
}
module.exports={sourceSchema,auditSchema,validSource,auditReading,uniqueFindings,formatFindings,normalize,fold};
