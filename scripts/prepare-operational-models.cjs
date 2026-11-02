'use strict';
// Preparacao unica dos seis modelos autorizados. Nao interpreta novas issues,
// nao consulta APIs e nao contem dados pessoais literais no codigo.
const fs=require('node:fs/promises'),path=require('node:path');
const Docxtemplater=require('docxtemplater'),PizZip=require('pizzip');
const {openDocx,paragraphsOf,parseXml,sha256}=require('../src/docx.cjs');
const {patchText}=require('../src/anonymize.cjs');
const {requireCondition:check}=require('../src/errors.cjs');
const root=path.resolve(__dirname,'..'),out=path.join(root,'storage/templates/operational');
const numberWords='zero|um|uma|dois|duas|três|tres|quatro|cinco|seis|sete|oito|nove|dez|onze|doze|treze|catorze|quatorze|quinze|dezesseis|dezassete|dezessete|dezoito|dezenove|vinte|trinta|quarenta|cinquenta|sessenta|setenta|oitenta|noventa|cem|cento|duzentos|duzentas|trezentos|trezentas|quatrocentos|quatrocentas|quinhentos|quinhentas|seiscentos|seiscentas|setecentos|setecentas|oitocentos|oitocentas|novecentos|novecentas|mil|milhão|milhões|bilhão|bilhões|vírgula|virgula';
const numeral=()=>new RegExp('(?<![\\p{L}])(?:'+numberWords+')(?:[\\s,]+(?:e\\s+)?(?:'+numberWords+'))*(?![\\p{L}])','giu');
const owners={'seller':'VENDEDOR','seller-spouse':'CONJ_VENDEDOR','buyer':'COMPRADOR','buyer-spouse':'CONJ_COMPRADOR','buyer1':'COMPRADOR1','buyer1-spouse':'CONJ_COMPRADOR1','buyer2':'COMPRADOR2','buyer2-spouse':'CONJ_COMPRADOR2','seller-company':'EMPRESA_VENDEDORA',administrator1:'ADMINISTRADOR1',administrator2:'ADMINISTRADOR2',shareholder:'SOCIO','minor-shareholder':'SOCIO_MENOR',property:'IMOVEL',transaction:'ATO','marriage-seller':'CAS_VENDEDORES','marriage-buyer':'CAS_COMPRADORES','marriage-buyer1':'CAS_COMPRADORES1','marriage-buyer2':'CAS_COMPRADORES2'};
const ascii=s=>s.normalize('NFD').replace(/\p{M}/gu,'').replace(/[^A-Za-z0-9]+/g,'_').replace(/^_|_$/g,'').toUpperCase();
function refine(profile,stories){
 const fields=[],retained=[];
 function add(f,a,b,property){
  while(a<b&&/\s/u.test(f.oldText[a]))a++;while(b>a&&/\s/u.test(f.oldText[b-1]))b--;
  if(a===b)return;
  const item={...f,id:f.entityId+'.'+property,property,start:f.start+a,end:f.start+b,oldText:f.oldText.slice(a,b)};
  const overlap=fields.find(x=>x.part===item.part&&x.paragraph===item.paragraph&&x.start<item.end&&x.end>item.start);
  if(overlap){check(overlap.start===item.start&&overlap.end===item.end&&overlap.entityId===item.entityId,'OPERATIONAL_FIELD_OVERLAP');return;}
  fields.push(item);
 }
 function capture(f,re,property){for(const m of f.oldText.matchAll(re)){const at=m.index+m[0].indexOf(m[1]);add(f,at,at+m[1].length,property);}}
 function address(f,prefix=f.property){
  const text=f.oldText,leading=/^(?:na|no|à|a)\s+/iu.exec(text)?.[0].length||0;let at=leading,index=0;
  for(const piece of text.slice(leading).split(',')){
   let start=at,end=at+piece.length;while(start<end&&/\s/u.test(text[start]))start++;while(end>start&&/\s/u.test(text[end-1]))end--;
   const s=text.slice(start,end);at+=piece.length+1;if(!s)continue;index++;
   const postal=/^CEP\s*:?\s*(.+)$/iu.exec(s),number=/^(?:n[º°.]|no(?=\s)|número\b)\s*\.?\s*(.+)$/iu.exec(s),lot=/^(?:lote|quadra)\s+(.+)$/iu.exec(s),floor=/^(\d+)º\s+(?:andar|pavimento)$/iu.exec(s),city=/^(.+)\/([A-Z]{2})$/u.exec(s);
   if(postal)add(f,start+s.indexOf(postal[1]),end,prefix+'.cep');
   else if(number)add(f,start+s.indexOf(number[1]),end,prefix+'.numero');
   else if(index===2&&/^\d+[A-Za-z]?$/u.test(s))add(f,start,end,prefix+'.numero');
   else if(lot)add(f,start+s.indexOf(lot[1]),end,prefix+'.'+(/^lote/iu.test(s)?'lote':'quadra'));
   else if(floor)add(f,start,start+floor[1].length,prefix+'.andar');
   else if(/^(?:sala|apartamento|apto\.?|bloco|conjunto|cj\.?)\s/iu.test(s))add(f,start,end,prefix+'.complemento');
   else if(city){add(f,start,start+city[1].length,prefix+'.municipio');add(f,end-2,end,prefix+'.uf');}
   else add(f,start,end,prefix+'.'+(index===1?'logradouro':/^s\/n[ºo]?$/iu.test(s)?'numero':index===2?'complemento':'bairro'));
  }
 }
 function amount(f,prefix=f.property){
  capture(f,/(\d(?:[\d.,]*\d)?)/gu,prefix+'.numero');
  let count=0;for(const group of f.oldText.matchAll(/\(([^)]*)\)/gu))for(const m of group[1].matchAll(numeral())){const at=group.index+1+m.index;add(f,at,at+m[0].length,prefix+'.extenso'+(++count));}
 }
 function description(f){
  const exclusions=[];let addressCount=0;
  const meanings=new Map();
  const meaning=(at,raw)=>{
   const before=f.oldText.slice(Math.max(0,at-120),at).normalize('NFD').replace(/\p{M}/gu,'').toLowerCase(),tail=f.oldText.slice(at+raw.length).replace(/^\s*(?:m²|m|ha)?\s*(?:\([^)]*\))?\s*/u,'').toLowerCase();
   for(const [pattern,label] of [[/area construida(?: de|:)?\s*$/u,'area_construida'],[/area(?: total)?(?: apurada)?(?: de|:)?\s*$/u,'area_total'],[/perimetro(?: de|:)?\s*$/u,'perimetro'],[/inscricao imobiliaria\s*n[º°o.]?\s*$/u,'inscricao_imobiliaria'],[/codigo de controle\s*n[º°o.]?\s*$/u,'certidao_codigo'],[/validade ate\s*$/u,'certidao_validade'],[/\bem\s*$/u,'certidao_emissao'],[/\bn[º°o.]\s*$/u,'certidao_numero']])if(pattern.test(before))return label;
   for(const [pattern,label] of [[/^de frente/u,'frente'],[/^de fundos/u,'fundos'],[/^lado direito/u,'lado_direito'],[/^lado esquerdo/u,'lado_esquerdo'],[/^sala de /u,'salas_tv'],[/^area de servico/u,'areas_servico'],[/^área de serviço/u,'areas_servico'],[/^área gourmet/u,'areas_gourmet'],[/^quartos/u,'quartos'],[/^su[ií]te/u,'suites'],[/^banheiros/u,'banheiros'],[/^salas/u,'salas'],[/^cozinha/u,'cozinhas'],[/^dep[oó]sito/u,'depositos'],[/^varandas/u,'varandas'],[/^garagem/u,'garagens']])if(pattern.test(tail))return label;
   return null;
  };
  const re=/((?:Rua|Avenida|Av\.|Praça|Estrada)\s+[\s\S]*?)(?=,\s*(?:CEP|inscri|com\s+inscri|feita\s+com)|$)/giu;
  for(const m of f.oldText.matchAll(re)){const start=m.index,end=start+m[1].length;exclusions.push([start,end]);address({...f,start:f.start+start,oldText:m[1]},'endereco'+(++addressCount));}
  const inside=(a,b)=>exclusions.some(([s,e])=>s<b&&e>a);
  let n=0;for(const m of f.oldText.matchAll(/\d(?:[\d.,/\-]*\d)?/gu))if(!inside(m.index,m.index+m[0].length)){const label=meaning(m.index,m[0])||'numero'+(++n);meanings.set(m.index,label);add(f,m.index,m.index+m[0].length,'descricao.'+label+'.numero');}
  let words=0;for(const group of f.oldText.matchAll(/\(([^)]*)\)/gu))if(!inside(group.index,group.index+group[0].length)){const before=f.oldText.slice(0,group.index),quantity=/\d(?:[\d.,]*\d)?\s*(?:m²|m|ha)?\s*$/u.exec(before),label=quantity&&meanings.get(quantity.index);let part=0;for(const m of group[1].matchAll(numeral()))add(f,group.index+1+m.index,group.index+1+m.index+m[0].length,'descricao.'+(label||'valor'+(++words))+'.extenso'+(++part));}
  for(const [property,pattern] of [
   ['pavimentos',/edificaç[aã]o possui\s+(um pavimento térreo)/giu],['tipo_edificacao',/uma\s+(casa residencial)\s+com/giu],
   ['telhado',/telhado\s+(cerâmico de variadas inclinações)/giu],['material_paredes',/paredes de\s+(bloco)/giu],
   ['pintura',/com\s+(pintura acrílica)\s+sobre/giu],['base_pintura',/sobre o\s+(reboco)/giu],
   ['piso_interno',/piso em\s+(porcelanato)/giu],['forro',/forro em\s+(gesso)/giu],
   ['esquadrias',/esquadrias em\s+(vidro e perfis metálicos)/giu],['revestimento_interno',/possuem revestimento em\s+(porcelanato)/giu],
   ['piso_externo',/possui\s+(contra piso)\s+em/giu],['revestimento_externo',/por\s+(revestimento cerâmico)/giu]
  ])capture(f,pattern,'descricao.'+property);
  // Nome do municipio emissor em uma certidao integrante da descricao.
  capture(f,/Prefeitura Municipal de\s+(.+?)(?=,|\s*\/)/giu,'descricao.certidao.municipio');
  capture(f,/Prefeitura Municipal de\s+.+?\/([A-Z]{2})/giu,'descricao.certidao.uf');
 }
 for(const f of profile.fields){
  const entity=profile.operation.entities.find(e=>e.id===f.entityId);
  if(f.property==='capacidade'||f.property==='estadoCivil'&&entity?.kind==='marriage'){retained.push({fieldId:f.id,reason:'Qualificacao juridica fixa da composicao do modelo, nao dado de preenchimento'});continue;}
  if(f.property==='descricao'){description(f);continue;}
  if(f.property==='pagamento'){
   capture(f,/(anteriormente a este ato|neste ato)/giu,'pagamento.momento');capture(f,/(à vista)/giu,'pagamento.condicao');
   capture(f,/(transferências bancárias(?: em pix)?|transferência bancária(?: em pix)?|espécie)/giu,'pagamento.meio');
   let i=0;for(const m of f.oldText.matchAll(/\d[\d.,]*\s*\([^)]*\)/gu))amount({...f,start:f.start+m.index,oldText:m[0]},'pagamento.valor'+(++i));continue;
  }
  const calendar=/^(.+?)\s+dias?\s+do mês de\s+(.+?)\s+do ano de\s+(.+?)\s*\((\d{2}\/\d{2}\/\d{4})/u.exec(f.oldText);
  if(calendar){let at=0;for(const [label,value] of [['dia_extenso',calendar[1]],['mes',calendar[2]],['ano_extenso',calendar[3]],['numerica',calendar[4]]]){const pos=f.oldText.indexOf(value,at);add(f,pos,pos+value.length,f.property+'.'+label);at=pos+value.length;}continue;}
  if(/endereco/.test(f.property)){address(f);continue;}
  if(/municipio$/.test(f.property)){
   const prefix=/^.*?Municipal\s+(?:do Munic[ií]pio\s+)?de\s+/iu.exec(f.oldText)||/^Munic[ií]pio de\s+/iu.exec(f.oldText),a=prefix?.[0].length||0,city=/^(.+)\/([A-Z]{2})$/u.exec(f.oldText.slice(a));
   if(city){add(f,a,a+city[1].length,f.property+'.nome');add(f,f.oldText.length-2,f.oldText.length,f.property+'.uf');}else add(f,a,f.oldText.length,f.property);continue;
  }
  if(/^(rg|cin|cnh)$/.test(f.property)){
   const document=/^([\d.\-/]+)(\s+)(.+?)(?:\/([A-Z]{2}))?$/u.exec(f.oldText);
   if(document){add(f,0,document[1].length,f.property+'.numero');const begin=document[1].length+document[2].length;add(f,begin,begin+document[3].length,f.property+'.orgao');if(document[4])add(f,f.oldText.length-2,f.oldText.length,f.property+'.uf');continue;}
  }
  if(/\d[^()]*\([^)]*\)/u.test(f.oldText)&&/preco|valor|base|total|area|perimetro/.test(f.property)){amount(f);continue;}
  add(f,0,f.oldText.length,f.property);
 }
 // O tipo documental e um dado; portador(a), parenteses e 'nº' ficam fixos.
 for(const p of stories)for(const m of p.text.matchAll(/(Carteira de Identidade Nacional|Carteira Nacional de Habilitaç[aã]o|Carteira de Identidade)(?:\s*\(?(CIN|CNH|RG)\)?)?/giu)){
  const owner=profile.fields.filter(f=>f.part===p.part&&f.paragraph===p.paragraph&&/^(cin|cnh|rg)$/.test(f.property)&&f.start>=m.index+m[0].length&&f.start<m.index+m[0].length+100).sort((a,b)=>a.start-b.start)[0];
  if(!owner)continue;const base={...owner,start:m.index,oldText:m[0]};add(base,0,m[1].length,'documento.tipo');if(m[2]){const at=m[0].lastIndexOf(m[2]);add(base,at,at+m[2].length,'documento.sigla');}
 }
 for(const f of fields){const p=stories.find(p=>p.part===f.part&&p.paragraph===f.paragraph);check(p?.text.slice(f.start,f.end)===f.oldText,'OPERATIONAL_SOURCE_MISMATCH');check(!profile.operation.protectedFields.some(p=>p.part===f.part&&p.paragraph===f.paragraph&&p.start<f.end&&p.end>f.start),'OPERATIONAL_FIXED_INSTITUTION_CHANGED');}
 return {fields:fields.sort((a,b)=>a.part.localeCompare(b.part)||a.paragraph-b.paragraph||a.start-b.start),retained};
}
function canon(n){return n?[n.uri+'|'+n.local,Object.values(n.attrs).filter(a=>a.uri!=='http://www.w3.org/2000/xmlns/').map(a=>[a.uri+'|'+a.local,a.value]).sort(),n.children.map(canon)]:null;}
function characterStyles(xml){return paragraphsOf(xml).map(p=>p.texts.flatMap(t=>{let r=t.parent;while(r&&r.local!=='r')r=r.parent;const key=JSON.stringify(canon(r?.children.find(n=>n.local==='rPr')));return Array(t.text.length).fill(key);}));}
function publicMetadata(privateData,publicFields){
 const labels={cpf:'CPF',cnpj:'CNPJ',uf:'UF',cep:'CEP',pai:'filiação paterna',mae:'filiação materna',extenso:'valor por extenso'};
 const fields=publicFields.map(({tag,entityId,property,occurrences,contextBefore,contextAfter})=>{const old=privateData.fields.find(f=>f.tag===tag);check(old&&old.entityId===entityId&&old.property===property&&old.occurrences===occurrences,'FIELD_METADATA_INVALID');const letters=old.originalValue.replace(/[^\p{L}]/gu,''),label=property.split('.').at(-1).replace(/\d+$/u,'');return {tag,entityId,property,occurrences,contextBefore,contextAfter,case:letters&&letters===letters.toLocaleUpperCase('pt-BR')?'upper':'source',label:labels[label]||property.replaceAll('.',' ').slice(0,55)};});
 return {schemaVersion:1,modelId:privateData.modelId,version:privateData.version,originalDocSha256:privateData.originalDocSha256,sourceDocxSha256:privateData.sourceDocxSha256,preparedSha256:privateData.preparedSha256,fields,entities:privateData.entities.filter(e=>e.kind!=='institution').map(({id,kind,role,spouseId})=>({id,kind,role,spouseId})),couples:privateData.couples.map(({id,a,b})=>({id,a,b}))};
}

module.exports={refine,publicMetadata};
