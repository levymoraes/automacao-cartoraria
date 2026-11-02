'use strict';
// Regras fechadas de apresentacao. Nao recebe texto livre para editar clausulas.
const claims=[
 {key:'capacidade',pattern:/\bcapaz(?:es)?\b/iu,purposes:['identity','contract_confirmation']},
 {key:'comparecimento',pattern:/\b(?:compareceram|compareceu|comparece|pessoalmente)\b/iu,purposes:['contract_confirmation']},
 {key:'pagamento',pattern:/recebeu|receberam|recebid[oa]s?|pag[oa]s? integral|pagamento(?![:\p{L}_])/iu,purposes:['contract_confirmation']},
 {key:'quitacao',pattern:/quita[cç][aã]o/iu,purposes:['contract_confirmation'],label:'quitação'},
 {key:'onus',pattern:/livre[s]?[\s\S]{0,45}de [oô]nus|inexist[eê]ncia de [oô]nus/iu,purposes:['property','certificate','contract_confirmation'],label:'ônus'},
 {key:'certificacoes',pattern:/certific(?:o|amos|ando)|dou f[eé]|consulta[\s\S]{0,70}(?:CNIB|indisponibilidade)|consultas oficiais/iu,purposes:['certificate','contract_confirmation'],label:'certificações'},
 {key:'leitura',pattern:/(?<![\p{L}\p{N}_])(?:lid[ao]s?|leitura)(?![\p{L}\p{N}_])/iu,purposes:['contract_confirmation']},
 {key:'assinatura',pattern:/\bassinad[ao]s?\b|\bassinaram\b/iu,purposes:['contract_confirmation']}
];
const visibleText=text=>text.replace(/\{\{[A-Z][A-Z0-9_]*\}\}/gu,tag=>' '.repeat(tag.length));
function documentFamily(value){
 const s=(value||'').normalize('NFD').replace(/\p{M}/gu,'').trim().toLowerCase();
 return /^(?:rg|registro geral|carteira de identidade)$/u.test(s)?'rg':/^(?:cin|carteira de identidade nacional)$/u.test(s)?'cin':/^(?:cnh|carteira nacional de habilitacao)$/u.test(s)?'cnh':/^passaporte$/u.test(s)?'passaporte':/^(?:crnm|carteira de registro nacional migratorio)$/u.test(s)?'crnm':null;
}
function notePosition(text,note){
 if(!note.anchor||note.anchor==='end')return text.length;
 const pattern=note.anchor==='preco'?/preço|valor[\s\S]{0,30}venda/iu:claims.find(c=>c.key===note.anchor)?.pattern;
 const matches=pattern?[...visibleText(text).matchAll(new RegExp(pattern.source,pattern.flags+'g'))]:[],match=matches[note.occurrence||0];if(!match)return text.length;
 // Para capacidade basta a qualificacao; nos demais casos, o fim da frase
 // ou clausula. Nao confundir decimais nem as abreviaturas com fim da frase.
 if(note.anchor==='capacidade'||note.anchor==='comparecimento')return match.index+match[0].length;
 const tail=text.slice(match.index+match[0].length),end=[...tail.matchAll(/;|\.(?=\s+(?:[A-ZÀ-Ý]|\d+[.)])|$)/gu)].find(end=>end[0]===';'||!/(?:^|\s)(?:Def|Art|Av|Dr|Dra|Sr|Sra|Inc|Pág|Fl)\.$/iu.test(text.slice(0,match.index+match[0].length+end.index+1)));
 return end?match.index+match[0].length+end.index+1:text.length;
}
function qualificationEdits(p,metadata,values,pending){
 const edits=[],fields=new Map(metadata.fields.map(f=>[f.tag,f])),literal=tag=>'{{'+tag+'}}';
 const add=(start,end,value,reason)=>{if(!edits.some(e=>start<e.end&&end>e.start))edits.push({start,end,value,reason});};
 for(const m of p.text.matchAll(/\{\{([A-Z][A-Z0-9_]*)\}\}/gu)){
  const f=fields.get(m[1]);if(!f)continue;
  if(pending.has(f.tag)&&/endereco.*\.andar$/u.test(f.property)){const suffix=/^[º°]\s*(?:andar|pavimento)\b/iu.exec(p.text.slice(m.index+m[0].length));if(suffix)add(m.index+m[0].length,m.index+m[0].length+suffix[0].length,'','ordinal de campo ausente');}
  if(values[f.tag]===''&&/^(?:rg|cin|cnh)\.validade$/u.test(f.property)){
   const prefix=/,\s*data de validade\s*$/iu.exec(p.text.slice(0,m.index));if(prefix)add(prefix.index,m.index+m[0].length,'','validade não indicada no documento alternativo');
  }
  if(f.property==='documento.tipo'&&documentFamily(values[f.tag])==='passaporte'){
   const prefix=/\bda\s*$/u.exec(p.text.slice(0,m.index));if(prefix)add(prefix.index,prefix.index+2,'do','concordância do tipo documental');
   const tail=p.text.slice(m.index+m[0].length),emission=/^[^;]{0,240}?,\s*(emitida)(?= em\s*\{\{([A-Z0-9_]+)\}\})/u.exec(tail);if(emission&&fields.get(emission[2])?.entityId===f.entityId){const at=m.index+m[0].length+emission.index+emission[0].lastIndexOf('emitida');add(at,at+7,'emitido','concordância do tipo documental');}
  }
 }
 for(const m of p.text.matchAll(/,\s*na qual consta o RG n[º°o.]?\s*\{\{([A-Z0-9_]+)\}\}\s+\{\{([A-Z0-9_]+)\}\}\/\{\{([A-Z0-9_]+)\}\}/gu)){
  const secondary=m.slice(1).map(tag=>fields.get(tag));if(secondary.some(f=>!f)||new Set(secondary.map(f=>f.entityId)).size!==1)continue;
  const entity=secondary[0].entityId,type=metadata.fields.find(f=>f.entityId===entity&&f.property==='documento.tipo');if(!type||documentFamily(values[type.tag])!=='rg')continue;
  const same=secondary.every(f=>/^rg\.(numero|orgao|uf)$/u.test(f.property)&&!pending.has(f.tag)&&metadata.fields.some(primary=>primary.entityId===entity&&/^(?:cin|cnh)\./u.test(primary.property)&&primary.property.split('.')[1]===f.property.split('.')[1]&&p.text.includes(literal(primary.tag))&&!pending.has(primary.tag)&&values[primary.tag]===values[f.tag]));
  if(same)add(m.index,m.index+m[0].length,'','referência duplicada ao mesmo RG');
 }
 return edits.sort((a,b)=>a.start-b.start);
}
function fieldLabel(f){
 const leaf=f.property.split('.').at(-1),labels={livro:'livro',folhas:'folha(s)',termo:'termo',numero:'número',orgao:'órgão emissor',uf:'UF',emissao:'data de emissão',validade:'validade',andar:'andar',logradouro:'logradouro',municipio:'município',bairro:'bairro',cep:'CEP',dia_extenso:'dia da lavratura',ano_extenso:'ano da lavratura',mes:'mês da lavratura',numerica:'data da lavratura'};
 return /^(?:rg|cin|cnh|passaporte|crnm)\./u.test(f.property)?(leaf==='numero'?'número do documento':labels[leaf]||leaf):labels[leaf]||f.label||f.property.replaceAll('.',' ');
}
function fieldValue(f,value,paragraphs){
 const before=paragraphs.find(p=>p.text.includes('{{'+f.tag+'}}'))?.text.split('{{'+f.tag+'}}')[0]||'';
 if(/endereco.*\.numero$/u.test(f.property)&&/n[º°.]\s*$/iu.test(before))return value.replace(/^n[º°.]\s*/iu,'');
 if(/endereco.*\.andar$/u.test(f.property))return value.replace(/^(\d+)[º°]?(?:\s+(?:andar|pavimento))?$/iu,'$1');
 if(/R\$\s*$/u.test(before))return value.replace(/^R\$\s*/u,'');
 return value;
}
module.exports={claims,visibleText,documentFamily,notePosition,qualificationEdits,fieldLabel,fieldValue};
