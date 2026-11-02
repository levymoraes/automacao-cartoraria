'use strict';
const {test}=require('node:test'),assert=require('node:assert/strict');
test('preparação: Nova Esperança Fictícia não é prefixo de número de endereço',()=>{
 const {refine}=require('../scripts/prepare-operational-models.cjs');
 const text='Rua Fictícia, nº 391, Sala 02, Centro, Nova Esperança Fictícia/BA, CEP: 00000-000';
 const {fields}=refine({fields:[{id:'endereco',entityId:'person',property:'endereco',oldText:text,start:0,end:text.length,part:'word/document.xml',paragraph:0}],operation:{entities:[{id:'person',kind:'person'}],protectedFields:[]}},[{part:'word/document.xml',paragraph:0,text}]);
 assert.deepEqual(fields.filter(f=>f.property==='endereco.numero').map(f=>f.oldText),['391']);
 assert.equal(fields.find(f=>f.property==='endereco.municipio')?.oldText,'Nova Esperança Fictícia');
 assert.equal(fields.find(f=>f.property==='endereco.complemento')?.oldText,'Sala 02');
 assert.deepEqual(fields.filter(f=>f.property==='endereco.bairro').map(f=>f.oldText),['Centro']);
});
