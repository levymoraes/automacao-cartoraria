'use strict';
const {openDocx,paragraphsOf}=require('../src/docx.cjs');
const {coverageOf}=require('../src/operation.cjs');
// Somente fixtures inventadas. Nao migrar automaticamente um modelo real.
async function reviewSynthetic(bytes,profile,plan){
  profile.schemaVersion=2;
  profile.operation={inventoryReviewed:true,coverage:coverageOf(await openDocx(bytes),paragraphsOf),entities:[{id:'person-slot',kind:'person',role:'seller',spouseId:null}]};
  for(const f of profile.fields){f.entityId='person-slot';f.property=f.id;}
  plan.entities=[{id:'new-person',slotId:'person-slot',kind:'person',role:'seller',spouseId:null}];
  for(const r of plan.replacements){r.entityId='new-person';r.property=r.fieldId;r.source.entityId='new-person';r.source.property=r.fieldId;}
  return {bytes,profile,plan};
}
module.exports={reviewSynthetic};
