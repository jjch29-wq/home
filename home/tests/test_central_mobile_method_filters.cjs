const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const root=path.resolve(__dirname,'../src/site_apps/central/google_apps_script');
class Select{
 constructor(multiple=false){this.options=[];this._value='';this.multiple=multiple;}
 replaceChildren(...options){this.options=options;this._value=options[0]?.value||'';this.options.forEach((o,i)=>o.selected=!this.multiple&&i===0);}
 add(option){this.options.push(option);}
 get selectedOptions(){return this.options.filter(o=>o.selected);}
 get value(){return this.multiple?(this.selectedOptions[0]?.value||''):this._value;}
 set value(value){this._value=this.options.some(o=>o.value===String(value))?String(value):'';this.options.forEach(o=>o.selected=o.value===this._value);}
}
const multiIds=new Set(['select-equip','select-aux-equip','select-paut-scanner','select-paut-probe','select-paut-wedge']);
const fields={};for(const id of ['workspace-material-id','workspace-warehouse','select-material','select-equip','select-aux-equip','select-paut-scanner','select-paut-probe','select-paut-wedge'])fields[id]=new Select(multiIds.has(id));
fields['input-custom-aux']={value:'',classList:{add(){}}};
fields['input-custom-mat']={value:'',classList:{contains:()=>true,add(){}}};
fields['material-input-wrap']={classList:{toggle(_name,hidden){this.hidden=hidden;}}};fields['material-history-notice']={classList:{toggle(_name,hidden){this.hidden=hidden;}}};
fields['workspace-manual-link']={checked:false};fields['workspace-manual-fields']={hidden:true};fields['workspace-warehouse-wrap']={hidden:true};
fields['paut-equipment-grid']={style:{},classList:{toggle(){}}};fields['main-equipment-wrap']={style:{}};fields['aux-equipment-wrap']={style:{},classList:{toggle(){}}};fields['aux-equipment-label']={textContent:''};
function Option(label,value){this.text=label;this.value=String(value);this.selected=false;}
const context=vm.createContext({console,Option,selectedMethod:'MT',document:{getElementById:id=>fields[id]||null}});
vm.runInContext(fs.readFileSync(path.join(root,'Reference.gs'),'utf8'),context);
context.appConfig=context.CENTRAL_REFERENCE.config;
const functions=fs.readFileSync(path.join(root,'MethodFilters.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/)[1];
vm.runInContext(functions,context);
let count=0;
for(const method of ['PAUT','RT','UT','MT','PT']){
 context.selectedMethod=method;
 context.refreshMethodChoices(method);context.applyMethodDefaults(method);
 const choices=context.methodChoices(context.appConfig,method);
 assert.equal(Boolean(fields['material-input-wrap'].classList.hidden),choices.materials.length===0);
 assert.deepEqual(fields['workspace-material-id'].options.slice(1).map(o=>o.value),Array.from(choices.inventory,r=>r.id));
 assert.deepEqual(fields['select-equip'].options.slice(1).map(o=>o.value),Array.from(choices.main));
 const categorized=Array.from(choices.aux).filter(value=>context.pautAuxCategory(value));
 const expectedAux=method==='PAUT'?Array.from(choices.aux).filter(value=>!context.pautAuxCategory(value)):Array.from(choices.aux);
 assert.deepEqual(fields['select-aux-equip'].options.slice(1,-1).map(o=>o.value),expectedAux);
 if(method==='PAUT')assert.deepEqual([
  ...fields['select-paut-scanner'].options.slice(1).map(o=>o.value),
  ...fields['select-paut-probe'].options.slice(1).map(o=>o.value),
  ...fields['select-paut-wedge'].options.slice(1).map(o=>o.value)
 ].sort(),categorized.sort());
 assert.ok(choices.inventory.every(r=>r.methods.includes(method)));
 if(method!=='RT')assert.ok(!choices.materials.some(x=>/CARESTREAM|FILM/i.test(x)));
 if(method!=='MT')assert.ok(!choices.main.some(x=>/YOKE/i.test(x)));
 if(method!=='PAUT')assert.ok(!choices.aux.some(x=>/COBRA|SCANNER|MICROBE/i.test(x)));
 count++;console.log('PASS '+method+': inventory '+choices.inventory.length+', main '+choices.main.length+', aux '+choices.aux.length);
}
context.selectedMethod='PAUT';context.refreshMethodChoices('PAUT');
const paut=context.methodChoices(context.appConfig,'PAUT');assert.ok(paut.main.length);
fields['select-equip'].value=paut.main[0];context.refreshMethodChoices('PAUT');assert.equal(fields['select-equip'].value,paut.main[0]);
context.selectedMethod='RT';context.refreshMethodChoices('RT');assert.notEqual(fields['select-equip'].value,paut.main[0]);count++;console.log('PASS valid selection retained, unrelated selection cleared on method change');
const shell=fs.readFileSync(path.join(root,'MobileShell.html'),'utf8');
vm.runInContext(shell.match(/function workspaceSelectValue[^\n]*/)[0],context);
context.workspaceSelectValue(fields['select-equip'],'old-model-custom');assert.equal(fields['select-equip'].value,'old-model-custom');assert.match(fields['select-equip'].options.at(-1).text,/기존 저장값/);
context.selectedMethod='MT';context.refreshMethodChoices('MT');assert.ok(!fields['select-equip'].options.some(o=>o.value==='old-model-custom'));count++;console.log('PASS historical custom value preserved only during explicit restore');

context.appConfig={method_catalog:{MT:{inventory:[
 {id:'unique',name:'현상제',model:'D-100',warehouse:'본점',company:'한국지역난방공사',role:'material'},
 {id:'multi-a',name:'필름',model:'T200',warehouse:'본점',company:'한국지역난방공사',role:'material'},
 {id:'multi-b',name:'필름',model:'T200',warehouse:'창고2',company:'한국지역난방공사',role:'material'}
],materials:['현상제 (D-100)','필름 (T200)'],main:[],aux:[]}}};context.selectedMethod='MT';context.refreshMethodChoices('MT');
fields['select-material'].value='현상제 (D-100)';assert.equal(context.syncMaterialConnection(),'unique');assert.equal(fields['workspace-warehouse-wrap'].hidden,true);count++;console.log('PASS unique material links automatically');
fields['select-material'].value='필름 (T200)';assert.equal(context.syncMaterialConnection(),'');assert.equal(fields['workspace-warehouse-wrap'].hidden,false);assert.deepEqual(fields['workspace-warehouse'].options.slice(1).map(o=>o.value),['multi-a','multi-b']);fields['workspace-warehouse'].value='multi-b';assert.equal(context.materialConnectionForSave('필름 (T200)'),'multi-b');count++;console.log('PASS duplicate material requires and preserves warehouse choice');
fields['workspace-manual-link'].checked=true;fields['workspace-material-id'].add(new Option('legacy','legacy'));fields['workspace-material-id'].value='legacy';assert.equal(context.syncMaterialConnection(),'legacy');assert.equal(fields['workspace-manual-fields'].hidden,false);context.restoreMaterialConnection('',true);assert.equal(fields['workspace-manual-link'].checked,true);count++;console.log('PASS explicit manual link mode is retained');
context.selectedMethod='PAUT';context.appConfig=context.CENTRAL_REFERENCE.config;context.refreshMethodChoices('PAUT');
assert.equal(context.pautAuxCategory('ST1018스텝웨지 표준시험편'),'');
fields['input-custom-mat'].value='과거 PAUT 자재';fields['input-custom-mat'].classList.contains=()=>false;context.restoreMaterialConnection('');assert.equal(fields['material-input-wrap'].classList.hidden,false);assert.equal(fields['material-history-notice'].classList.hidden,false);
const scanner=fields['select-paut-scanner'].options[1]?.value,probe=fields['select-paut-probe'].options[1]?.value,wedge=fields['select-paut-wedge'].options[1]?.value;
assert.ok(scanner&&probe&&wedge);fields['select-paut-scanner'].value=scanner;fields['select-paut-probe'].value=probe;fields['select-paut-wedge'].value=wedge;
const combined=context.collectAuxEquipment();assert.equal(combined,[scanner,probe,wedge].join('; '));
fields['select-paut-scanner'].value='';fields['select-paut-probe'].value='';fields['select-paut-wedge'].value='';context.restoreAuxEquipment(combined);
assert.equal(fields['select-paut-scanner'].value,scanner);assert.equal(fields['select-paut-probe'].value,probe);assert.equal(fields['select-paut-wedge'].value,wedge);count++;console.log('PASS PAUT scanner, probe and wedge save and restore in order');
const extraScanner=fields['select-paut-scanner'].options[2]?.value;
if(extraScanner){context.setMethodValues(fields['select-paut-scanner'],[scanner,extraScanner]);assert.deepEqual(Array.from(context.selectedMethodValues(fields['select-paut-scanner'])),[scanner,extraScanner]);assert.match(context.collectAuxEquipment(),new RegExp(extraScanner.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')));}
const mainChoices=Array.from(context.methodChoices(context.appConfig,'PAUT').main).slice(0,2);context.setMethodValues(fields['select-equip'],mainChoices);assert.equal(context.collectMainEquipment(),mainChoices.join('; '));count++;console.log('PASS main and auxiliary equipment allow multiple selections');
console.log('Completed '+count+' filter checks');
