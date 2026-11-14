/* Browser-independent client queue tests; simulated transport, no Google writes. */
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict');
const source=fs.readFileSync(path.resolve(__dirname,'../src/site_apps/central/google_apps_script/MobileShell.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/)[1].split('(function workspaceInit()')[0];
function fixture(){
 const store={},elements={};let sequence=0;
 elements['workspace-material-id']={value:''};elements['workspace-manual-link']={checked:false};
 elements['workspace-warehouse-wrap']={hidden:true};elements['workspace-manual-fields']={hidden:true};
 elements['workspace-warehouse']={value:'',replaceChildren(){this.value='';},add(){}};
 elements['select-material']={value:''};elements['input-custom-mat']={value:'',classList:{contains:()=>true}};
 const field=id=>elements[id]||(elements[id]={value:'',textContent:''});
 const context=vm.createContext({console,JSON,Math,Number,String,Object,Array,Set,Date,Promise,setTimeout,clearTimeout,
  crypto:{randomUUID:()=>`request-${++sequence}`},window:{crypto:{},scrollTo(){}},
  document:{getElementById:field,querySelectorAll:()=>[]},selectedMethod:'MT',workerRows:[],currentEditingRecord:null,
  appConfig:{material_catalog:[],method_catalog:{MT:{inventory:[],materials:[],main:[],aux:[]}}},loadRecentRecords(){},switchMethod(){},renderWorkers(){},showToast(){},
  getFormData:()=>({date:'2026-11-14',method:'MT',company:'한국지역난방공사',loc_type:'열수송관',qty:3,unit_price:75433,travel_cost:0,total_fee:226299,workers:[]}),
  localStorage:{setItem:(k,v)=>store[k]=v,getItem:k=>store[k]||null}});
 vm.runInContext(source,context);
 vm.runInContext(fs.readFileSync(path.resolve(__dirname,'../src/site_apps/central/google_apps_script/MethodFilters.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/)[1],context);
 vm.runInContext('workspaceRenderPending=()=>{};workspaceSwitch=()=>{};',context);
 return {context,store,field,run:code=>vm.runInContext(code,context)};
}
async function main(){let passed=0;
 {
  const t=fixture();t.run('workspaceQueueCurrent()');const id=t.run('workspacePending[0].client_request_id');
  t.run('workspaceEditingId=workspacePending[0].client_request_id;workspaceQueueCurrent()');assert.equal(t.run('workspacePending.length'),1);assert.equal(t.run('workspacePending[0].client_request_id'),id);
  assert.equal(JSON.parse(Object.values(t.store)[0]).pending.length,1);passed++;console.log('PASS local queue editing preserves stable retry ID');
 }
 {
  const t=fixture();t.run('workspaceQueueCurrent();workspaceQueueCurrent();workspaceCall=async()=>{throw new Error("offline")};');
  const ids=t.run('JSON.stringify(workspacePending.map(r=>r.client_request_id))');await t.run('workspaceSavePending()');assert.equal(t.run('JSON.stringify(workspacePending.map(r=>r.client_request_id))'),ids);assert.equal(t.run('workspaceSaving'),false);passed++;console.log('PASS network failure preserves pending work and IDs');
 }
 {
  const t=fixture();t.run('workspaceQueueCurrent();workspaceQueueCurrent();workspaceCall=async(name,payload)=>({status:"partial",results:[{status:"success",request_id:payload.records[0].client_request_id},{status:"error",request_id:payload.records[1].client_request_id,detail:"retry"}]});');
  const remaining=t.run('workspacePending[1].client_request_id');await t.run('workspaceSavePending()');assert.equal(t.run('workspacePending.length'),1);assert.equal(t.run('workspacePending[0].client_request_id'),remaining);
  t.run('workspaceCall=async(name,payload)=>({status:"success",results:payload.records.map(r=>({status:"success",request_id:r.client_request_id}))});');await t.run('workspaceSavePending()');assert.equal(t.run('workspacePending.length'),0);passed++;console.log('PASS partial success removes only confirmed saved rows');
 }
 {
  const t=fixture();t.run('workspaceSkipDraftCapture=true;workspacePersist();');assert.equal(JSON.parse(Object.values(t.store)[0]).draft,null);t.run('workspaceQueueCurrent()');assert.notEqual(JSON.parse(Object.values(t.store)[0]).draft,null);passed++;console.log('PASS erased draft stays erased until new entry');
 }
 {
  const t=fixture();t.context.localStorage.setItem=()=>{throw new Error('quota')};t.run('workspacePersist()');assert.equal(t.run('workspaceStorageOK'),false);assert.match(t.field('workspace-local-status').textContent,/사용할 수 없습니다/);passed++;console.log('PASS storage failure is visible');
 }
 {
  const t=fixture();assert.throws(()=>t.run('workspaceValidate({date:"2026-11-14",method:"RT",qty:1.5,unit_price:1,travel_cost:0,total_fee:1})'),/정수/);passed++;console.log('PASS client RT validation');
 }
 console.log(`Completed ${passed} client checks`);
}
main().catch(error=>{console.error(error);process.exitCode=1;});
