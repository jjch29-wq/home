/* Execute the GAS sources against disposable Sheets doubles; no network/account. */
const fs=require('node:fs'),path=require('node:path'),vm=require('node:vm'),assert=require('node:assert/strict'),crypto=require('node:crypto');
const base=path.resolve(__dirname,'../src/site_apps/central/google_apps_script');
let sheet;
class Sheet{
 constructor(){this.rows=[];}
 getLastRow(){return this.rows.length;}
 getLastColumn(){return Math.max(0,...this.rows.map(r=>r.length));}
 getRange(row,col,n=1,m=1){return {
  getValues:()=>Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>this.rows[row+i-1]?.[col+j-1]??'')),
  setValue:v=>{while(this.rows.length<row)this.rows.push([]);this.rows[row-1][col-1]=v;},
  setValues:values=>{values.forEach((r,i)=>r.forEach((v,j)=>{while(this.rows.length<row+i)this.rows.push([]);this.rows[row+i-1][col+j-1]=v;}));},
  setFontWeight(){return this;},setBackground(){return this;}
 };}
 appendRow(row){this.rows.push([...row]);}
 setFrozenRows(){}
 deleteRow(row){this.rows.splice(row-1,1);}
}
const context=vm.createContext({console,Date,Math,JSON,Number,String,Array,Object,isFinite,
 SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:()=>sheet,insertSheet:()=>sheet=new Sheet()}),flush(){}},
 LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){}})},
 Utilities:{formatDate:(d,tz,fmt)=>fmt==='yyyy-MM-dd'?d.toISOString().slice(0,10):d.toISOString().replace('T',' ').replace('Z',''),
 DigestAlgorithm:{SHA_256:'sha256'},computeDigest:(alg,text)=>crypto.createHash(alg).update(text).digest(),base64Encode:b=>Buffer.from(b).toString('base64')}
});
for(const file of ['Code.gs','Reference.gs','Workspace.gs'])vm.runInContext(fs.readFileSync(path.join(base,file),'utf8'),context,{filename:file});
let tests=0;
function test(name,fn){sheet=new Sheet();sheet.appendRow([...context.HEADERS]);fn();tests++;console.log('PASS '+name);}
const record=(id,extra={})=>({client_request_id:id,date:'2026-11-14',company:'한국지역난방공사',method:'MT',qty:2.774,unit:'M',unit_price:75433,total_fee:209251,workers:[{user:'부장 주진철',time:'09:00~18:00',ot:0,meal:12000}],...extra});
test('managed catalogs exclude historical entry-company aliases',()=>{
 assert.equal(context.getConfig().users.length,32);assert.equal(context.getConfig().vehicles.length,16);
 assert.deepEqual(Array.from(context.getConfig().companies),['한국지역난방공사']);
});
test('retry is idempotent and retains worker meals / exact quantities',()=>{
 const r=record('retry');context.saveRecord(r);const res=context.saveRecord({...r});assert.equal(res.duplicate,true);assert.equal(sheet.getLastRow(),2);
 const headers=sheet.rows[0],row=sheet.rows[1];for(const key of ['Usage','검사량','수량'])assert.equal(row[headers.indexOf(key)],2.774);
 assert.equal(row[headers.indexOf('Meal')],12000);
 assert.throws(()=>context.saveRecord(record('retry',{qty:5})),/이미 다른 내용/);
});
test('batch retries do not duplicate jobs',()=>{
 const batch={records:[record('a'),record('b',{method:'RT',qty:27})]};
 assert.equal(context.saveRecordsBatch(batch).status,'success');assert.equal(context.saveRecordsBatch(batch).status,'success');assert.equal(sheet.getLastRow(),3);
});
test('invalid batch validates before any append',()=>{
 assert.throws(()=>context.saveRecordsBatch({records:[record('good'),record('bad',{method:'RT',qty:1.5})]}),/정수/);assert.equal(sheet.getLastRow(),1);
});
test('negative and nonfinite quantities rejected',()=>{
 assert.throws(()=>context.saveRecord(record('bad',{qty:-1})),/0 이상/);
 assert.throws(()=>context.saveRecord(record('bad',{qty:Infinity})),/0 이상/);
 assert.throws(()=>context.saveRecord(record('bad',{date:'2026-02-30'})),/날짜/);
});
test('old sheet gets missing fields without rewriting existing data',()=>{
 sheet.rows=[['Date','Usage'],['2026-11-01',10]];
 context.getDailyUsageSheet();assert.equal(sheet.rows[1][1],10);assert.ok(sheet.rows[0].includes('Meal'));assert.ok(sheet.rows[0].includes('ClientRequestID'));
});
test('partial batch failure keeps precise request outcomes',()=>{
 const append=sheet.appendRow.bind(sheet);let calls=0;
 sheet.appendRow=row=>{calls++;if(calls===2)throw new Error('simulated write failure');append(row);};
 const response=context.saveRecordsBatch({records:[record('a'),record('b')]});
 assert.equal(response.status,'partial');assert.equal(response.results[0].request_id,'a');assert.equal(response.results[1].request_id,'b');assert.equal(sheet.getLastRow(),2);
 sheet.appendRow=append;assert.equal(context.saveRecordsBatch({records:[record('a'),record('b')]}).status,'success');assert.equal(sheet.getLastRow(),3);
});
test('query filters date, method, company, search and paginates',()=>{
 context.saveRecord(record('a',{date:'2026-11-01',note:'first'}));context.saveRecord(record('b',{date:'2026-11-14',method:'RT',qty:12,note:'찾기'}));
 const q=context.queryRecords({start:'2026-11-10',method:'RT',search:'찾기'});assert.equal(q.total,1);assert.equal(q.records[0].qty,12);
 assert.equal(context.queryRecords({start:'2026-11-15'}).total,0);
});
test('snapshot inventory is dated, searchable, read only',()=>{
 const result=context.getInventory({search:'세척제'});assert.ok(result.generated_at);assert.equal(result.snapshot,true);assert.ok(result.records.every(r=>JSON.stringify(r).includes('세척제')));assert.equal(sheet.getLastRow(),1);
});
test('same-second records use client IDs for edit and reject ambiguous timestamps',()=>{
 const a=record('a',{entry_time:'2026-11-14 10:00:00.100'}),b=record('b',{entry_time:'2026-11-14 10:00:00.900',qty:5});context.saveRecord(a);context.saveRecord(b);
 assert.equal(context.findMobileRow(sheet,sheet.rows[0],{entry_time:'2026-11-14 10:00:00'}),-1);
 assert.equal(context.updateRecord({...b,qty:0,total_fee:0}).status,'success');
 const rows=context.queryRecords({}).records;assert.equal(rows.find(r=>r.client_request_id==='a').qty,2.774);assert.equal(rows.find(r=>r.client_request_id==='b').qty,0);
 assert.equal(context.deleteRecord({client_request_id:'a'}).status,'success');assert.equal(sheet.getLastRow(),2);
});
const html=fs.readFileSync(path.join(base,'index.html'),'utf8');
for(const match of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g))new vm.Script(match[1].replace(/<\?[\s\S]*?\?>/g,'{}'));
const shell=fs.readFileSync(path.join(base,'MobileShell.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/)[1];new vm.Script(shell);
assert.ok(!shell.includes("confirm('이 기기에 임시저장한 입력을 복원할까요?')"));
assert.ok(shell.includes('임시 입력을 자동 복원했습니다.'));
new vm.Script(fs.readFileSync(path.join(base,'MethodFilters.html'),'utf8').match(/<script>([\s\S]*?)<\/script>/)[1]);
console.log('PASS all browser scripts parse');tests++;
const packed=path.resolve(__dirname,'../../outputs/central-mobile-gas/Code.gs');
if(fs.existsSync(packed)){
 let template='';
 context.HtmlService={XFrameOptionsMode:{ALLOWALL:'all'},createTemplate:source=>{
  template=source;return {evaluate:()=>({setTitle(){return this;},addMetaTag(){return this;},setXFrameOptionsMode(){return this;}})};
 }};
 vm.runInContext(fs.readFileSync(packed,'utf8'),context,{filename:'packed-Code.gs'});
 context.doGet({parameter:{}});
 assert.ok(template.includes('workspaceInit'));assert.ok(!template.includes("include('MobileShell')"));assert.ok(!template.includes("include('MethodFilters')"));assert.ok(template.includes('function refreshMethodChoices(method)'));
 console.log('PASS single-file deployment serves embedded page without HTML file lookup');tests++;
}
console.log('Completed '+tests+' checks');
if(process.argv.includes('--preview')){
 const shellHtml=fs.readFileSync(path.join(base,'MobileShell.html'),'utf8');
 const fixtures={getConfig:context.getConfig(),getHistoricalCompanies:['한국지역난방공사','한국지역난방공사 중앙지사'],queryRecords:context.queryRecords({}),getInventory:context.getInventory({}),getRecentRecords:context.getRecentRecords('2026-11-14')};
 const bridge=`<script>window.confirm=()=>false;const previewFixtures=${JSON.stringify(fixtures).replaceAll('<','\\u003c')};var google={script:{get run(){let success=()=>{},failure=()=>{};let proxy;proxy=new Proxy({}, {get(target,key){if(key==='withSuccessHandler')return fn=>{success=fn;return proxy};if(key==='withFailureHandler')return fn=>{failure=fn;return proxy};return payload=>setTimeout(()=>success(previewFixtures[key]||{status:'success',results:((payload||{}).records||[]).map(r=>({status:'success',request_id:r.client_request_id}))}),20);}});return proxy;}}};</script>`;
 let preview=html.replace(/<\?!= JSON.stringify[\s\S]*?\?>/,JSON.stringify(context.getConfig()).replaceAll('<','\\u003c')).replace(/<\?!= include\('MobileShell'\); \?>/,shellHtml).replace("<?!= include('MethodFilters'); ?>",fs.readFileSync(path.join(base,'MethodFilters.html'),'utf8'));
 preview=preview.replace('</head>',bridge+'</head>');
 preview=preview.replace('<main ', '<div style="padding:8px;text-align:center;background:#fff2c2;color:#734700;font-size:12px">검증용 미리보기 · 실제 클라우드에 저장되지 않습니다</div><main ');
 const out=path.resolve(__dirname,'../../outputs/central-mobile-review');fs.mkdirSync(out,{recursive:true});fs.writeFileSync(path.join(out,'preview.html'),preview);console.log('Preview prepared');
}
