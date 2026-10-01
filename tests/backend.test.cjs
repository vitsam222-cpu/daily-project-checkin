const {test}=require('node:test');const assert=require('node:assert/strict');const fs=require('node:fs');const vm=require('node:vm');const crypto=require('node:crypto');
function env(){
 let now='2026-09-30T17:59:00Z';class Clock extends Date{constructor(...a){super(...(a.length?a:[now]));}};
 const fmt=d=>new Intl.DateTimeFormat('en-CA',{timeZone:'Asia/Omsk',year:'numeric',month:'2-digit',day:'2-digit'}).format(d);
 class Sheet{constructor(id,rows){this.id=id;this.rows=rows;}getLastRow(){return this.rows.length}getMaxRows(){return 1000}getRange(row,col,n=1,m=1){const self=this;return {getValues(){return Array.from({length:n},(_,i)=>Array.from({length:m},(_,j)=>self.rows[row+i-1]?.[col+j-1]??''))},setValues(values){values.forEach((r,i)=>r.forEach((v,j)=>{self.rows[row+i-1]??=[];self.rows[row+i-1][col+j-1]=typeof v==='string'&&v.startsWith("'")?v.slice(1):v}));return this},setValue(v){return this.setValues([[v]])},setNumberFormat(){return this},setVerticalAlignment(){return this},setWrap(){return this}}}getDataRange(){return this.getRange(1,1,this.rows.length,4)}}
 const registry=new Sheet(0,[['ID','Имя','Цель','ID листа'],['u1','Человек один','',1],['u2','Человек два','',2],['u3','Человек три','',3]]);
 const sheets=[registry,...[1,2,3].map(i=>new Sheet(i,[['date','goal','today','tomorrow','insight','timestamp','revision','requestId']]))];
 const book={getSheetByName:()=>registry,getSheetById:id=>sheets[id]};let locked=false;
 const context=vm.createContext({Date:Clock,console,PropertiesService:{getScriptProperties:()=>({getProperty:()=> 'book'})},SpreadsheetApp:{openById:()=>book,flush(){}},Utilities:{formatDate:fmt,parseDate:d=>new Clock(d+'T00:00:00+06:00'),getUuid:()=>crypto.randomUUID()},LockService:{getScriptLock:()=>({waitLock(){assert.equal(locked,false);locked=true},releaseLock(){locked=false}})}});
 vm.runInContext(fs.readFileSync(require('node:path').join(__dirname,'../Code.gs'),'utf8'),context);
 const goal=(id='u1',g='Цель',expected='')=>context.saveGoal({participantId:id,goal:g,expectedGoal:expected});
 const payload=(extra={})=>({participantId:'u1',date:fmt(new Clock()),goal:'Цель',today:'Сегодня',tomorrow:'Завтра',insight:'Инсайт',requestId:crypto.randomUUID(),expectedRevision:null,...extra});
 return {context,sheets,goal,payload,setTime:x=>{now=x}};
}
test('required fields and unknown participant are rejected',()=>{const e=env();e.goal();for(const f of ['goal','today','tomorrow','insight'])assert.throws(()=>e.context.submitReport(e.payload({[f]:'  '})));assert.throws(()=>e.context.getParticipant({participantId:'wrong'}));assert.equal(e.sheets[1].rows.length,1)});
test('one row per day; retry idempotent; explicit update',()=>{const e=env();e.goal();const p=e.payload();const a=e.context.submitReport(p);const b=e.context.submitReport(p);assert.equal(a.report.revision,b.report.revision);assert.equal(e.sheets[1].rows.length,2);const c=e.context.submitReport(e.payload({expectedRevision:a.report.revision,today:'Исправлено'}));assert.equal(c.report.today,'Исправлено');assert.equal(e.sheets[1].rows.length,2)});
test('concurrent stale revision cannot overwrite report',()=>{const e=env();e.goal();e.context.submitReport(e.payload());assert.equal(e.context.submitReport(e.payload({today:'Stale'})).code,'CONFLICT');assert.equal(e.sheets[1].rows[1][2],'Сегодня')});
test('Omsk midnight creates new row and preserves historical goal',()=>{const e=env();e.goal();const p=e.payload();e.context.submitReport(p);e.setTime('2026-09-30T18:01:00Z');assert.equal(e.context.submitReport(p).code,'DATE_CHANGED');e.goal('u1','Новая цель','Цель');e.context.submitReport(e.payload({goal:'Новая цель'}));assert.equal(e.sheets[1].rows.length,3);assert.equal(e.sheets[1].rows[1][1],'Цель');assert.equal(e.sheets[1].rows[2][1],'Новая цель');assert.equal(e.context.getParticipant({participantId:'u1'}).report.date,'2026-10-01')});
test('participants and goals isolated; goal conflict detected',()=>{const e=env();e.goal();e.goal('u2','Другая цель');e.context.submitReport(e.payload());assert.equal(e.context.getParticipant({participantId:'u2'}).goal,'Другая цель');assert.equal(e.context.getParticipant({participantId:'u2'}).report,null);assert.equal(e.goal('u1','Stale','').conflict,true);assert.equal(e.context.submitReport(e.payload({goal:'Stale'})).code,'CONFLICT')});
test('formula-looking input is stored as literal text',()=>{const e=env();e.goal('u1','=SUM(A1:A2)');const r=e.context.submitReport(e.payload({goal:'=SUM(A1:A2)',today:'=IMPORTXML("url")'}));assert.equal(r.report.today,'=IMPORTXML("url")');assert.equal(e.context.literal_('=1+1'),"'=1+1");assert.equal(e.context.getParticipant({participantId:'u1'}).goal,'=SUM(A1:A2)')});
test('yesterday plan uses Omsk calendar day and does not borrow older reports',()=>{const e=env();e.goal();e.context.submitReport(e.payload({tomorrow:'План на сегодня'}));e.setTime('2026-09-30T18:01:00Z');let p=e.context.getParticipant({participantId:'u1'});assert.equal(p.yesterdayPlan,'План на сегодня');assert.equal(p.report,null);assert.equal(e.context.getParticipant({participantId:'u2'}).yesterdayPlan,'');e.setTime('2026-10-01T18:01:00Z');assert.equal(e.context.getParticipant({participantId:'u1'}).yesterdayPlan,'')});
test('returning participant skips setup; sent report requires explicit edit; new participant sees goal',async()=>{
 const source=fs.readFileSync(require('node:path').join(__dirname,'../Index.html'),'utf8').split('<script>')[1].split('</script>')[0].replace(/init\(\);\s*$/,'');
 const data=new Map(),element={hidden:false,textContent:'',innerHTML:''};
 const c=vm.createContext({Intl,Date,console,crypto,setTimeout,clearTimeout,localStorage:{getItem:k=>data.get(k)??null,setItem:(k,v)=>data.set(k,v),removeItem:k=>data.delete(k)},document:{getElementById:()=>element,addEventListener(){}}});
 vm.runInContext(source+'\nrender=()=>{};',c);
 await vm.runInContext("state.id='u1';write('database',{goals:{u1:'Моя цель'},reports:{}});init()",c);
 assert.equal(vm.runInContext('state.step',c),2);assert.equal(vm.runInContext('state.success',c),false);
 await vm.runInContext("write('database',{goals:{u1:'Моя цель'},reports:{['u1:'+today()]:{today:'Сделано',tomorrow:'План',insight:'Мысль',revision:'rev',submittedAt:new Date().toISOString()}}});init()",c);
 assert.equal(vm.runInContext('state.success',c),true);assert.equal(vm.runInContext('state.revisited',c),true);assert.equal(vm.runInContext('state.today',c),'Сделано');
 await vm.runInContext("state.id='u2';loadPerson().then(enterPerson)",c);assert.equal(vm.runInContext('state.step',c),1);assert.equal(vm.runInContext('state.success',c),false);assert.equal(vm.runInContext('state.today',c),'');
});
test('history is participant-scoped, newest first and paginated without overlap',()=>{
 const e=env();e.goal();for(let i=1;i<=12;i++){e.setTime(`2026-09-${String(i).padStart(2,'0')}T10:00:00Z`);e.context.submitReport(e.payload({today:'День '+i}))}
 const a=e.context.getHistory({participantId:'u1'});assert.equal(a.reports.length,10);assert.equal(a.reports[0].date,'2026-09-12');assert.equal(a.nextBefore,'2026-09-03');assert.equal(a.reports[0].revision,undefined);
 const b=e.context.getHistory({participantId:'u1',before:a.nextBefore});assert.equal(b.reports.length,2);assert.equal(b.reports[0].date,'2026-09-02');assert.equal(b.nextBefore,null);assert.equal(e.context.getHistory({participantId:'u2'}).reports.length,0);assert.throws(()=>e.context.getHistory({participantId:'bad'}));assert.throws(()=>e.context.getHistory({participantId:'u1',before:42}));
});
function frontend(){
 const source=fs.readFileSync(require('node:path').join(__dirname,'../Index.html'),'utf8').split('<script>')[1].split('</script>')[0].replace(/init\(\);\s*$/,'');
 const data=new Map(),element={hidden:false,textContent:'',innerHTML:'',scrollIntoView(){}};
 const c=vm.createContext({Intl,Date,console,crypto,setTimeout,clearTimeout,navigator:{onLine:true},localStorage:{getItem:k=>data.get(k)??null,setItem:(k,v)=>data.set(k,v),removeItem:k=>data.delete(k)},document:{getElementById:()=>element,addEventListener(){}}});
 vm.runInContext(source+'\nrender=()=>{};',c);return c;
}
test('draft resumes exact step and stays isolated by participant and day',async()=>{
 const c=frontend();await vm.runInContext("state.id='u1';write('database',{goals:{u1:'Цель'},reports:{}});loadPerson();",c);
 vm.runInContext("state.step=3;state.today='Сделал';state.tomorrow='План';saveDraft();",c);
 await vm.runInContext('loadPerson().then(enterPerson)',c);assert.equal(vm.runInContext('state.view',c),'resume');vm.runInContext('resumeDraft()',c);assert.equal(vm.runInContext('state.step',c),3);assert.equal(vm.runInContext('state.today',c),'Сделал');
 await vm.runInContext("state.id='u2';loadPerson().then(enterPerson)",c);assert.equal(vm.runInContext('state.hasDraft',c),false);assert.equal(vm.runInContext('state.today',c),'');
});
test('offline retry keeps request ID, retains answers, succeeds once and clears draft',async()=>{
 const c=frontend();vm.runInContext("state.id='u1';state.step=4;state.goal=state.goalSaved='Цель';state.today='Сегодня';state.tomorrow='Завтра';state.insight='Мысль';write('database',{goals:{u1:'Цель'},reports:{}});navigator.onLine=false",c);
 await vm.runInContext('submit()',c);assert.equal(vm.runInContext('state.success',c),false);assert.equal(vm.runInContext('state.retrySend',c),true);assert.equal(vm.runInContext('state.draftSaved',c),true);const request=vm.runInContext('pending.requestId',c);
 await vm.runInContext('loadPerson().then(enterPerson)',c);assert.equal(vm.runInContext('pending.requestId',c),request);vm.runInContext('resumeDraft();navigator.onLine=true',c);await vm.runInContext('submit()',c);assert.equal(vm.runInContext('state.success',c),true);assert.equal(vm.runInContext("Object.keys(read('database').reports).length",c),1);assert.equal(vm.runInContext('read(draftKey())',c),null);
});
test('lost acknowledgement is recognized on reload, stale local drafts do not overwrite updated reports',async()=>{
 const c=frontend();vm.runInContext("state.id='u1';state.step=4;state.goal=state.goalSaved='Цель';state.today='Сегодня';state.tomorrow='Завтра';state.insight='Мысль';write('database',{goals:{u1:'Цель'},reports:{}});navigator.onLine=false",c);await vm.runInContext('submit()',c);vm.runInContext('demoSubmit(pending)',c);await vm.runInContext('loadPerson().then(enterPerson)',c);assert.equal(vm.runInContext('state.success',c),true);assert.equal(vm.runInContext('state.hasDraft',c),false);assert.equal(vm.runInContext('read(draftKey())',c),null);
 vm.runInContext("state.success=false;state.today='Устаревший черновик';saveDraft();const db=read('database');db.reports[state.id+':'+state.date].revision='new-revision';db.reports[state.id+':'+state.date].today='Свежий отчёт';write('database',db)",c);await vm.runInContext('loadPerson().then(enterPerson)',c);assert.equal(vm.runInContext('state.today',c),'Свежий отчёт');assert.equal(vm.runInContext('state.hasDraft',c),false);
});
test('viewing another participant history leaves author, draft and form step unchanged',async()=>{
 const c=frontend();vm.runInContext("state.id='u1';state.step=3;state.goal=state.goalSaved='Цель';state.today='Мой черновик';state.tomorrow='Мой план';write('selected','u1');write('database',{goals:{u1:'Цель'},reports:{['u2:'+today()]:{date:today(),today:'Чужой отчёт',goal:'Другая цель',tomorrow:'Другой план',insight:'Инсайт',submittedAt:new Date().toISOString()}}})",c);
 await vm.runInContext('openTeamHistory()',c);await vm.runInContext("selectHistoryPerson('u2')",c);
 assert.equal(vm.runInContext('state.history[0].today',c),'Чужой отчёт');assert.equal(vm.runInContext('state.id',c),'u1');assert.equal(vm.runInContext('state.step',c),3);assert.equal(vm.runInContext('state.today',c),'Мой черновик');assert.equal(vm.runInContext("read('selected')",c),'u1');assert.equal(vm.runInContext('read(draftKey()).today',c),'Мой черновик');
 await vm.runInContext("selectHistoryPerson('unknown')",c);assert.equal(vm.runInContext('state.historyId',c),'u2');
});
