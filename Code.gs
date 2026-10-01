/** Daily project reports. All helpers ending in _ are private to Apps Script. */
const TIME_ZONE = 'Asia/Omsk';
const PARTICIPANTS = [
  { id: 'p1', name: 'Человек один' },
  { id: 'p2', name: 'Человек два' },
  { id: 'p3', name: 'Человек три' }
];
const HEADERS = ['Дата заполнения', 'Цель на проект', 'Что сделал сегодня', 'Что планирую завтра', 'Инсайт дня', 'Время отправки', 'Версия', 'ID отправки'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Мой день — отчёт по проекту')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Run once from the editor; the trailing _ keeps setup inaccessible to visitors. */
function setup_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('SPREADSHEET_ID') || '1X-wT68o9aM6ELnyjciI8SmgezJ8eGncLR_yzIULMwqs';
  const book = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!book) throw new Error('Добавьте SPREADSHEET_ID в свойства скрипта.');
  props.setProperty('SPREADSHEET_ID', book.getId());
  book.setSpreadsheetTimeZone(TIME_ZONE);
  book.setSpreadsheetLocale('ru_RU');
  const registry = book.getSheetByName('_Участники') || book.insertSheet('_Участники');
  if (!registry.getLastRow()) registry.appendRow(['ID', 'Имя', 'Цель', 'ID листа']);
  PARTICIPANTS.forEach(p => {
    const records = registry.getDataRange().getValues();
    const i = records.findIndex((r, n) => n > 0 && r[0] === p.id);
    let sheet = i > 0 ? book.getSheetById(Number(records[i][3])) : null;
    if (!sheet && i > 0) throw new Error('Не найден лист участника ' + p.id + '. Восстановите его перед настройкой.');
    sheet = sheet || book.getSheetByName(p.name) || book.insertSheet(p.name);
    if (sheet.getName() !== p.name) sheet.setName(p.name);
    if (!sheet.getLastRow()) sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    if(i < 0) registry.appendRow([p.id,p.name,'',sheet.getSheetId()]);
    else registry.getRange(i+1,2).setValue(p.name);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#eef0f3').setWrap(true);
    sheet.setRowHeight(1,48);
    sheet.setColumnWidth(1,145);sheet.setColumnWidths(2,4,300);sheet.setColumnWidth(6,175);
    sheet.hideColumns(7,2);
  });
  registry.hideSheet();
  SpreadsheetApp.flush();
  console.log(book.getUrl());
}

function getBootstrap() {
  return { participants: PARTICIPANTS, date: day_(), serverTime: new Date().toISOString() };
}
function getParticipant(input) {
  return locked_(() => {
    const p = participant_(input && input.participantId), book = book_(), entry = entry_(book,p.id);
    const sheet=book.getSheetById(entry.sheetId), date=day_(), match=findReport_(sheet,date);
    const yesterday=dayBefore_(date), previous=findReport_(sheet,yesterday);
    return {goal:entry.goal,report:match ? report_(match.values) : null,yesterdayPlan:previous ? String(previous.values[3]) : '',date:date};
  });
}
function saveGoal(input) {
  const p=participant_(input && input.participantId), goal=text_(input.goal,'Цель');
  return locked_(() => {
    const book=book_(),entry=entry_(book,p.id);
    if(entry.goal!==input.expectedGoal) return {conflict:true};
    book.getSheetByName('_Участники').getRange(entry.row,3).setNumberFormat('@').setValue(literal_(goal));
    SpreadsheetApp.flush();return {ok:true};
  });
}
function submitReport(input) {
  const p=participant_(input && input.participantId);
  const data={goal:text_(input.goal,'Цель'),today:text_(input.today,'Сегодня'),tomorrow:text_(input.tomorrow,'Завтра'),insight:text_(input.insight,'Инсайт')};
  if(typeof input.requestId!=='string'||!/^[a-zA-Z0-9-]{16,80}$/.test(input.requestId))throw new Error('Некорректный ID отправки. Перезагрузите форму.');
  return locked_(() => {
    const date=day_();if(input.date!==date)return {code:'DATE_CHANGED',date:date};
    const book=book_(), entry=entry_(book,p.id),sheet=book.getSheetById(entry.sheetId),match=findReport_(sheet,date);
    // An identical retry after a lost response is a successful read, not a second write.
    if(match&&match.values[7]===input.requestId)return {report:report_(match.values)};
    if((match?match.values[6]:null)!==(input.expectedRevision||null))return {code:'CONFLICT'};
    if(entry.goal!==data.goal)return {code:'GOAL_CHANGED'};
    const now=new Date(), revision=Utilities.getUuid(), row=match?match.row:sheet.getLastRow()+1;
    if(row>sheet.getMaxRows())sheet.insertRowsAfter(sheet.getMaxRows(),100);
    const values=[Utilities.parseDate(date,TIME_ZONE,'yyyy-MM-dd'),data.goal,data.today,data.tomorrow,data.insight,now,revision,input.requestId];
    sheet.getRange(row,2,1,4).setNumberFormat('@');
    sheet.getRange(row,1,1,8).setValues([values.map(literal_)]).setVerticalAlignment('top').setWrap(true);
    sheet.getRange(row,1).setNumberFormat('dd.MM.yyyy');sheet.getRange(row,6).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    SpreadsheetApp.flush();return {report:report_(values)};
  });
}
function text_(v,label){if(typeof v!=='string'||!v.trim())throw new Error('Заполните поле «'+label+'».');if(v.length>5000)throw new Error('Поле «'+label+'»: максимум 5000 символов.');return v.trim();}
function literal_(v){return typeof v==='string'&&/^[=+\-@]/.test(v)?"'"+v:v;}
function participant_(id){const p=PARTICIPANTS.find(p=>p.id===id);if(!p)throw new Error('Выберите участника из списка.');return p;}
function book_(){const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');if(!id)throw new Error('Таблица ещё не подключена. Владелец должен выполнить setup_.');return SpreadsheetApp.openById(id);}
function entry_(book,id){const registry=book.getSheetByName('_Участники');if(!registry)throw new Error('Владелец должен выполнить setup_.');const data=registry.getDataRange().getValues();const i=data.findIndex((r,n)=>n>0&&r[0]===id);if(i<0)throw new Error('Участник ещё не настроен.');return {row:i+1,goal:String(data[i][2]||''),sheetId:Number(data[i][3])};}
function day_(){return Utilities.formatDate(new Date(),TIME_ZONE,'yyyy-MM-dd');}
function dayBefore_(date){return Utilities.formatDate(new Date(new Date(date+'T12:00:00+06:00').getTime()-86400000),TIME_ZONE,'yyyy-MM-dd');}
function findReport_(sheet,date){if(!sheet)throw new Error('Лист участника не найден.');if(sheet.getLastRow()<2)return null;const values=sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues();const i=values.findIndex(r=>(r[0] instanceof Date?Utilities.formatDate(r[0],TIME_ZONE,'yyyy-MM-dd'):String(r[0]))===date);return i<0?null:{row:i+2,values:values[i]};}
function report_(r){return {date:r[0] instanceof Date?Utilities.formatDate(r[0],TIME_ZONE,'yyyy-MM-dd'):String(r[0]),goal:String(r[1]),today:String(r[2]),tomorrow:String(r[3]),insight:String(r[4]),submittedAt:r[5] instanceof Date?r[5].toISOString():String(r[5]),revision:r[6],requestId:r[7]};}
function locked_(fn){const lock=LockService.getScriptLock();lock.waitLock(25000);try{return fn();}finally{lock.releaseLock();}}
