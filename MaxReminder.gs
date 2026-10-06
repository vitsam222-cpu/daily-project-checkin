/** MAX group reminders. Credentials stay only in Script Properties. */
const MAX_START_DATE = '2026-10-06';
const MAX_FORM_URL = 'https://vitsam222-cpu.github.io/daily-project-checkin/';
// This endpoint has a publicly trusted TLS chain; keep certificate validation enabled.
const MAX_API_URL = 'https://platform-api.max.ru';

/** Owner-only editor setup. Does not send messages. */
function configureMaxReminder() {
  const active=Session.getActiveUser().getEmail(),effective=Session.getEffectiveUser().getEmail();
  if(!active||active!==effective)throw new Error('Настройку может запускать только владелец из редактора.');
  const props=PropertiesService.getScriptProperties();
  const bot=maxApi_('/me');
  if(bot.username!=='id550622896807_1_bot')throw new Error('Токен относится к другому боту.');
  let chatId=props.getProperty('MAX_CHAT_ID');
  if(!chatId){
    const updates=maxApi_('/updates?timeout=0&limit=100');
    const candidates=[...new Set((updates.updates||[]).map(u=>u.chat_id||(u.message&&u.message.recipient&&u.message.recipient.chat_id)).filter(v=>v!==null&&v!==undefined).map(String))];
    const groups=candidates.map(id=>maxApi_('/chats/'+encodeURIComponent(id))).filter(c=>c.type==='chat'&&c.status==='active');
    if(groups.length!==1){console.log(JSON.stringify(groups.map(c=>({chat_id:String(c.chat_id),title:c.title}))));throw new Error('Укажите MAX_CHAT_ID выбранной группы в свойствах скрипта.');}
    chatId=String(groups[0].chat_id);
  }
  if(!/^-?\d+$/.test(chatId))throw new Error('MAX_CHAT_ID должен быть числовым ID группы.');
  const chat=maxApi_('/chats/'+encodeURIComponent(chatId));
  if(chat.type!=='chat'||chat.status!=='active')throw new Error('Бот не подключён к выбранной группе.');
  const membership=maxApi_('/chats/'+encodeURIComponent(chatId)+'/members/me');
  if(!membership.user_id)throw new Error('Не удалось подтвердить участие бота в группе.');
  props.setProperty('MAX_CHAT_ID',chatId);
  const existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='maxReminderTick');
  let trigger=existing[0];
  if(!trigger)trigger=ScriptApp.newTrigger('maxReminderTick').timeBased().everyMinutes(1).create();
  props.setProperty('MAX_TRIGGER_ID',trigger.getUniqueId());
  existing.slice(1).forEach(t=>ScriptApp.deleteTrigger(t));
  props.setProperty('MAX_REMINDERS_ENABLED','true');
  console.log('MAX подключён: '+chat.title+'. Начало: '+MAX_START_DATE+', 21:00 по Омску. Сообщения сейчас не отправлялись.');
}

/** Only the installed clock trigger can send; remote visitors cannot invoke it. */
function maxReminderTick(event) {
  const props=PropertiesService.getScriptProperties();
  if(!event||String(event.triggerUid)!==props.getProperty('MAX_TRIGGER_ID')||props.getProperty('MAX_REMINDERS_ENABLED')!=='true')return;
  const now=new Date(),date=Utilities.formatDate(now,TIME_ZONE,'yyyy-MM-dd'),time=Utilities.formatDate(now,TIME_ZONE,'HH:mm');
  // Minute-based clock avoids a daily trigger's random hour. Apps Script can still be delayed.
  if(date<MAX_START_DATE||time<'21:00'||time>='21:15')return;
  const lock=LockService.getScriptLock();if(!lock.tryLock(1000))return;
  try{
    const attempt=props.getProperty('MAX_LAST_ATTEMPT_DATE');
    if(attempt===date)return;
    const chatId=props.getProperty('MAX_CHAT_ID');
    if(!chatId||!props.getProperty('MAX_BOT_TOKEN'))throw new Error('Не заполнены настройки MAX.');
    const book=book_();
    const missing=PARTICIPANTS.filter(p=>{const entry=entry_(book,p.id);return !findReport_(book.getSheetById(entry.sheetId),date);});
    const text=missing.length===0?'Все молодцы! Отчеты сданы.':'Пора подвести итоги дня!\nЕщё не сдали отчёт за сегодня: '+missing.map(p=>p.name.split(' ')[0]).join(', ')+'.\nЗаполните форму:\n'+MAX_FORM_URL;
    // Claim before HTTP to prevent a duplicate after a lost acknowledgement.
    // An uncertain delivery is recorded for owner review, not automatically resent.
    props.setProperty('MAX_LAST_ATTEMPT_DATE',date);
    props.setProperty('MAX_LAST_STATUS','sending');
    try{
      const result=maxApi_('/messages?chat_id='+encodeURIComponent(chatId)+'&disable_link_preview=true',{text,notify:true});
      if(!result.message)throw new Error('MAX не подтвердил отправку.');
      props.setProperty('MAX_LAST_SENT_DATE',date);props.setProperty('MAX_LAST_STATUS','sent');
    }catch(e){props.setProperty('MAX_LAST_STATUS','delivery_uncertain');throw new Error('Напоминание MAX: отправка не подтверждена. Проверьте группу перед повтором.');}
  }finally{lock.releaseLock();}
}
function maxApi_(path,payload) {
  const token=PropertiesService.getScriptProperties().getProperty('MAX_BOT_TOKEN');
  if(!token)throw new Error('Добавьте MAX_BOT_TOKEN в свойства скрипта.');
  let response;
  try{response=UrlFetchApp.fetch(MAX_API_URL+path,{method:payload?'post':'get',headers:{Authorization:token},contentType:'application/json',...(payload?{payload:JSON.stringify(payload)}:{}),muteHttpExceptions:true});}
  catch(e){throw new Error('Нет соединения с API MAX. Проверьте доступность сервиса.');}
  const status=response.getResponseCode();
  if(status<200||status>=300)throw new Error('API MAX: HTTP '+status+'. Проверьте токен и права бота.');
  try{return JSON.parse(response.getContentText());}catch(e){throw new Error('API MAX вернул некорректный ответ.');}
}
