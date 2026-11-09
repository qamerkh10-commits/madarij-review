/**
 * نظام متابعة المراجعة — Google Sheet + بوت تليجرام + موقع (GitHub Pages)
 * ------------------------------------------------------------------------
 * الأوراق:  المحاضرات · الفريق · الصلاحيات · السجل · الإعدادات
 * الأدوار:  مالك (أنت، كل الصلاحيات) · مدير · مراجع — وصلاحيات المدير والمراجع بتتظبط من الموقع.
 * الدخول للموقع: أي عضو يبعت /login للبوت فيوصله رابط دخول شخصي لمرة واحدة.
 * البوت: بيبعت المحاضرة PDF بموعدها، وتذكير قبلها بيوم، ويستقبل الملاحظات ويحوّلها للمدراء باسم المراجع،
 *        وملخص الصبح في أيام الشغل، وتذكير المالك بالليل قبل يوم الشغل.
 */

const TZ = 'Africa/Cairo';
const SITE_URL = 'https://qamerkh10-commits.github.io/madarij-review/';
const SH = { LECT: 'المحاضرات', TEAM: 'الفريق', PROJ: 'المشاريع', PERM: 'الصلاحيات', LOG: 'السجل', SET: 'الإعدادات' };

// أعمدة ورقة المحاضرات (1 = A)
const L = { ID: 1, TITLE: 2, FILE: 3, REVIEWER: 4, SENT: 5, DUE: 6, STATUS: 7,
            NOTES: 8, NOTE_FILES: 9, DONE_AT: 10, REMINDED: 11, MSGS: 12, PROJECT: 13 };
const L_HEADERS = ['الرقم', 'عنوان المحاضرة', 'ملف PDF (رابط Drive)', 'المراجع',
  'تاريخ الإرسال', 'موعد التسليم', 'الحالة', 'ملاحظات المراجع', 'ملفات الملاحظات',
  'تاريخ التسليم', 'تنبيه قبلها بيوم', '(داخلي) رسائل تليجرام', 'المشروع'];

// أعمدة ورقة الفريق
const T = { NAME: 1, CHAT: 2, USER: 3, JOINED: 4, ROLE: 5, ACTIVE: 6, NOTIFY: 7, PROJECTS: 8 };
const T_HEADERS = ['الاسم', 'معرّف تليجرام', 'حساب تليجرام', 'تاريخ الانضمام', 'الدور', 'نشط', 'يستلم الإشعارات', 'المشاريع (فاضي = الكل)'];

// المشاريع: [الاسم، اسم العنصر، اللون]
const P_HEADERS = ['المشروع', 'اسم العنصر', 'اللون', 'أيام الشغل (فاضي = العامة)', 'نشط'];
const DEFAULT_PROJECTS = [
  ['أخصر', 'محاضرة', '#2E8B85'], ['إناسة', 'حلقة', '#7B68C8'], ['مقدمة', 'محاضرة', '#D08A2E'],
  ['الروض', 'محاضرة', '#3B7FC0'], ['مشروع الكتب', 'كتاب', '#B3362B'], ['المقالات', 'مقالة', '#2F9E66']];
const ROLE = { OWNER: 'مالك', MANAGER: 'مدير', REVIEWER: 'مراجع' };

// الصلاحيات: [المفتاح، الوصف، افتراضي المدير، افتراضي المراجع] — المالك عنده كله دايمًا
const PERMS = [
  ['view_all', 'رؤية لوحة التحكم وكل المحاضرات والمراجعين والسجل', true, false],
  ['send', 'إرسال محاضرات للمراجعة', true, false],
  ['status', 'تغيير حالة المحاضرات', true, false],
  ['team', 'إدارة المراجعين (التفعيل والإشعارات والأسماء)', true, false],
  ['settings', 'تعديل الإعدادات (أيام الشغل ومواعيد التنبيه)', false, false],
];

const ST = { NEW: 'جديدة', SENT: 'عند المراجع', LATE: 'متأخرة', DONE: 'سُلّمت',
             FIXED: 'عُدّلت', PUB: 'نُشرت' };
const ST_COLORS = [[ST.NEW, '#eeeeee'], [ST.SENT, '#fff4cc'], [ST.LATE, '#f8d0cc'],
  [ST.DONE, '#d4edda'], [ST.FIXED, '#d0e4f5'], [ST.PUB, '#c9c3e6']];

const DAY_NAMES = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
const MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس',
  'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
const MAX_UPLOAD = 50 * 1024 * 1024;   // حد تليجرام لرفع البوت
const NOTES_FOLDER = 'ملاحظات المراجعة';
const SESSION_DAYS = 30;

// ================================================================
// القائمة
// ================================================================
function onOpen() {
  SpreadsheetApp.getUi().createMenu('🤖 بوت المراجعة')
    .addItem('📤 إرسال المحاضرات الجديدة', 'sendNewLectures')
    .addItem('📊 ابعت لي الملخص دلوقتي', 'sendDigestNow')
    .addSeparator()
    .addItem('⚙️ إعداد النظام', 'setup')
    .addItem('🔗 ربط الموقع', 'connectSite')
    .addItem('🔄 تحديث الأوراق', 'upgradeSheets')
    .addItem('⛔ إيقاف البوت', 'stopBot')
    .addToUi();
}

function setup() {
  fresh_();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('TOKEN');
  const r = ui.prompt('إعداد البوت',
    'الصق توكن البوت اللي أخدته من @BotFather' +
    (token ? '\n(سيبه فاضي لو عايز تحتفظ بالتوكن الحالي)' : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const typed = r.getResponseText().trim();
  if (typed) token = typed;
  if (!token) { ui.alert('لازم تحط توكن البوت.'); return; }
  props.setProperty('TOKEN', token);

  let me;
  try { me = tg_('getMe'); } catch (e) { ui.alert('التوكن مش صحيح: ' + e.message); return; }
  // لو الموقع مربوط: استقبال فوري. غير كده: سحب الرسايل كل دقيقة لحد ما يتربط
  let instant = false;
  try { instant = enableWebhook_(); } catch (e) { instant = false; }
  if (!instant) { tg_('deleteWebhook'); props.deleteProperty('WEBHOOK'); }
  props.setProperty('BOT_USERNAME', me.username);
  props.setProperty('SHEET_ID', SpreadsheetApp.getActiveSpreadsheet().getId());

  ensureSheets_();
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'tick')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(1).create();
  try {
    tg_('setMyCommands', { commands: [
      { command: 'login', description: 'رابط الدخول للموقع' },
      { command: 'my', description: 'المحاضرات اللي عليّ' },
      { command: 'help', description: 'طريقة الاستخدام' }] });
  } catch (e) {}

  const code = String(Math.floor(1000 + Math.random() * 9000));
  props.setProperty('ADMIN_CODE', code);
  props.setProperty('ADMIN_CODE_AT', String(nowDate_().getTime()));   // الكود بيشتغل ٣٠ دقيقة بس
  ui.alert('تم الإعداد ✅',
    'البوت: @' + me.username + '\n\n' +
    '١) عشان تبقى المالك: افتح البوت من حسابك وابعت:\n/admin ' + code + '\n\n' +
    '٢) كل عضو في الفريق يفتح البوت ويضغط Start، فاسمه ينزل لوحده في ورقة «الفريق» كمراجع.\n\n' +
    '٣) اربط الموقع من القائمة: 🔗 ربط الموقع.', ui.ButtonSet.OK);
}

/** بعد تحديث الكود: يضيف الأوراق والأعمدة الجديدة من غير ما يلمس البيانات أو التوكن */
function upgradeSheets() {
  fresh_();
  ensureSheets_();
  CacheService.getScriptCache().remove('SETTINGS_V1');
  SpreadsheetApp.getUi().alert('تم تحديث الأوراق ✅\nاتضافت ورقة «المشاريع» والأعمدة الجديدة، وكل البيانات زي ما هي.');
}

function stopBot() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'tick')
    .forEach(t => ScriptApp.deleteTrigger(t));
  SpreadsheetApp.getUi().alert('اتوقف البوت. تقدر تشغله تاني من «⚙️ إعداد النظام».');
}

/** بياخد رابط الـ Web app مرة واحدة، وبعدها أي عضو يدخل الموقع بـ /login من البوت */
function connectSite() {
  fresh_();
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  let api = props.getProperty('API_URL') || '';
  const r = ui.prompt('ربط الموقع',
    'الصق رابط الـ Web app (اللي آخره /exec) من Deploy ← Manage deployments' +
    (api ? '\n\nالرابط الحالي:\n' + api + '\n(سيبه فاضي عشان تحتفظ بيه)' : ''), ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const typed = r.getResponseText().trim();
  if (typed) api = typed;
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/.test(api)) {
    ui.alert('الرابط ده مش رابط Web app. لازم يبدأ بـ https://script.google.com/macros/s/ وآخره /exec');
    return;
  }
  props.setProperty('API_URL', api);
  let instant = '';
  try { instant = enableWebhook_() ? '\n\n⚡ البوت بقى بيرد فورًا.' : ''; }
  catch (e) { instant = '\n\n⚠️ الرد الفوري ما اشتغلش (' + e.message + ')، والبوت هيفضل يرد في خلال دقيقة.'; }
  ui.alert('تم الربط ✅',
    'دلوقتي أي حد في الفريق يبعت للبوت:\n/login\nفيوصله زرار «🔐 ادخل الموقع» بحسابه وصلاحياته.\n\n' +
    'الموقع: ' + SITE_URL + instant, ui.ButtonSet.OK);
}

function sendDigestNow() {
  fresh_();
  const to = managersToNotify_();
  if (!to.length) {
    SpreadsheetApp.getUi().alert('لسه مفيش مالك. ابعت للبوت /admin والكود اللي ظهر في الإعداد.');
    return;
  }
  to.forEach(m => send_(m.chat, digestText_(m)));
  SpreadsheetApp.getUi().alert('اتبعت الملخص على تليجرام ✅');
}

// ================================================================
// إرسال المحاضرات من الشيت (الصفوف اللي حالتها «جديدة»)
// ================================================================
function sendNewLectures() {
  fresh_();
  const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const res = sendNewLectures_();
    if (!res.ok.length && !res.bad.length) {
      ui.alert('مفيش محاضرات جديدة. اكتب العنوان والملف والمراجع، وسيب «الحالة» فاضية أو «جديدة».');
      return;
    }
    let msg = '';
    if (res.ok.length) msg += '✅ اتبعت (' + res.ok.length + '):\n' + res.ok.join('\n');
    if (res.bad.length) msg += (msg ? '\n\n' : '') + '⚠️ ما اتبعتش (' + res.bad.length + '):\n' + res.bad.join('\n');
    ui.alert(msg);
  } finally {
    lock.releaseLock();
  }
}

function sendNewLectures_() {
  const s = getSettings_();
  const lects = readLectures_();
  const revs = readReviewers_();
  let nextId = lects.reduce((m, l) => Math.max(m, l.id), 0) + 1;
  const ok = [], bad = [];

  lects.filter(l => l.title && (!l.status || l.status === ST.NEW)).forEach(l => {
    const name = '«' + l.title + '»';
    if (!l.reviewer) { bad.push(name + ': مفيش مراجع'); return; }
    const rev = revs.find(r => r.name === l.reviewer);
    if (!rev) { bad.push(name + ': «' + l.reviewer + '» مش مسجّل في البوت أو مش نشط'); return; }
    const id = l.id || nextId;
    const due = ymdOf_(l.due) || computeDue_(todayYmd_(), projWorkDays_(l.project));
    let msg;
    try { msg = deliver_(l.title, l.file, rev, id, due, l.project); }
    catch (e) { bad.push(name + ': ' + e.message); return; }
    if (id === nextId) nextId++;

    const sh = sheet_(SH.LECT);
    sh.getRange(l.row, L.ID).setValue(id);
    sh.getRange(l.row, L.SENT).setValue(nowDate_());
    sh.getRange(l.row, L.DUE).setValue(dateOf_(due));
    sh.getRange(l.row, L.STATUS).setValue(ST.SENT);
    sh.getRange(l.row, L.REMINDED).setValue('');
    sh.getRange(l.row, L.MSGS).setValue(String(msg.message_id));
    sh.getRange(l.row, L.PROJECT).setValue(l.project);
    log_('الشيت', 'أرسل للمراجعة ← ' + rev.name, l.title, l.project);
    ok.push(name + ' ← ' + rev.name + ' (لحد ' + arDate_(due) + ')');
  });
  return { ok: ok, bad: bad };
}

/** يبعت ملف المحاضرة للمراجع ويرجّع رسالة تليجرام، أو يرمي خطأ برسالة عربية مفهومة */
function deliver_(title, link, rev, id, due, project) {
  const P = projectOf_(project || defaultProject_());
  const fileId = driveId_(link);
  if (!fileId) throw new Error('رابط الملف ناقص أو مش مفهوم (لازم رابط من Google Drive)');
  let blob;
  try {
    const f = DriveApp.getFileById(fileId);
    if (f.getSize() > MAX_UPLOAD) throw new Error('BIG');
    blob = f.getBlob().setName(f.getName());
  } catch (e) {
    throw new Error(e.message === 'BIG' ? 'الملف أكبر من 50 ميجا (حد تليجرام)'
      : 'مش قادر أفتح الملف من Drive (اتأكد إنه على حسابك أو متشارك معاك)');
  }
  const caption =
    '📚 <b>' + esc_(P.item) + ' للمراجعة — ' + esc_(P.name) + '</b>\n«' + esc_(title) + '»\n\n' +
    '🗓 موعد التسليم: <b>' + arDate_(due) + '</b>\n\n' +
    '✍️ ابعت ملاحظاتك <b>ردًّا (Reply) على الرسالة دي</b> — نص أو ملف أو صورة أو رسالة صوتية — ' +
    'ولما تخلص اضغط «تم التسليم».';
  try {
    return tg_('sendDocument', { chat_id: rev.chat, document: blob, caption: caption,
      parse_mode: 'HTML', reply_markup: doneKeyboard_(id) });
  } catch (e) { throw new Error('تليجرام رفض — ' + e.message); }
}

// ================================================================
// API للموقع (Web App) — كل طلب لازم يكون معاه جلسة دخول صالحة
// ================================================================
const API = {
  bootstrap: apiBootstrap_, send: apiSend_, setStatus: apiSetStatus_, addNote: apiAddNote_,
  markDone: apiMarkDone_, updateMember: apiUpdateMember_, setPermissions: apiSetPermissions_,
  saveSettings: apiSaveSettings_, saveProjects: apiSaveProjects_, logout: apiLogout_,
};

function doGet() { return json_({ ok: true, result: 'API شغال ✅' }); }

// العمليات اللي بتكتب: بتاخد قفل، وبترجّع البيانات الجديدة في نفس الرد (من غير رحلة تانية)
const API_WRITES = { send: 1, setStatus: 1, addNote: 1, markDone: 1, updateMember: 1, setPermissions: 1, saveSettings: 1, saveProjects: 1 };

function doPost(e) {
  fresh_();
  if (e && e.parameter && e.parameter.tg) return telegramWebhook_(e);
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (req.action === 'login') return json_({ ok: true, result: apiLogin_(String(req.token || '')) });
    const ses = String(req.session || '');
    const user = sessionUser_(ses);
    if (!user) return json_({ ok: false, auth: true, error: 'انتهت الجلسة — ابعت /login للبوت وادخل تاني' });
    const fn = API[req.action];
    if (!fn) return json_({ ok: false, error: 'طلب مش معروف' });
    if (!API_WRITES[req.action]) return json_({ ok: true, result: fn(user, req.args || {}, ses) });

    const lock = LockService.getScriptLock();
    let result;
    lock.waitLock(20000);
    try { result = fn(user, req.args || {}, ses); } finally { lock.releaseLock(); }
    fresh_();
    const u2 = sessionUser_(ses);
    return json_({ ok: true, result: result, boot: u2 ? apiBootstrap_(u2) : null });
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ---------- استقبال فوري من تليجرام (webhook) ----------
/** تليجرام بيبعت كل رسالة هنا في نفس اللحظة بدل ما نسحبها كل دقيقة */
function telegramWebhook_(e) {
  const ok = HtmlService.createHtmlOutput('ok');   // ردّ 200 مباشر (من غير تحويل) عشان تليجرام ما يعيدش الإرسال
  const secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_SECRET');
  if (!secret || e.parameter.tg !== secret) return ok;
  let u;
  try { u = JSON.parse(e.postData.contents); } catch (err) { return ok; }
  if (!u || u.update_id === undefined) return ok;
  const cache = CacheService.getScriptCache();
  if (cache.get('U_' + u.update_id)) return ok;    // نفس الرسالة اتعالجت قبل كده
  cache.put('U_' + u.update_id, '1', 21600);
  // الأوامر اللي بتقرا بس (/login /my /help /status) بترد على طول من غير قفل
  const readOnly = u.message && /^\/(login|my|help|status)\b/.test((u.message.text || '').trim());
  const lock = readOnly ? null : LockService.getScriptLock();
  try {
    if (lock) lock.waitLock(25000);
    handleUpdate_(u);
  } catch (err) {
    console.error('webhook ' + u.update_id + ': ' + (err && err.stack));
  } finally {
    if (lock) try { lock.releaseLock(); } catch (x) {}
  }
  return ok;
}

/** يشغّل الاستقبال الفوري لو الموقع مربوط (محتاج رابط الـ Web app) */
function enableWebhook_() {
  const props = PropertiesService.getScriptProperties();
  const api = props.getProperty('API_URL');
  if (!api || !props.getProperty('TOKEN')) return false;
  let secret = props.getProperty('WEBHOOK_SECRET');
  if (!secret) { secret = newToken_(); props.setProperty('WEBHOOK_SECRET', secret); }
  tg_('setWebhook', { url: api + '?tg=' + secret, allowed_updates: ['message', 'callback_query'], max_connections: 5 });
  props.setProperty('WEBHOOK', '1');
  return true;
}

// ---------- الدخول والجلسات ----------
function apiLogin_(token) {
  if (!/^[a-f0-9]{32}$/.test(token)) throw new Error('رابط الدخول مش صحيح');
  const cache = CacheService.getScriptCache();
  const chat = cache.get('LOGIN_' + token);
  if (!chat) throw new Error('رابط الدخول انتهى أو اتستخدم قبل كده — ابعت /login للبوت تاني');
  cache.remove('LOGIN_' + token);
  const m = readTeam_().find(x => x.chat === chat && x.active);
  if (!m) throw new Error('حسابك مش مفعّل في الفريق — كلّم المسؤول');
  const ses = newToken_() + newToken_();
  PropertiesService.getScriptProperties().setProperty('SES_' + ses,
    JSON.stringify({ chat: chat, exp: nowDate_().getTime() + SESSION_DAYS * 86400000 }));
  return { session: ses, name: m.name };
}

function sessionUser_(ses) {
  if (!/^[a-f0-9]{64}$/.test(ses)) return null;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty('SES_' + ses);
  if (!raw) return null;
  const s = JSON.parse(raw);
  if (s.exp < nowDate_().getTime()) { props.deleteProperty('SES_' + ses); return null; }
  const m = readTeam_().find(x => x.chat === s.chat && x.active);
  return m ? userOf_(m) : null;
}

function apiLogout_(u, args, ses) {
  PropertiesService.getScriptProperties().deleteProperty('SES_' + ses);
  return true;
}

function cleanSessions_() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties(), now = nowDate_().getTime();
  Object.keys(all).forEach(k => {
    if (k.indexOf('SES_') !== 0) return;
    try { if (JSON.parse(all[k]).exp < now) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); }
  });
}

function newToken_() { return Utilities.getUuid().replace(/-/g, ''); }

function userOf_(m) {
  const owner = m.chat === adminChat_();
  const role = owner ? ROLE.OWNER : (m.role === ROLE.MANAGER ? ROLE.MANAGER : ROLE.REVIEWER);
  const perms = {};
  const matrix = readPerms_();
  PERMS.forEach(p => { perms[p[0]] = owner ? true : !!matrix[role === ROLE.MANAGER ? 'manager' : 'reviewer'][p[0]]; });
  return { chat: m.chat, name: m.name, role: role, isOwner: owner, perms: perms, projects: owner ? [] : (m.projects || []) };
}

function need_(u, perm) {
  if (!u.isOwner && !u.perms[perm]) throw new Error('مش عندك صلاحية للعملية دي');
}

// ---------- البيانات ----------
function apiBootstrap_(u) {
  const s = getSettings_();
  const props = PropertiesService.getScriptProperties();
  let bot = props.getProperty('BOT_USERNAME') || '';
  if (!bot && props.getProperty('TOKEN')) {
    try { bot = tg_('getMe').username; props.setProperty('BOT_USERNAME', bot); } catch (e) {}
  }
  const today = todayYmd_();
  const projs = readProjects_();
  const all = readLectures_().filter(l => l.title);
  const visible = u.perms.view_all ? all.filter(l => sees_(u, l.project) || l.reviewer === u.name) : all.filter(l => l.reviewer === u.name);
  const canTeam = u.isOwner || u.perms.team || u.perms.view_all;
  return {
    me: u,
    bot: bot,
    today: today,
    nextWorkDay: nextWorkDay_(today, s.workDays),
    defaultDue: computeDue_(today, s.workDays),
    settings: { workDays: s.workDays, hour: s.hour, adminHour: s.adminHour },
    projects: projs.filter(p => u.isOwner || u.perms.settings || sees_(u, p.name)).map(p => {
      const wd = p.workDays || s.workDays;
      return { name: p.name, item: p.item, color: p.color, workDays: p.workDays, active: p.active, mine: sees_(u, p.name),
        nextWorkDay: nextWorkDay_(today, wd), defaultDue: computeDue_(today, wd) };
    }),
    statuses: Object.keys(ST).map(k => ST[k]),
    reviewers: u.perms.send ? readReviewers_().map(r => r.name) : [],
    lectures: visible.reverse().map(l => lecOut_(l, today, u)),
    team: canTeam ? readTeam_().map(m => ({
      chat: m.chat, name: m.name, username: m.username, joined: ymdOf_(m.joined),
      role: m.chat === adminChat_() ? ROLE.OWNER : (m.role === ROLE.MANAGER ? ROLE.MANAGER : ROLE.REVIEWER),
      active: m.active, notify: m.notify, projects: m.projects })) : [],
    permsDef: PERMS.map(p => ({ key: p[0], label: p[1] })),
    permissions: (u.isOwner || u.perms.team) ? readPerms_() : null,
    log: u.perms.view_all ? readLog_(300).filter(a => !a.project || sees_(u, a.project)) : [],
  };
}

function lecOut_(l, today, u) {
  const due = ymdOf_(l.due);
  return {
    id: l.id, title: l.title, file: l.file, reviewer: l.reviewer, status: l.status || ST.NEW, project: l.project,
    sent: ymdOf_(l.sent), due: due, done: ymdOf_(l.doneAt), reminded: !!ymdOf_(l.reminded),
    lateDays: isOpen_(l) && due && due < today ? daysBetween_(due, today) : 0,
    notes: l.notes.trim(), files: l.noteFiles.split('\n').map(x => x.trim()).filter(Boolean),
    mine: l.reviewer === u.name,
  };
}

/** {project, items: [{title, link, reviewer, due}]} — كل عنصر يتبعت الأول، ولو نجح بس ينزل في الشيت */
function apiSend_(u, args) {
  need_(u, 'send');
  const project = String(args.project || '').trim();
  const P = readProjects_().find(x => x.name === project);
  if (!P) throw new Error('اختار المشروع');
  if (!P.active) throw new Error('المشروع ده متوقّف');
  if (!sees_(u, project)) throw new Error('المشروع ده مش من مشاريعك');
  const wd = P.workDays || getSettings_().workDays;
  const revs = readReviewers_();
  let nextId = readLectures_().reduce((m, l) => Math.max(m, l.id), 0) + 1;
  const ok = [], bad = [];
  (args.items || []).forEach(it => {
    const title = String(it.title || '').trim(), link = String(it.link || '').trim();
    if (!title && !link) return;
    const name = '«' + (title || 'بدون عنوان') + '»';
    const keep = { title: title, link: link, reviewer: it.reviewer, due: it.due };
    if (!title) { bad.push(Object.assign(keep, { error: name + ': اكتب عنوان ال' + P.item })); return; }
    const rev = revs.find(r => r.name === String(it.reviewer || '').trim());
    if (!rev) { bad.push(Object.assign(keep, { error: name + ': اختار المراجع' })); return; }
    const due = ymdOf_(it.due) || computeDue_(todayYmd_(), wd);
    let msg;
    try { msg = deliver_(title, link, rev, nextId, due, project); }
    catch (e) { bad.push(Object.assign(keep, { due: due, error: name + ': ' + e.message })); return; }
    sheet_(SH.LECT).appendRow([nextId, title, link, rev.name, nowDate_(), dateOf_(due), ST.SENT, '', '', '', '', String(msg.message_id), project]);
    log_(u.name, 'أرسل للمراجعة ← ' + rev.name, title, project);
    ok.push(name + ' ← ' + rev.name + ' (لحد ' + arDate_(due) + ')');
    nextId++;
  });
  return { ok: ok, bad: bad };
}

function apiSetStatus_(u, args) {
  need_(u, 'status');
  const status = String(args.status || '');
  if (Object.keys(ST).map(k => ST[k]).indexOf(status) === -1) throw new Error('حالة مش معروفة');
  const lec = findLecture_(args.id);
  if (!sees_(u, lec.project)) throw new Error('المشروع ده مش من مشاريعك');
  if (lec.status === status) return true;
  sheet_(SH.LECT).getRange(lec.row, L.STATUS).setValue(status);
  log_(u.name, 'غيّر الحالة: ' + (lec.status || ST.NEW) + ' ← ' + status, lec.title, lec.project);
  return true;
}

function apiAddNote_(u, args) {
  const lec = findLecture_(args.id);
  if (lec.reviewer !== u.name && !u.isOwner && !(u.perms.status && sees_(u, lec.project))) throw new Error('المحاضرة دي مش متسجّلة باسمك');
  const text = String(args.text || '').trim();
  if (!text) throw new Error('اكتب الملاحظة الأول');
  if (text.length > 3000) throw new Error('الملاحظة طويلة قوي — قسّمها على أكتر من ملاحظة');
  const first = !lec.notes.trim() && !lec.noteFiles.trim();
  const old = lec.notes.trim();
  sheet_(SH.LECT).getRange(lec.row, L.NOTES).setValue((old ? old + '\n' : '') + '• ' + text);
  if (first) log_(u.name, 'بدأ يكتب ملاحظات', lec.title, lec.project);
  managersToNotify_().filter(m => m.chat !== u.chat && sees_(m, lec.project)).forEach(m =>
    send_(m.chat, '📝 <b>' + esc_(u.name) + '</b> — «' + esc_(lec.title) + '» · ' + esc_(lec.project) + ' <i>(من الموقع)</i>\n\n' + esc_(text)));
  return true;
}

function apiMarkDone_(u, args) {
  const lec = findLecture_(args.id);
  if (lec.reviewer !== u.name && !u.isOwner && !(u.perms.status && sees_(u, lec.project))) throw new Error('المحاضرة دي مش متسجّلة باسمك');
  if (!isOpen_(lec)) throw new Error('المحاضرة دي اتسلّمت قبل كده');
  markDone_(lec, lec.reviewer, u.name === lec.reviewer ? 'من الموقع' : 'سجّلها ' + u.name);
  return true;
}

function apiUpdateMember_(u, args) {
  need_(u, 'team');
  const team = readTeam_();
  const t = team.find(m => m.chat === String(args.chat));
  if (!t) throw new Error('العضو مش موجود');
  const targetOwner = t.chat === adminChat_();
  const targetRole = targetOwner ? ROLE.OWNER : (t.role === ROLE.MANAGER ? ROLE.MANAGER : ROLE.REVIEWER);
  if (!u.isOwner && targetRole !== ROLE.REVIEWER) throw new Error('تعديل المدراء للمالك بس');
  const sh = sheet_(SH.TEAM);
  const changes = [];

  if (args.name !== undefined) {
    const name = String(args.name).trim();
    if (!name) throw new Error('الاسم مينفعش يبقى فاضي');
    if (name !== t.name) {
      if (team.some(m => m.name === name)) throw new Error('الاسم ده مستخدم لعضو تاني');
      sh.getRange(t.row, T.NAME).setValue(name);
      readLectures_().filter(l => l.reviewer === t.name)        // المحاضرات بتتربط بالاسم
        .forEach(l => sheet_(SH.LECT).getRange(l.row, L.REVIEWER).setValue(name));
      changes.push('غيّر الاسم من «' + t.name + '» إلى «' + name + '»');
    }
  }
  if (args.role !== undefined && args.role !== targetRole) {
    if (!u.isOwner) throw new Error('تغيير الأدوار للمالك بس');
    if (targetOwner) throw new Error('المالك ميتغيّرش دوره من هنا (بيتنقل بكود /admin جديد)');
    if ([ROLE.MANAGER, ROLE.REVIEWER].indexOf(args.role) === -1) throw new Error('دور مش معروف');
    sh.getRange(t.row, T.ROLE).setValue(args.role);
    changes.push('خلّى «' + t.name + '» ' + args.role);
  }
  if (args.active !== undefined && !!args.active !== t.active) {
    if (targetOwner) throw new Error('المالك ميتوقّفش');
    sh.getRange(t.row, T.ACTIVE).setValue(!!args.active);
    changes.push((args.active ? 'فعّل' : 'وقّف') + ' «' + t.name + '»');
  }
  if (args.notify !== undefined && !!args.notify !== t.notify) {
    sh.getRange(t.row, T.NOTIFY).setValue(!!args.notify);
    changes.push((args.notify ? 'شغّل' : 'قفل') + ' إشعارات «' + t.name + '»');
  }
  if (args.projects !== undefined) {
    if (!u.isOwner) throw new Error('تحديد مشاريع الأعضاء للمالك بس');
    const names = readProjects_().map(x => x.name);
    const list = (args.projects || []).map(String).filter(x => names.indexOf(x) !== -1);
    const val = list.join('، ');
    if (val !== t.projects.join('، ')) {
      sh.getRange(t.row, T.PROJECTS).setValue(val);
      changes.push('حدّد مشاريع «' + t.name + '»: ' + (list.length ? val : 'كل المشاريع'));
    }
  }
  changes.forEach(c => log_(u.name, c, ''));
  return true;
}

/** إضافة/تعديل/ترتيب المشاريع. اسم جديد لمشروع قديم بيتطبّق على محاضراته وعلى أعضائه */
function apiSaveProjects_(u, args) {
  need_(u, 'settings');
  const list = (args.projects || []).map(x => ({
    name: String(x.name || '').trim(), old: String(x.old || '').trim(), item: String(x.item || '').trim() || 'محاضرة',
    color: /^#[0-9a-f]{6}$/i.test(String(x.color || '')) ? x.color : '#2E8B85',
    workDays: (x.workDays || []).map(Number).filter(d => d >= 0 && d <= 6), active: x.active !== false }));
  if (!list.length) throw new Error('لازم يبقى فيه مشروع واحد على الأقل');
  const seen = {};
  list.forEach(x => {
    if (!x.name) throw new Error('اسم المشروع مينفعش يبقى فاضي');
    if (seen[x.name]) throw new Error('اسم المشروع «' + x.name + '» متكرر');
    seen[x.name] = 1;
  });
  list.filter(x => x.old && x.old !== x.name).forEach(x => {
    readLectures_().filter(l => l.project === x.old).forEach(l => sheet_(SH.LECT).getRange(l.row, L.PROJECT).setValue(x.name));
    readTeam_().filter(m => m.projects.indexOf(x.old) !== -1).forEach(m =>
      sheet_(SH.TEAM).getRange(m.row, T.PROJECTS).setValue(m.projects.map(y => y === x.old ? x.name : y).join('، ')));
    log_(u.name, 'غيّر اسم المشروع من «' + x.old + '» إلى «' + x.name + '»', '', x.name);
  });
  const sh = sheet_(SH.PROJ);
  const n = sh.getLastRow();
  if (n > 1) sh.getRange(2, 1, n - 1, P_HEADERS.length).clearContent();
  sh.getRange(2, 1, list.length, P_HEADERS.length).setValues(list.map(x =>
    [x.name, x.item, x.color, x.workDays.map(d => DAY_NAMES[d]).join('، '), x.active]));
  log_(u.name, 'عدّل المشاريع', '');
  fresh_();
  return true;
}

function apiSetPermissions_(u, args) {
  if (!u.isOwner) throw new Error('تعديل الصلاحيات للمالك بس');
  const m = args.matrix || {};
  const sh = sheet_(SH.PERM);
  PERMS.forEach((p, i) => {
    sh.getRange(i + 2, 3).setValue(!!(m.manager && m.manager[p[0]]));
    sh.getRange(i + 2, 4).setValue(!!(m.reviewer && m.reviewer[p[0]]));
  });
  log_(u.name, 'عدّل الصلاحيات', '');
  fresh_();
  return readPerms_();
}

function apiSaveSettings_(u, args) {
  need_(u, 'settings');
  const days = (args.workDays || []).map(Number).filter(d => d >= 0 && d <= 6);
  if (!days.length) throw new Error('اختار يوم شغل واحد على الأقل');
  const hour = parseInt(args.hour, 10), adminHour = parseInt(args.adminHour, 10);
  if (!(hour >= 0 && hour <= 23) || !(adminHour >= 0 && adminHour <= 23)) throw new Error('الساعة لازم تبقى من 0 لـ 23');
  setSetting_('أيام الشغل', days.map(d => DAY_NAMES[d]).join('، '));
  setSetting_('ساعة التنبيه', hour);
  setSetting_('ساعة تنبيهي', adminHour);
  log_(u.name, 'عدّل الإعدادات', '');
  return getSettings_();
}

function findLecture_(id) {
  const lec = readLectures_().find(l => l.id === Number(id));
  if (!lec) throw new Error('المحاضرة مش موجودة');
  return lec;
}

// ================================================================
// المؤقّت: كل دقيقة يسحب رسائل البوت، ومرة في اليوم يبعت التنبيهات
// ================================================================
function tick() {
  fresh_();
  // في وضع الاستقبال الفوري: المؤقّت مالوش شغل غير التنبيهات اليومية، فما يقفلش النظام من غير داعي
  if (PropertiesService.getScriptProperties().getProperty('WEBHOOK') === '1' && !dailyDue_()) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    pollUpdates_();
    maybeDaily_();
  } finally {
    lock.releaseLock();
  }
}

function pollUpdates_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('WEBHOOK') === '1') return;    // الرسايل بتوصل فورًا عن طريق doPost
  const offset = Number(props.getProperty('OFFSET') || 0);
  const updates = tg_('getUpdates', { offset: offset, timeout: 0,
    allowed_updates: ['message', 'callback_query'] });
  updates.forEach(u => {
    props.setProperty('OFFSET', String(u.update_id + 1));   // قبل المعالجة: الرسالة المعطوبة ما تتكررش
    try { handleUpdate_(u); } catch (e) { console.error('update ' + u.update_id + ': ' + e.stack); }
  });
}

function dailyDue_() {
  const props = PropertiesService.getScriptProperties();
  const now = nowDate_(), today = ymdOf_(now), hour = Number(Utilities.formatDate(now, TZ, 'H'));
  const s = getSettings_();
  return (props.getProperty('DAILY') !== today && hour >= s.hour) ||
         (props.getProperty('ADMIN_EVE') !== today && hour >= s.adminHour);
}

function maybeDaily_() {
  const props = PropertiesService.getScriptProperties();
  const now = nowDate_();
  const today = ymdOf_(now);
  const hour = Number(Utilities.formatDate(now, TZ, 'H'));
  if (props.getProperty('DAILY') !== today && hour >= getSettings_().hour) {
    props.setProperty('DAILY', today);
    dailyJob_(today);
  }
  if (props.getProperty('ADMIN_EVE') !== today && hour >= getSettings_().adminHour) {
    props.setProperty('ADMIN_EVE', today);
    adminEveJob_(today);
  }
}

/** تنبيه المالك (مشرف المونتاج) قبل يوم الشغل بيوم: إيه اللي مستنيه بكرة */
function adminEveJob_(today) {
  const tomorrow = addDays_(today, 1);
  if (!adminChat_()) return;
  const projs = readProjects_().filter(p => p.active && projWorkDays_(p.name).indexOf(dow_(tomorrow)) !== -1);
  if (!projs.length) return;
  send_(adminChat_(), workDayText_(tomorrow, projs));
}

function workDayText_(day, projs) {
  const ls = readLectures_().filter(l => l.title);
  const out = ['🎬 <b>تذكير: بكرة يوم شغل — ' + arDate_(day) + '</b>'];
  projs.forEach(p => {
    const pl = ls.filter(l => l.project === p.name);
    const by = st => pl.filter(l => l.status === st);
    const done = by(ST.DONE), fixed = by(ST.FIXED), open = pl.filter(isOpen_);
    out.push('', '━━━━━━━━━━', '📁 <b>' + esc_(p.name) + '</b>');
    if (done.length) out.push('✂️ تعديل ورندر ورفع (' + done.length + ') — الملاحظات وصلت:', ...done.map(l => '  • «' + esc_(l.title) + '»'));
    if (fixed.length) out.push('📤 متعدّلة ومستنية الرفع (' + fixed.length + '):', ...fixed.map(l => '  • «' + esc_(l.title) + '»'));
    if (open.length) out.push('⏳ لسه عند المراجعين (' + open.length + '):', ...open.map(l => '  • «' + esc_(l.title) + '» — ' + esc_(l.reviewer) + ' — ' +
      (l.status === ST.LATE ? '⚠️ متأخرة' : 'لحد ' + arDate_(ymdOf_(l.due)))));
    out.push('🆕 تجهيز الدفعة الجديدة وإرسالها للمراجعة.');
  });
  return out.join('\n');
}

function dailyJob_(today) {
  const s = getSettings_();
  const tomorrow = addDays_(today, 1);
  const revs = readReviewers_();
  const sh = sheet_(SH.LECT);

  readLectures_().forEach(l => {
    if (l.status !== ST.SENT) return;
    const due = ymdOf_(l.due);
    const rev = revs.find(r => r.name === l.reviewer);
    if (!due || !rev) return;

    if (due === tomorrow && !l.reminded) {
      const m = send_(rev.chat,
        '⏰ <b>تذكير</b>\nميعاد تسليم مراجعة «' + esc_(l.title) + '» (' + esc_(l.project) + ') <b>بكرة، ' + arDate_(due) + '</b>.\n\n' +
        'لو خلصت: ابعت الملاحظات ردًّا على رسالة المحاضرة واضغط «تم التسليم».',
        { reply_markup: doneKeyboard_(l.id), reply_to: l.msgs[0] });
      sh.getRange(l.row, L.REMINDED).setValue(nowDate_());
      if (m) sh.getRange(l.row, L.MSGS).setValue(l.msgs.concat(String(m.message_id)).join(','));
    } else if (due < today) {
      sh.getRange(l.row, L.STATUS).setValue(ST.LATE);
      log_('النظام', 'المراجعة اتأخرت عند ' + l.reviewer, l.title, l.project);
      send_(rev.chat,
        '⏳ ميعاد مراجعة «' + esc_(l.title) + '» (' + esc_(l.project) + ') كان ' + arDate_(due) + '.\n' +
        'لو خلصت ابعت الملاحظات واضغط «تم التسليم»، ولو محتاج وقت قول للمسؤول.',
        { reply_markup: doneKeyboard_(l.id), reply_to: l.msgs[0] });
    }
  });

  fresh_();
  const workToday = readProjects_().some(p => p.active && projWorkDays_(p.name).indexOf(dow_(today)) !== -1);
  if (workToday) managersToNotify_().forEach(m => send_(m.chat, digestText_(m)));
  cleanSessions_();
}

// ================================================================
// رسائل البوت
// ================================================================
function handleUpdate_(u) {
  if (u.callback_query) return onCallback_(u.callback_query);
  const m = u.message;
  if (!m || !m.chat || m.chat.type !== 'private') return;
  const text = (m.text || '').trim();
  if (/^\/start\b/.test(text)) return onStart_(m, text.replace(/^\/start\s*/, ''));
  if (/^\/admin\b/.test(text)) return onAdmin_(m, text);
  if (/^\/login\b/.test(text)) return onLogin_(m);
  if (/^\/status\b/.test(text)) return onStatus_(m);
  if (/^\/my\b/.test(text)) return onMy_(m);
  if (/^\/help\b/.test(text)) return onHelp_(m);
  if (/^\//.test(text)) return send_(m.chat.id, 'أمر مش معروف. ابعت /help');
  return onNote_(m);
}

function onStart_(m, payload) {
  const chat = String(m.chat.id);
  const member = readTeam_().find(r => r.chat === chat);
  if (member) {
    if (payload === 'login') return onLogin_(m);
    send_(chat, 'أهلًا ' + esc_(member.name) + ' 👋\nأنت متسجّل. أي محاضرة عليك مراجعتها هتوصلك هنا بميعاد تسليمها.\n\n' +
      '/my — المحاضرات اللي عليك\n/login — رابط الدخول للموقع');
    return;
  }
  const full = [m.from.first_name, m.from.last_name].filter(Boolean).join(' ').trim() || ('عضو ' + chat);
  const name = uniqueName_(full);
  sheet_(SH.TEAM).appendRow([name, chat, m.from.username ? '@' + m.from.username : '', nowDate_(), ROLE.REVIEWER, true, false]);
  fresh_();
  log_(name, 'انضم للفريق', '');
  send_(chat, 'أهلًا ' + esc_(name) + ' 👋\nاتسجّلت في نظام المراجعة ✅\n' +
    'أي محاضرة عليك مراجعتها هتوصلك هنا كملف PDF بميعاد تسليمها، وهيوصلك تذكير قبل الميعاد بيوم.\n\n' +
    'الملاحظات: اعمل رد (Reply) على رسالة المحاضرة بالنص أو الملف أو الصوت، وفي الآخر اضغط «تم التسليم».\n\n' +
    '/login — رابط الدخول للموقع');
  managersToNotify_().forEach(x => send_(x.chat, '👤 عضو جديد انضم للفريق: <b>' + esc_(name) + '</b> (مراجع)'));
  if (payload === 'login') onLogin_(m);
}

function onAdmin_(m, text) {
  const props = PropertiesService.getScriptProperties();
  const code = props.getProperty('ADMIN_CODE');
  const given = text.replace(/^\/admin\b/, '').trim();
  const age = nowDate_().getTime() - Number(props.getProperty('ADMIN_CODE_AT') || 0);
  if (code && age > 30 * 60000) { props.deleteProperty('ADMIN_CODE'); send_(m.chat.id, 'الكود ده انتهى (بيشتغل ٣٠ دقيقة بس). خد كود جديد من «⚙️ إعداد النظام» في الشيت.'); return; }
  if (!code || given !== code) { send_(m.chat.id, 'الكود غلط. خده من «⚙️ إعداد النظام» في الشيت.'); return; }
  const chat = String(m.chat.id);
  const prev = adminChat_();
  props.setProperty('ADMIN_CHAT', chat);
  props.deleteProperty('ADMIN_CODE');

  const sh = sheet_(SH.TEAM);
  let team = readTeam_();
  if (!team.some(x => x.chat === chat)) {
    const full = [m.from.first_name, m.from.last_name].filter(Boolean).join(' ').trim() || 'المالك';
    sh.appendRow([uniqueName_(full), chat, m.from.username ? '@' + m.from.username : '', nowDate_(), ROLE.OWNER, true, true]);
    fresh_();
    team = readTeam_();
  }
  team.forEach(x => {
    if (x.chat === chat) { sh.getRange(x.row, T.ROLE).setValue(ROLE.OWNER); sh.getRange(x.row, T.ACTIVE).setValue(true); sh.getRange(x.row, T.NOTIFY).setValue(true); }
    else if (x.chat === prev || x.role === ROLE.OWNER) sh.getRange(x.row, T.ROLE).setValue(ROLE.MANAGER);
  });
  fresh_();
  log_(readTeam_().find(x => x.chat === chat).name, 'بقى مالك النظام', '');
  send_(chat, '✅ أنت مالك النظام دلوقتي.\nهيوصلك: انضمام أي عضو، وكل ملاحظة باسم صاحبها، وكل تسليم، وملخص الصبح في أيام الشغل، وتذكير بالليل قبل يوم الشغل.\n\n' +
    '/login — ادخل الموقع وظبّط الفريق والصلاحيات\n/status — الملخص في أي وقت');
}

function onLogin_(m) {
  const chat = String(m.chat.id);
  const member = readTeam_().find(r => r.chat === chat);
  if (!member) { send_(chat, 'ابعت /start الأول عشان تتسجّل.'); return; }
  if (!member.active) { send_(chat, 'حسابك متوقّف حاليًا — كلّم المسؤول.'); return; }
  const api = PropertiesService.getScriptProperties().getProperty('API_URL');
  if (!api) { send_(chat, 'الموقع لسه ما اتربطش. المسؤول يربطه من الشيت: 🔗 ربط الموقع.'); return; }
  const token = newToken_();
  CacheService.getScriptCache().put('LOGIN_' + token, chat, 600);
  const url = SITE_URL + '#login=' + token + '&api=' + encodeURIComponent(api);
  send_(chat, '🔐 <b>رابط دخولك للموقع</b>\nصالح ١٠ دقايق ولمرة واحدة، وهيفضل فاكرك ' + arNum_(SESSION_DAYS) + ' يوم.\nمتبعتهوش لحد.',
    { reply_markup: { inline_keyboard: [[{ text: '🔐 ادخل الموقع', url: url }]] } });
}

function onStatus_(m) {
  const member = readTeam_().find(r => r.chat === String(m.chat.id) && r.active);
  if (!member || !userOf_(member).perms.view_all) { send_(m.chat.id, 'الأمر ده للمدراء بس. جرّب /my'); return; }
  send_(m.chat.id, digestText_(member));
}

function onMy_(m) {
  const rev = readReviewers_().find(r => r.chat === String(m.chat.id));
  if (!rev) { send_(m.chat.id, 'ابعت /start الأول عشان تتسجّل.'); return; }
  const open = readLectures_().filter(l => l.reviewer === rev.name && isOpen_(l));
  if (!open.length) { send_(m.chat.id, 'مفيش محاضرات عليك دلوقتي 👌'); return; }
  send_(m.chat.id, '📚 <b>اللي عليك دلوقتي:</b>\n' + open.map(l =>
    '• «' + esc_(l.title) + '» (' + esc_(l.project) + ') — لحد ' + arDate_(ymdOf_(l.due)) + (l.status === ST.LATE ? ' ⚠️ متأخرة' : '')).join('\n'));
}

function onHelp_(m) {
  send_(m.chat.id, '• المحاضرة بتوصلك هنا ملف PDF بميعاد تسليمها.\n' +
    '• ملاحظاتك: اعمل رد (Reply) على رسالة المحاضرة — نص أو ملف أو صورة أو صوت، وابعت قد ما تحب.\n' +
    '• لما تخلص اضغط «تم التسليم».\n\n/my — اللي عليك دلوقتي\n/login — رابط الدخول للموقع');
}

function onNote_(m) {
  const chat = String(m.chat.id);
  const rev = readReviewers_().find(r => r.chat === chat);
  if (!rev) { send_(chat, 'ابعت /start الأول عشان تتسجّل.'); return; }
  const mine = readLectures_().filter(l => l.reviewer === rev.name && l.msgs.length);

  let lec = null;
  if (m.reply_to_message) lec = mine.find(l => l.msgs.indexOf(String(m.reply_to_message.message_id)) !== -1);
  if (!lec) {
    const open = mine.filter(isOpen_);
    if (open.length === 1) lec = open[0];
    else if (open.length > 1) {
      send_(chat, 'عليك أكتر من محاضرة — اعمل رد (Reply) على رسالة المحاضرة اللي الملاحظة دي تبعها، عشان تتسجّل في مكانها.');
      return;
    } else {
      send_(chat, 'مفيش محاضرة مفتوحة عليك دلوقتي. لو دي ملاحظة على محاضرة قديمة، اعمل رد (Reply) على رسالتها.');
      return;
    }
  }

  const firstNote = !lec.notes.trim() && !lec.noteFiles.trim();
  saveNote_(lec, m);
  forwardNote_(m, rev, lec);
  react_(chat, m.message_id);
  if (firstNote) {
    log_(rev.name, 'بدأ يبعت ملاحظات', lec.title, lec.project);
    send_(chat, 'اتسجّلت ملاحظاتك على «' + esc_(lec.title) + '» 👍\nكمّل براحتك، ولما تخلص اضغط «تم التسليم».',
      { reply_markup: isOpen_(lec) ? doneKeyboard_(lec.id) : undefined });
  }
}

function saveNote_(lec, m) {
  const sh = sheet_(SH.LECT);
  const text = (m.text || m.caption || '').trim();
  if (text) {
    const old = lec.notes.trim();
    sh.getRange(lec.row, L.NOTES).setValue((old ? old + '\n' : '') + '• ' + text);
  }
  const f = m.document || (m.photo && m.photo[m.photo.length - 1]) || m.voice || m.audio || m.video || m.video_note;
  if (!f) return;
  let link;
  try {
    link = saveTelegramFile_(f, lec.title, fileName_(m, f));
  } catch (e) {
    link = '(ملف ما اتحفظش على Drive: ' + e.message + ' — اتبعت للمدراء على البوت)';
  }
  const old = lec.noteFiles.trim();
  sh.getRange(lec.row, L.NOTE_FILES).setValue((old ? old + '\n' : '') + link);
}

/** كل ملاحظة توصل المدراء اللي مفعّلين الإشعارات فورًا، باسم المراجع واسم المحاضرة */
function forwardNote_(m, rev, lec) {
  const to = managersToNotify_().filter(x => x.chat !== String(m.chat.id) && sees_(x, lec.project));
  if (!to.length) return;
  const head = '📝 <b>' + esc_(rev.name) + '</b> — «' + esc_(lec.title) + '» · ' + esc_(lec.project);
  to.forEach(x => {
    if (m.text) { send_(x.chat, head + '\n\n' + esc_(m.text)); return; }
    if (m.video_note || m.sticker) {          // أنواع ما بتقبلش caption
      send_(x.chat, head);
      try { tg_('copyMessage', { chat_id: x.chat, from_chat_id: m.chat.id, message_id: m.message_id }); } catch (e) {}
      return;
    }
    const original = (m.caption || '').trim();
    let caption = head + (original ? '\n\n' + esc_(original) : '');
    if (caption.length > 1000) caption = caption.slice(0, 990) + '…';
    try {
      tg_('copyMessage', { chat_id: x.chat, from_chat_id: m.chat.id, message_id: m.message_id,
        caption: caption, parse_mode: 'HTML' });
    } catch (e) {
      send_(x.chat, head + '\n(بعت ملف ومقدرتش أحوّله — شوفه في الموقع أو الشيت)');
    }
  });
}

function saveTelegramFile_(f, title, name) {
  const info = tg_('getFile', { file_id: f.file_id });   // حد التحميل من تليجرام 20 ميجا
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  const blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot' + token + '/' + info.file_path).getBlob().setName(name);
  const file = folder_(folder_(DriveApp.getRootFolder(), NOTES_FOLDER), title.slice(0, 100)).createFile(blob);
  return file.getUrl();
}

function onCallback_(q) {
  const chat = String(q.message.chat.id);
  const mt = /^done:(\d+)$/.exec(q.data || '');
  if (!mt) { answer_(q.id); return; }
  const rev = readReviewers_().find(r => r.chat === chat);
  const lec = readLectures_().find(l => l.id === Number(mt[1]));
  if (!rev || !lec || lec.reviewer !== rev.name) { answer_(q.id, 'المحاضرة دي مش متسجّلة باسمك.'); return; }

  clearButtons_(chat, q.message.message_id);
  if (!isOpen_(lec)) { answer_(q.id, 'اتسلّمت قبل كده ✅'); return; }
  markDone_(lec, rev.name, '');
  answer_(q.id, 'تم التسليم ✅');
  send_(chat, 'جزاك الله خيرًا 🌿 اتسجّل تسليم مراجعة «' + esc_(lec.title) + '».');
}

function markDone_(lec, reviewerName, via) {
  const sh = sheet_(SH.LECT);
  sh.getRange(lec.row, L.STATUS).setValue(ST.DONE);
  sh.getRange(lec.row, L.DONE_AT).setValue(nowDate_());
  log_(reviewerName, 'سلّم المراجعة' + (via ? ' (' + via + ')' : ''), lec.title, lec.project);

  const due = ymdOf_(lec.due), today = todayYmd_();
  const late = due && today > due ? daysBetween_(due, today) : 0;
  const nText = (lec.notes.match(/^• /gm) || []).length;
  const nFiles = lec.noteFiles.trim() ? lec.noteFiles.trim().split('\n').length : 0;
  const text = '✅ <b>' + esc_(reviewerName) + '</b> سلّم مراجعة «' + esc_(lec.title) + '» · ' + esc_(lec.project) + (via ? ' <i>(' + esc_(via) + ')</i>' : '') + '\n' +
    (late ? '⚠️ متأخر ' + late + ' يوم' : 'في الميعاد') + ' — ' +
    (nText || nFiles ? nText + ' ملاحظة مكتوبة، ' + nFiles + ' ملف (كلها في الموقع والشيت)' : 'من غير ملاحظات');
  managersToNotify_().filter(m => sees_(m, lec.project)).forEach(m => send_(m.chat, text));
}

// ================================================================
// الملخص
// ================================================================
function digestText_(m) {
  const today = todayYmd_();
  const ls = readLectures_().filter(l => l.title && (!m || sees_(m, l.project)));
  const out = ['📊 <b>ملخص المراجعة — ' + arDate_(today) + '</b>'];
  let any = false;
  readProjects_().filter(p => !m || sees_(m, p.name)).forEach(p => {
    const pl = ls.filter(l => l.project === p.name);
    const by = st => pl.filter(l => l.status === st);
    const done = by(ST.DONE), sent = by(ST.SENT), late = by(ST.LATE), fixed = by(ST.FIXED);
    if (!done.length && !sent.length && !late.length && !fixed.length) return;
    any = true;
    out.push('', '━━━━━━━━━━', '📁 <b>' + esc_(p.name) + '</b>');
    if (late.length) out.push('⚠️ متأخرة (' + late.length + '):', ...late.map(l => '  • «' + esc_(l.title) + '» — ' + esc_(l.reviewer) + ' — من ' + daysBetween_(ymdOf_(l.due), today) + ' يوم'));
    if (sent.length) out.push('⏳ عند المراجع (' + sent.length + '):', ...sent.map(l => '  • «' + esc_(l.title) + '» — ' + esc_(l.reviewer) + ' — لحد ' + arDate_(ymdOf_(l.due))));
    if (done.length) out.push('✅ اتسلّمت ومستنية التعديل (' + done.length + '):', ...done.map(l => '  • «' + esc_(l.title) + '» — ' + esc_(l.reviewer)));
    if (fixed.length) out.push('📤 متعدّلة ومستنية الرفع: ' + fixed.length);
  });
  if (!any) out.push('', 'مفيش حاجة مفتوحة في أي مشروع 👌');
  return out.join('\n');
}

// ================================================================
// إنشاء الأوراق
// ================================================================
function ensureSheets_() {
  const book = book_();
  book.setSpreadsheetTimeZone(TZ);
  const rows = 1000;

  // الفريق (ولو فيه ورقة «المراجعون» من النسخة القديمة بتتحوّل)
  let team = book.getSheetByName(SH.TEAM);
  if (!team && book.getSheetByName('المراجعون')) {
    team = book.getSheetByName('المراجعون').setName(SH.TEAM);
    const n = Math.max(team.getLastRow() - 1, 0);
    team.getRange(1, 1, 1, T_HEADERS.length).setValues([T_HEADERS]);
    if (n) team.getRange(2, T.ROLE, n, 3).setValues(Array.from({ length: n }, () => [ROLE.REVIEWER, true, false]));
  }
  if (!team) {
    team = book.insertSheet(SH.TEAM);
    team.getRange(1, 1, 1, T_HEADERS.length).setValues([T_HEADERS]);
  }
  styleHeader_(team, T_HEADERS.length);
  team.getRange('B:B').setNumberFormat('@');
  team.getRange(2, T.JOINED, rows).setNumberFormat('yyyy/mm/dd');
  team.getRange(2, T.ROLE, rows).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList([ROLE.OWNER, ROLE.MANAGER, ROLE.REVIEWER], true).setAllowInvalid(false).build());
  // مربعات اختيار من غير ما تمسح القيم الموجودة (insertCheckboxes بيصفّرها)
  team.getRange(2, T.ACTIVE, rows, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  [170, 130, 150, 120, 90, 70, 110].forEach((w, i) => team.setColumnWidth(i + 1, w));

  let lec = book.getSheetByName(SH.LECT);
  if (!lec) {
    lec = book.insertSheet(SH.LECT, 0);
    lec.getRange(1, 1, 1, L_HEADERS.length).setValues([L_HEADERS]);
    styleHeader_(lec, L_HEADERS.length);
    lec.getRange(2, L.STATUS, rows).setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(ST).map(k => ST[k]), true).setAllowInvalid(false).build());
    lec.getRange(2, L.SENT, rows).setNumberFormat('yyyy/mm/dd');
    lec.getRange(2, L.DUE, rows).setNumberFormat('yyyy/mm/dd');
    lec.getRange(2, L.DONE_AT, rows).setNumberFormat('yyyy/mm/dd hh:mm');
    lec.getRange(2, L.REMINDED, rows).setNumberFormat('yyyy/mm/dd hh:mm');
    lec.getRange(2, L.MSGS, rows).setNumberFormat('@');
    lec.getRange(2, L.NOTES, rows, 2).setWrap(true);
    lec.getRange(2, 1, rows, L_HEADERS.length).setVerticalAlignment('top');
    [60, 260, 200, 140, 100, 110, 110, 360, 220, 120, 120, 80].forEach((w, i) => lec.setColumnWidth(i + 1, w));
    lec.setConditionalFormatRules(ST_COLORS.map(([st, color]) => SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(st).setBackground(color).setRanges([lec.getRange(2, L.STATUS, rows)]).build()));
    lec.hideColumns(L.MSGS);
  }
  lec.getRange(2, L.REVIEWER, rows).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(team.getRange('A2:A'), true).setAllowInvalid(true).build());

  let perm = book.getSheetByName(SH.PERM);
  if (!perm) {
    perm = book.insertSheet(SH.PERM);
    perm.getRange(1, 1, PERMS.length + 1, 4).setValues(
      [['المفتاح', 'الصلاحية', ROLE.MANAGER, ROLE.REVIEWER]].concat(PERMS.map(p => [p[0], p[1], p[2], p[3]])));
    styleHeader_(perm, 4);
    perm.getRange(2, 3, PERMS.length, 2).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    perm.setColumnWidth(2, 380); perm.setColumnWidth(3, 90); perm.setColumnWidth(4, 90);
    perm.hideColumns(1);
  }

  let log = book.getSheetByName(SH.LOG);
  if (!log) {
    log = book.insertSheet(SH.LOG);
    log.getRange(1, 1, 1, 4).setValues([['الوقت', 'الشخص', 'الحدث', 'المحاضرة']]);
    styleHeader_(log, 4);
    log.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
    [140, 150, 320, 260].forEach((w, i) => log.setColumnWidth(i + 1, w));
  }

  let set = book.getSheetByName(SH.SET);
  if (!set) {
    set = book.insertSheet(SH.SET);
    set.getRange(1, 1, 4, 3).setValues([
      ['الإعداد', 'القيمة', 'شرح'],
      ['أيام الشغل', 'السبت، الأربعاء', 'موعد التسليم = اليوم اللي قبل يوم الشغل الجاي (محاضرات السبت ← الثلاثاء، والأربعاء ← الجمعة).'],
      ['ساعة التنبيه', 9, 'الساعة (بتوقيت القاهرة، من 0 لـ 23) اللي يتبعت فيها تذكير «التسليم بكرة» للمراجعين وملخص الصبح في أيام الشغل.'],
      ['ساعة تنبيهي', 20, 'الساعة اللي يوصل فيها المالك تذكير «بكرة يوم شغل» باللي مستنيه.'],
    ]);
    styleHeader_(set, 3);
    set.setColumnWidth(1, 130); set.setColumnWidth(2, 160); set.setColumnWidth(3, 520);
    set.getRange('C2:C4').setWrap(true);
  }

  // المشاريع + الأعمدة الجديدة (من غير ما نمسح أي بيانات موجودة)
  ensureProjectsSheet_(book);
  const projSheet = book.getSheetByName(SH.PROJ);
  lec.getRange(1, L.PROJECT).setValue('المشروع');
  styleHeader_(lec, L_HEADERS.length);
  lec.setColumnWidth(L.PROJECT, 120);
  lec.getRange(2, L.PROJECT, rows).setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInRange(projSheet.getRange('A2:A'), true).setAllowInvalid(true).build());
  team.getRange(1, T.PROJECTS).setValue(T_HEADERS[T.PROJECTS - 1]);
  styleHeader_(team, T_HEADERS.length);
  team.setColumnWidth(T.PROJECTS, 200);
  log.getRange(1, 5).setValue('المشروع');
  styleHeader_(log, 5);
  log.setColumnWidth(5, 120);

  // شيل الورقة الفاضية الافتراضية
  const mine = [SH.LECT, SH.TEAM, SH.PROJ, SH.PERM, SH.LOG, SH.SET];
  book.getSheets().forEach(s => {
    if (mine.indexOf(s.getName()) === -1 && s.getLastRow() === 0 && book.getSheets().length > mine.length) book.deleteSheet(s);
  });
  book.setActiveSheet(lec);
}

function ensureProjectsSheet_(book) {
  if (book.getSheetByName(SH.PROJ)) return;
  const s = book.insertSheet(SH.PROJ);
  s.getRange(1, 1, DEFAULT_PROJECTS.length + 1, P_HEADERS.length).setValues(
    [P_HEADERS].concat(DEFAULT_PROJECTS.map(x => [x[0], x[1], x[2], '', true])));
  styleHeader_(s, P_HEADERS.length);
  s.getRange(2, 5, 100).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  [160, 110, 90, 220, 70].forEach((w, i) => s.setColumnWidth(i + 1, w));
}

function styleHeader_(s, n) {
  s.setRightToLeft(true);
  s.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#182D3B').setFontColor('#ffffff')
    .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
  s.setFrozenRows(1);
  s.setRowHeight(1, 36);
}

// ================================================================
// قراءة وكتابة البيانات
// ================================================================
function book_() {
  let b = null;
  try { b = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (b) return b;
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('شغّل «⚙️ إعداد النظام» من الشيت الأول');
  return SpreadsheetApp.openById(id);
}
// ---------- ذاكرة مؤقتة: كل ورقة بتتقري مرة واحدة بس في العملية الواحدة ----------
let MEMO_ = {};
function fresh_() { MEMO_ = {}; }
function readLectures_() { return MEMO_.lect || (MEMO_.lect = readLecturesRaw_()); }
function readProjects_() { return MEMO_.proj || (MEMO_.proj = readProjectsRaw_()); }
function readTeam_() { return MEMO_.team || (MEMO_.team = readTeamRaw_()); }
function readPerms_() { return MEMO_.perms || (MEMO_.perms = readPermsRaw_()); }
/** الإعدادات بتتحفظ ٥ دقايق في الكاش (المؤقّت بيسأل عليها كل دقيقة) */
function getSettings_() {
  if (MEMO_.settings) return MEMO_.settings;
  const cache = CacheService.getScriptCache();
  const hit = cache.get('SETTINGS_V1');
  if (hit) return (MEMO_.settings = JSON.parse(hit));
  MEMO_.settings = getSettingsRaw_();
  cache.put('SETTINGS_V1', JSON.stringify(MEMO_.settings), 300);
  return MEMO_.settings;
}

function sheet_(name) { return (MEMO_.book || (MEMO_.book = book_())).getSheetByName(name); }

function readLecturesRaw_() {
  const s = sheet_(SH.LECT);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  const vals = s.getRange(2, 1, n, L_HEADERS.length).getValues();
  const links = s.getRange(2, L.FILE, n, 1).getRichTextValues();
  return vals.map((r, i) => {
    const rt = links[i][0];
    const link = rt && (rt.getLinkUrl() || (rt.getRuns && rt.getRuns().map(x => x.getLinkUrl()).filter(Boolean)[0]));
    return {
      row: i + 2, id: Number(r[L.ID - 1]) || 0, title: String(r[L.TITLE - 1]).trim(),
      file: link || String(r[L.FILE - 1]).trim(), reviewer: String(r[L.REVIEWER - 1]).trim(),
      sent: r[L.SENT - 1], due: r[L.DUE - 1], status: String(r[L.STATUS - 1]).trim(),
      notes: String(r[L.NOTES - 1]), noteFiles: String(r[L.NOTE_FILES - 1]),
      doneAt: r[L.DONE_AT - 1], reminded: r[L.REMINDED - 1],
      msgs: String(r[L.MSGS - 1]).split(',').map(x => x.trim()).filter(Boolean),
      project: String(r[L.PROJECT - 1] || '').trim() || defaultProject_(),
    };
  });
}

function readProjectsRaw_() {
  let s = sheet_(SH.PROJ);
  if (!s) { ensureProjectsSheet_(book_()); s = sheet_(SH.PROJ); }
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  return s.getRange(2, 1, n, P_HEADERS.length).getValues().map((r, i) => {
    const d = parseDays_(String(r[3] || ''));
    return { row: i + 2, name: String(r[0]).trim(), item: String(r[1]).trim() || 'محاضرة',
      color: /^#[0-9a-f]{6}$/i.test(String(r[2]).trim()) ? String(r[2]).trim() : '#2E8B85',
      workDays: d.length ? d : null, active: r[4] === '' || r[4] === true || /^(true|نعم)$/i.test(String(r[4])) };
  }).filter(p => p.name);
}
function defaultProject_() { const ps = readProjects_(); return ps.length ? ps[0].name : 'عام'; }
function projectOf_(name) { return readProjects_().find(p => p.name === name) || { name: name, item: 'محاضرة', color: '#2E8B85', workDays: null, active: true }; }
function projWorkDays_(name) { const p = projectOf_(name); return p.workDays || getSettings_().workDays; }
/** العضو يشوف المشروع؟ (قايمة مشاريعه فاضية = كل المشاريع) */
function sees_(m, project) { return !m.projects || !m.projects.length || m.projects.indexOf(project) !== -1; }

function readTeamRaw_() {
  const s = sheet_(SH.TEAM);
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  const bool = v => v === true || /^(true|نعم|✓)$/i.test(String(v).trim());
  return s.getRange(2, 1, n, T_HEADERS.length).getValues().map((r, i) => ({
    row: i + 2, name: String(r[0]).trim(), chat: String(r[1]).trim(), username: String(r[2]).trim(),
    joined: r[3], role: String(r[4]).trim() || ROLE.REVIEWER,
    active: r[5] === '' || bool(r[5]), notify: bool(r[6]),
    projects: String(r[7] || '').split(/[،,]/).map(x => x.trim()).filter(Boolean),
  })).filter(m => m.name && m.chat);
}

/** كل الأعضاء النشطين (أي حد نشط ينفع يتوزّع عليه مراجعة) */
function readReviewers_() { return readTeam_().filter(m => m.active); }

/** المالك + المدراء النشطين اللي مفعّلين «يستلم الإشعارات» */
function managersToNotify_() {
  const owner = adminChat_();
  return readTeam_().filter(m => m.active && m.notify && (m.chat === owner || m.role === ROLE.MANAGER || m.role === ROLE.OWNER));
}

function readPermsRaw_() {
  const out = { manager: {}, reviewer: {} };
  PERMS.forEach(p => { out.manager[p[0]] = p[2]; out.reviewer[p[0]] = p[3]; });
  const s = sheet_(SH.PERM);
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 4).getValues().forEach(r => {
    const k = String(r[0]).trim();
    if (!(k in out.manager)) return;
    out.manager[k] = r[2] === true; out.reviewer[k] = r[3] === true;
  });
  return out;
}

function log_(who, what, title, project) {
  try { sheet_(SH.LOG).appendRow([nowDate_(), who, what, title || '', project || '']); } catch (e) { console.error('log: ' + e.message); }
}

function readLog_(limit) {
  const s = sheet_(SH.LOG);
  if (!s) return [];
  const n = s.getLastRow() - 1;
  if (n < 1) return [];
  const from = Math.max(2, n + 2 - limit);
  return s.getRange(from, 1, n + 2 - from, 5).getValues().reverse().map(r => ({
    at: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm') : String(r[0]),
    who: String(r[1]), what: String(r[2]), lecture: String(r[3]), project: String(r[4] || ''),
  }));
}

function getSettingsRaw_() {
  const out = { workDays: [6, 3], hour: 9, adminHour: 20 };
  const s = sheet_(SH.SET);
  if (!s || s.getLastRow() < 2) return out;
  s.getRange(2, 1, s.getLastRow() - 1, 2).getValues().forEach(r => {
    const k = String(r[0]).trim(), v = r[1];
    if (k === 'أيام الشغل') { const d = parseDays_(String(v)); if (d.length) out.workDays = d; }
    if (k === 'ساعة التنبيه') { const h = parseInt(v, 10); if (h >= 0 && h <= 23) out.hour = h; }
    if (k === 'ساعة تنبيهي') { const h = parseInt(v, 10); if (h >= 0 && h <= 23) out.adminHour = h; }
  });
  return out;
}

function setSetting_(key, value) {
  const s = sheet_(SH.SET);
  const n = s.getLastRow();
  const keys = n >= 2 ? s.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]).trim()) : [];
  const i = keys.indexOf(key);
  if (i === -1) s.appendRow([key, value, '']);
  else s.getRange(i + 2, 2).setValue(value);
  delete MEMO_.settings;
  CacheService.getScriptCache().remove('SETTINGS_V1');
}

function parseDays_(text) {
  const norm = s => s.replace(/[أإآ]/g, 'ا').replace(/ى/g, 'ي').replace(/^ال/, '').trim();
  const names = DAY_NAMES.map(norm);
  return text.split(/[،,\s\-و]+/).map(norm).filter(Boolean)
    .map(w => names.indexOf(w)).filter(i => i !== -1);
}

function isOpen_(l) { return l.status === ST.SENT || l.status === ST.LATE; }
function adminChat_() { return PropertiesService.getScriptProperties().getProperty('ADMIN_CHAT') || ''; }

function uniqueName_(name) {
  const taken = readTeam_().map(r => r.name);
  let n = name, i = 2;
  while (taken.indexOf(n) !== -1) n = name + ' ' + i++;
  return n;
}

function driveId_(s) {
  s = String(s || '').trim();
  const m = /\/d\/([\w-]{20,})/.exec(s) || /[?&]id=([\w-]{20,})/.exec(s);
  if (m) return m[1];
  return /^[\w-]{20,}$/.test(s) ? s : '';
}

function folder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function fileName_(m, f) {
  if (m.document && m.document.file_name) return m.document.file_name;
  const stamp = Utilities.formatDate(nowDate_(), TZ, 'yyyy-MM-dd_HH-mm');
  if (m.photo) return 'صورة_' + stamp + '.jpg';
  if (m.voice) return 'صوت_' + stamp + '.ogg';
  if (m.audio) return (m.audio.file_name || 'صوت_' + stamp + '.mp3');
  return 'فيديو_' + stamp + '.mp4';
}

// ================================================================
// التواريخ (كلها بصيغة yyyy-MM-dd بتوقيت القاهرة)
// ================================================================
function nowDate_() { return new Date(); }
function todayYmd_() { return ymdOf_(nowDate_()); }

function ymdOf_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  const m = /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/.exec(String(v || '').trim());
  return m ? m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2) : '';
}
function utc_(ymd) { const p = ymd.split('-').map(Number); return new Date(Date.UTC(p[0], p[1] - 1, p[2], 10)); }
function dateOf_(ymd) { return utc_(ymd); }   // الساعة 12 ظهرًا بالقاهرة: التاريخ ثابت مهما كان التوقيت
function addDays_(ymd, n) { const d = utc_(ymd); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); }
function dow_(ymd) { return utc_(ymd).getUTCDay(); }
function daysBetween_(a, b) { return Math.round((utc_(b) - utc_(a)) / 86400000); }

/** موعد التسليم = اليوم اللي قبل أول يوم شغل جاي (ويكون بعد يوم الإرسال) */
function computeDue_(sentYmd, workDays) {
  for (let i = 2; i <= 8; i++) {
    const d = addDays_(sentYmd, i);
    if (workDays.indexOf(dow_(d)) !== -1) return addDays_(d, -1);
  }
  return addDays_(sentYmd, 3);
}

/** أقرب يوم شغل من النهارده (النهارده نفسه لو يوم شغل) */
function nextWorkDay_(today, workDays) {
  for (let i = 0; i < 7; i++) { const d = addDays_(today, i); if (workDays.indexOf(dow_(d)) !== -1) return d; }
  return '';
}

function arDate_(ymd) {
  if (!ymd) return '—';
  const d = utc_(ymd);
  return DAY_NAMES[d.getUTCDay()] + ' ' + arNum_(d.getUTCDate()) + ' ' + MONTHS[d.getUTCMonth()];
}
function arNum_(n) { return String(n).replace(/\d/g, x => '٠١٢٣٤٥٦٧٨٩'[x]); }

// ================================================================
// تليجرام
// ================================================================
function tg_(method, params) {
  const token = PropertiesService.getScriptProperties().getProperty('TOKEN');
  const opts = { method: 'post', muteHttpExceptions: true };
  if (params) {
    const hasBlob = Object.keys(params).some(k => params[k] && typeof params[k].getBytes === 'function');
    if (hasBlob) {                         // رفع ملف: multipart
      const p = {};
      Object.keys(params).forEach(k => {
        const v = params[k];
        if (v === undefined || v === null) return;
        p[k] = (v && typeof v.getBytes === 'function') ? v : (typeof v === 'object' ? JSON.stringify(v) : String(v));
      });
      opts.payload = p;
    } else {
      opts.contentType = 'application/json';
      opts.payload = JSON.stringify(params);
    }
  }
  const res = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, opts);
  const j = JSON.parse(res.getContentText());
  if (!j.ok) throw new Error(j.description || ('HTTP ' + res.getResponseCode()));
  return j.result;
}

function send_(chat, text, extra) {
  extra = extra || {};
  const p = { chat_id: String(chat), text: text, parse_mode: 'HTML', link_preview_options: { is_disabled: true } };
  if (extra.reply_markup) p.reply_markup = extra.reply_markup;
  if (extra.reply_to) p.reply_parameters = { message_id: Number(extra.reply_to), allow_sending_without_reply: true };
  try { return tg_('sendMessage', p); } catch (e) { console.error('send to ' + chat + ': ' + e.message); return null; }
}

function answer_(id, text) { try { tg_('answerCallbackQuery', text ? { callback_query_id: id, text: text } : { callback_query_id: id }); } catch (e) {} }
function clearButtons_(chat, msgId) { try { tg_('editMessageReplyMarkup', { chat_id: chat, message_id: msgId, reply_markup: { inline_keyboard: [] } }); } catch (e) {} }
function react_(chat, msgId) {
  try { tg_('setMessageReaction', { chat_id: chat, message_id: msgId, reaction: [{ type: 'emoji', emoji: '👍' }] }); } catch (e) {}
}
function doneKeyboard_(id) { return { inline_keyboard: [[{ text: '✅ تم التسليم', callback_data: 'done:' + id }]] }; }
function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
