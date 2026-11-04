/**
 * 洗車予約 バックエンド(Google Apps Script / スプレッドシート連携)
 * - 予約・キャンセル・キャンセル待ち・メール送信・管理API
 * - 画面(index.html / admin.html)から fetch(POST) で呼ばれます
 */
const TZ = 'Asia/Tokyo';
const SH = { RES: '予約', WAIT: 'キャンセル待ち', MODEL: '車種', HOL: '休業日', SET: '設定' };
const COLS = {
  RES: ['id', '予約番号', '日付', '開始', '終了', '分', '氏名', '電話', 'メール', 'ナンバー', 'ナンバー表示', '車種', '大型', '状態', '作成', '保留期限', '待ちID', 'トークン'],
  WAIT: ['id', '作成', '氏名', '電話', 'メール', 'ナンバー', 'ナンバー表示', '車種', '大型', '対象', '状態', 'スキップ', '最終通知'],
  MODEL: ['車種', '大型'],
  HOL: ['日付', 'メモ'],
  SET: ['項目', '値', '説明']
};
const DEF_SETTINGS = [
  ['店舗名', '洗車予約', '画面上部とメールに出る名前'],
  ['同時洗車台数', '2', '同じ時間に洗車できる台数(1〜10)'],
  ['開始', '09:30', '最初の枠の開始時刻'],
  ['昼休憩開始', '12:30', ''],
  ['昼休憩終了', '13:30', ''],
  ['最終開始(普通)', '16:45', '普通車の最後の予約開始時刻'],
  ['最終開始(大型)', '16:30', '大型車の最後の予約開始時刻'],
  ['洗車分(普通)', '45', ''],
  ['洗車分(大型)', '60', ''],
  ['刻み分', '15', '開始時刻の間隔(分)'],
  ['先行月数', '2', '何か月先まで予約できるか'],
  ['解禁時刻', '00:00', '毎日この時刻に、2か月先の日が1日ずつ開きます'],
  ['最短日数', '1', '1=明日から予約可 / 0=当日も可'],
  ['定休曜日', '', '日=0 月=1 火=2 水=3 木=4 金=5 土=6 (例: 3 または 2,3)'],
  ['保留時間(分)', '120', 'キャンセル待ちの案内メール後、その枠を確保しておく時間'],
  ['サイトURL', '', '予約サイトのURL(キャンセル待ちメールのリンクに使います)'],
  ['問い合わせ先', '', 'メール末尾に入れる店名・電話番号']
];
const DEF_MODELS = [['UX', 'FALSE'], ['NX', 'FALSE'], ['RX', 'FALSE'], ['IS', 'FALSE'], ['ES', 'FALSE'], ['LBX', 'FALSE'], ['LM', 'TRUE'], ['LX', 'TRUE'], ['その他(普通車)', 'FALSE'], ['その他(大型車)', 'TRUE']];

/* ========== 初期設定(最初に1回だけ手動実行) ========== */
function setup() { Object.keys(SH).forEach(k => sh_(k)); }
function setupTrigger() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'processAll').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processAll').timeBased().everyMinutes(10).create();
}

/* ========== Web API ========== */
function doGet() { return ContentService.createTextOutput('OK'); }
function doPost(e) {
  let out;
  try {
    const p = JSON.parse(e.postData.contents);
    out = { ok: true, data: route_(p) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
function route_(p) {
  const pub = { config: config_, month: month_, slots: slots_, holdInfo: holdInfo_, lookup: lookup_ };
  const wr = { reserve: reserve_, cancel: cancel_, waitAdd: waitAdd_, waitStop: waitStop_, holdConfirm: holdConfirm_ };
  const admR = { adminRes: adminRes_, adminWaits: adminWaits_, adminGet: adminGet_ };
  const admW = { adminCancel: adminCancel_, adminWaitStop: adminWaitStop_, adminSave: adminSave_ };
  if (pub[p.action]) return pub[p.action](p);
  if (wr[p.action]) return withLock_(() => wr[p.action](p));
  if (admR[p.action]) { auth_(p); return admR[p.action](p); }
  if (admW[p.action]) { auth_(p); return withLock_(() => admW[p.action](p)); }
  throw new Error('不明な操作です');
}
function withLock_(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(25000);
  try { return fn(); } finally { l.releaseLock(); }
}
function auth_(p) {
  const pw = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!pw) throw new Error('管理者パスワード(ADMIN_PASSWORD)が未設定です');
  if (String(p.pw || '') !== pw) throw new Error('パスワードが違います');
}

/* ========== シート入出力 ========== */
function sh_(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SH[key]);
  if (!s) {
    s = ss.insertSheet(SH[key]);
    const cols = COLS[key];
    s.getRange(1, 1, s.getMaxRows(), cols.length).setNumberFormat('@');
    s.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (key === 'SET') s.getRange(2, 1, DEF_SETTINGS.length, 3).setValues(DEF_SETTINGS);
    if (key === 'MODEL') s.getRange(2, 1, DEF_MODELS.length, 2).setValues(DEF_MODELS);
  }
  return s;
}
function table_(key) {
  const s = sh_(key), n = COLS[key].length, last = s.getLastRow();
  if (last < 2) return [];
  const v = s.getRange(1, 1, last, n).getValues(), head = v[0], out = [];
  for (let i = 1; i < v.length; i++) {
    if (v[i].every(c => c === '' || c === null)) continue;
    const o = { _row: i + 1 };
    head.forEach((h, j) => { o[h] = (v[i][j] === null || v[i][j] === undefined) ? '' : String(v[i][j]); });
    out.push(o);
  }
  return out;
}
function appendRow_(key, o) {
  const s = sh_(key);
  s.appendRow(COLS[key].map(c => o[c] === undefined ? '' : o[c]));
  o._row = s.getLastRow();
}
function saveRow_(key, o) {
  sh_(key).getRange(o._row, 1, 1, COLS[key].length).setValues([COLS[key].map(c => o[c] === undefined ? '' : o[c])]);
}
function rewrite_(key, rows) {
  const s = sh_(key), n = COLS[key].length, last = s.getLastRow();
  if (last > 1) s.getRange(2, 1, last - 1, n).clearContent();
  if (rows.length) s.getRange(2, 1, rows.length, n).setValues(rows);
}

/* ========== 共通ヘルパー ========== */
const pad2_ = n => ('0' + n).slice(-2);
const nowStr_ = () => Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss");
const fmtD_ = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
const parseD_ = s => { const a = s.split('-'); return new Date(+a[0], +a[1] - 1, +a[2]); };
const toMin_ = s => { const m = String(s).match(/(\d+):(\d+)/); return m ? (+m[1]) * 60 + (+m[2]) : 0; };
const fmtMin_ = n => pad2_(Math.floor(n / 60)) + ':' + pad2_(n % 60);
const isTrue_ = v => String(v).toUpperCase() === 'TRUE';
const WD_ = ['日', '月', '火', '水', '木', '金', '土'];
const dateJa_ = d => { const a = d.split('-'); return (+a[1]) + '月' + (+a[2]) + '日(' + WD_[parseD_(d).getDay()] + ')'; };
const ymJa_ = ym => { const a = ym.split('-'); return a[0] + '年' + (+a[1]) + '月'; };
const normPlate_ = s => String(s || '').normalize('NFKC').replace(/[\s\-‐‑‒–—−・･]/g, '').toUpperCase();
const validMail_ = s => /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || ''));

function cfg_() {
  const raw = {};
  DEF_SETTINGS.forEach(r => raw[r[0]] = r[1]);
  table_('SET').forEach(r => { if (r['値'] !== '') raw[r['項目']] = r['値']; else if (r['項目'] === '定休曜日' || r['項目'] === 'サイトURL' || r['項目'] === '問い合わせ先') raw[r['項目']] = ''; });
  return {
    name: raw['店舗名'],
    cap: Math.min(10, Math.max(1, Math.floor(+raw['同時洗車台数']) || 1)),
    open: toMin_(raw['開始']), lunchS: toMin_(raw['昼休憩開始']), lunchE: toMin_(raw['昼休憩終了']),
    lastStd: toMin_(raw['最終開始(普通)']), lastLarge: toMin_(raw['最終開始(大型)']),
    durStd: +raw['洗車分(普通)'], durLarge: +raw['洗車分(大型)'], step: Math.max(5, +raw['刻み分'] || 15),
    months: +raw['先行月数'], unlock: toMin_(raw['解禁時刻']), minDays: +raw['最短日数'],
    closedWd: String(raw['定休曜日']).split(/[,、\s]+/).filter(x => x !== '').map(Number),
    holdMin: +raw['保留時間(分)'], siteUrl: String(raw['サイトURL'] || ''), contact: String(raw['問い合わせ先'] || '')
  };
}
function ctx_(withWaits) {
  return { c: cfg_(), hols: table_('HOL').map(r => r['日付']), res: table_('RES'), waits: withWaits ? table_('WAIT') : [], models: table_('MODEL') };
}
/** 予約できる日の範囲。毎日「解禁時刻」に2か月先の日が1日ずつ開く */
function window_(c) {
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const min = new Date(today); min.setDate(min.getDate() + c.minDays);
  const eff = new Date(today);
  if (now.getHours() * 60 + now.getMinutes() < c.unlock) eff.setDate(eff.getDate() - 1);
  const t = eff.getMonth() + c.months;
  let max = new Date(eff.getFullYear(), t, eff.getDate());
  if (max.getMonth() !== ((t % 12) + 12) % 12) max = new Date(eff.getFullYear(), t + 1, 0); // 月末の補正
  return { min: fmtD_(min), max: fmtD_(max), today: fmtD_(today) };
}
function closed_(x, d) { return x.c.closedWd.includes(parseD_(d).getDay()) || x.hols.includes(d); }
function okDate_(x, d) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(d))) return false;
  const w = window_(x.c);
  return d >= w.min && d <= w.max && !closed_(x, d);
}
function busy_(x, date, exceptToken) {
  const now = nowStr_();
  return x.res.filter(r => r['日付'] === date && (r['状態'] === '予約' ||
    (r['状態'] === '保留' && r['保留期限'] > now && r['トークン'] !== exceptToken)))
    .map(r => [toMin_(r['開始']), toMin_(r['終了'])]);
}
/** 開始できる時刻(分)の一覧。1台ずつなので、他の予約・昼休憩と重ならない開始時刻のみ */
function starts_(x, date, large, busy) {
  const c = x.c, dur = large ? c.durLarge : c.durStd, last = large ? c.lastLarge : c.lastStd;
  const w = window_(c);
  const now = new Date();
  const minStart = date === w.today ? now.getHours() * 60 + now.getMinutes() + 1 : 0;
  const out = [];
  for (let s = c.open; s <= last; s += c.step) {
    const e = s + dur;
    if (s < minStart) continue;
    if (s < c.lunchE && e > c.lunchS) continue;
    if (maxDepth_(busy, s, e) >= c.cap) continue; // 同時洗車台数に達している時間帯は不可
    out.push(s);
  }
  return out;
}
/** [s,e) の間に同時に入っている予約の最大数 */
function maxDepth_(busy, s, e) {
  const ov = busy.filter(b => s < b[1] && e > b[0]);
  if (!ov.length) return 0;
  const pts = [s].concat(ov.map(b => b[0]).filter(t => t > s && t < e));
  let m = 0;
  pts.forEach(t => { m = Math.max(m, ov.filter(b => b[0] <= t && t < b[1]).length); });
  return m;
}
/** その日にあと何台入れられるか(カレンダーの「残りわずか」判定用) */
function remaining_(x, date, large, busy) {
  const dur = large ? x.c.durLarge : x.c.durStd, b = busy.slice();
  let n = 0;
  while (n < 80) {
    const st = starts_(x, date, large, b);
    if (!st.length) break;
    b.push([st[0], st[0] + dur]); n++;
  }
  return n;
}
function monthCount_(x, plate, ym, exceptToken) {
  const now = nowStr_();
  return x.res.filter(r => r['ナンバー'] === plate && String(r['日付']).slice(0, 7) === ym &&
    (r['状態'] === '予約' || (r['状態'] === '保留' && r['保留期限'] > now && r['トークン'] !== exceptToken))).length;
}
function newNo_(x) {
  let n;
  do { n = String(Math.floor(100000 + Math.random() * 900000)); } while (x.res.some(r => r['予約番号'] === n));
  return n;
}
function need_(p, keys) { keys.forEach(k => { if (!String(p[k] === undefined ? '' : p[k]).trim()) throw new Error('入力されていない項目があります'); }); }
function modelOf_(x, name) {
  const m = x.models.find(m => m['車種'] === name);
  if (!m) throw new Error('車種を選択してください');
  return m;
}
function mail_(c, to, subject, body) {
  if (!to) return;
  const foot = '\n\n――――――――――\n※このメールは送信専用です。ご返信いただいても届きません。\n' + (c.contact || c.name);
  try { MailApp.sendEmail({ to: to, subject: '【' + c.name + '】' + subject, body: body + foot, name: c.name }); }
  catch (e) { console.error('mail failed: ' + e); }
}
function monthLimitMsg_(ym) { return ymJa_(ym) + 'のご予約はすでにあります(1台につき月1回までです)。別の月の日付をお選びください。'; }

/* ========== お客様向けAPI ========== */
function config_() {
  const x = ctx_(false), w = window_(x.c);
  return {
    name: x.c.name, min: w.min, max: w.max, today: w.today,
    models: x.models.map(m => ({ name: m['車種'], large: isTrue_(m['大型']) })),
    durStd: x.c.durStd, durLarge: x.c.durLarge, closedWd: x.c.closedWd, contact: x.c.contact,
    lunch: fmtMin_(x.c.lunchS) + '〜' + fmtMin_(x.c.lunchE)
  };
}
function month_(p) {
  const x = ctx_(false), large = !!p.large, dur = large ? x.c.durLarge : x.c.durStd;
  const a = String(p.ym).split('-').map(Number), last = new Date(a[0], a[1], 0).getDate(), w = window_(x.c), days = {};
  for (let d = 1; d <= last; d++) {
    const ds = a[0] + '-' + pad2_(a[1]) + '-' + pad2_(d);
    if (ds < w.min || ds > w.max) days[ds] = { s: 'out' };
    else if (closed_(x, ds)) days[ds] = { s: 'closed' };
    else {
      const n = remaining_(x, ds, large, busy_(x, ds, ''));
      days[ds] = { s: n === 0 ? 'full' : (n <= 2 ? 'few' : 'free'), n: n };
    }
  }
  return { days: days };
}
function slots_(p) {
  const x = ctx_(false), large = !!p.large, dur = large ? x.c.durLarge : x.c.durStd;
  if (!okDate_(x, p.date)) return { starts: [] };
  return { starts: starts_(x, p.date, large, busy_(x, p.date, '')).map(s => ({ s: fmtMin_(s), e: fmtMin_(s + dur) })) };
}
function reserve_(p) {
  need_(p, ['date', 'start', 'name', 'phone', 'email', 'plate', 'model']);
  if (!validMail_(p.email)) throw new Error('メールアドレスを確認してください');
  const x = ctx_(true), large = isTrue_(modelOf_(x, p.model)['大型']);
  const plate = normPlate_(p.plate);
  if (plate.length < 3) throw new Error('車両ナンバーを確認してください');
  if (!okDate_(x, p.date)) throw new Error('この日は予約できません');
  const ym = p.date.slice(0, 7);
  if (monthCount_(x, plate, ym, '') > 0) throw new Error(monthLimitMsg_(ym));
  const s = toMin_(p.start);
  if (!starts_(x, p.date, large, busy_(x, p.date, '')).includes(s)) throw new Error('その時間は埋まってしまいました。別の時間をお選びください。');
  const dur = large ? x.c.durLarge : x.c.durStd;
  const o = {
    id: Utilities.getUuid().slice(0, 8), '予約番号': newNo_(x), '日付': p.date, '開始': fmtMin_(s), '終了': fmtMin_(s + dur), '分': dur,
    '氏名': String(p.name).trim(), '電話': String(p.phone).trim(), 'メール': String(p.email).trim(), 'ナンバー': plate,
    'ナンバー表示': String(p.plate).trim(), '車種': p.model, '大型': large ? 'TRUE' : 'FALSE', '状態': '予約', '作成': nowStr_()
  };
  appendRow_('RES', o); x.res.push(o);
  sendConfirm_(x.c, o);
  return { no: o['予約番号'], date: o['日付'], start: o['開始'], end: o['終了'], model: o['車種'] };
}
function sendConfirm_(c, o) {
  mail_(c, o['メール'], '洗車のご予約を承りました',
    o['氏名'] + ' 様\n\n洗車のご予約を承りました。\n\n予約番号: ' + o['予約番号'] + '\n日時: ' + dateJa_(o['日付']) + ' ' + o['開始'] + '〜' + o['終了'] +
    '\nお車: ' + o['車種'] + '(' + o['ナンバー表示'] + ')\n\nキャンセルは予約サイトの「予約の確認・キャンセル」からお願いいたします。');
}
function lookup_(p) {
  const x = ctx_(true), plate = normPlate_(p.plate), mail = String(p.email || '').trim().toLowerCase(), w = window_(x.c);
  if (!plate || !mail) throw new Error('メールアドレスと車両ナンバーを入力してください');
  const res = x.res.filter(r => r['状態'] === '予約' && r['日付'] >= w.today && r['ナンバー'] === plate && r['メール'].toLowerCase() === mail)
    .sort((a, b) => (a['日付'] + a['開始']) < (b['日付'] + b['開始']) ? -1 : 1)
    .map(r => ({ id: r.id, no: r['予約番号'], date: r['日付'], start: r['開始'], end: r['終了'], model: r['車種'] }));
  const waits = x.waits.filter(r => (r['状態'] === '待機' || r['状態'] === '通知中') && r['ナンバー'] === plate && r['メール'].toLowerCase() === mail)
    .map(r => ({ id: r.id, state: r['状態'], summary: waitSummary_(x, r) }));
  return { res: res, waits: waits };
}
function waitSummary_(x, w) {
  const t = JSON.parse(w['対象']), months = (t.ym || []).map(ymJa_).join('・');
  if (t.mode === 'dates') return (t.dates || []).slice().sort().map(dateJa_).join('、');
  if (t.mode === 'weekdays') return months + ' の ' + (t.wd || []).map(i => WD_[i]).join('・') + '曜日';
  if (t.mode === 'weekend') return months + ' の土日';
  return months + ' 全体';
}
function cancel_(p) {
  const x = ctx_(true), plate = normPlate_(p.plate), mail = String(p.email || '').trim().toLowerCase();
  const r = x.res.find(r => r.id === p.id && r['状態'] === '予約' && r['ナンバー'] === plate && r['メール'].toLowerCase() === mail);
  if (!r) throw new Error('予約が見つかりません');
  cancelRes_(x, r);
  return { ok: 1 };
}
function cancelRes_(x, r) {
  r['状態'] = 'キャンセル'; saveRow_('RES', r);
  mail_(x.c, r['メール'], '洗車のご予約をキャンセルしました',
    r['氏名'] + ' 様\n\n以下のご予約をキャンセルいたしました。\n\n予約番号: ' + r['予約番号'] + '\n日時: ' + dateJa_(r['日付']) + ' ' + r['開始'] + '〜' + r['終了']);
  processDate_(x, r['日付']);
}

/* ========== キャンセル待ち ========== */
function targetDates_(x, t) {
  const only = t.only ? d => t.only.includes(d) : () => true; // 登録時点で満席だった日だけが対象
  if (t.mode === 'dates') return (t.dates || []).filter(d => okDate_(x, d) && only(d)).sort();
  const w = window_(x.c), ymset = t.ym || [], out = [];
  for (let d = parseD_(w.min); fmtD_(d) <= w.max; d.setDate(d.getDate() + 1)) {
    const ds = fmtD_(d), wd = d.getDay();
    if (!ymset.includes(ds.slice(0, 7))) continue;
    if (t.mode === 'weekdays' && !(t.wd || []).includes(wd)) continue;
    if (t.mode === 'weekend' && wd !== 0 && wd !== 6) continue;
    if (!okDate_(x, ds) || !only(ds)) continue;
    out.push(ds);
  }
  return out;
}
function waitAdd_(p) {
  need_(p, ['name', 'phone', 'email', 'plate', 'model', 'mode']);
  if (!validMail_(p.email)) throw new Error('メールアドレスを確認してください');
  if (['dates', 'weekdays', 'weekend', 'month'].indexOf(p.mode) < 0) throw new Error('対象の指定が正しくありません');
  const x = ctx_(true), large = isTrue_(modelOf_(x, p.model)['大型']), plate = normPlate_(p.plate);
  if (plate.length < 3) throw new Error('車両ナンバーを確認してください');
  const t = { mode: p.mode, ym: (p.ym || []).map(String), wd: (p.wd || []).map(Number), dates: (p.dates || []).map(String) };
  if (p.mode === 'dates' && !t.dates.length) throw new Error('日付を選んでください');
  if (p.mode !== 'dates' && !t.ym.length) throw new Error('月を選んでください');
  if (p.mode === 'weekdays' && !t.wd.length) throw new Error('曜日を選んでください');
  const days = targetDates_(x, t);
  if (!days.length) throw new Error('対象になる日がありません(休業日・予約できる期間外の可能性があります)');
  // キャンセル待ちは「いま満席の日」だけが対象(空きのある日は通常の予約でお取りいただく)
  t.only = days.filter(d => starts_(x, d, large, busy_(x, d, '')).length === 0);
  if (!t.only.length) throw new Error('選ばれた日にはすでに空きがあります。「洗車を予約する」から直接お取りください。');
  if (x.waits.some(w => (w['状態'] === '待機' || w['状態'] === '通知中') && w['ナンバー'] === plate && w['メール'].toLowerCase() === String(p.email).trim().toLowerCase() && w['対象'] === JSON.stringify(t)))
    throw new Error('同じ内容のキャンセル待ちがすでに登録されています');
  const o = {
    id: Utilities.getUuid().slice(0, 8), '作成': nowStr_(), '氏名': String(p.name).trim(), '電話': String(p.phone).trim(), 'メール': String(p.email).trim(),
    'ナンバー': plate, 'ナンバー表示': String(p.plate).trim(), '車種': p.model, '大型': large ? 'TRUE' : 'FALSE', '対象': JSON.stringify(t), '状態': '待機'
  };
  appendRow_('WAIT', o); x.waits.push(o);
  mail_(x.c, o['メール'], 'キャンセル待ちを受け付けました',
    o['氏名'] + ' 様\n\nキャンセル待ちを受け付けました。\n対象: ' + waitSummary_(x, o) + '(登録時点で満席の' + t.only.length + '日が対象)\n\n空きが出ましたら、ご登録の順にメールでご案内します。\n※1台につき月1回のため、すでにご予約のある月はご案内の対象外です。');
  return { days: t.only.length, skipped: days.length - t.only.length };
}
function waitStop_(p) {
  const x = ctx_(true), plate = normPlate_(p.plate), mail = String(p.email || '').trim().toLowerCase();
  const w = x.waits.find(w => w.id === p.id && w['ナンバー'] === plate && w['メール'].toLowerCase() === mail && (w['状態'] === '待機' || w['状態'] === '通知中'));
  if (!w) throw new Error('キャンセル待ちが見つかりません');
  stopWait_(x, w);
  return { ok: 1 };
}
function stopWait_(x, w) {
  const hold = x.res.find(r => r['待ちID'] === w.id && r['状態'] === '保留');
  w['状態'] = '停止'; saveRow_('WAIT', w);
  if (hold) { hold['状態'] = '解除'; saveRow_('RES', hold); processDate_(x, hold['日付']); }
}
/** 空きが出た日について、キャンセル待ちを登録順に確認し、空き枠を保留して案内メールを送る */
function processDate_(x, date) {
  if (!okDate_(x, date)) return;
  const ym = date.slice(0, 7);
  const waits = x.waits.filter(w => w['状態'] === '待機').sort((a, b) => a['作成'] < b['作成'] ? -1 : 1);
  for (const w of waits) {
    if (!targetDates_(x, JSON.parse(w['対象'])).includes(date)) continue;
    if (monthCount_(x, w['ナンバー'], ym, '') > 0) continue;
    const large = isTrue_(w['大型']), skip = String(w['スキップ'] || '').split(',');
    const st = starts_(x, date, large, busy_(x, date, '')).filter(s => skip.indexOf(date + '|' + fmtMin_(s)) < 0);
    if (!st.length) continue;
    hold_(x, w, date, st[0], large);
  }
}
function hold_(x, w, date, s, large) {
  const c = x.c, dur = large ? c.durLarge : c.durStd, token = Utilities.getUuid().replace(/-/g, '');
  const until = Utilities.formatDate(new Date(Date.now() + c.holdMin * 60000), TZ, "yyyy-MM-dd'T'HH:mm:ss");
  const o = {
    id: Utilities.getUuid().slice(0, 8), '予約番号': '', '日付': date, '開始': fmtMin_(s), '終了': fmtMin_(s + dur), '分': dur,
    '氏名': w['氏名'], '電話': w['電話'], 'メール': w['メール'], 'ナンバー': w['ナンバー'], 'ナンバー表示': w['ナンバー表示'], '車種': w['車種'], '大型': w['大型'],
    '状態': '保留', '作成': nowStr_(), '保留期限': until, '待ちID': w.id, 'トークン': token
  };
  appendRow_('RES', o); x.res.push(o);
  w['状態'] = '通知中'; w['最終通知'] = nowStr_(); saveRow_('WAIT', w);
  const link = c.siteUrl ? c.siteUrl + (c.siteUrl.indexOf('?') >= 0 ? '&' : '?') + 'hold=' + token : '(管理者にお問い合わせください)';
  const lim = until.slice(5, 10).replace('-', '月') + '日 ' + until.slice(11, 16);
  mail_(c, w['メール'], 'キャンセル待ち:空きが出ました',
    w['氏名'] + ' 様\n\nキャンセル待ちをいただいていた日程に空きが出ました。\n\n日時: ' + dateJa_(date) + ' ' + fmtMin_(s) + '〜' + fmtMin_(s + dur) +
    '\nお車: ' + w['車種'] + '(' + w['ナンバー表示'] + ')\n\n下記のリンクから ' + lim + ' までに「予約を確定する」を押してください。\n' + link +
    '\n\n期限を過ぎますと、次のお客様へご案内します。');
}
function holdInfo_(p) {
  const x = ctx_(false), r = x.res.find(r => r['トークン'] === String(p.token) && r['状態'] === '保留');
  if (!r || r['保留期限'] <= nowStr_()) throw new Error('この案内は期限が過ぎているか、すでに処理済みです。');
  return { date: r['日付'], start: r['開始'], end: r['終了'], model: r['車種'], plate: r['ナンバー表示'], name: r['氏名'], until: r['保留期限'].slice(11, 16) };
}
function holdConfirm_(p) {
  const x = ctx_(true), token = String(p.token), r = x.res.find(r => r['トークン'] === token && r['状態'] === '保留');
  if (!r || r['保留期限'] <= nowStr_()) throw new Error('この案内は期限が過ぎているか、すでに処理済みです。');
  const ym = r['日付'].slice(0, 7);
  if (monthCount_(x, r['ナンバー'], ym, token) > 0) throw new Error(monthLimitMsg_(ym));
  r['状態'] = '予約'; r['予約番号'] = newNo_(x); r['保留期限'] = ''; saveRow_('RES', r);
  const w = x.waits.find(w => w.id === r['待ちID']);
  if (w) { w['状態'] = '完了'; saveRow_('WAIT', w); }
  sendConfirm_(x.c, r);
  return { no: r['予約番号'], date: r['日付'], start: r['開始'], end: r['終了'], model: r['車種'] };
}
/** 10分おきに実行: 期限切れの保留を解除して次の人へ案内 / 新しく開いた日の確認 */
function processAll() {
  withLock_(() => {
    const x = ctx_(true), now = nowStr_(), dates = {};
    x.res.filter(r => r['状態'] === '保留' && r['保留期限'] <= now).forEach(r => {
      r['状態'] = '期限切れ'; saveRow_('RES', r);
      const w = x.waits.find(w => w.id === r['待ちID']);
      if (w && w['状態'] === '通知中') {
        w['状態'] = '待機';
        w['スキップ'] = (w['スキップ'] ? w['スキップ'] + ',' : '') + r['日付'] + '|' + r['開始'];
        saveRow_('WAIT', w);
      }
    });
    x.waits.filter(w => w['状態'] === '待機').forEach(w => targetDates_(x, JSON.parse(w['対象'])).forEach(d => dates[d] = 1));
    Object.keys(dates).sort().forEach(d => processDate_(x, d));
  });
}

/* ========== 管理者API ========== */
function adminRes_(p) {
  const x = ctx_(false), w = window_(x.c), from = p.from || w.today, to = p.to || '9999-12-31', now = nowStr_();
  return x.res.filter(r => (r['状態'] === '予約' || (r['状態'] === '保留' && r['保留期限'] > now)) && r['日付'] >= from && r['日付'] <= to)
    .sort((a, b) => (a['日付'] + a['開始']) < (b['日付'] + b['開始']) ? -1 : 1)
    .map(r => ({ id: r.id, no: r['予約番号'], date: r['日付'], start: r['開始'], end: r['終了'], name: r['氏名'], phone: r['電話'], email: r['メール'], plate: r['ナンバー表示'], model: r['車種'], large: isTrue_(r['大型']), status: r['状態'] }));
}
function adminCancel_(p) {
  const x = ctx_(true), r = x.res.find(r => r.id === p.id && (r['状態'] === '予約' || r['状態'] === '保留'));
  if (!r) throw new Error('見つかりません');
  if (r['状態'] === '保留') { const w = x.waits.find(w => w.id === r['待ちID']); if (w) stopWait_(x, w); }
  else cancelRes_(x, r);
  return { ok: 1 };
}
function adminWaits_() {
  const x = ctx_(true);
  return x.waits.filter(w => w['状態'] === '待機' || w['状態'] === '通知中')
    .sort((a, b) => a['作成'] < b['作成'] ? -1 : 1)
    .map(w => ({ id: w.id, created: w['作成'].replace('T', ' ').slice(0, 16), name: w['氏名'], phone: w['電話'], email: w['メール'], plate: w['ナンバー表示'], model: w['車種'], state: w['状態'], summary: waitSummary_(x, w), count: targetDates_(x, JSON.parse(w['対象'])).length }));
}
function adminWaitStop_(p) {
  const x = ctx_(true), w = x.waits.find(w => w.id === p.id);
  if (!w) throw new Error('見つかりません');
  stopWait_(x, w);
  return { ok: 1 };
}
function adminGet_() {
  const x = ctx_(false);
  const known = {}; table_('SET').forEach(r => known[r['項目']] = r['値']);
  return {
    settings: DEF_SETTINGS.map(r => ({ k: r[0], v: known[r[0]] === undefined ? r[1] : known[r[0]], d: r[2] })),
    models: x.models.map(m => ({ name: m['車種'], large: isTrue_(m['大型']) })),
    holidays: table_('HOL').map(h => ({ date: h['日付'], note: h['メモ'] }))
  };
}
function checkSettings_(s) {
  const t = k => String(s[k] === undefined ? '' : s[k]).trim();
  const timeKeys = ['開始', '昼休憩開始', '昼休憩終了', '最終開始(普通)', '最終開始(大型)', '解禁時刻'];
  timeKeys.forEach(k => { if (!/^\d{1,2}:\d{2}$/.test(t(k))) throw new Error('「' + k + '」は 09:30 のような時刻で入力してください'); });
  const numKeys = [['同時洗車台数', 1], ['洗車分(普通)', 5], ['洗車分(大型)', 5], ['刻み分', 5], ['先行月数', 1], ['最短日数', 0], ['保留時間(分)', 1]];
  numKeys.forEach(a => { const v = +t(a[0]); if (!(v >= a[1]) || v !== Math.floor(v)) throw new Error('「' + a[0] + '」は' + a[1] + '以上の整数で入力してください'); });
  if (+t('同時洗車台数') > 10) throw new Error('「同時洗車台数」は10台までです');
  if (toMin_(t('昼休憩開始')) > toMin_(t('昼休憩終了'))) throw new Error('昼休憩の開始と終了の時刻が逆です');
  if (t('定休曜日') && !/^[0-6]([,、\s]+[0-6])*$/.test(t('定休曜日'))) throw new Error('「定休曜日」は 日=0 〜 土=6 の数字で入力してください(例: 2,3)');
}
function adminSave_(p) {
  if (p.settings) {
    // 送られてきた項目だけを現在の設定に重ねる(送られなかった項目は今の値のまま)
    const cur = {}; table_('SET').forEach(r => cur[r['項目']] = r['値']);
    const merged = {};
    DEF_SETTINGS.forEach(r => { merged[r[0]] = p.settings[r[0]] !== undefined ? String(p.settings[r[0]]).trim() : (cur[r[0]] !== undefined ? cur[r[0]] : r[1]); });
    checkSettings_(merged);
    rewrite_('SET', DEF_SETTINGS.map(r => [r[0], merged[r[0]], r[2]]));
  }
  if (p.models) rewrite_('MODEL', p.models.filter(m => String(m.name).trim()).map(m => [String(m.name).trim(), m.large ? 'TRUE' : 'FALSE']));
  if (p.holidays) rewrite_('HOL', p.holidays.filter(h => /^\d{4}-\d{2}-\d{2}$/.test(h.date)).sort((a, b) => a.date < b.date ? -1 : 1).map(h => [h.date, h.note || '']));
  return { ok: 1 };
}
