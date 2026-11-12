/**
 * 家族ボードBot  (スプレッドシート紐づけGAS) ― 個人/グループ/複数人トーク対応版
 * スクリプトプロパティ: TOKEN(Messaging APIのチャネルアクセストークン長期)
 *                       LOGIN_CHANNEL_ID(LINEログインチャネルのID) / LIFF_ID
 * 初回に setup() を1回実行 → シート作成 + 5分おきのトリガー登録
 * appsscript.json の timeZone は "Asia/Tokyo"
 * 「スペース」= 通知先1つ分。個人=ユーザーID(U…) / グループ=C… / 複数人トーク=R…
 */
const P = PropertiesService.getScriptProperties();
const TZ = 'Asia/Tokyo';
const SH = {
  members: ['k', 'space', 'userId', 'name', 'required'],
  rules:   ['id', 'space', 'title', 'category', 'recur', 'ans', 'remind', 'deadline', 'active'],
  answers: ['k', 'date', 'ruleId', 'userId', 'value'],
  log:     ['k', 'reminded', 'closed'],
  spaces:  ['id', 'name']
};
const OPT = { yesno: ['いる', 'いらない'], yn: ['YES', 'NO'], maru: ['〇', '×'] };
// 叱り文句（{n}=名前）。自由に差し替え可
const SCOLD = [
  // ① 大阪のおかん（関西弁）
  '{n}！スマホ見とるやろ！いつまで放置しとんねん、ええ加減にしなささい、どつきまわすぞ！',
  
  // ② 博多のおかん（博多弁）
  '{n}！なんば無視しとっとね！いい加減にせんと怒るよ、返事ば出さんと、くらすぞ！',
  
  // ③ 津軽のおかん（東北弁）
  '{n}！何回言えば解かるば！さっさと返事さねば、おどがす（脅す）ぞ！しっかりせ！',
  
  // ④ 名古屋のおかん（名古屋弁）
  '{n}！たわけ！いつまで放置しとるの！さっさと返事せんか、ごじゃはげに怒るでね！',
  
  // ⑤ 土佐のおかん（高知弁）
  '{n}！なに無視しちゅうがぜ！返事出さんと、おんびん（お説教）食らわすきね！'
];

function setup() {
  const ss = SpreadsheetApp.getActive();
  for (const n in SH) {
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    s.getRange(1, 1, 1000, SH[n].length).setNumberFormat('@'); // 日付/時刻の自動変換を防ぐ
    s.getRange(1, 1, 1, SH[n].length).setValues([SH[n]]);
  }
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(5).create();
}

/* ---------- 小さなDB層 ---------- */
const S = n => SpreadsheetApp.getActive().getSheetByName(n);
function read(n) {
  const v = S(n).getDataRange().getValues(), h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}
function put(n, key, o) { // keyで検索して更新、なければ追加（渡した項目のみ上書き）
  const s = S(n), v = s.getDataRange().getValues(), h = v[0], ki = h.indexOf(key);
  const i = v.findIndex((r, j) => j && r[ki] == o[key]), old = i > 0 ? v[i] : [];
  const row = h.map((k, c) => k in o ? (typeof o[k] == 'object' ? JSON.stringify(o[k]) : String(o[k])) : (old[c] ?? ''));
  i > 0 ? s.getRange(i + 1, 1, 1, h.length).setValues([row]) : s.appendRow(row);
}
function del(n, key, val) {
  const s = S(n), v = s.getDataRange().getValues(), i = v.findIndex((r, j) => j && r[v[0].indexOf(key)] == val);
  if (i > 0) s.deleteRow(i + 1);
}
const rules = () => read('rules').map(r => ({ ...r, recur: JSON.parse(r.recur || '{}'), ans: JSON.parse(r.ans || '{}') }));
const mems = space => read('members').filter(m => m.space == space);
const fmt = d => Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
const liffUrl = qs => 'https://liff.line.me/' + P.getProperty('LIFF_ID') + (qs ? '?' + qs : '');
const getLog = k => read('log').find(l => l.k == k) || {};

/* ---------- LINE API ---------- */
function line(p, b) {
  const o = { headers: { Authorization: 'Bearer ' + P.getProperty('TOKEN') }, muteHttpExceptions: true };
  if (b) { o.method = 'post'; o.contentType = 'application/json'; o.payload = JSON.stringify(b); }
  return UrlFetchApp.fetch('https://api.line.me/v2/bot/' + p, o);
}
const push = (to, msgs) => line('message/push', { to, messages: msgs });
const reply = (t, msgs) => line('message/reply', { replyToken: t, messages: msgs });

// スペースのメンバーを取得。未登録ならLINEで所属確認して登録(回答必須=ON)。所属していなければnull
function member(space, u, name) {
  const k = space + '|' + u, m = read('members').find(x => x.k == k);
  if (m) return m;
  if (!name) {
    const path = { C: 'group/' + space + '/member/' + u, R: 'room/' + space + '/member/' + u, U: 'profile/' + u }[space[0]];
    const r = line(path);
    if (r.getResponseCode() != 200) return null;
    name = JSON.parse(r.getContentText()).displayName;
  }
  if (!read('spaces').some(s => s.id == space)) {
    let sn = { U: '個人', R: 'トークルーム' }[space[0]];
    if (space[0] == 'C') {
      const r = line('group/' + space + '/summary');
      sn = r.getResponseCode() == 200 ? JSON.parse(r.getContentText()).groupName : 'グループ';
    }
    put('spaces', 'id', { id: space, name: sn });
  }
  const nm = { k, space, userId: u, name, required: '1' };
  put('members', 'k', nm);
  return nm;
}

/* ---------- 繰り返し判定 ---------- */
function matches(r, d) {
  const c = r.recur, wd = d.getDay(), dom = d.getDate();
  switch (c.type) {
    case 'daily': return true;
    case 'weekly': return (c.days || []).includes(wd);
    case 'monthly': return dom == c.dom;
    case 'nth': return wd == c.wd && Math.ceil(dom / 7) == c.nth;
    case 'yearly': return d.getMonth() + 1 == c.m && dom == c.dom;
    case 'once': return fmt(d) == c.date;
  }
  return false;
}

/* ---------- 定期実行(5分おき) ---------- */
function tick() {
  const n = new Date(), d = fmt(n), hm = Utilities.formatDate(n, TZ, 'HH:mm');
  rules().filter(r => r.active == '1' && r.space && matches(r, n)).forEach(r => {
    const k = d + '|' + r.id, l = getLog(k);
    if (!l.reminded && hm >= r.remind && hm < r.deadline) { put('log', 'k', { k, reminded: '1' }); remind(r, d); }
    else if (l.reminded && !l.closed && hm >= r.deadline) { put('log', 'k', { k, closed: '1' }); closeRule(r, d); }
  });
}

function remind(r, d) {
  const ms = mems(r.space), nm = id => (ms.find(m => m.userId == id) || {}).name || '?';
  const ans = read('answers').filter(a => a.date == d && a.ruleId == r.id);
  const body = [{ type: 'text', text: r.title, weight: 'bold', size: 'lg', wrap: true }]
    .concat(ans.map(a => ({ type: 'text', size: 'sm', wrap: true, color: '#555555',
      text: nm(a.userId) + 'さんは「' + a.value + '」と登録されています。変更が必要あれば教えてください。' })));
  const t = r.ans.type, opts = OPT[t] || (t == 'choice' ? r.ans.choices : null);
  const cal = liffUrl('s=' + r.space + '&q=' + encodeURIComponent(d + '|' + r.id));
  const btns = opts
    ? opts.map(o => ({ type: 'button', style: 'primary', color: '#06c755', margin: 'sm',
        action: { type: 'postback', label: String(o).slice(0, 20), data: [d, r.id, o].join('|') } }))
    : [{ type: 'button', style: 'primary', color: '#06c755',
        action: { type: 'uri', label: t == 'time' ? '時間を入力' : '入力する', uri: cal } }];
  btns.push({ type: 'button', style: 'secondary', margin: 'md', action: { type: 'uri', label: 'カレンダーを開く', uri: cal } });
  push(r.space, [{ type: 'flex', altText: r.title, contents: { type: 'bubble',
    body: { type: 'box', layout: 'vertical', spacing: 'sm', contents: body },
    footer: { type: 'box', layout: 'vertical', contents: btns } } }]);
  checkDone(d, r);
}

function sum(d, r, ans) {
  const ms = mems(r.space), nm = id => (ms.find(m => m.userId == id) || {}).name || '?';
  return { type: 'text', text: '📋 ' + r.title + '（' + d + '）\n' + ans.map(a => nm(a.userId) + '：' + a.value).join('\n') };
}
function checkDone(d, r) { // 回答必須者が全員回答済みなら集計を送る
  if (!getLog(d + '|' + r.id).reminded) return;
  const mem = mems(r.space).filter(m => m.required == '1');
  const ans = read('answers').filter(a => a.date == d && a.ruleId == r.id);
  if (mem.length && mem.every(m => ans.some(a => a.userId == m.userId))) push(r.space, [sum(d, r, ans)]);
}
function closeRule(r, d) { // 期限: 未回答者を叱る + その時点の集計
  const mem = mems(r.space).filter(m => m.required == '1');
  const ans = read('answers').filter(a => a.date == d && a.ruleId == r.id);
  const miss = mem.filter(m => !ans.some(a => a.userId == m.userId));
  if (!miss.length) return;
  const msgs = [{ type: 'text', text: miss.map(m => SCOLD[Math.floor(Math.random() * SCOLD.length)].replace('{n}', m.name)).join('\n') }];
  if (ans.length) msgs.push(sum(d, r, ans));
  push(r.space, msgs);
}
function saveAns(d, rid, u, v) {
  put('answers', 'k', { k: [d, rid, u].join('|'), date: d, ruleId: rid, userId: u, value: v });
  const r = rules().find(x => x.id == rid);
  if (r && d == fmt(new Date())) checkDone(d, r);
}

/* ---------- Webhook / LIFF API ---------- */
function doPost(e) {
  const b = JSON.parse(e.postData.contents);
  if (b.events) return hook(b);
  let out;
  try {
    const u = who(b.idToken);
    if (b.a != 'init' && !read('members').some(m => m.k == b.space + '|' + u.userId)) throw 'このスペースのメンバーではありません';
    out = { ok: 1, d: A[b.a](u, b) };
  } catch (x) { out = { ok: 0, e: String(x) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
// ※GASはX-Line-Signatureヘッダを読めないため署名検証は省略（回答はスペース一致とLINE所属確認で保護）
function hook(b) {
  b.events.forEach(ev => {
    const s = ev.source, space = s.groupId || s.roomId || s.userId;
    if (!space || !s.userId) return;
    const m = member(space, s.userId);
    if (!m) return;
    if (ev.type == 'follow')
      reply(ev.replyToken, [{ type: 'text', text: '友だち追加ありがとう！「メニュー」と送ると、予定の管理画面を開けるよ。' }, menuMsg(space)]);
    if (ev.type == 'message' && ev.message.type == 'text' && /^(メニュー|めにゅー|menu)$/i.test(ev.message.text.trim()))
      reply(ev.replyToken, [menuMsg(space)]);
    if (ev.type == 'postback') {
      const p = ev.postback.data.split('|'), v = p.slice(2).join('|'), r = rules().find(x => x.id == p[1]);
      if (!r || r.space != space) return; // 別スペースの質問には回答させない
      saveAns(p[0], p[1], m.userId, v);
      reply(ev.replyToken, [{ type: 'text', text: m.name + 'さんの回答「' + v + '」を受け付けました' }]);
    }
  });
  return ContentService.createTextOutput('ok');
}
function menuMsg(space) {
  const b = (label, tab) => ({ type: 'button', style: 'primary', color: '#06c755', margin: 'sm',
    action: { type: 'uri', label, uri: liffUrl('s=' + space + '&tab=' + tab) } });
  return { type: 'flex', altText: 'メニュー', contents: { type: 'bubble',
    body: { type: 'box', layout: 'vertical', contents: [{ type: 'text', text: 'メニュー', weight: 'bold', size: 'lg' }] },
    footer: { type: 'box', layout: 'vertical', contents: [b('カレンダー', 'cal'), b('一覧', 'list'), b('設定', 'set')] } } };
}
function who(t) { // IDトークン検証（LINEログイン済みの本人確認）
  const r = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', { method: 'post', muteHttpExceptions: true,
    payload: { id_token: t, client_id: P.getProperty('LOGIN_CHANNEL_ID') } });
  const j = JSON.parse(r.getContentText());
  if (!j.sub) throw '認証に失敗しました';
  return { userId: j.sub, name: j.name || '自分' };
}
function month(space, ym) {
  const [y, m] = ym.split('-').map(Number), rs = rules().filter(r => r.active == '1' && r.space == space), an = read('answers'), out = {};
  for (let d = 1, n = new Date(y, m, 0).getDate(); d <= n; d++) {
    const dt = new Date(y, m - 1, d), ds = fmt(dt);
    const l = rs.filter(r => matches(r, dt)).map(r => ({ ruleId: r.id,
      answers: an.filter(a => a.date == ds && a.ruleId == r.id).map(a => ({ userId: a.userId, value: a.value })) }));
    if (l.length) out[ds] = l;
  }
  return out;
}
const A = {
  init: u => { // 個人スペースを自動作成し、登録済みグループ/トークの所属も確認 → 使えるスペース一覧を返す
    member(u.userId, u.userId, u.name);
    read('spaces').filter(s => s.id[0] != 'U').forEach(s => member(s.id, u.userId));
    const mine = read('members').filter(m => m.userId == u.userId).map(m => m.space);
    return { spaces: read('spaces').filter(s => mine.includes(s.id)) };
  },
  data: (u, p) => ({ me: u.userId, members: mems(p.space), rules: rules().filter(r => r.space == p.space) }),
  month: (u, p) => month(p.space, p.ym),
  saveRule: (u, p) => {
    const r = p.rule; r.id = r.id || Utilities.getUuid().slice(0, 8);
    const old = rules().find(x => x.id == r.id);
    if (old && old.space != p.space) throw '権限がありません';
    r.space = p.space; put('rules', 'id', r); return 1;
  },
  delRule: (u, p) => { const old = rules().find(x => x.id == p.id); if (old && old.space == p.space) del('rules', 'id', p.id); return 1; },
  setReq: (u, p) => { put('members', 'k', { k: p.space + '|' + p.userId, required: p.v ? '1' : '0' }); return 1; },
  answer: (u, p) => {
    const r = rules().find(x => x.id == p.ruleId);
    if (!r || r.space != p.space) throw '不正なリクエストです';
    saveAns(p.date, p.ruleId, u.userId, p.value); return 1;
  }
};
