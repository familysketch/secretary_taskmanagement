const LIFF_ID = '2012515520-dMDeNwhB';   
const $ = s => document.querySelector(s);
const pad = n => String(n).padStart(2, '0');
const btns = () => document.querySelectorAll('button[data-v]');
const say = t => { $('#m').textContent = t; };

// 初期値は「今から1時間後」
const d = new Date(Date.now() + 3600e3);
$('#t').value = pad(d.getHours()) + ':' + pad(d.getMinutes());

async function send(dinner) {
  const t = $('#t').value;
  if (!t) return say('帰る時間を選んでください');
  const messages = [{ type: 'text', text: `帰る時間：${t}\n夕食：${dinner}` }];
  btns().forEach(b => b.disabled = true);
  say('送信中…');
  try {
    const c = liff.getContext();
    if (c && ['utou', 'room', 'group'].includes(c.type)) {
      await liff.sendMessages(messages);                       // 開いたトークに自分の名前で送信
    } else if (liff.isApiAvailable('shareTargetPicker')) {
      const r = await liff.shareTargetPicker(messages);        // トーク以外から開いた場合は送信先を選ぶ
      if (!r) { btns().forEach(b => b.disabled = false); return say(''); }
    } else throw new Error('LINEのトークからこのページを開いてください');
    say('✔ 送信しました');
    setTimeout(() => liff.closeWindow(), 700);
  } catch (e) {
    say('送信できませんでした：' + (e.message || e));
    btns().forEach(b => b.disabled = false);
  }
}

btns().forEach(b => b.addEventListener('click', () => send(b.dataset.v)));
liff.init({ liffId: LIFF_ID }).catch(e => say('LIFFの初期化に失敗：' + e.message));
