const LIFF_ID='2012515520-dMDeNwhB', GAS='https://script.google.com/macros/s/AKfycbzffSbSy7Gb05bADVtV26LqBIRhKak-MtdKivI1VYIvl5BRI4Z9Hu0gIpQzb9hMeACM4Q/exec';  
const $=s=>document.querySelector(s);
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const WD='日月火水木金土';
const TYPES={yesno:'いる / いらない',yn:'YES / NO',maru:'〇 / ×',time:'時間で答える',text:'自由記述',choice:'選択肢（最大10）'};
const RT={daily:'毎日',weekly:'曜日固定',monthly:'毎月〇日',nth:'第〇×曜日',yearly:'毎年〇日',once:'単発'};
const OPT={yesno:['いる','いらない'],yn:['YES','NO'],maru:['〇','×']};
const ym=d=>d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0');
const S={tab:'cal',ym:ym(new Date()),data:{},sel:null,edit:null,rules:[],members:[],me:''};
const rt=c=>({daily:'毎日',weekly:'毎週'+(c.days||[]).map(i=>WD[i]).join(''),monthly:`毎月${c.dom}日`,nth:`第${c.nth}${WD[c.wd]}曜日`,yearly:`毎年${c.m}/${c.dom}`,once:'単発 '+c.date}[c.type]||'');
const nm=id=>S.members.find(m=>m.userId==id)?.name||'?';
const opts=r=>OPT[r.ans.type]||r.ans.choices||[];

async function api(a,p={}){
  const r=await fetch(GAS,{method:'POST',body:JSON.stringify({a,idToken:liff.getIDToken(),...p})});
  const j=await r.json(); if(!j.ok) throw j.e; return j.d;
}
async function load(){
  const [i,m]=await Promise.all([api('init'),api('month',{ym:S.ym})]);
  Object.assign(S,i,{data:m}); render();
}
async function main(){
  await liff.init({liffId:LIFF_ID});
  if(!liff.isLoggedIn()) return liff.login();
  const u=new URLSearchParams(location.search), st=new URLSearchParams((u.get('liff.state')||'').replace(/^\?/,''));
  const gp=k=>u.get(k)||st.get(k), q=gp('q'), t=gp('tab');
  if(['set','cal','list'].includes(t)) S.tab=t;
  if(q){S.sel=q.split('|')[0];S.ym=S.sel.slice(0,7);S.tab='cal';}
  await load();
}
main().catch(e=>{document.querySelector('main').textContent='エラー: '+e});

function go(t){S.tab=t;S.edit=null;render()}
function render(){
  document.querySelectorAll('nav button').forEach(b=>b.classList.toggle('on',b.dataset.t==S.tab));
  $('main').innerHTML={set:vSet,cal:vCal,list:vList}[S.tab]();
}

/* ---- カレンダー ---- */
function vCal(){
  const [y,m]=S.ym.split('-').map(Number),f=new Date(y,m-1,1).getDay(),n=new Date(y,m,0).getDate();
  let g=[...WD].map(w=>`<span class=sm>${w}</span>`).join('')+'<span></span>'.repeat(f);
  for(let d=1;d<=n;d++){const ds=S.ym+'-'+String(d).padStart(2,'0'),k=S.data[ds];
    g+=`<div class="d ${S.sel==ds?'sel':''}" onclick="pick('${ds}')">${d}${k?`<i>●${k.length}</i>`:''}</div>`}
  return `<div style="display:flex;justify-content:space-between;align-items:center"><button onclick="mv(-1)">＜</button><b>${y}年${m}月</b><button onclick="mv(1)">＞</button></div><div class=g>${g}</div>`+(S.sel?day():'<p class=sm>日付を選ぶと、事前登録や振り返りができます。</p>');
}
function day(){
  const L=S.data[S.sel]||[];
  return `<h3>${S.sel}</h3>`+(L.map(x=>{
    const r=S.rules.find(r=>r.id==x.ruleId); if(!r) return '';
    const mine=x.answers.find(a=>a.userId==S.me);
    return `<div class=row><b>${esc(r.title)}</b><div class=sm>${x.answers.map(a=>esc(nm(a.userId))+'：'+esc(a.value)).join(' ／ ')||'まだ回答がありません'}</div>${inp(r,mine?.value)}</div>`;
  }).join('')||'<p class=sm>この日の予定はありません。</p>');
}
function inp(r,v){
  const t=r.ans.type;
  if(t=='time'||t=='text') return `<input id="i_${r.id}" type="${t=='time'?'time':'text'}" value="${esc(v)}"> <button class=p onclick="ans('${r.id}',document.getElementById('i_${r.id}').value)">登録</button>`;
  return opts(r).map((x,i)=>`<button class="${x==v?'p':''}" onclick="ans('${r.id}',${i})">${esc(x)}</button>`).join('');
}
async function ans(id,v){
  const r=S.rules.find(r=>r.id==id); if(typeof v=='number') v=opts(r)[v]; if(v==='') return;
  await api('answer',{date:S.sel,ruleId:id,value:v}); S.data=await api('month',{ym:S.ym}); render();
}
function pick(ds){S.sel=ds;render()}
async function mv(n){const [y,m]=S.ym.split('-').map(Number);S.ym=ym(new Date(y,m-1+n,1));S.sel=null;S.data=await api('month',{ym:S.ym});render()}

/* ---- 一覧（カテゴリ別） ---- */
function vList(){
  const by={}; S.rules.filter(r=>r.active=='1').forEach(r=>(by[r.category||'未分類']??=[]).push(r));
  return Object.entries(by).map(([c,rs])=>`<h3>${esc(c)}</h3>`+rs.map(r=>{
    const ds=Object.keys(S.data).filter(d=>S.data[d].some(x=>x.ruleId==r.id));
    return `<div class=row><b>${esc(r.title)}</b><div class=sm>${rt(r.recur)}／${r.remind} リマインド／${r.deadline} 締切／${TYPES[r.ans.type]}</div>
    <div class=sm>${S.ym} の予定： ${ds.map(d=>`<a href="#" onclick="S.tab='cal';S.sel='${d}';render();return false">${+d.slice(8)}日(${S.data[d].find(x=>x.ruleId==r.id).answers.length})</a>`).join(' ')||'なし'}</div></div>`;
  }).join('')).join('')||'<p class=sm>「設定」タブでリマインドを追加してください。</p>';
}

/* ---- 設定 ---- */
function vSet(){
  if(S.edit) return form();
  return `<button class=p onclick="ed()">＋ 新しいリマインド</button>`+
    S.rules.map(r=>`<div class=row onclick="ed('${r.id}')"><b>${esc(r.title)}</b>${r.active=='1'?'':' <span class=sm>(停止中)</span>'}<div class=sm>${esc(r.category)}／${rt(r.recur)}／${r.remind}／締切 ${r.deadline}</div></div>`).join('')+
    `<h3>回答必須メンバー</h3><p class=sm>グループのメンバーがアプリを開く／回答すると自動で追加されます。</p>`+
    S.members.map(m=>`<label style="color:var(--ink)"><input type=checkbox ${m.required=='1'?'checked':''} onchange="req('${m.userId}',this.checked)"> ${esc(m.name)}</label>`).join('');
}
const req=(id,v)=>api('setReq',{userId:id,v});
function ed(id){
  const r=S.rules.find(r=>r.id==id);
  S.edit=r?structuredClone(r):{title:'',category:'',recur:{type:'daily'},ans:{type:'yesno',choices:[]},remind:'18:00',deadline:'20:00',active:'1'};
  render();
}
function rx(c){
  const n=(id,v,mx)=>`<input id=${id} type=number min=1 max=${mx} value="${v??1}" style="width:68px">`;
  return {daily:'',
    weekly:[...WD].map((w,i)=>`<label style="display:inline-block;margin:4px 8px 0 0;color:var(--ink)"><input type=checkbox class=wd value=${i} ${(c.days||[]).includes(i)?'checked':''}>${w}</label>`).join(''),
    monthly:`毎月 ${n('f_dom',c.dom,31)} 日`,
    nth:`第 ${n('f_n',c.nth,5)} <select id=f_w>${[...WD].map((w,i)=>`<option value=${i} ${c.wd==i?'selected':''}>${w}</option>`).join('')}</select> 曜日`,
    yearly:`毎年 ${n('f_m',c.m,12)} 月 ${n('f_dom',c.dom,31)} 日`,
    once:`<input id=f_date type=date value="${c.date||''}">`}[c.type];
}
function form(){
  const e=S.edit,c=e.recur,a=e.ans;
  const sel=(id,o,v)=>`<select id=${id} onchange="sync()">${Object.entries(o).map(([k,t])=>`<option value=${k} ${k==v?'selected':''}>${t}</option>`).join('')}</select>`;
  return `<label>タイトル（例：夕食はいりますか？）<br><input id=f_title value="${esc(e.title)}" style="width:100%"></label>
  <label>カテゴリ<br><input id=f_cat value="${esc(e.category)}" placeholder="食事 / ゴミ出し など"></label>
  <label>くり返し<br>${sel('f_rt',RT,c.type)}</label><div id=rx>${rx(c)}</div>
  <label>聞き方<br>${sel('f_at',TYPES,a.type)}</label>
  ${a.type=='choice'?`<label>選択肢（1行に1つ、10個まで）<br><textarea id=f_ch rows=5 style="width:100%">${esc((a.choices||[]).join('\n'))}</textarea></label>`:''}
  <label>リマインド時刻 <input id=f_rm type=time value="${e.remind}"></label>
  <label>回答期限 <input id=f_dl type=time value="${e.deadline}"></label>
  <label style="color:var(--ink)"><input type=checkbox id=f_ac ${e.active=='1'?'checked':''}> 有効にする</label>
  <p><button class=p onclick="save()">保存</button><button onclick="S.edit=null;render()">戻る</button>${e.id?'<button onclick="rm()">削除</button>':''}</p>`;
}
function collect(){
  const v=id=>document.getElementById(id)?.value,t=v('f_rt'),c={type:t};
  if(t=='weekly') c.days=[...document.querySelectorAll('.wd:checked')].map(x=>+x.value);
  if(t=='monthly'||t=='yearly') c.dom=+v('f_dom');
  if(t=='yearly') c.m=+v('f_m');
  if(t=='nth'){c.nth=+v('f_n');c.wd=+v('f_w')}
  if(t=='once') c.date=v('f_date');
  return {...S.edit,title:v('f_title'),category:v('f_cat'),recur:c,
    ans:{type:v('f_at'),choices:(v('f_ch')??(S.edit.ans.choices||[]).join('\n')).split('\n').map(s=>s.trim()).filter(Boolean).slice(0,10)},
    remind:v('f_rm'),deadline:v('f_dl'),active:document.getElementById('f_ac').checked?'1':'0'};
}
function sync(){S.edit=collect();render()}
async function save(){
  const r=collect(); if(!r.title) return alert('タイトルを入力してください');
  if(r.ans.type=='choice'&&r.ans.choices.length<2) return alert('選択肢を2つ以上入力してください');
  await api('saveRule',{rule:r}); S.edit=null; await load();
}
async function rm(){if(!confirm('削除しますか？'))return;await api('delRule',{id:S.edit.id});S.edit=null;await load()}
