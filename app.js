"use strict";
// ================= ZIP + XLSX (без бібліотек) =================
async function inflateRaw(u8){
  const ds=new DecompressionStream('deflate-raw');
  const stream=new Blob([u8]).stream().pipeThrough(ds);
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
function readZip(buf){
  const d=new DataView(buf), u8=new Uint8Array(buf);
  let e=-1;
  for(let i=buf.byteLength-22;i>=Math.max(0,buf.byteLength-22-65535);i--){ if(d.getUint32(i,true)===0x06054b50){e=i;break;} }
  if(e<0) throw new Error('Це не схоже на .xlsx (не знайдено zip-структуру)');
  const cnt=d.getUint16(e+10,true); let p=d.getUint32(e+16,true);
  const map=new Map(), dec=new TextDecoder();
  for(let k=0;k<cnt;k++){
    if(d.getUint32(p,true)!==0x02014b50) throw new Error('Пошкоджений zip');
    const method=d.getUint16(p+10,true), csize=d.getUint32(p+20,true);
    const nl=d.getUint16(p+28,true), xl=d.getUint16(p+30,true), cl=d.getUint16(p+32,true), off=d.getUint32(p+42,true);
    const name=dec.decode(u8.subarray(p+46,p+46+nl));
    map.set(name.toLowerCase(),{method,csize,off});
    p+=46+nl+xl+cl;
  }
  return {
    has:n=>map.has(n.toLowerCase()),
    async text(n){
      const en=map.get(n.toLowerCase()); if(!en) return null;
      const st=en.off+30+d.getUint16(en.off+26,true)+d.getUint16(en.off+28,true);
      const raw=u8.subarray(st,st+en.csize);
      const out=en.method===0?raw:await inflateRaw(raw);
      return new TextDecoder('utf-8').decode(out);
    }
  };
}
const ENT={amp:'&',lt:'<',gt:'>',quot:'"',apos:"'"};
function unesc(s){
  if(s.indexOf('&')<0) return s;
  return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi,(m,e)=>{
    if(e[0]==='#') return String.fromCodePoint(e[1]==='x'||e[1]==='X'?parseInt(e.slice(2),16):parseInt(e.slice(1),10));
    return ENT[e]??m;
  });
}
function attr(s,k){ const m=new RegExp('(?:^|\\s)(?:\\w+:)?'+k+'\\s*=\\s*"([^"]*)"').exec(s)||new RegExp('(?:^|\\s)(?:\\w+:)?'+k+"\\s*=\\s*'([^']*)'").exec(s); return m?unesc(m[1]):null; }
async function openBook(buf){
  const z=readZip(buf);
  const wb=await z.text('xl/workbook.xml'); if(!wb) throw new Error('У файлі немає xl/workbook.xml');
  const rels=await z.text('xl/_rels/workbook.xml.rels')||'';
  const relMap={}; for(const m of rels.matchAll(/<(?:\w+:)?Relationship\b([^>]*)>/g)){ relMap[attr(m[1],'Id')]=attr(m[1],'Target'); }
  const sheets=[]; for(const m of wb.matchAll(/<(?:\w+:)?sheet\b([^>]*)>/g)){
    const a=m[1], name=attr(a,'name'), rid=attr(a,'id');
    let t=relMap[rid]||''; t=t.startsWith('/')?t.slice(1):'xl/'+t.replace(/^\.\//,'');
    sheets.push({name,path:t});
  }
  let shared=null;
  async function getShared(){
    if(shared) return shared; shared=[];
    const ss=await z.text('xl/sharedStrings.xml'); if(!ss) return shared;
    for(const m of ss.matchAll(/<(?:\w+:)?si>([\s\S]*?)<\/(?:\w+:)?si>|<(?:\w+:)?si\/>/g)){
      if(m[1]===undefined){shared.push('');continue;}
      const inner=m[1].replace(/<(?:\w+:)?rPh\b[\s\S]*?<\/(?:\w+:)?rPh>/g,'');
      let s=''; for(const t of inner.matchAll(/<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g)) s+=t[1];
      shared.push(unesc(s));
    }
    return shared;
  }
  async function readSheet(name, wantRows){
    const sh=sheets.find(s=>s.name.trim().toLowerCase()===name.trim().toLowerCase()); if(!sh) throw new Error('Аркуш '+name+' не знайдено');
    const xml=await z.text(sh.path); if(xml==null) throw new Error('Не вдалося прочитати '+sh.path);
    const sst=await getShared();
    const rows=[]; let autoRow=0;
    const re=/<(?:\w+:)?c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:\w+:)?c>)/g;
    let m;
    while((m=re.exec(xml))){
      const a=m[1], inner=m[2]||'';
      const r=attr(a,'r'); let row, col;
      if(r){ const mm=/^([A-Z]+)(\d+)$/i.exec(r); col=0; for(const ch of mm[1].toUpperCase()) col=col*26+(ch.charCodeAt(0)-64); col--; row=+mm[2]-1; }
      else { row=autoRow; col=(rows[row]||[]).length; }
      autoRow=row;
      if(wantRows && row>=wantRows) break;
      const t=attr(a,'t')||'n';
      let v=null;
      const vm=/<(?:\w+:)?v>([\s\S]*?)<\/(?:\w+:)?v>/.exec(inner);
      if(t==='s'){ v=vm?(sst[+vm[1]]??''):''; }
      else if(t==='inlineStr'){ let s=''; for(const tt of inner.matchAll(/<(?:\w+:)?t(?:\s[^>]*)?>([\s\S]*?)<\/(?:\w+:)?t>/g)) s+=tt[1]; v=unesc(s); }
      else if(t==='str'||t==='e'||t==='d'){ v=vm?unesc(vm[1]):''; }
      else if(t==='b'){ v=vm?(+vm[1]):null; }
      else { if(vm){ const f=parseFloat(vm[1]); v=Number.isFinite(f)?f:unesc(vm[1]); } }
      (rows[row]||(rows[row]=[]))[col]=v;
    }
    return rows;
  }
  return {sheets, readSheet};
}

// ================= Розбір значень =================
const DAY=86400000;
function parseDt(v){
  if(v==null||v==='') return null;
  if(typeof v==='number'){ if(v>20000&&v<100000) return Math.round((v-25569)*DAY/1000)*1000; return null; }
  const s=String(v).trim(); if(!s) return null;
  const n=s.match(/\d+/g); if(!n||n.length<3) return null;
  const N=n.map(Number); let y,mo,d,r;
  if(s.includes('-')&&N[0]>1900){[y,mo,d]=N;r=N.slice(3);}
  else if(s.includes('/')){mo=N[0];d=N[1];y=N[2];r=N.slice(3);}   // американський формат
  else if(s.includes('.')&&N[2]>1900){d=N[0];mo=N[1];y=N[2];r=N.slice(3);}
  else return null;
  let h=r[0]||0; const mi=r[1]||0, se=r[2]||0, up=s.toUpperCase();
  if(up.includes('PM')&&h<12) h+=12; if(up.includes('AM')&&h===12) h=0;
  if(mo<1||mo>12||d<1||d>31||h>23||mi>59) return null;
  return Date.UTC(y,mo-1,d,h,mi,se);
}
function parsePct(v){
  if(v==null||v==='') return null;
  if(typeof v==='number') return v*100;
  const t=String(v).replace(/%/g,'').replace(/,/g,'.').replace(/\s+/g,'').replace(/−/g,'-');
  if(t==='') return null; const f=Number(t); return Number.isFinite(f)?f:null;
}
const pad=x=>String(x).padStart(2,'0');
const fDT=t=>{const d=new Date(t);return `${pad(d.getUTCDate())}.${pad(d.getUTCMonth()+1)}.${d.getUTCFullYear()} ${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`};
const fD=t=>{const d=new Date(t);return `${pad(d.getUTCDate())}.${pad(d.getUTCMonth()+1)}.${d.getUTCFullYear()}`};
const fHM=t=>{const d=new Date(t);return `${pad(d.getUTCHours())}:${pad(d.getUTCMinutes())}`};

// ================= Налаштування =================
const DEF={
  dateFrom:'2026-06-01', cutFrom:'2026-08-22 08:07:00', cutTo:'2026-08-22 08:43:59',
  exclude:'Test_Program,MITOUSDT,PTBUSDT,SGCUSDT',
  good:21, k:3, m:0.5, top19:65, topMed:25, topCv:0.9, topNeg:12, bl19:60, blMed:19, blCv:1.4, blMin:-70, blNeg:20,
  minN:5, top:20, days:10, side:'all',
  spike:{Events:4,_other:1}
};
const SET_FIELDS=[
  ['dateFrom','Дані з дати','text'],['cutFrom','Вирізати шматок: від','text'],['cutTo','Вирізати шматок: до','text'],
  ['exclude','Виключені монети (через кому)','text'],
  ['good','Поріг доброго відкату, % (нижче — луз)','number'],
  ['k','k (стоп / тейк)','number'],['m','m (збиток на недовідкаті)','number'],
  ['top19','Топ: частка добрих не менше, %','number'],['topMed','Топ: медіана не менше, %','number'],['topCv','Топ: CV не більше','number'],['topNeg','Топ: мінусів не більше, %','number'],
  ['bl19','Blacklist: частка добрих менше, %','number'],['blMed','Blacklist: медіана менше, %','number'],['blCv','Blacklist: CV більше','number'],['blMin','Blacklist: мінімум нижче, %','number'],['blNeg','Blacklist: мінусів не менше, %','number'],
];
let S=structuredClone(DEF);
function loadSettings(){ try{ const s=VAULT.settings||null; if(s) return Object.assign(structuredClone(DEF),s,{spike:Object.assign({},DEF.spike,s.spike||{})}); }catch(e){} return structuredClone(DEF); }
function saveSettings(){ VAULT.settings=S; persist(); }

// ================= Стан =================
let BOOK=null, FILE='', TAB=null, TABS=[], CACHE={}, sortKey='score', sortDir=-1;
const PERIODS=[[1,'1д'],[3,'3д'],[7,'7д'],[10,'10д'],[14,'14д'],[30,'30д'],[0,'Все']];
const TOPS=[[10,'10'],[20,'20'],[30,'30'],[50,'50'],[0,'Всі']];
const SIDES=[['all','Всі'],['Long','Long'],['Short','Short']];
const spikeDef=t=>S.spike[t]??(/^events$/i.test(t)?DEF.spike.Events:DEF.spike._other);
const isMain=t=>/^events$/i.test(t);
function bucketEdges(tab,min,max){
  let e=isMain(tab)?[4,5,6,8,10,12,15]:[1,1.5,2,3,4,5,8];
  e=e.filter(x=>x>min); e.unshift(min); if(max>0) e=e.filter(x=>x<max);
  return e;
}

// ================= Завантаження вкладки =================
async function loadTab(name){
  if(CACHE[name]) return CACHE[name];
  const rows=await BOOK.readSheet(name);
  let hi=-1; for(let i=0;i<Math.min(rows.length,20);i++){ if((rows[i]||[]).some(c=>String(c??'').trim().toLowerCase()==='coin')){hi=i;break;} }
  if(hi<0) throw new Error(`На аркуші ${name} немає рядка заголовків (Date, Coin, ...)`);
  const h=rows[hi].map(c=>String(c??'').trim().toLowerCase());
  const f=p=>h.findIndex(x=>x&&x.startsWith(p));
  const ci={d:f('date'),c:f('coin'),l:f('leverage'),s:f('side'),sp:f('spike'),rb:f('rollback')};
  for(const [k,n] of [['d','Date'],['c','Coin'],['s','Side'],['sp','SpikePercent'],['rb','RollbackPercent']]) if(ci[k]<0) throw new Error(`На аркуші ${name} немає колонки ${n}`);
  const ev=[]; let bad=0;
  for(let i=hi+1;i<rows.length;i++){
    const r=rows[i]; if(!r) continue;
    const coin=String(r[ci.c]??'').trim(); if(!coin) continue;
    const t=parseDt(r[ci.d]), sp=parsePct(r[ci.sp]), rb=parsePct(r[ci.rb]);
    if(t==null||sp==null||rb==null){bad++;continue;}
    let side=String(r[ci.s]??'').replace(/\s+/g,''); const sl=side.toLowerCase(); if(sl==='long') side='Long'; else if(sl==='short') side='Short';
    let lev=ci.l>=0?String(r[ci.l]??'').trim().replace(/x$/i,''):'';
    ev.push({t,coin,cu:coin.toUpperCase(),side,lev,sp,rb});
  }
  return CACHE[name]={ev,bad};
}
function cleaned(name){
  const raw=CACHE[name]; const ex=S.exclude.split(',').map(x=>x.trim().toUpperCase()).filter(Boolean);
  const from=parseDt(S.dateFrom), cf=parseDt(S.cutFrom), ct=parseDt(S.cutTo);
  let rmEx=0,rmOld=0,rmCut=0; const out=[];
  for(const e of raw.ev){
    if(ex.some(x=>x===e.cu||x+'USDT'===e.cu)){rmEx++;continue;}
    if(from!=null&&e.t<from){rmOld++;continue;}
    if(cf!=null&&ct!=null&&e.t>=cf&&e.t<=ct){rmCut++;continue;}
    out.push(e);
  }
  return {ev:out,rmEx,rmOld,rmCut,bad:raw.bad};
}

// ================= Метрики =================
function metrics(a){
  const n=a.length; const v=a.slice().sort((x,y)=>x-y);
  const mean=v.reduce((s,x)=>s+x,0)/n;
  const med=n%2?v[(n-1)/2]:(v[n/2-1]+v[n/2])/2;
  const sd=n>1?Math.sqrt(v.reduce((s,x)=>s+(x-mean)**2,0)/(n-1)):0;
  const cv=n>1&&mean>0?sd/mean:null;
  const G=+S.good||19; let g0=0,g19=0,g35=0,g50=0,neg=0; for(const x of v){ if(x>=0)g0++; if(x>=G)g19++; if(x>=35)g35++; if(x>=50)g50++; if(x<0)neg++; }
  const mid=n-g19-neg;
  const e=(g19-S.m*mid-S.k*neg)/n;
  const score=0.4*g19/n*100+0.3*g35/n*100+0.3*med-2*neg/n*100-20*(cv??9);
  return {n,mean,med,cv,g0,g19,g35,g50,neg,mid,min:v[0],e,score};
}
function status(m){
  const a19=m.g19/m.n*100, ng=m.neg/m.n*100, cv=m.cv??9;
  if(a19>=S.top19&&ng<=S.topNeg&&cv<=S.topCv&&m.med>=S.topMed){
    const edge=a19<S.top19+3||ng>S.topNeg-2||cv>S.topCv-0.1||m.med<S.topMed+2;
    return edge?'edge':'top';
  }
  if(a19<S.bl19||m.med<S.blMed||cv>S.blCv||m.min<S.blMin||ng>=S.blNeg) return 'bl';
  return 'mid';
}
const STN={top:'топ',edge:'на межі',mid:'середня',bl:'blacklist'};
const P=(k,n)=>`${Math.round(100*k/n)}% (${k})`;
const p1=x=>x.toFixed(1).replace('.',',')+'%';
const sg=x=>(x<0?'−':'+')+Math.abs(x).toFixed(1).replace('.',',')+'%';
const sg2=x=>(x<-1e-9?'−':'+')+Math.abs(x).toFixed(2).replace('.',',');
const cvs=c=>c==null?'—':c.toFixed(2).replace('.',',');
const short=c=>c.replace(/USDT$/i,'');
const esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));

// ================= Розрахунок =================
let LAST=null;
function compute(){
  const c=cleaned(TAB); const all=c.ev;
  if(!all.length) return {err:'Після очистки немає подій'};
  const last=all.reduce((m,e)=>e.t>m?e.t:m,-Infinity);
  const days=+S.days||0, from=days>0?last-days*DAY:-Infinity;
  const smin=+(S.spike[TAB]??spikeDef(TAB)), smax=+(S['smax_'+TAB]||0);
  const q=(document.getElementById('coinq').value||'').trim().toUpperCase();
  const ev=all.filter(e=>e.t>=from && e.sp>=smin && (!(smax>0)||e.sp<smax) && (S.side==='all'||e.side===S.side));
  return {c,last,from,days,smin,smax,ev,q};
}

function render(){
  if(!BOOK||!TAB||!CACHE[TAB]) return;
  const R=compute(); if(R.err){ document.getElementById('sum').textContent=R.err; return; }
  LAST=R;
  const {ev,last,from,days,smin,smax,c}=R; const n=ev.length;
  const first=ev.length?ev.reduce((m,e)=>e.t<m?e.t:m,Infinity):last;
  const coinsAll=new Map(); for(const e of ev){ (coinsAll.get(e.coin)||coinsAll.set(e.coin,[]).get(e.coin)).push(e); }
  const minN=Math.max(1,+S.minN||1);
  let rows=[]; let skC=0,skE=0;
  for(const [coin,es] of coinsAll){
    if(es.length<minN){skC++;skE+=es.length;continue;}
    const m=metrics(es.map(e=>e.rb));
    const sd=s=>{const x=es.filter(e=>e.side===s); return x.length?{k:x.filter(e=>e.rb>=S.good).length,n:x.length}:null;};
    const levs=[...new Set(es.map(e=>e.lev))].sort((a,b)=>(+a||1e9)-(+b||1e9));
    rows.push({coin,m,st:status(m),L:sd('Long'),S:sd('Short'),lev:levs.join(', ')});
  }
  // ранг за score (від найкращої до найгіршої)
  rows.sort((a,b)=>b.m.score-a.m.score); rows.forEach((r,i)=>r.rank=i+1);
  const white=rows.filter(r=>r.st==='top'), edge=rows.filter(r=>r.st==='edge'), black=rows.filter(r=>r.st==='bl');

  // підсумок
  const spikeTxt=(smax>0?`${smin}–${smax}%`:`від ${smin}%`).replace(/\./g,',');
  document.getElementById('sum').innerHTML=
    `<b>${esc(TAB)}</b> · спайки <b>${spikeTxt}</b> · ${days>0?`<b>${days} дн.</b>: ${fDT(from)} — ${fDT(last)}`:`<b>весь період</b>: ${n?fDT(first):'—'} — ${fDT(last)}`}`+
    ` · відлік від останньої події у файлі${S.side!=='all'?` · сторона <b>${S.side}</b>`:''}<br>`+
    `<span class="note">Очистка: дані з ${esc(S.dateFrom)}; вирізано ${c.rmCut} подій (${esc(S.cutFrom)} — ${esc(S.cutTo)}); виключено ${c.rmEx} рядків монет (${esc(S.exclude)}); старіших за дату — ${c.rmOld}; нерозпізнаних — ${c.bad}.</span>`;
  const mAll=n?metrics(ev.map(e=>e.rb)):null;
  document.getElementById('cards').innerHTML=!n?'<div class="card"><div class="v">0</div><div class="l">подій за цими умовами</div></div>':[
    [n,'подій'],[coinsAll.size,'монет'],[p1(mAll.med),'медіана відкату'],[P(mAll.neg,n),'<0%'],[P(mAll.g19,n),'≥'+S.good+'%'],[sg2(mAll.e),'E (умовне)'],
    [white.length+' / '+edge.length+' / '+black.length,'топ / на межі / blacklist']
  ].map(([v,l])=>`<div class="card"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('');

  // таблиця монет
  const q=R.q; let view=rows.filter(r=>!q||r.coin.toUpperCase().includes(q));
  const keyf={rank:r=>r.rank,coin:r=>r.coin,n:r=>r.m.n,med:r=>r.m.med,cv:r=>r.m.cv??99,g0:r=>r.m.g0/r.m.n,g19:r=>r.m.g19/r.m.n,g35:r=>r.m.g35/r.m.n,g50:r=>r.m.g50/r.m.n,neg:r=>r.m.neg/r.m.n,min:r=>r.m.min,e:r=>r.m.e,score:r=>r.m.score,L:r=>r.L?r.L.k/r.L.n:-1,S:r=>r.S?r.S.k/r.S.n:-1,st:r=>({top:0,edge:1,mid:2,bl:3})[r.st]};
  const kf=keyf[sortKey]||keyf.score;
  view.sort((a,b)=>{const x=kf(a),y=kf(b); return (x<y?-1:x>y?1:0)*sortDir || a.rank-b.rank;});
  const topN=+S.top||0; const shown=topN>0?view.slice(0,topN):view;
  document.getElementById('toptitle').textContent=`Топ-лист${topN?' — топ '+topN:''} (${shown.length} з ${rows.length} монет)`;
  const cols=[['rank','#'],['coin','Монета','l'],['n','N'],['med','Медіана'],['cv','CV'],['neg','<0%'],['g0','≥0%'],['g19','≥'+S.good+'%'],['g35','≥35%'],['g50','≥50%'],['min','Мін.'],['e','E'],['L','Long'],['S','Short'],['lev','Плече'],['st','Статус','l']];
  const sd=(x)=>x?`${Math.round(100*x.k/x.n)}% (${x.n})`:'—';
  document.getElementById('coins').innerHTML=
    `<thead><tr>${cols.map(([k,t,cl])=>`<th class="${cl||''} ${sortKey===k?'sorted':''}" data-k="${k}">${t}${sortKey===k?(sortDir<0?' ↓':' ↑'):''}</th>`).join('')}</tr></thead><tbody>`+
    shown.map(r=>`<tr class="click" data-c="${esc(r.coin)}"><td>${r.rank}</td><td class="l"><b>${esc(short(r.coin))}</b></td><td>${r.m.n}</td><td>${p1(r.m.med)}</td><td>${cvs(r.m.cv)}</td><td class="${r.m.neg?'neg':''}">${P(r.m.neg,r.m.n)}</td><td>${P(r.m.g0,r.m.n)}</td><td>${P(r.m.g19,r.m.n)}</td><td>${P(r.m.g35,r.m.n)}</td><td>${P(r.m.g50,r.m.n)}</td><td class="${r.m.min<0?'neg':'pos'}">${sg(r.m.min)}</td><td class="${r.m.e<0?'neg':''}">${sg2(r.m.e)}</td><td>${sd(r.L)}</td><td>${sd(r.S)}</td><td>${esc(r.lev)}</td><td class="l"><span class="st ${r.st}">${STN[r.st]}</span>${r.m.n<10?' <span class="muted" title="мала вибірка">мало</span>':''}</td></tr>`).join('')+'</tbody>';
  document.querySelectorAll('#coins th').forEach(th=>th.onclick=()=>{const k=th.dataset.k; if(k==='lev')return; if(sortKey===k) sortDir=-sortDir; else {sortKey=k; sortDir=(k==='coin'||k==='rank'||k==='cv'||k==='neg'||k==='st')?1:-1;} render();});
  document.querySelectorAll('#coins tr.click').forEach(tr=>tr.onclick=()=>coinModal(tr.dataset.c));
  const small=rows.filter(r=>r.m.n<10&&(r.st==='top'||r.st==='edge')).map(r=>short(r.coin));
  document.getElementById('coinnote').innerHTML=`Монети з N &lt; ${minN} не показані: ${skC} монет, ${skE} подій. Клік по монеті — детальна статистика. Клік по заголовку — сортування.`+
    (small.length?`<br><span class="warn">Мала вибірка (&lt;10 подій), місцю в топі довіряти мало: ${esc(small.join(', '))}.</span>`:'');
  const tot=rows.reduce((s,r)=>s+r.m.n,0)+skE;

  // списки
  const L=(arr,us)=>arr.length?arr.map(r=>us?r.coin:short(r.coin)).join(','):'—';
  document.getElementById('lists').innerHTML=[['Whitelist',white,'top'],['На межі',edge,'edge'],['Blacklist',black,'bl']].map(([t,a,cl],i)=>
    `<div class="list"><div class="head"><span class="st ${cl}">${t} — ${a.length}</span><span><button data-l="${i}" data-u="1">Копіювати з USDT</button> <button data-l="${i}" data-u="0">без USDT</button></span></div><pre>${esc(L(a,true))}</pre></div>`).join('')+
    `<div class="note" style="grid-column:1/-1">Правила: топ = ≥${S.good}% ≥ ${S.top19}%, медіана ≥ ${S.topMed}%, CV ≤ ${S.topCv}, мінусів ≤ ${S.topNeg}%. «На межі» — топ, але хоч один показник близько до порогу. Blacklist = ≥${S.good}% &lt; ${S.bl19}% або медіана &lt; ${S.blMed}% або CV &gt; ${S.blCv} або мінімум &lt; ${S.blMin}% або мінусів ≥ ${S.blNeg}%. Списки рахуються з усіх монет з N ≥ ${minN}, незалежно від «Топ» і пошуку.</div>`;
  const LS=[white,edge,black];
  document.querySelectorAll('#lists button').forEach(b=>b.onclick=()=>copy(L(LS[+b.dataset.l],b.dataset.u==='1')));

  // ринок
  const mk=[]; let checks=[];
  if(n){
    const g=[['Усі',ev]]; const sides={}; for(const e of ev)(sides[e.side]||(sides[e.side]=[])).push(e);
    for(const k of Object.keys(sides).sort()) g.push([k,sides[k]]);
    mk.push(tbl('Загалом і по сторонах','Зріз',g)); checks.push(['сторони',Object.values(sides).reduce((s,a)=>s+a.length,0)]);
    const edges=bucketEdges(TAB,smin,smax); const bg=[];
    for(let i=0;i<edges.length;i++){ const a=edges[i], b=i+1<edges.length?edges[i+1]:(smax>0?smax:Infinity);
      bg.push([(b===Infinity?`${a}%+`:`${a}–${b}%`).replace(/\./g,','),ev.filter(e=>e.sp>=a&&e.sp<b)]); }
    mk.push(tbl('За розміром спайку','Спайк',bg)); checks.push(['розмір спайку',bg.reduce((s,x)=>s+x[1].length,0)]);
    const lv={}; for(const e of ev)(lv[e.lev]||(lv[e.lev]=[])).push(e);
    const lg=[]; const sm=[],smn=[];
    for(const k of Object.keys(lv).sort((a,b)=>(+a||1e9)-(+b||1e9))){ if(lv[k].length<20){smn.push((k||'?')+'x');sm.push(...lv[k]);} else lg.push([(k||'?')+'x',lv[k]]); }
    if(sm.length) lg.push([smn.join(', ')+' (мало даних)',sm]);
    mk.push(tbl('За плечем (групи <20 подій об\'єднані)','Плече',lg)); checks.push(['плече',lg.reduce((s,x)=>s+x[1].length,0)]);
    const monthly=(last-first)/DAY>31; const by=new Map();
    for(const e of ev){ const d=new Date(e.t); const k=monthly?d.getUTCFullYear()*100+d.getUTCMonth()+1:Math.floor(e.t/DAY); (by.get(k)||by.set(k,[]).get(k)).push(e); }
    const keys=[...by.keys()].sort((a,b)=>a-b);
    const dg=keys.map((k,i)=>{ let nm;
      if(monthly) nm=`${pad(k%100)}.${Math.floor(k/100)}`;
      else { nm=fD(k*DAY); if(i===0&&days>0&&k===Math.floor(from/DAY)) nm+=` (з ${fHM(from)})`; if(i===keys.length-1) nm+=` (до ${fHM(last)})`; }
      return [nm,by.get(k)]; });
    mk.push(tbl(monthly?'По місяцях':'По днях',monthly?'Місяць':'День',dg)); checks.push([monthly?'місяці':'дні',dg.reduce((s,x)=>s+x[1].length,0)]);
  }
  checks.push(['монети + відкинуті',tot]);
  const ok=checks.every(x=>x[1]===n);
  mk.push(`<div><h2>Перевірка сум (має бути ${n})</h2>${checks.map(([k,v])=>`<div class="${v===n?'':'neg'}">${k}: ${v} — ${v===n?'зійшлося':'НЕ ЗІЙШЛОСЯ'}</div>`).join('')}</div>`);
  document.getElementById('market').innerHTML=mk.join('');
  if(!ok) toast('Суми не зійшлися — перевір файл');
}
function tbl(title,first,groups){
  let i=0;
  return `<div><h2>${title}</h2><div class="tw"><table><thead><tr><th>#</th><th class="l">${first}</th><th>N</th><th>Медіана</th><th>CV</th><th>&lt;0%</th><th>≥0%</th><th>≥${S.good}%</th><th>≥35%</th><th>≥50%</th><th>Мін.</th><th>E</th></tr></thead><tbody>`+
    groups.filter(g=>g[1].length).map(([nm,es])=>{const m=metrics(es.map(e=>e.rb)); i++;
      return `<tr><td>${i}</td><td class="l">${esc(nm)}</td><td>${m.n}</td><td>${p1(m.med)}</td><td>${cvs(m.cv)}</td><td class="${m.neg?'neg':''}">${P(m.neg,m.n)}</td><td>${P(m.g0,m.n)}</td><td>${P(m.g19,m.n)}</td><td>${P(m.g35,m.n)}</td><td>${P(m.g50,m.n)}</td><td class="${m.min<0?'neg':'pos'}">${sg(m.min)}</td><td class="${m.e<0?'neg':''}">${sg2(m.e)}</td></tr>`;}).join('')+'</tbody></table></div></div>';
}

function coinModal(coin){
  window._lastCoin=coin;
  const {ev,smin,smax,last,from,days}=LAST; const es=ev.filter(e=>e.coin===coin).sort((a,b)=>b.t-a.t);
  if(!es.length) return;
  const m=metrics(es.map(e=>e.rb));
  const g=[['Усі',es]]; for(const s of ['Long','Short']){ const x=es.filter(e=>e.side===s); if(x.length) g.push([s,x]); }
  const edges=bucketEdges(TAB,smin,smax); const bg=[];
  for(let i=0;i<edges.length;i++){ const a=edges[i], b=i+1<edges.length?edges[i+1]:(smax>0?smax:Infinity); bg.push([(b===Infinity?`${a}%+`:`${a}–${b}%`).replace(/\./g,','),es.filter(e=>e.sp>=a&&e.sp<b)]); }
  const first=es[es.length-1].t; const monthly=(last-first)/DAY>31; const by=new Map();
  for(const e of es){ const d=new Date(e.t); const k=monthly?d.getUTCFullYear()*100+d.getUTCMonth()+1:Math.floor(e.t/DAY); (by.get(k)||by.set(k,[]).get(k)).push(e); }
  const dg=[...by.keys()].sort((a,b)=>a-b).map(k=>[monthly?`${pad(k%100)}.${Math.floor(k/100)}`:fD(k*DAY),by.get(k)]);
  const worst=es.slice().sort((a,b)=>a.rb-b.rb).slice(0,5);
  const evRows=a=>a.map(e=>`<tr><td class="l">${fDT(e.t)}</td><td class="l">${e.side}</td><td>${e.lev}x</td><td>${p1(e.sp)}</td><td class="${e.rb<0?'neg':e.rb>=S.good?'pos':''}">${p1(e.rb)}</td></tr>`).join('');
  document.getElementById('mbox').innerHTML=
    `<div style="display:flex;justify-content:space-between;align-items:center"><h2 style="font-size:18px">${esc(coin)} · ${esc(TAB)} · спайки ${(smax>0?`${smin}–${smax}%`:`від ${smin}%`).replace(/\./g,',')} · ${days>0?days+' дн.':'весь період'}</h2><button id="mclose">Закрити</button></div>`+
    `<div class="cards">${[[m.n,'подій'],[p1(m.med),'медіана'],[cvs(m.cv),'CV'],[P(m.neg,m.n),'<0%'],[P(m.g0,m.n),'≥0%'],[P(m.g19,m.n),'≥'+S.good+'%'],[P(m.g35,m.n),'≥35%'],[P(m.g50,m.n),'≥50%'],[sg(m.min),'мінімум'],[sg2(m.e),'E'],[STN[status(m)],'статус']].map(([v,l])=>`<div class="card"><div class="v">${v}</div><div class="l">${l}</div></div>`).join('')}</div>`+
    `<div class="grid2">${tbl('По сторонах','Зріз',g)}${tbl('За розміром спайку','Спайк',bg)}${tbl(monthly?'По місяцях':'По днях',monthly?'Місяць':'День',dg)}`+
    `<div><h2>Найгірші події</h2><div class="tw"><table><thead><tr><th class="l">Дата</th><th class="l">Сторона</th><th>Плече</th><th>Спайк</th><th>Відкат</th></tr></thead><tbody>${evRows(worst)}</tbody></table></div>`+
    `<h2 style="margin-top:12px">Останні 15 подій</h2><div class="tw"><table><thead><tr><th class="l">Дата</th><th class="l">Сторона</th><th>Плече</th><th>Спайк</th><th>Відкат</th></tr></thead><tbody>${evRows(es.slice(0,15))}</tbody></table></div></div></div>`+
    `<div style="margin-top:12px"><button id="mcopy">Копіювати рядок монети (Markdown)</button></div>`;
  const md=document.getElementById('modal'); md.style.display='flex';
  document.getElementById('mclose').onclick=()=>{md.style.display='none';window._lastCoin=null;};
  md.onclick=e=>{if(e.target===md) md.style.display='none';};
  document.getElementById('mcopy').onclick=()=>{
    const sd=s=>{const x=es.filter(e=>e.side===s);return x.length?`${Math.round(100*x.filter(e=>e.rb>=S.good).length/x.length)}% (${x.length})`:'—';};
    copy(`| # | Монета | N | Медіана | CV | <0% | ≥0% | ≥${S.good}% | ≥35% | ≥50% | Мін. | E | Long | Short | Плече | Статус |\n|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|\n| 1 | ${short(coin)} | ${m.n} | ${p1(m.med)} | ${cvs(m.cv)} | ${P(m.neg,m.n)} | ${P(m.g0,m.n)} | ${P(m.g19,m.n)} | ${P(m.g35,m.n)} | ${P(m.g50,m.n)} | ${sg(m.min)} | ${sg2(m.e)} | ${sd('Long')} | ${sd('Short')} | ${[...new Set(es.map(e=>e.lev))].join(', ')} | ${STN[status(m)]} |`);
  };
}

function copyMarkdown(){
  const trs=[...document.querySelectorAll('#coins tbody tr')];
  const head=`| # | Монета | N | Медіана | CV | <0% | ≥0% | ≥${S.good}% | ≥35% | ≥50% | Мін. | E | Long | Short | Плече | Статус |\n|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|---|`;
  copy(head+'\n'+trs.map(tr=>'| '+[...tr.children].map(td=>td.textContent.trim().replace(/\s*мало$/,' (мало)')).join(' | ')+' |').join('\n'));
}
function copy(t){
  const done=()=>toast('Скопійовано');
  if(navigator.clipboard&&window.isSecureContext) navigator.clipboard.writeText(t).then(done,()=>fallback(t));
  else fallback(t);
  function fallback(t){ const ta=document.createElement('textarea'); ta.value=t; document.body.appendChild(ta); ta.select(); try{document.execCommand('copy');done();}catch(e){prompt('Скопіюй вручну:',t);} ta.remove(); }
}
function toast(t){ const el=document.getElementById('toast'); el.textContent=t; el.style.display='block'; clearTimeout(el._t); el._t=setTimeout(()=>el.style.display='none',1600); }

// ================= Інтерфейс =================
function buttons(id,items,cur,onPick){
  const box=document.getElementById(id); box.querySelectorAll('button').forEach(b=>b.remove());
  for(const [v,t] of items){ const b=document.createElement('button'); b.textContent=t; if(String(v)===String(cur)) b.classList.add('on'); b.onclick=()=>onPick(v); box.appendChild(b); }
}
function syncControls(){
  document.getElementById('smin').value=S.spike[TAB]??spikeDef(TAB);
  document.getElementById('smax').value=S['smax_'+TAB]||'';
  document.getElementById('days').value=S.days;
  document.getElementById('minn').value=S.minN;
  buttons('periods',PERIODS,S.days,v=>{S.days=v;saveSettings();syncControls();render();});
  buttons('sides',SIDES,S.side,v=>{S.side=v;saveSettings();syncControls();render();});
  buttons('tops',TOPS,S.top,v=>{S.top=v;saveSettings();syncControls();render();});
}
function buildSettings(){
  const box=document.getElementById('settings');
  box.innerHTML=SET_FIELDS.map(([k,t,ty])=>`<label>${t}<input data-k="${k}" type="${ty}" ${ty==='number'?'step="any"':''} value="${esc(S[k])}" style="${ty==='text'?'width:200px':''}"></label>`).join('');
  box.querySelectorAll('input').forEach(inp=>inp.onchange=()=>{ const k=inp.dataset.k; S[k]=inp.type==='number'?+inp.value:inp.value; saveSettings(); render(); });
}
function buildTabs(){
  const nav=document.getElementById('tabs'); nav.innerHTML='';
  for(const t of TABS){ const b=document.createElement('button'); b.textContent=t.replace(/_Events$/i,'').replace(/^Events$/i,'Events'); b.title=t; if(t===TAB) b.classList.add('on'); b.onclick=()=>selectTab(t); nav.appendChild(b); }
}
async function selectTab(t){
  TAB=t; buildTabs(); syncControls();
  if(!CACHE[t]){ document.getElementById('sum').textContent='Читаю аркуш '+t+'...'; document.getElementById('cards').innerHTML=''; document.getElementById('coins').innerHTML=''; document.getElementById('market').innerHTML='';
    await new Promise(r=>setTimeout(r,20));
    try{ await loadTab(t); }catch(e){ document.getElementById('sum').innerHTML='<span class="neg">Помилка: '+esc(e.message)+'</span>'; return; } }
  render();
}
async function openBuffer(buf,name,silent){
  if(!silent) document.getElementById('fname').textContent='Відкриваю '+name+'...';
  const book=await openBook(buf);
  const names=book.sheets.map(s=>s.name);
  const tabs=names.filter(n=>/events$/i.test(n.trim())&&!/^arhiv/i.test(n));
  if(!tabs.length) throw new Error('У файлі немає аркушів *Events. Аркуші: '+names.join(', '));
  tabs.sort((a,b)=>(isMain(b)-isMain(a)));
  BOOK=book; FILE=name; CACHE={}; TABS=tabs;
  document.getElementById('fname').textContent=name;
  document.getElementById('empty').style.display='none'; document.getElementById('app').style.display='';
  const modalOpen=document.getElementById('modal').style.display==='flex';
  await selectTab(TABS.includes(TAB)?TAB:TABS[0]);
  if(modalOpen&&LAST&&window._lastCoin) coinModal(window._lastCoin);
}
async function openFile(f){
  if(!f) return;
  stopLive();
  try{ await openBuffer(await f.arrayBuffer(),f.name); }
  catch(e){ document.getElementById('fname').innerHTML='<span class="neg">Помилка: '+esc(e.message)+'</span>'; }
}

// ---------- Живе оновлення: файл з ПК або Google Sheets ----------
let SRC=null, liveTimer=null, watchTimer=null, busy=false;
const hm=()=>{const d=new Date();return pad(d.getHours())+':'+pad(d.getMinutes())+':'+pad(d.getSeconds());};
function setLive(txt,bad){ const el=document.getElementById('livest'); el.innerHTML=bad?'<span class="neg">'+esc(txt)+'</span>':esc(txt); }
function stopLive(){ clearInterval(liveTimer); clearInterval(watchTimer); SRC=null; document.getElementById('live').style.display='none'; }
function startLive(){
  document.getElementById('live').style.display='';
  clearInterval(liveTimer); clearInterval(watchTimer);
  const min=+document.getElementById('every').value||5;
  liveTimer=setInterval(()=>refreshNow(false),min*60000);
  if(SRC&&SRC.type==='handle') watchTimer=setInterval(checkFileChanged,15000);
}
async function checkFileChanged(){
  if(!SRC||SRC.type!=='handle'||busy) return;
  try{ const f=await SRC.handle.getFile(); if(f.lastModified!==SRC.lastMod) refreshNow(false); }catch(e){}
}
function gsToUrl(u){
  u=u.trim(); if(!u) return null;
  if(/\/pub(html)?\b|output=/.test(u)){
    u=u.replace('/pubhtml','/pub');
    if(/[?&]output=/.test(u)) u=u.replace(/([?&])output=[^&]*/,'$1output=xlsx'); else u+=(u.includes('?')?'&':'?')+'output=xlsx';
    u=u.replace(/[?&](gid|single)=[^&]*/g,'');
    return u;
  }
  const m=/\/spreadsheets\/d\/([a-zA-Z0-9_-]{20,})/.exec(u);
  if(m) return 'https://docs.google.com/spreadsheets/d/'+m[1]+'/export?format=xlsx';
  return u;
}
async function refreshNow(manual){
  if(!SRC||busy) return; busy=true;
  setLive('оновлюю...');
  try{
    let buf,name,stamp;
    if(SRC.type==='handle'){
      const f=await SRC.handle.getFile();
      if(!manual&&f.lastModified===SRC.lastMod&&BOOK){ setLive('без змін, перевірено о '+hm()); busy=false; return; }
      buf=await f.arrayBuffer(); name=f.name; SRC.lastMod=f.lastModified;
      stamp='файл змінено '+new Date(f.lastModified).toLocaleString('uk-UA');
    } else {
      const sep=SRC.url.includes('?')?'&':'?';
      let r;
      try{ r=await fetch(SRC.url+sep+'_t='+Date.now(),{cache:'no-store'}); }
      catch(e){ throw new Error('не вдалося завантажити таблицю (немає інтернету або доступу). Перевір посилання; якщо не допоможе — опублікуй таблицю як .xlsx (Файл → Поділитися → Опублікувати в Інтернеті) і встав те посилання.'); }
      if(!r.ok) throw new Error('Google відповів '+r.status+'. Перевір, що таблиця опублікована в Інтернеті як .xlsx.');
      buf=await r.arrayBuffer();
      const head=new Uint8Array(buf.slice(0,2));
      if(!(head[0]===0x50&&head[1]===0x4b)) throw new Error('за посиланням не xlsx (схоже, сторінка входу Google). Увійди в Google в цьому Chrome під акаунтом, що має доступ до таблиці, або опублікуй таблицю як .xlsx і встав те посилання.');
      name='Moon_Scanner (Google Sheets)'; stamp='завантажено з Google';
    }
    await openBuffer(buf,name,true);
    const lastEv=LAST?fDT(LAST.last):'—';
    setLive('оновлено о '+hm()+' · '+stamp+' · остання подія '+lastEv);
  }catch(e){ setLive('Помилка: '+e.message,true); if(!BOOK) document.getElementById('fname').innerHTML='<span class="neg">'+esc(e.message)+'</span>'; }
  busy=false;
}
document.getElementById('linkfile').onclick=async()=>{
  if(!window.showOpenFilePicker){ toast('Цей браузер не вміє підв\'язувати файл'); return; }
  try{
    const [h]=await window.showOpenFilePicker({types:[{description:'Excel',accept:{'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}],multiple:false});
    SRC={type:'handle',handle:h,lastMod:null}; startLive(); await refreshNow(true);
  }catch(e){ if(e.name!=='AbortError') toast('Не вдалося: '+e.message); }
};
document.getElementById('gsbtn').onclick=()=>{ const b=document.getElementById('gsbox'); b.style.display=b.style.display==='none'?'':'none'; try{ document.getElementById('gsurl').value=VAULT.gs||''; }catch(e){} };
document.getElementById('gsclose').onclick=()=>document.getElementById('gsbox').style.display='none';
document.getElementById('gsgo').onclick=async()=>{
  const raw=document.getElementById('gsurl').value; const url=gsToUrl(raw); if(!url){toast('Встав посилання');return;}
  VAULT.gs=raw; persist();
  SRC={type:'url',url}; startLive(); await refreshNow(true);
  if(BOOK&&FILE==='Google Sheets') document.getElementById('gsbox').style.display='none';
};
document.getElementById('refresh').onclick=()=>refreshNow(true);
document.getElementById('unlink').onclick=()=>{ stopLive(); toast('Відв\'язано'); };
document.getElementById('every').onchange=()=>{ if(SRC) startLive(); };
// автопідключення Google Sheets, якщо посилання вже збережене
document.getElementById('file').onchange=e=>openFile(e.target.files[0]);
const drop=document.body;
drop.addEventListener('dragover',e=>{e.preventDefault();document.getElementById('drop').classList.add('over');});
drop.addEventListener('dragleave',()=>document.getElementById('drop').classList.remove('over'));
drop.addEventListener('drop',e=>{e.preventDefault();document.getElementById('drop').classList.remove('over');openFile(e.dataTransfer.files[0]);});
let deb; const later=f=>{clearTimeout(deb);deb=setTimeout(f,250);};
document.getElementById('smin').oninput=e=>later(()=>{const v=parseFloat(String(e.target.value).replace(',','.')); if(Number.isFinite(v)){S.spike[TAB]=v;saveSettings();render();}});
document.getElementById('smax').oninput=e=>later(()=>{const v=parseFloat(String(e.target.value).replace(',','.')); S['smax_'+TAB]=Number.isFinite(v)?v:0; saveSettings(); render();});
document.getElementById('days').oninput=e=>later(()=>{const v=parseInt(e.target.value,10); if(Number.isFinite(v)&&v>=0){S.days=v;saveSettings();buttons('periods',PERIODS,S.days,v=>{S.days=v;saveSettings();syncControls();render();});render();}});
document.getElementById('minn').oninput=e=>later(()=>{const v=parseInt(e.target.value,10); if(v>=1){S.minN=v;saveSettings();render();}});
document.getElementById('coinq').oninput=()=>later(render);
document.getElementById('copymd').onclick=copyMarkdown;
document.getElementById('reset').onclick=()=>{S=structuredClone(DEF);saveSettings();buildSettings();syncControls();render();};
document.addEventListener('keydown',e=>{if(e.key==='Escape') document.getElementById('modal').style.display='none';});
if(typeof DecompressionStream==='undefined'){ document.getElementById('empty').innerHTML='<span class="neg">Цей браузер застарий: потрібен Chrome, Edge або Firefox свіжої версії.</span>'; }


// ================= Вхід за паролем (сайт) =================
let VAULT={settings:null,gs:''}, KEY=null, LOGCFG=null;
const te=new TextEncoder(), td=new TextDecoder();
const b64=u=>btoa(String.fromCharCode(...u));
const unb64=s=>Uint8Array.from(atob(s),c=>c.charCodeAt(0));
async function deriveKey(pw,salt){
  const base=await crypto.subtle.importKey('raw',te.encode(pw),'PBKDF2',false,['deriveKey']);
  return crypto.subtle.deriveKey({name:'PBKDF2',salt,iterations:310000,hash:'SHA-256'},base,{name:'AES-GCM',length:256},true,['encrypt','decrypt']);
}
async function open_(key,box){ const pt=await crypto.subtle.decrypt({name:'AES-GCM',iv:unb64(box.iv)},key,unb64(box.ct)); return JSON.parse(td.decode(pt)); }
function persist(){ try{ localStorage.setItem('moonstats_site',JSON.stringify(VAULT.settings)); }catch(e){} }
const KEYSTORE='moonstats_site_key';
function storedKey(){ try{ return sessionStorage.getItem(KEYSTORE)||localStorage.getItem(KEYSTORE); }catch(e){ return null; } }
function forgetKey(){ try{ sessionStorage.removeItem(KEYSTORE); localStorage.removeItem(KEYSTORE); }catch(e){} }
async function tryKey(key){
  const cfg=await open_(key,window.MOON_CFG);
  if(!cfg||!cfg.url) throw new Error('bad');
  KEY=key; VAULT.gs=cfg.url; LOGCFG=cfg.log&&cfg.k?{u:cfg.log,k:cfg.k}:null;
  try{ VAULT.settings=JSON.parse(localStorage.getItem('moonstats_site')||'null'); }catch(e){ VAULT.settings=null; }
}
const NAMEKEY='moonstats_name';
function getName(){ try{ return localStorage.getItem(NAMEKEY)||''; }catch(e){ return ''; } }
function tabId(){ try{ let t=sessionStorage.getItem('moonstats_tab'); if(!t){ t=Math.random().toString(36).slice(2,10); sessionStorage.setItem('moonstats_tab',t); } return t; }catch(e){ return 'x'; } }
function device(){
  const ua=navigator.userAgent;
  const os=/Android/i.test(ua)?'Android':/iPhone|iPad|iPod/i.test(ua)?'iOS':/Windows/i.test(ua)?'Windows':/Mac OS X/i.test(ua)?'Mac':/Linux/i.test(ua)?'Linux':'?';
  const mob=/Mobi|Android|iPhone|iPad/i.test(ua)?'телефон':'ПК';
  const m=ua.match(/(Edg|OPR|YaBrowser|Firefox|Chrome|Version)\/(\d+)/);
  const br=m?(({Edg:'Edge',OPR:'Opera',YaBrowser:'Yandex',Version:'Safari'})[m[1]]||m[1])+' '+m[2]:'?';
  return {d:mob+', '+os, b:br, s:screen.width+'x'+screen.height};
}
function ping(ev){
  if(!LOGCFG) return;
  try{
    const q=new URLSearchParams(Object.assign({k:LOGCFG.k,n:getName()||'?',e:ev,i:tabId()},device()));
    fetch(LOGCFG.u+'?'+q.toString(),{mode:'no-cors',cache:'no-store'}).catch(()=>{});
  }catch(e){}
}
let hbT=null;
function startHeartbeat(){
  clearInterval(hbT);
  hbT=setInterval(()=>{ if(document.visibilityState==='visible') ping('онлайн'); },4*60000);
  document.addEventListener('visibilitychange',()=>{ if(document.visibilityState==='visible') ping('онлайн'); });
}
function afterUnlock(ev){
  document.getElementById('lock').style.display='none';
  document.getElementById('shell').style.display='';
  S=loadSettings(); buildSettings();
  SRC={type:'url',url:gsToUrl(VAULT.gs)}; startLive(); refreshNow(true);
  const who=document.getElementById('whoami'); if(who) who.textContent=getName()||'?';
  let first=false; try{ first=!sessionStorage.getItem('moonstats_logged'); sessionStorage.setItem('moonstats_logged','1'); }catch(e){}
  if(ev==='вхід'||first) ping(ev==='вхід'?'вхід':'повернення'); else ping('онлайн');
  startHeartbeat();
}
async function initLock(){
  const L=document.getElementById('lock'), msg=document.getElementById('lmsg');
  document.getElementById('lpw2').style.display='none';
  if(!window.MOON_CFG||!window.MOON_CFG.ct){
    L.style.display='flex';
    document.getElementById('ltitle').textContent='Сайт ще не налаштовано: поклади поруч файл config.js, створений на сторінці setup.html.';
    document.getElementById('lform').querySelectorAll('input,button,label').forEach(e=>e.style.display='none');
    return;
  }
  const SALT=unb64(window.MOON_CFG.salt);
  const k=storedKey();
  if(k){ try{ await tryKey(await crypto.subtle.importKey('raw',unb64(k),{name:'AES-GCM'},true,['encrypt','decrypt'])); return afterUnlock('повернення'); }catch(e){ forgetKey(); } }
  L.style.display='flex'; document.getElementById('ltitle').textContent='Введи своє ім\'я і пароль';
  const nm=document.getElementById('lname'); nm.style.display=''; nm.value=getName(); (nm.value?document.getElementById('lpw'):nm).focus();
  document.getElementById('lform').onsubmit=async ev=>{
    ev.preventDefault(); msg.textContent='Перевіряю...';
    const pw=document.getElementById('lpw').value;
    const nmv=document.getElementById('lname').value.trim();
    if(!nmv){ msg.textContent='Вкажи ім\'я'; document.getElementById('lname').focus(); return; }
    try{ localStorage.setItem(NAMEKEY,nmv.slice(0,40)); }catch(e){}
    try{ await tryKey(await deriveKey(pw,SALT)); }
    catch(e){ msg.textContent='Невірний пароль'; document.getElementById('lpw').select(); return; }
    const raw=b64(new Uint8Array(await crypto.subtle.exportKey('raw',KEY)));
    try{ sessionStorage.setItem(KEYSTORE,raw); if(document.getElementById('lrem').checked) localStorage.setItem(KEYSTORE,raw); }catch(e){}
    document.getElementById('lpw').value=''; msg.textContent='';
    afterUnlock('вхід');
  };
}
document.getElementById('lockbtn').onclick=()=>{ forgetKey(); location.reload(); };
initLock();
