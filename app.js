/* =====================================================================
   خدمة براعم — الموقع
   - قاعدة بيانات الأطفال: شيت «قاعدة البيانات» في ملف الإكسيل (على GitHub).
   - الحضور والملاحظات: Google Sheet منفصل، الموقع بيقرا منه وبيكتب عليه
     عن طريق Google Apps Script (ملف Code.gs).
   - كل الحسابات (النسب، الافتقاد، الإحصائيات) بيعملها الموقع نفسه.
   ===================================================================== */
const FILE='Bara3em_System_Final.xlsx';

/* ---------------- الإعدادات (كانت في شيت «الإعدادات» في الإكسيل) ---------------- */
const CONFIG={
 api:'https://script.google.com/macros/s/AKfycbyy8W0KpADyXaZnFK0j8F4X5XW2efz_YGJb4bnmCvQ8txw4xTGAdNBsfYFP3F5TZzk9/exec',                       // رابط Web App بتاع Google Apps Script (بينتهي بـ /exec)
 key:'Mina_Bra3em',                       // نفس كلمة KEY اللي في Code.gs
 classes:['سوريال','غبريال','رفائيل','ميخائيل'],            // الفصول (الملائكة) — ماتغيّرش الأسماء بعد تسجيل الأطفال
 years:['سنة أولى','سنة تانية','سنة تالتة','سنة رابعة','سنة خامسة','سنة سادسة'],
 start:'2026-10-01',           // بداية العام الخدمي (YYYY-MM-DD)
 months:12,                    // مدة العام الخدمي بالشهور
 day:'الجمعة',                 // يوم الاجتماع
 th:{ex:0.85,good:0.7,streak:3}, // «ممتاز» من / «جيد» من / غياب متتالي = افتقاد ضروري من
 cap:30,                       // سعة الفصل (للعرض بس)
 extra:['2026-11-04']                      // اجتماعات إضافية خارج اليوم الثابت، مثال: ['2026-12-24']
};
const DAYS=['الأحد','الاثنين','الثلاثاء','الأربعاء','الخميس','الجمعة','السبت'];

const $=(s,e=document)=>e.querySelector(s),$$=(s,e=document)=>[...e.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const ic=n=>`<svg class="ic" aria-hidden="true"><use href="#i-${n}"/></svg>`;
// [أيقونة، الوصف، اللون]
const TIER=[['clock','لسه مفيش حضور','var(--none)'],['star','ممتاز','var(--ok)'],['checkc','جيد','var(--good)'],['alert','محتاج متابعة','var(--warn)'],['phone','افتقاده ضروري','var(--bad)'],['ban','غير نشط','var(--off)']];
const CC=['var(--c1)','var(--c2)','var(--c3)','var(--c4)','var(--c5)','var(--c6)'];
const NAV=[['home','home','الرئيسية'],['kids','list','الأطفال'],['class','users','الفصول'],['att','calendar','الحضور'],['rec','check','تسجيل'],['follow','phone','الافتقاد'],['settings','settings','الإعدادات',1],['guide','book','الدليل',1]];
const cell=(ws,r,c)=>{const x=ws&&ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:undefined};
const iso2d=s=>{const [y,m,d]=String(s).split('-').map(Number);return new Date(Date.UTC(y,m-1,d))};
const addM=(d,n)=>{const t=new Date(Date.UTC(d.getUTCFullYear(),d.getUTCMonth()+n,1)),last=new Date(Date.UTC(t.getUTCFullYear(),t.getUTCMonth()+1,0)).getUTCDate();t.setUTCDate(Math.min(d.getUTCDate(),last));return t};
const fd=d=>d.getUTCDate()+'/'+(d.getUTCMonth()+1),fdl=d=>fd(d)+'/'+d.getUTCFullYear();
const ph=v=>{if(v==null||v==='')return'';let s=String(v).replace(/\D/g,'');if(s.length===10&&s[0]!=='0')s='0'+s;return s};
const wa=p=>'https://wa.me/20'+p.slice(-10),pc=p=>p==null?'—':Math.round(p*100)+'%';
let D,view=$('#view'),LM='',lastFocus=null,RAW,RC={c:'',j:-1};
const dk=d=>d.toISOString().slice(0,10);
const noteDate=iso=>{const [y,m,d]=String(iso).split('-');return +d+'/'+ +m+'/'+y};
const uid=()=>Date.now().toString(36)+Math.random().toString(36).slice(2,6);

/* ---------------- شيت الحضور على جوجل: القراءة والكتابة ---------------- */
// SH = آخر نسخة اتقرت من الشيت | SS = حالة الاتصال | Q = تسجيلات لسه ما اتبعتتش (محفوظة على الجهاز)
const LSQ='b3_q',LSH='b3_sh';
let SH={m:{},has:{},notes:[]},SS='off',SERR='',LASTF=0,Q=[],flushing=false;
const emptySH=()=>({m:{},has:{},notes:[]});
const loadQ=()=>{try{const a=JSON.parse(localStorage.getItem(LSQ));return Array.isArray(a)?a:[]}catch(e){return[]}};
const saveQ=()=>{try{localStorage.setItem(LSQ,JSON.stringify(Q.map(o=>{const c={...o};delete c.fl;return c})))}catch(e){}};
const loadSH=()=>{try{const x=JSON.parse(localStorage.getItem(LSH));return x&&x.m?x:null}catch(e){return null}};
const saveSH=()=>{try{localStorage.setItem(LSH,JSON.stringify(SH))}catch(e){}};
const normSheet=j=>{const m={},has={};
 for(const c in j.tabs||{}){const t=j.tabs[c];has[c]={};
  t.rows.forEach(r=>{const id=r[0];has[c][id]=r[1];const v={};t.dates.forEach((d,i)=>{if(d&&r[2+i])v[d]=r[2+i]});(m[id]=m[id]||{})[c]={v}})}
 return{m,has,notes:(j.notes||[]).map(n=>({...n,ts:+n.ts||0}))}};
// قيمة الحضور من الشيت (فصل الطفل الحالي أولًا، ولو مفيش بنبص في أي فصل كان فيه قبل كده)
function sv(id,cls,date){const m=SH.m[id];if(!m)return'';const a=m[cls]&&m[cls].v[date];if(a)return a;for(const c in m){const x=m[c].v[date];if(x)return x}return''}
function applyLocal(o){
 if(o.t==='att'){const e=(SH.m[o.id]=SH.m[o.id]||{}),x=(e[o.cls]=e[o.cls]||{v:{}});if(o.v)x.v[o.date]=o.v;else delete x.v[o.date];const h=(SH.has[o.cls]=SH.has[o.cls]||{});if(h[o.id]===undefined)h[o.id]=o.name}
 else if(o.t==='note'){if(!SH.notes.some(n=>n.nid===o.nid))SH.notes.push({nid:o.nid,id:o.id,cls:o.cls,name:o.name,date:o.date,text:o.text,ts:o.ts})}
 else if(o.t==='del'){SH.notes=SH.notes.filter(n=>n.nid!==o.nid)}}
const post=async body=>{const ac=new AbortController(),tm=setTimeout(()=>ac.abort(),30000);
 try{const r=await fetch(CONFIG.api,{method:'POST',body:JSON.stringify({key:CONFIG.key,...body}),signal:ac.signal});return await r.json()}finally{clearTimeout(tm)}};
async function loadSheet(){
 SERR='';
 if(!CONFIG.api){SH=emptySH();SS='off';return}
 const ac=new AbortController(),tm=setTimeout(()=>ac.abort(),20000);
 try{const r=await fetch(CONFIG.api+(CONFIG.api.includes('?')?'&':'?')+'key='+encodeURIComponent(CONFIG.key)+'&_='+Date.now(),{signal:ac.signal});
  const j=await r.json();if(!j||!j.ok)throw Error((j&&j.err)||'bad');
  SH=normSheet(j);SS='ok';LASTF=Date.now();saveSH()}
 catch(e){console.warn('sheet',e);SERR=e.message;const c=loadSH();if(c){SH=c;SS='cache'}else{SH=emptySH();SS='fail'}}
 finally{clearTimeout(tm)}}
// تسجيلات الحضور اللي لسه ما اتبعتتش (آخر واحدة لكل طفل/تاريخ هي اللي بتتحسب)
const overlay=()=>{const o={};Q.forEach(x=>{if(x.t==='att')o[x.id+'|'+x.date]=x.v});return o};
function allNotes(){const del=new Set(Q.filter(o=>o.t==='del').map(o=>o.nid)),have=new Set(SH.notes.map(n=>n.nid));
 return [...SH.notes,...Q.filter(o=>o.t==='note'&&!have.has(o.nid)).map(o=>({nid:o.nid,id:o.id,date:o.date,text:o.text,ts:o.ts,pd:1}))].filter(n=>!del.has(n.nid)).sort((a,b)=>String(a.date).localeCompare(String(b.date))||(a.ts||0)-(b.ts||0))}
function qAtt(k,d,v){Q=Q.filter(o=>!(o.t==='att'&&o.id===k.id&&o.date===d&&!o.fl));
 if(sv(k.id,k.cls,d)===v&&!Q.some(o=>o.t==='att'&&o.id===k.id&&o.date===d))return;
 Q.push({t:'att',id:k.id,cls:k.cls,name:k.name,date:d,v,ts:Date.now()})}
async function flush(){
 if(flushing||!Q.length||!CONFIG.api)return;flushing=true;setSt();
 const ops=Q.slice(0,40);ops.forEach(o=>o.fl=1);
 try{const r=await post({ops});if(!r||!r.ok)throw Error((r&&r.err)||'bad');
  r.results.forEach((x,i)=>{if(x.ok)applyLocal(ops[i])});
  Q=Q.filter(o=>!ops.includes(o));saveQ();saveSH();SERR='';flushing=false;build();setSt();const e=$('#rl');if(e)renderRec();
  if(Q.length)return flush()}
 catch(e){console.warn('flush',e);SERR=e.message;ops.forEach(o=>delete o.fl);flushing=false;setSt()}}
// بيضيف الأطفال الجداد (اللي اتضافوا في الإكسيل) لتابات الحضور في جوجل شيت
async function syncKids(){
 if(!CONFIG.api||SS!=='ok'||!D)return;
 const need=D.kids.filter(k=>CONFIG.classes.includes(k.cls)&&String((SH.has[k.cls]||{})[k.id]??'\u0000').trim()!==k.name).map(k=>({id:k.id,name:k.name,cls:k.cls}));
 if(!need.length)return;
 try{const r=await post({sync:need});if(r&&r.ok){need.forEach(k=>{(SH.has[k.cls]=SH.has[k.cls]||{})[k.id]=k.name});saveSH()}}catch(e){console.warn('sync',e)}}
const stBanner=()=>{const n=Q.length,w=t=>`<div class="note" style="margin:0 0 14px">${ic('alert')}<span>${t}</span></div>`;
 if(!CONFIG.api)return w('الموقع لسه مش متربط بشيت الحضور على جوجل. حط رابط الـ Apps Script في <b dir="ltr">CONFIG.api</b> جوه app.js — لحد ما تعمل كده الحضور والملاحظات مش هيتسجلوا.');
 if(SERR==='auth')return w('كلمة <b dir="ltr">KEY</b> في app.js مش مطابقة للي في Code.gs.');
 const retry=n?` <a href="javascript:void 0" data-retry style="font-weight:600;text-decoration:underline">إعادة المحاولة</a>`:'';
 if(SS==='fail')return w('مقدرتش أوصل لشيت الحضور دلوقتي (النت؟)، فالحضور والملاحظات مش ظاهرين. جرّب زر التحديث.');
 if(SS==='cache')return w('مقدرتش أوصل لشيت الحضور — بعرض آخر نسخة محفوظة على الجهاز.'+(n?` فيه ${n} تسجيل مستني يتبعت.`:'')+retry);
 return n&&!flushing?w(`فيه ${n} تسجيل لسه ما اتبعتوش لجوجل شيت — هيتبعتوا أول ما النت يرجع.`+retry):''};
const setSt=()=>$$('.stb').forEach(e=>e.innerHTML=stBanner());

async function photosOf(buf){const out={};try{
 const z=await JSZip.loadAsync(buf),rd=async p=>{const f=z.file(p);return f?new DOMParser().parseFromString(await f.async('string'),'text/xml'):null};
 const wbx=await rd('xl/workbook.xml'),rels=await rd('xl/_rels/workbook.xml.rels');
 const rid=[...wbx.getElementsByTagName('sheet')].find(s=>s.getAttribute('name')==='قاعدة البيانات').getAttribute('r:id');
 const t=[...rels.getElementsByTagName('Relationship')].find(r=>r.getAttribute('Id')===rid).getAttribute('Target').replace(/^\/?(xl\/)?/,'');
 const xml=await z.file('xl/'+t).async('string');
 const meta=await rd('xl/metadata.xml'),rv=await rd('xl/richData/rdrichvalue.xml'),rr=await rd('xl/richData/richValueRel.xml'),rl=await rd('xl/richData/_rels/richValueRel.xml.rels');
 const vmb=[...meta.getElementsByTagName('valueMetadata')[0].getElementsByTagName('bk')].map(b=>+b.getElementsByTagName('rc')[0].getAttribute('v'));
 const fut=[...meta.getElementsByTagName('futureMetadata')].find(f=>f.getAttribute('name')==='XLRICHVALUE');
 const rvb=[...fut.getElementsByTagName('bk')].map(b=>+b.getElementsByTagName('xlrd:rvb')[0].getAttribute('i'));
 const rvs=[...rv.getElementsByTagName('rv')].map(r=>+r.getElementsByTagName('v')[0].textContent);
 const ids=[...rr.getElementsByTagName('rel')].map(r=>r.getAttribute('r:id')),tg={};
 [...rl.getElementsByTagName('Relationship')].forEach(r=>tg[r.getAttribute('Id')]=r.getAttribute('Target'));
 const cache={};
 for(const m of xml.matchAll(/<c r="K(\d+)"([^>]*)>/g)){const vm=/\bvm="(\d+)"/.exec(m[2]);if(!vm)continue;
  const p='xl/media/'+tg[ids[rvs[rvb[vmb[vm[1]-1]]]]].split('/').pop();
  if(!cache[p]){const f=z.file(p);if(f)cache[p]=URL.createObjectURL(await f.async('blob'))}
  out[+m[1]]=cache[p]}
}catch(e){console.warn('photos',e)}return out}


async function load(){
 Q=loadQ();
 const db=(async()=>{const res=await fetch(FILE+'?v='+Date.now());if(!res.ok)throw Error('تعذّر تحميل ملف الإكسيل ('+res.status+'). تأكد إن اسم الملف '+FILE+' ومرفوع جنب index.html');
  const lm=res.headers.get('last-modified');LM=lm?new Date(lm).toLocaleString('ar-EG',{dateStyle:'medium',timeStyle:'short'}):'';
  const buf=await res.arrayBuffer();RAW={buf,wb:XLSX.read(buf,{type:'array'}),photos:await photosOf(buf)}})();
 await Promise.all([db,loadSheet()]);
 build();syncKids();flush()}
function build(){const wb=RAW.wb,photos=RAW.photos,B=wb.Sheets['قاعدة البيانات'],C=CONFIG;
 const classes=C.classes.slice(),years=C.years.slice(),th=C.th,dayName=C.day,want=DAYS.indexOf(dayName)<0?5:DAYS.indexOf(dayName);
 const start=iso2d(C.start),end=addM(start,C.months);let d=new Date(start);while(d.getUTCDay()!==want)d=new Date(+d+864e5);
 const dates=[];while(d<=end&&dates.length<53){dates.push(new Date(d));d=new Date(+d+6048e5)}
 (C.extra||[]).forEach(x=>{const e=iso2d(x);if(!dates.some(y=>+y===+e))dates.push(e)});dates.sort((a,b)=>a-b);
 const dkeys=dates.map(dk),ov=overlay(),notes=allNotes(),kids=[];
 for(let r=6;r<=305;r++){const name=cell(B,r,3);if(!name)continue;
  const cls=String(cell(B,r,6)??'').trim(),id='B-'+String(r-5).padStart(3,'0');
  const k={row:r,id,name:String(name).trim(),g:cell(B,r,4)||'',year:cell(B,r,5)||'',cls,alone:cell(B,r,7)||'',mom:ph(cell(B,r,8)),dad:ph(cell(B,r,9)),addr:cell(B,r,10)||'',notes:cell(B,r,12)||'',active:cell(B,r,13)!=='غير نشط',photo:photos[r]};
  k.pd=[];k.att=dkeys.map((dd,j)=>{const key=id+'|'+dd;if(key in ov){k.pd[j]=1;return ov[key]}return sv(id,cls,dd)});
  let P=0,Ab=0,E=0,last=-1;k.att.forEach((v,j)=>{if(v==='ح'){P++;last=j}else if(v==='غ')Ab++;else if(v==='ع')E++});
  Object.assign(k,{P,A:Ab,E,pct:P+Ab?P/(P+Ab):null,streak:k.att.filter((v,j)=>v==='غ'&&j>last).length,last:last>=0?dates[last]:null});
  k.tier=!k.active?5:k.pct==null?0:k.streak>=th.streak?4:k.pct>=th.ex?1:k.pct>=th.good?2:3;
  k.pn=notes.filter(n=>n.id===id).map(n=>({id:n.nid,date:noteDate(n.date),text:n.text,pd:n.pd}));kids.push(k)}
 const cls={};classes.forEach(c=>{const ks=kids.filter(k=>k.cls===c),act=ks.filter(k=>k.active),P=ks.reduce((s,k)=>s+k.P,0),Ab=ks.reduce((s,k)=>s+k.A,0);
  const per=dates.map((_,j)=>{const p=ks.filter(k=>k.att[j]==='ح').length,t=p+ks.filter(k=>k.att[j]==='غ').length;return{p,t,any:ks.some(k=>k.att[j])}});
  cls[c]={ks,n:act.length,boys:act.filter(k=>k.g==='ولد').length,girls:act.filter(k=>k.g==='بنت').length,avg:P+Ab?P/(P+Ab):null,per,meet:per.filter(x=>x.any).length,need:ks.filter(k=>k.tier==3||k.tier==4).length}});
 const tot=dates.map((_,j)=>({p:classes.reduce((s,c)=>s+cls[c].per[j].p,0),any:classes.some(c=>cls[c].per[j].any)}));
 const tk=iso2d(dayKey()),now=dates.reduce((m,x,i)=>x<=tk&&tk-x<6048e5?i:m,-1);
 D={kids,classes,cls,dates,th,dayName,years,start,end,tot,meet:tot.filter(x=>x.any).length,now,tj:dkeys.indexOf(dayKey())}}

/* ---------------- مكوّنات العرض ---------------- */
const av=k=>k.photo?`<img src="${k.photo}" alt="" loading="lazy">`:`<span>${esc([...k.name][0])}</span>`;
const ring=(p,c,sz=46)=>{const L=2*Math.PI*18;return`<svg class="ring" width="${sz}" height="${sz}" viewBox="0 0 44 44" role="img" aria-label="${p==null?'لا يوجد حضور':'نسبة الحضور '+Math.round(p*100)+'%'}"><circle cx="22" cy="22" r="18" class="tr"/><circle cx="22" cy="22" r="18" class="fg" style="stroke:${c};--L:${L};--o:${L*(1-(p||0))}"/><text x="22" y="26.5" text-anchor="middle">${p==null?'—':Math.round(p*100)}</text></svg>`};
const tag=t=>`<span class="tag" style="--c:${t[2]}">${ic(t[0])}${t[1]}</span>`;
const btns=k=>[['أم',k.mom],['أب',k.dad]].filter(x=>x[1]).map(([l,p])=>`<a class="b sm" href="tel:${p}" aria-label="اتصال بال${l}">${ic('phone')}${l}</a><a class="b sm wa" target="_blank" rel="noopener" href="${wa(p)}" aria-label="واتساب ال${l}">${ic('chat')}${l}</a>`).join('');
const card=k=>{const t=TIER[k.tier];return`<article class="card kid rv ${k.g==='بنت'?'f':''} ${k.active?'':'off'}"><div class="av">${av(k)}</div><div class="inf"><h3>${esc(k.name)}</h3><p class="meta">${esc(k.cls)} · ${esc(k.year)} · ${k.g==='بنت'?'بنت':'ولد'}</p>${tag(t)}</div>${ring(k.pct,t[2])}<div class="acts">${btns(k)}<button class="b sm pri" data-open="${k.row}">التفاصيل</button></div></article>`};
const chips=(sel,pre,suf='')=>`<div class="chips">${D.classes.map((c,i)=>`<a href="#${pre}/${encodeURIComponent(c)}${suf}" class="${c===sel?'on':''}" style="--c:${CC[i]}">${esc(c)}</a>`).join('')}</div>`;
const sw=(a,c)=>D.classes.includes(a)?a:D.classes[0];


function recBody(){const {c,j}=RC,o=D.cls[c],ks=o.ks.filter(k=>k.active);
 const cn=v=>ks.filter(k=>k.att[j]===v).length,left=ks.length-cn('ح')-cn('غ')-cn('ع');
 return`<div class="card info"><span>حاضر: <b>${cn('ح')}</b></span><span>غائب: <b>${cn('غ')}</b></span><span>معتذر: <b>${cn('ع')}</b></span><span>لسه: <b>${left}</b></span></div><div class="rl">${ks.map(k=>`<div class="card rr ${k.g==='بنت'?'f':''}"><div class="av">${av(k)}</div><div class="rn"><b>${esc(k.name)}</b>${k.pd[j]?'<small>لسه ما اتبعتش لجوجل شيت</small>':''}<button class="b sm nb" data-note="${k.row}" aria-label="ملاحظة عن ${esc(k.name)}">${ic('note')}ملاحظة${k.pn.length?' ('+k.pn.length+')':''}</button></div><div class="seg" role="group" aria-label="حضور ${esc(k.name)}">${[['ح','ok','حاضر'],['غ','no','غائب'],['ع','ex','معتذر']].map(([v,cl,t])=>`<button class="sg ${cl}${k.att[j]===v?' on':''}" data-set="${v}" data-r="${k.row}" data-j="${j}" aria-pressed="${k.att[j]===v}" aria-label="${t}: ${esc(k.name)}">${v}</button>`).join('')}</div></div>`).join('')||`<div class="empty">${ic('users')}مفيش أطفال نشطين في الفصل ده</div>`}</div>`}
const renderRec=()=>{const e=$('#rl');if(e&&RC.j>=0)e.innerHTML=recBody();};
const V={
rec(a){const c=sw(a);
 const j=D.tj;RC={c,j};
 const head=`<h2>${ic('check')}تسجيل الحضور</h2>${chips(c,'rec')}<div class="stb">${stBanner()}</div>`;
 if(j<0){const nx=D.dates.find(x=>dk(x)>dayKey());
  return head+`<div class="card empty">${ic('lock')}<b>التسجيل مقفول النهارده</b><span>الحضور بيتسجّل في يوم الاجتماع بس (${esc(D.dayName)}) أو في يوم إضافي محدد.</span>${nx?`<span>الاجتماع الجاي: <b>${fdl(nx)}</b></span>`:''}</div>`}
 return head+`<div class="tb"><span class="chip">${ic('calendar')}اجتماع النهارده: ${fdl(D.dates[j])}</span></div><p class="sub" style="margin-bottom:12px">سجّل الحضور في يوم الاجتماع فقط، بما في ذلك الأيام الإضافية المحددة في الإعدادات. كل طفل: ح حاضر، غ غائب، ع معتذر. التغييرات تتحفظ فورًا في جوجل شيت.</p><div id="rl">${recBody()}</div>`},
home(){const a=D.kids.filter(k=>k.active),P=D.kids.reduce((s,k)=>s+k.P,0),Ab=D.kids.reduce((s,k)=>s+k.A,0),avg=P+Ab?P/(P+Ab):null;
 const st=[['users','المسجلين',D.kids.length],['checkc','النشطين',a.length],['user','أولاد',a.filter(k=>k.g==='ولد').length],['user','بنات',a.filter(k=>k.g==='بنت').length],['calendar','جمع مسجلة',D.meet],['alert','محتاجين افتقاد',D.kids.filter(k=>k.tier==3||k.tier==4).length,'warn']];
 let last=D.tot.map(x=>x.any).lastIndexOf(true);const mx=Math.max(1,...D.tot.map(x=>x.p));
 return`<section class="hero"><div class="card hero-main"><h1>${ic('sprout')}خدمة براعم</h1><p>نظام الافتقاد والحضور — كل أطفال الخدمة وفصولهم ومتابعتهم في مكان واحد.</p><div class="meta-row"><span class="chip">${ic('calendar')}العام الخدمي ${fdl(D.start)} ← ${fdl(D.end)}</span><span class="chip">${ic('clock')}اجتماع كل ${esc(D.dayName)}</span></div></div>
 <div class="card hero-side rv">${ring(avg,'var(--p)',132)}<div><b>متوسط الحضور</b><p class="sub">من ${D.meet} اجتماع مسجّل</p></div></div></section>
 <div class="stb">${stBanner()}</div><div class="stats">${st.map(s=>`<div class="card st ${s[3]||''}"><div class="e">${ic(s[0])}</div><div><b data-count="${s[2]}">0</b><small>${s[1]}</small></div></div>`).join('')}</div>
 <h2>${ic('users')}الفصول</h2><div class="grid">${D.classes.map((c,i)=>{const o=D.cls[c];return`<div class="card ct rv" style="--c:${CC[i]}"><h3><span>ملاك ${esc(c)}</span>${ring(o.avg,CC[i],54)}</h3><ul><li>${ic('users')}${o.n} طفل</li><li>${ic('alert')}${o.need} محتاج افتقاد</li><li>${ic('user')}${o.boys} ولد · ${o.girls} بنت</li><li>${ic('plus')}${Math.max(0,CONFIG.cap-o.ks.length)} مكان فاضي</li></ul><div class="row"><a class="b pri" href="#class/${encodeURIComponent(c)}">${ic('users')}الكروت</a><a class="b" href="#att/${encodeURIComponent(c)}">${ic('calendar')}الحضور</a></div></div>`}).join('')}</div>
 <h2>${ic('chart')}الحضور في كل اجتماع</h2><div class="card rv">${last<0?`<div class="empty">${ic('calendar')}لسه مفيش حضور مسجّل</div>`:`<div class="bars">${D.tot.slice(0,last+1).map((x,i)=>`<div class="bar ${i==D.now?'now':''}" title="${fdl(D.dates[i])}: ${x.p}" style="--h:${x.p/mx*78}%;--i:${i}"><b>${x.p}</b><i></i><span>${fd(D.dates[i])}</span></div>`).join('')}</div>`}</div>
 <h2>${ic('map')}اختصارات</h2><div style="display:flex;gap:10px;flex-wrap:wrap"><a class="b pri" href="#rec">${ic('check')}تسجيل الحضور</a><a class="b" href="#kids">${ic('list')}قاعدة البيانات</a><a class="b" href="#follow">${ic('phone')}الافتقاد</a><a class="b" href="#settings">${ic('settings')}الإعدادات</a><a class="b" href="#guide">${ic('book')}الدليل</a><a class="b" href="${FILE}" download>${ic('plus')}إضافة طفل (حمّل الإكسيل)</a></div>
 ${LM?`<p class="sub" style="margin-top:26px;text-align:center">آخر تحديث لقاعدة البيانات: ${LM}</p>`:''}`},
kids(){return`<h2>${ic('list')}قاعدة بيانات الأطفال</h2><div class="tb"><input id="q" type="search" aria-label="بحث" placeholder="ابحث بالاسم أو الرقم أو العنوان…"><select id="fc" aria-label="الفصل"><option value="">كل الفصول</option>${D.classes.map(c=>`<option>${esc(c)}</option>`).join('')}</select><select id="fg" aria-label="النوع"><option value="">أولاد وبنات</option><option>ولد</option><option>بنت</option></select><select id="fs" aria-label="الحالة"><option value="">كل الحالات</option><option value="1">نشط</option><option value="0">غير نشط</option></select><select id="fy" aria-label="السنة الدراسية"><option value="">كل السنين</option>${D.years.map(y=>`<option>${esc(y)}</option>`).join('')}</select></div><p id="cn" class="sub" style="margin-bottom:12px" aria-live="polite"></p><div class="grid" id="kg"></div>`},
class(a){const c=sw(a),o=D.cls[c];return`<h2>${ic('users')}كروت أطفال ملاك ${esc(c)}</h2>${chips(c,'class')}<div class="card info rv"><span>العدد: <b>${o.n}</b></span><span>${o.boys} ولد · ${o.girls} بنت</span><span>متوسط الحضور: <b>${pc(o.avg)}</b></span><span>جمع مسجلة: <b>${o.meet}</b></span><span>محتاجين افتقاد: <b>${o.need}</b></span><span>أماكن فاضية: <b>${Math.max(0,CONFIG.cap-o.ks.length)}</b></span></div><div class="grid">${o.ks.map(card).join('')||`<div class="empty">${ic('users')}مفيش أطفال في الفصل ده</div>`}</div>`},
att(a){const c=sw(a),o=D.cls[c],dt=D.dates;return`<h2>${ic('calendar')}قايمة حضور ملاك ${esc(c)}</h2>${chips(c,'att')}<div class="stb">${stBanner()}</div><div class="card info rv"><span>جمع مسجلة: <b>${o.meet}</b></span><span>المتوسط: <b>${pc(o.avg)}</b></span><span class="lgd"><i class="lg-ok">ح</i>حاضر</span><span class="lgd"><i class="lg-no">غ</i>غائب</span><span class="lgd"><i class="lg-ex">ع</i>معتذر</span></div>
 <div class="card tw rv"><table><thead><tr><th>الاسم</th><th>حضور</th><th>غياب</th><th>بعذر</th><th>النسبة</th><th>غياب متتالي</th><th>آخر حضور</th><th>الحالة</th>${dt.map((x,i)=>`<th class="${i==D.now?'now':''}" ${i==D.now?'id="nowc"':''}>${fd(x)}</th>`).join('')}</tr></thead><tbody>${o.ks.map(k=>`<tr><td><a href="javascript:void 0" data-open="${k.row}">${esc(k.name)}</a></td><td>${k.P}</td><td>${k.A}</td><td>${k.E}</td><td>${pc(k.pct)}</td><td>${k.streak}</td><td>${k.last?fd(k.last):'—'}</td><td style="color:${TIER[k.tier][2]}">${TIER[k.tier][1]}</td>${k.att.map((v,i)=>`<td class="${v==='ح'?'ok':v==='غ'?'no':v==='ع'?'ex':''} ${i==D.now?'now':''} ${k.pd[i]?'pd':''}">${v}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr><td>الحاضرين</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.any?x.p:''}</td>`).join('')}</tr><tr><td>النسبة</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.t?Math.round(x.p/x.t*100)+'%':''}</td>`).join('')}</tr></tfoot></table></div>`},
follow(){const l=D.kids.filter(k=>k.tier==3||k.tier==4).sort((a,b)=>b.streak-a.streak||(a.pct??0)-(b.pct??0));
 return`<h2>${ic('phone')}قايمة الافتقاد</h2><div class="card info rv"><span>الأطفال اللي نسبة حضورهم أقل من <b>${Math.round(D.th.ex*100)}%</b> أو غايبين <b>${D.th.streak}</b> اجتماعات ورا بعض أو أكتر — العدد: <b>${l.length}</b></span></div><div class="fl">${l.map((k,i)=>`<div class="card fi rv" style="--c:${TIER[k.tier][2]}"><div class="rank">${i+1}</div><div><h3>${esc(k.name)}</h3><p class="meta">${esc(k.cls)} · نسبة ${pc(k.pct)} · غياب متتالي ${k.streak} · آخر حضور ${k.last?fdl(k.last):'—'}<br>${ic('pin')} ${esc(k.addr||'—')}</p>${tag(TIER[k.tier])}</div><div class="acts">${btns(k)||'<span class="sub">— مفيش رقم —</span>'}<button class="b sm pri" data-open="${k.row}">التفاصيل</button></div></div>`).join('')||`<div class="card empty">${ic('checkc')}مفيش حد محتاج افتقاد دلوقتي</div>`}</div>`},
settings(){const k=(a,b)=>`<div><small>${a}</small><b>${b}</b></div>`;return`<h2>${ic('settings')}الإعدادات</h2><div class="card kv rv">${k('بداية العام الخدمي',fdl(D.start))+k('نهاية العام الخدمي',fdl(D.end))+k('يوم الاجتماع',esc(D.dayName))+k('«ممتاز» من',Math.round(D.th.ex*100)+'%')+k('«جيد» من',Math.round(D.th.good*100)+'%')+k('افتقاد ضروري بعد غياب',D.th.streak+' اجتماعات')+k('الفصول (الملائكة)',D.classes.map(esc).join(' · '))+k('السنين الدراسية',D.years.map(esc).join(' · '))+k('شيت الحضور على جوجل',!CONFIG.api?'مش متربط':SS==='ok'?'متربط ✓':'متربط (مفيش اتصال دلوقتي)')}</div><div class="note">${ic('alert')}<span>الإعدادات دي موجودة في أول ملف <b dir="ltr">app.js</b> (<b dir="ltr">CONFIG</b>) — لو غيّرتها ارفع الملف على GitHub. ماتغيّرش أسماء الفصول بعد تسجيل الأطفال.</span></div>`},
guide(){const s=[['plus','إضافة طفل جديد','افتح ملف الإكسيل (قاعدة البيانات) وأضف الطفل في أول صف فاضي (الاسم، النوع، السنة، الفصل، التليفونات…) وارفع الملف على GitHub بنفس الاسم. الطفل يظهر في الموقع، وأول ما حد يفتح الموقع بيتضاف لوحده في شيت الحضور على جوجل.'],['camera','صورة الطفل','حطّ الصورة في عمود «الصورة» بـ Insert ← Pictures ← Place in Cell. بتظهر في كارت الطفل، ولو مفيش صورة بيظهر أول حرف من الاسم.'],['check','تسجيل الحضور','من تبويب «تسجيل» اختار الفصل وادوس ح / غ / ع جنب كل طفل. كل الأطفال بيتسجلوا غايبين تلقائي. التسجيل متاح يوم الاجتماع بس، والأيام اللي فاتت مقفولة ومينفعش تتعدّل. بيتحفظ فورًا في شيت الحضور على جوجل ويظهر عند كل المدرسين.'],['note','الملاحظات','من كارت الطفل (التفاصيل أو زر «ملاحظة» في التسجيل) اكتب ملاحظة بتاريخها، وبتتحفظ في تاب «الملاحظات» في جوجل شيت.'],['calendar','قايمة الحضور','تبويب «الحضور» بيعرض كل اجتماعات العام لكل فصل، والنسب والألوان بتتحسب في الموقع أوتوماتيك.'],['phone','الافتقاد','القايمة بتتجمّع أوتوماتيك حسب الحدود في «الإعدادات». اضغط على زر الاتصال أو واتساب للتواصل مع الأم أو الأب.'],['ban','طفل غير نشط','بدل ما تمسح الصف، غيّر «الحالة» لـ «غير نشط». ماتعملش Sort ولا تمسح صفوف.'],['refresh','تحديث ولو النت قطع','زر التحديث في أعلى الصفحة بيجيب أحدث بيانات (والحضور بيتحدّث لوحده لما ترجع للموقع). لو النت قطع وقت التسجيل، التسجيل بيتحفظ على الجهاز وبيتبعت لوحده لما النت يرجع.']];return`<h2>${ic('book')}دليل الاستخدام</h2><div class="steps">${s.map(x=>`<div class="card rv"><div class="e">${ic(x[0])}</div><h3>${x[1]}</h3><p>${x[2]}</p></div>`).join('')}</div>`}};

function renderKids(){const q=$('#q').value.trim().toLowerCase(),f=id=>$(id).value;
 const l=D.kids.filter(k=>(!f('#fc')||k.cls===f('#fc'))&&(!f('#fg')||k.g===f('#fg'))&&(!f('#fy')||k.year===f('#fy'))&&(!f('#fs')||k.active==(f('#fs')==='1'))&&(!q||[k.name,k.mom,k.dad,k.addr,k.id,k.notes,...k.pn.map(n=>n.text)].join(' ').toLowerCase().includes(q)));
 $('#cn').textContent='النتائج: '+l.length+' من '+D.kids.length;$('#kg').innerHTML=l.map(card).join('')||`<div class="empty">${ic('list')}مفيش نتائج — جرّب تغيّر البحث أو الفلاتر</div>`;fx($('#kg'))}


function openKid(row,again,fn){const k=D.kids.find(x=>x.row==row);if(!k)return;const t=TIER[k.tier],f=(a,b)=>b!==''&&b!=null?`<div><small>${a}</small><b>${b}</b></div>`:'';
 if(!again)lastFocus=document.activeElement;
 const modal=$('#modal'),wasHidden=modal.hidden;
 $('#mc').innerHTML=`<div class="mh in"><div class="av ${k.g==='بنت'?'f':''}">${av(k)}</div><div><h2>${esc(k.name)}</h2><p class="meta">${k.id} · ${esc(k.cls)} · ${esc(k.year)}</p>${tag(t)}</div></div>
 <div class="kv">${f('حضور',k.P)+f('غياب',k.A)+f('بعذر',k.E)+f('النسبة',pc(k.pct))+f('غياب متتالي',k.streak)+f('آخر حضور',k.last?fdl(k.last):'—')+f('يروح لوحده؟',esc(k.alone))+f('النوع',k.g==='بنت'?'بنت':'ولد')}</div>
 ${k.addr?`<div class="mp">${ic('pin')}<span>${esc(k.addr)}</span><a class="b sm" target="_blank" rel="noopener" href="https://www.google.com/maps/search/${encodeURIComponent(k.addr)}">${ic('map')}الخريطة</a></div>`:''}
 ${k.notes?`<p class="mp">${ic('note')}<span>${esc(k.notes)}</span></p>`:''}<div style="display:flex;gap:6px;flex-wrap:wrap">${btns(k)}</div>
 <div class="nt"><h3>${ic('note')}ملاحظات جديدة</h3>${k.pn.map(n=>`<div class="pnote"><span class="d">${esc(n.date)}</span><span>${esc(n.text)}${n.pd?' <small class="sub">(لسه ما اتبعتتش)</small>':''}</span><button class="b sm" data-delnote="${n.id}" data-r="${k.row}" aria-label="حذف الملاحظة">${ic('close')}</button></div>`).join('')}<div class="addn"><input type="date" id="nd" value="${dayKey()}" aria-label="تاريخ الملاحظة"><textarea id="ntx" rows="2" placeholder="اكتب ملاحظة…" aria-label="الملاحظة"></textarea><button class="b pri" data-addnote="${k.row}">${ic('plus')}إضافة ملاحظة</button></div><p class="sub" style="margin-top:8px">بتتحفظ في تاب «الملاحظات» في جوجل شيت.</p></div>
 <div class="tl">${k.att.map((v,i)=>v?`<span class="${v==='ح'?'ok':v==='غ'?'no':'ex'}${k.pd[i]?' pd':''}">${fd(D.dates[i])} ${v}</span>`:'').join('')}</div>`;
 if(wasHidden){const state=history.state&&typeof history.state==='object'?history.state:{};history.pushState(Object.assign({},state,{__b3Modal:true}),'',location.href)}
 modal.hidden=false;document.body.style.overflow='hidden';if(fn){const t=$('#ntx');t.scrollIntoView({block:'center'});t.focus()}else if(!again)$('#mx').focus()}

const closeM=fromHistory=>{if($('#modal').hidden)return;const goBack=!fromHistory&&history.state&&history.state.__b3Modal;$('#modal').hidden=true;document.body.style.overflow='';if($('#rl'))renderRec();else if(lastFocus&&lastFocus.focus)lastFocus.focus();if(goBack)history.back()};

const io=new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}}),{threshold:.08});
function fx(r=view){$$('.rv:not(.in)',r).forEach(e=>io.observe(e));
 $$('[data-count]',r).forEach(e=>{const to=+e.dataset.count,s=e.dataset.suf||'',t0=performance.now();(function f(t){const p=Math.min(1,(t-t0)/1100);e.textContent=Math.round(to*(1-Math.pow(1-p,4)))+s;p<1&&requestAnimationFrame(f)})(t0)})}
function route(){const [r,a,b]=decodeURIComponent(location.hash.slice(1)||'home').split('/'),n=V[r]?r:'home';
 view.innerHTML=`<div class="page">${V[n](a,b)}</div>`;scrollTo(0,0);fx();
 if(n==='kids'){['#q','#fc','#fg','#fs','#fy'].forEach(i=>$(i).addEventListener('input',renderKids));renderKids()}
  if(n==='att')setTimeout(()=>{const c=$('#nowc');c&&c.scrollIntoView({inline:'center',block:'nearest',behavior:'smooth'})},500);
 const href='#'+n;$$('#nav a,#tabs a').forEach(x=>{const on=x.getAttribute('href')===href;x.classList.toggle('on',on);on?x.setAttribute('aria-current','page'):x.removeAttribute('aria-current')})}
function init(){
 $('#nav').innerHTML=NAV.map(n=>`<a href="#${n[0]}">${ic(n[1])}${n[2]}</a>`).join('');
 $('#tabs').innerHTML=NAV.filter(n=>!n[3]).map(n=>`<a href="#${n[0]}">${ic(n[1])}<span>${n[2]}</span></a>`).join('');
 $('#dl').href=FILE;addEventListener('hashchange',route)}
document.addEventListener('click',e=>{const o=e.target.closest('[data-open]');if(o)openKid(o.dataset.open);if(e.target.id==='modal'||e.target.closest('#mx'))closeM()});
addEventListener('keydown',e=>e.key==='Escape'&&closeM());
addEventListener('popstate',()=>{if(!$('#modal').hidden)closeM(true)});

async function boot(first){try{await load();if(first)init();route()}catch(e){view.innerHTML=`<div class="card empty" style="padding:34px">${ic('alert')}<h2 style="justify-content:center;margin:0">حصلت مشكلة</h2><p class="sub" style="line-height:2">${esc(e.message)}<br>لو بتفتح الملف من الكمبيوتر مباشرة (file://) لازم يتفتح من GitHub Pages أو سيرفر محلي.</p></div>`}$('#splash').classList.add('off')}
/* ---------------- كلمة السر اليومية: (تاريخ اليوم)@Bra3em ---------------- */
const PW_TAG='@bra3em';
const cairoNow=()=>{const p={};new Intl.DateTimeFormat('en-GB',{timeZone:'Africa/Cairo',day:'2-digit',month:'2-digit',year:'numeric'}).formatToParts(new Date()).forEach(x=>p[x.type]=x.value);return p};
const dayKey=()=>{const t=cairoNow();return t.year+'-'+t.month+'-'+t.day};
const toEn=s=>s.replace(/[٠-٩]/g,d=>'٠١٢٣٤٥٦٧٨٩'.indexOf(d)).replace(/[۰-۹]/g,d=>'۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
// الصيغة الأساسية: 26102026@Bra3em  (اليوم+الشهر+السنة) — بيقبل كمان 26-10-2026 و 2026-10-26 و 26/10
function pwOk(inp){const s=toEn(String(inp||'').trim());if(!s.toLowerCase().endsWith(PW_TAG))return false;
 const head=s.slice(0,-PW_TAG.length).trim();if(!head||!/^[\d\s\/\-.]+$/.test(head))return false;
 const t=cairoNow(),d=t.day,m=t.month,y=t.year,dn=String(+d),mn=String(+m);
 return [d+m+y,y+m+d,d+m,dn+mn+y].includes(head.replace(/\D/g,''))}
const unlocked=()=>{try{return localStorage.getItem('b3_unlock')===dayKey()}catch(e){return false}};
const setUnlock=()=>{try{localStorage.setItem('b3_unlock',dayKey())}catch(e){}};
const lock=()=>{try{localStorage.removeItem('b3_unlock')}catch(e){}location.reload()};
function start(){$('#gate').hidden=true;$('#splash').classList.remove('off');boot(true);
 // لو اليوم اتغيّر والموقع مفتوح، يقفل ويطلب الباسورد الجديد
 setInterval(()=>{if(!unlocked())location.reload()},60000)}
function showGate(){$('#splash').classList.add('off');$('#gate').hidden=false;$('#gp').focus();
 $('#gf').onsubmit=e=>{e.preventDefault();const v=$('#gp').value;
  if(pwOk(v)){setUnlock();start()}
  else{const f=$('#gf');$('#ge').textContent='كلمة السر غلط، جرّب تاني';$('#gp').select();f.classList.remove('shake');void f.offsetWidth;f.classList.add('shake')}}}

$('#rf').onclick=()=>{$('#splash').classList.remove('off');boot(false)};
$('#lk').onclick=lock;

/* ---------------- التسجيل والملاحظات ---------------- */
document.addEventListener('click',e=>{
 const s=e.target.closest('[data-set]');
 if(s){const k=D.kids.find(x=>x.row==s.dataset.r),j=+s.dataset.j;if(!k||D.tj<0||j!==D.tj||j!==RC.j)return;
  let v=s.dataset.set;if(k.att[j]===v)v='';qAtt(k,dk(D.dates[j]),v);saveQ();build();renderRec();flush();return}
 const an=e.target.closest('[data-addnote]');
 if(an){const k=D.kids.find(x=>x.row==an.dataset.addnote),t=$('#ntx').value.trim();if(!t){$('#ntx').focus();return}
  Q.push({t:'note',nid:uid(),id:k.id,cls:k.cls,name:k.name,date:$('#nd').value||dayKey(),text:t,ts:Date.now()});saveQ();build();openKid(k.row,true);flush();return}
 const dn=e.target.closest('[data-delnote]');
 if(dn){if(!confirm('تمسح الملاحظة دي؟'))return;const nid=dn.dataset.delnote,qi=Q.findIndex(o=>o.t==='note'&&o.nid===nid&&!o.fl);
  if(qi>=0)Q.splice(qi,1);else Q.push({t:'del',nid,ts:Date.now()});saveQ();build();openKid(dn.dataset.r,true);flush();return}
 if(e.target.closest('[data-retry]')){loadSheet().then(()=>{build();setSt();renderRec();syncKids();flush()})}
});
document.addEventListener('click',e=>{const n=e.target.closest('[data-note]');if(n)openKid(n.dataset.note,false,true)});
// أول ما ترجع للموقع بيجيب أحدث حضور وملاحظات (من غير ما يحمّل الإكسيل تاني)
document.addEventListener('visibilitychange',async()=>{
 if(document.visibilityState!=='visible'||!D||!CONFIG.api||!$('#gate').hidden||Date.now()-LASTF<30000)return;
 await loadSheet();build();setSt();if($('#modal').hidden)renderRec();flush()});
addEventListener('online',()=>{if(D)flush()});

unlocked()?start():showGate();
