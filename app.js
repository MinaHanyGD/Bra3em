const FILE='Bara3em_System_Final.xlsx';
const $=(s,e=document)=>e.querySelector(s),$$=(s,e=document)=>[...e.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const ic=n=>`<svg class="ic" aria-hidden="true"><use href="#i-${n}"/></svg>`;
// [أيقونة، الوصف، اللون]
const TIER=[['clock','لسه مفيش حضور','var(--none)'],['star','ممتاز','var(--ok)'],['checkc','جيد','var(--good)'],['alert','محتاج متابعة','var(--warn)'],['phone','افتقاده ضروري','var(--bad)'],['ban','غير نشط','var(--off)']];
const CC=['var(--c1)','var(--c2)','var(--c3)','var(--c4)','var(--c5)','var(--c6)'];
const NAV=[['home','home','الرئيسية'],['kids','list','الأطفال'],['class','users','الفصول'],['att','calendar','الحضور'],['rec','check','تسجيل'],['follow','phone','الافتقاد'],['sync','download','للإكسيل',1],['settings','settings','الإعدادات',1],['guide','book','الدليل',1]];
const cell=(ws,r,c)=>{const x=ws&&ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:undefined};
const ds=n=>new Date(Math.round((n-25569)*864e5)),fd=d=>d.getUTCDate()+'/'+(d.getUTCMonth()+1),fdl=d=>fd(d)+'/'+d.getUTCFullYear();
const ph=v=>{if(v==null||v==='')return'';let s=String(v).replace(/\D/g,'');if(s.length===10&&s[0]!=='0')s='0'+s;return s};
const wa=p=>'https://wa.me/20'+p.slice(-10),pc=p=>p==null?'—':Math.round(p*100)+'%';
let D,view=$('#view'),LM='',lastFocus=null;

/* ---------- تسجيل محلي (حضور + ملاحظات) محفوظ على الجهاز لحد ما يتنقل للإكسيل ---------- */
const LSK='b3_pending',dk=d=>d.toISOString().slice(0,10);
let PEND={att:{},notes:[]},RAW,RC={c:'',j:0},CP={};
const loadPend=()=>{try{const p=JSON.parse(localStorage.getItem(LSK));if(p&&p.att&&Array.isArray(p.notes))return p}catch(e){}return{att:{},notes:[]}};
const savePend=()=>{try{localStorage.setItem(LSK,JSON.stringify(PEND))}catch(e){alert('مش قادر أحفظ على الجهاز ده — تأكد إن المتصفح مش في الوضع الخاص')}};
const noteDate=iso=>{const [y,m,d]=iso.split('-');return +d+'/'+ +m+'/'+y};
const defJ=()=>{if(D.now>=0)return D.now;const t=Date.now();return D.dates.reduce((m,x,i)=>x<=t?i:m,0)};
const q=s=>/[\n"\t]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;
const comb=k=>[k.notes,...k.pn.map(n=>n.date+': '+n.text)].filter(x=>x!==''&&x!=null).join('\n');
const copyTxt=async t=>{try{await navigator.clipboard.writeText(t)}catch(e){const x=document.createElement('textarea');x.value=t;x.style.cssText='position:fixed;opacity:0';document.body.appendChild(x);x.select();try{document.execCommand('copy')}catch(_){}x.remove()}};
const pendBanner=()=>{const a=Object.keys(PEND.att).length,n=PEND.notes.length;return a+n?`<div class="note" style="margin:0 0 14px">${ic('alert')}<span>فيه ${a} تسجيل حضور و${n} ملاحظة لسه ما اتنقلوش للإكسيل. <a href="#sync" style="font-weight:600;text-decoration:underline">افتح صفحة «للإكسيل»</a></span></div>`:''};

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
 const res=await fetch(FILE+'?v='+Date.now());if(!res.ok)throw Error('تعذّر تحميل ملف الإكسيل ('+res.status+'). تأكد إن اسم الملف '+FILE+' ومرفوع جنب index.html');
 const lm=res.headers.get('last-modified');LM=lm?new Date(lm).toLocaleString('ar-EG',{dateStyle:'medium',timeStyle:'short'}):'';
 const buf=await res.arrayBuffer();RAW={buf,wb:XLSX.read(buf,{type:'array'}),photos:await photosOf(buf)};PEND=loadPend();build()}
function build(){const wb=RAW.wb,photos=RAW.photos;
 const S=wb.Sheets['الإعدادات'],B=wb.Sheets['قاعدة البيانات'];
 const classes=[],days=[],years=[];
 for(let r=4;r<=8;r++){const v=cell(S,r,2);if(typeof v==='string')classes.push(v)}
 for(let r=4;r<=10;r++){const v=cell(S,r,8);if(v)days.push(v);const y=cell(S,r,4);if(y)years.push(y)}
 const th={ex:cell(S,16,2),good:cell(S,17,2),streak:cell(S,18,2)},dayName=cell(S,12,2),want=days.indexOf(dayName)<0?5:days.indexOf(dayName);
 const end=ds(cell(S,11,2));let d=ds(cell(S,10,2));while(d.getUTCDay()!==want)d=new Date(+d+864e5);
 const dates=[];while(d<=end&&dates.length<53){dates.push(new Date(d));d=new Date(+d+6048e5)}
 const kids=[],cnt={};let dirty=0;
 for(let r=6;r<=305;r++){const name=cell(B,r,3);if(!name)continue;
  const cls=String(cell(B,r,6)??''),i=cnt[cls]=(cnt[cls]||0)+1,A=wb.Sheets['حضور '+cls];
  const k={row:r,id:'B-'+String(r-5).padStart(3,'0'),name:String(name).trim(),g:cell(B,r,4)||'',year:cell(B,r,5)||'',cls,alone:cell(B,r,7)||'',mom:ph(cell(B,r,8)),dad:ph(cell(B,r,9)),addr:cell(B,r,10)||'',notes:cell(B,r,12)||'',active:cell(B,r,13)!=='غير نشط',photo:photos[r]};
  k.att=dates.map((_,c)=>(A&&i<=30&&cell(A,6+i,10+c))||'');k.xl=k.att.slice();k.pd=[];
  k.att=k.att.map((v,j)=>{const key=k.cls+'|'+k.name+'|'+dk(dates[j]),p=PEND.att[key];if(!p)return v;if(p.v===v){delete PEND.att[key];dirty=1;return v}k.pd[j]=1;return p.v});
  let P=0,Ab=0,E=0,last=-1;k.att.forEach((v,j)=>{if(v==='ح'){P++;last=j}else if(v==='غ')Ab++;else if(v==='ع')E++});
  Object.assign(k,{P,A:Ab,E,pct:P+Ab?P/(P+Ab):null,streak:k.att.filter((v,j)=>v==='غ'&&j>last).length,last:last>=0?dates[last]:null});
  k.tier=!k.active?5:k.pct==null?0:k.streak>=th.streak?4:k.pct>=th.ex?1:k.pct>=th.good?2:3;kids.push(k)}
 PEND.notes=PEND.notes.filter(n=>{const k=kids.find(x=>x.cls+'|'+x.name===n.key);if(k&&String(k.notes).includes(n.text)){dirty=1;return false}return true});
 kids.forEach(k=>k.pn=PEND.notes.filter(n=>n.key===k.cls+'|'+k.name));if(dirty)savePend();
 const cls={};classes.forEach(c=>{const ks=kids.filter(k=>k.cls===c),act=ks.filter(k=>k.active),P=ks.reduce((s,k)=>s+k.P,0),Ab=ks.reduce((s,k)=>s+k.A,0);
  const per=dates.map((_,j)=>{const p=ks.filter(k=>k.att[j]==='ح').length,t=p+ks.filter(k=>k.att[j]==='غ').length;return{p,t,any:ks.some(k=>k.att[j])}});
  cls[c]={ks,n:act.length,boys:act.filter(k=>k.g==='ولد').length,girls:act.filter(k=>k.g==='بنت').length,avg:P+Ab?P/(P+Ab):null,per,meet:per.filter(x=>x.any).length,need:ks.filter(k=>k.tier==3||k.tier==4).length}});
 const tot=dates.map((_,j)=>({p:classes.reduce((s,c)=>s+cls[c].per[j].p,0),any:classes.some(c=>cls[c].per[j].any)}));
 D={kids,classes,cls,dates,th,dayName,years,start:ds(cell(S,10,2)),end,tot,meet:tot.filter(x=>x.any).length,now:(()=>{const t=Date.now();return dates.findIndex(x=>x<=t&&t<+x+6048e5)})()}}

/* ---------------- مكوّنات العرض ---------------- */
const av=k=>k.photo?`<img src="${k.photo}" alt="" loading="lazy">`:`<span>${esc([...k.name][0])}</span>`;
const ring=(p,c,sz=46)=>{const L=2*Math.PI*18;return`<svg class="ring" width="${sz}" height="${sz}" viewBox="0 0 44 44" role="img" aria-label="${p==null?'لا يوجد حضور':'نسبة الحضور '+Math.round(p*100)+'%'}"><circle cx="22" cy="22" r="18" class="tr"/><circle cx="22" cy="22" r="18" class="fg" style="stroke:${c};--L:${L};--o:${L*(1-(p||0))}"/><text x="22" y="26.5" text-anchor="middle">${p==null?'—':Math.round(p*100)}</text></svg>`};
const tag=t=>`<span class="tag" style="--c:${t[2]}">${ic(t[0])}${t[1]}</span>`;
const btns=k=>[['أم',k.mom],['أب',k.dad]].filter(x=>x[1]).map(([l,p])=>`<a class="b sm" href="tel:${p}" aria-label="اتصال بال${l}">${ic('phone')}${l}</a><a class="b sm wa" target="_blank" rel="noopener" href="${wa(p)}" aria-label="واتساب ال${l}">${ic('chat')}${l}</a>`).join('');
const card=k=>{const t=TIER[k.tier];return`<article class="card kid rv ${k.g==='بنت'?'f':''} ${k.active?'':'off'}"><div class="av">${av(k)}</div><div class="inf"><h3>${esc(k.name)}</h3><p class="meta">${esc(k.cls)} · ${esc(k.year)} · ${k.g==='بنت'?'بنت':'ولد'}</p>${tag(t)}</div>${ring(k.pct,t[2])}<div class="acts">${btns(k)}<button class="b sm pri" data-open="${k.row}">التفاصيل</button></div></article>`};
const chips=(sel,pre,suf='')=>`<div class="chips">${D.classes.map((c,i)=>`<a href="#${pre}/${encodeURIComponent(c)}${suf}" class="${c===sel?'on':''}" style="--c:${CC[i]}">${esc(c)}</a>`).join('')}</div>`;
const sw=(a,c)=>D.classes.includes(a)?a:D.classes[0];

function autoAbsent(c,j){if(dk(D.dates[j])>dayKey())return;let ch=0;
 D.cls[c].ks.slice(0,30).forEach(k=>{if(k.active&&k.att[j]===''){PEND.att[k.cls+'|'+k.name+'|'+dk(D.dates[j])]={v:'غ',t:Date.now()};ch=1}});if(ch){savePend();build()}}
function recBody(){const {c,j}=RC,o=D.cls[c],ks=o.ks.filter(k=>k.active);
 const cn=v=>ks.filter(k=>k.att[j]===v).length,left=ks.length-cn('ح')-cn('غ')-cn('ع');
 return`<div class="card info"><span>حاضر: <b>${cn('ح')}</b></span><span>غائب: <b>${cn('غ')}</b></span><span>معتذر: <b>${cn('ع')}</b></span><span>لسه: <b>${left}</b></span></div><div class="rl">${ks.map(k=>`<div class="card rr ${k.g==='بنت'?'f':''}"><div class="av">${av(k)}</div><div class="rn"><b>${esc(k.name)}</b>${k.pd[j]?'<small>لسه ما اتنقلش للإكسيل</small>':''}<button class="b sm nb" data-note="${k.row}" aria-label="ملاحظة عن ${esc(k.name)}">${ic('note')}ملاحظة${k.pn.length?' ('+k.pn.length+')':''}</button></div><div class="seg" role="group" aria-label="حضور ${esc(k.name)}">${[['ح','ok','حاضر'],['غ','no','غائب'],['ع','ex','معتذر']].map(([v,cl,t])=>`<button class="sg ${cl}${k.att[j]===v?' on':''}" data-set="${v}" data-r="${k.row}" data-j="${j}" aria-pressed="${k.att[j]===v}" aria-label="${t}: ${esc(k.name)}">${v}</button>`).join('')}</div></div>`).join('')||`<div class="empty">${ic('users')}مفيش أطفال نشطين في الفصل ده</div>`}</div>`}
const renderRec=()=>{const e=$('#rl');if(e)e.innerHTML=recBody()};
const V={
rec(a,b){const c=sw(a),j=(b!=null&&D.dates[+b])?+b:defJ();autoAbsent(c,j);RC={c,j};
 return`<h2>${ic('check')}تسجيل الحضور</h2>${chips(c,'rec','/'+j)}<div class="tb"><select id="rd" aria-label="تاريخ الاجتماع">${D.dates.map((x,i)=>`<option value="${i}" ${i===j?'selected':''}>${fdl(x)}${i===D.now?' (الأسبوع ده)':''}</option>`).join('')}</select><a class="b" href="#sync">${ic('download')}نقل للإكسيل</a></div><p class="sub" style="margin-bottom:12px">كل الأطفال بيتسجلوا غايبين تلقائي (للتواريخ اللي جات)، وانت دوس ح لمن حضر أو ع للمعتذر. بيتحفظ فورًا على الجهاز ده. دوسة تانية على نفس الحرف بتمسحه.</p><div id="rl">${recBody()}</div>`},
sync(){CP={};let n=0;const cp=t=>{const id='c'+(++n);CP[id]=t;return id};
 const na=Object.keys(PEND.att).length,nn=PEND.notes.length,empty=`<div class="card empty">${ic('checkc')}مفيش حاجة مستنية</div>`;let att='';
 D.classes.forEach(c=>{const o=D.cls[c];D.dates.forEach((dt,j)=>{const ch=o.ks.filter(k=>k.pd[j]);if(!ch.length)return;const col=XLSX.utils.encode_col(9+j),ks=o.ks.slice(0,30);
  att+=`<div class="card sy rv"><div><h3>ملاك ${esc(c)} — ${fdl(dt)}</h3><p class="sub">شيت «حضور ${esc(c)}» · عمود <b dir="ltr">${col}</b> · الصق من الخلية <b dir="ltr">${col}7</b></p></div><div class="tl">${ch.map(k=>`<span class="${k.att[j]==='ح'?'ok':k.att[j]==='غ'?'no':k.att[j]==='ع'?'ex':''}">${esc(k.name)} · ${k.att[j]||'مسح'}</span>`).join('')}</div><button class="b pri" data-copy="${cp(ks.map(k=>k.att[j]).join('\n'))}">${ic('list')}نسخ عمود ${fd(dt)} (${ks.length} اسم)</button></div>`})});
 const kn=D.kids.filter(k=>k.pn.length);let nt='';
 if(kn.length){const B=RAW.wb.Sheets['قاعدة البيانات'],last=Math.max(...D.kids.map(k=>k.row)),rows=[];for(let r=6;r<=last;r++){const k=D.kids.find(x=>x.row===r);rows.push(q(k?comb(k):String(cell(B,r,12)??'')))}
  nt=`<div class="card sy rv"><div><h3>عمود الملاحظات كله</h3><p class="sub">شيت «قاعدة البيانات» · عمود <b dir="ltr">L</b> · الصق من <b dir="ltr">L6</b> (فيه القديم + الجديد)</p></div><button class="b pri" data-copy="${cp(rows.join('\n'))}">${ic('list')}نسخ العمود</button></div>`+kn.map(k=>`<div class="card sy rv"><div><h3>${esc(k.name)}</h3><p class="sub">${esc(k.cls)} · الخلية <b dir="ltr">L${k.row}</b></p></div><div class="pn">${k.pn.map(x=>`<p><span class="d">${esc(x.date)}</span> ${esc(x.text)}</p>`).join('')}</div><button class="b" data-copy="${cp(q(comb(k)))}">${ic('note')}نسخ الملاحظة كاملة</button></div>`).join('')}
 return`<h2>${ic('download')}نقل التسجيلات للإكسيل</h2><div class="card info rv"><span>حضور لسه ما اتنقلش: <b>${na}</b></span><span>ملاحظات: <b>${nn}</b></span></div>${na+nn?`<div class="card sy rv"><div><h3>${ic('download')}الطريقة السريعة</h3><p class="sub" style="line-height:1.9">حمّل الإكسيل وفيه كل التسجيلات والملاحظات متحطوطة في أماكنها (الصور والتنسيق زي ما هم)، وارفعه على GitHub بنفس الاسم. بعدها التسجيلات بتتشال من هنا لوحدها. ملحوظة: التسجيلات محفوظة على الجهاز والمتصفح ده بس.</p></div><button class="b pri" data-xlsx>${ic('download')}تحميل الإكسيل محدّث</button></div>${recCard()}<div class="note">${ic('alert')}<span>أو لو عايز تلصق بإيدك: انسخ من الأزرار اللي تحت والصق في الخلية المكتوبة.</span></div>`:''}<h2>${ic('calendar')}الحضور</h2>${att||empty}<h2>${ic('note')}الملاحظات</h2>${nt||empty}${na+nn?`<div style="margin-top:22px"><button class="b" data-clearall>${ic('close')}مسح كل التسجيلات</button></div>`:''}`},
home(){const a=D.kids.filter(k=>k.active),P=D.kids.reduce((s,k)=>s+k.P,0),Ab=D.kids.reduce((s,k)=>s+k.A,0),avg=P+Ab?P/(P+Ab):null;
 const st=[['users','المسجلين',D.kids.length],['checkc','النشطين',a.length],['user','أولاد',a.filter(k=>k.g==='ولد').length],['user','بنات',a.filter(k=>k.g==='بنت').length],['calendar','جمع مسجلة',D.meet],['alert','محتاجين افتقاد',D.kids.filter(k=>k.tier==3||k.tier==4).length,'warn']];
 let last=D.tot.map(x=>x.any).lastIndexOf(true);const mx=Math.max(1,...D.tot.map(x=>x.p));
 return`<section class="hero"><div class="card hero-main"><h1>${ic('sprout')}خدمة براعم</h1><p>نظام الافتقاد والحضور — كل أطفال الخدمة وفصولهم ومتابعتهم في مكان واحد.</p><div class="meta-row"><span class="chip">${ic('calendar')}العام الخدمي ${fdl(D.start)} ← ${fdl(D.end)}</span><span class="chip">${ic('clock')}اجتماع كل ${esc(D.dayName)}</span></div></div>
 <div class="card hero-side rv">${ring(avg,'var(--p)',132)}<div><b>متوسط الحضور</b><p class="sub">من ${D.meet} اجتماع مسجّل</p></div></div></section>
 ${pendBanner()}<div class="stats">${st.map(s=>`<div class="card st ${s[3]||''}"><div class="e">${ic(s[0])}</div><div><b data-count="${s[2]}">0</b><small>${s[1]}</small></div></div>`).join('')}</div>
 <h2>${ic('users')}الفصول</h2><div class="grid">${D.classes.map((c,i)=>{const o=D.cls[c];return`<div class="card ct rv" style="--c:${CC[i]}"><h3><span>ملاك ${esc(c)}</span>${ring(o.avg,CC[i],54)}</h3><ul><li>${ic('users')}${o.n} طفل</li><li>${ic('alert')}${o.need} محتاج افتقاد</li><li>${ic('user')}${o.boys} ولد · ${o.girls} بنت</li><li>${ic('plus')}${30-o.ks.length} مكان فاضي</li></ul><div class="row"><a class="b pri" href="#class/${encodeURIComponent(c)}">${ic('users')}الكروت</a><a class="b" href="#att/${encodeURIComponent(c)}">${ic('calendar')}الحضور</a></div></div>`}).join('')}</div>
 <h2>${ic('chart')}الحضور في كل اجتماع</h2><div class="card rv">${last<0?`<div class="empty">${ic('calendar')}لسه مفيش حضور مسجّل</div>`:`<div class="bars">${D.tot.slice(0,last+1).map((x,i)=>`<div class="bar ${i==D.now?'now':''}" title="${fdl(D.dates[i])}: ${x.p}" style="--h:${x.p/mx*78}%;--i:${i}"><b>${x.p}</b><i></i><span>${fd(D.dates[i])}</span></div>`).join('')}</div>`}</div>
 <h2>${ic('map')}اختصارات</h2><div style="display:flex;gap:10px;flex-wrap:wrap"><a class="b pri" href="#rec">${ic('check')}تسجيل الحضور</a><a class="b" href="#sync">${ic('download')}للإكسيل</a><a class="b" href="#kids">${ic('list')}قاعدة البيانات</a><a class="b" href="#follow">${ic('phone')}الافتقاد</a><a class="b" href="#settings">${ic('settings')}الإعدادات</a><a class="b" href="#guide">${ic('book')}الدليل</a><a class="b" href="${FILE}" download>${ic('plus')}إضافة طفل (حمّل الإكسيل)</a></div>
 ${LM?`<p class="sub" style="margin-top:26px;text-align:center">آخر تحديث للبيانات: ${LM}</p>`:''}`},
kids(){return`<h2>${ic('list')}قاعدة بيانات الأطفال</h2><div class="tb"><input id="q" type="search" aria-label="بحث" placeholder="ابحث بالاسم أو الرقم أو العنوان…"><select id="fc" aria-label="الفصل"><option value="">كل الفصول</option>${D.classes.map(c=>`<option>${esc(c)}</option>`).join('')}</select><select id="fg" aria-label="النوع"><option value="">أولاد وبنات</option><option>ولد</option><option>بنت</option></select><select id="fs" aria-label="الحالة"><option value="">كل الحالات</option><option value="1">نشط</option><option value="0">غير نشط</option></select><select id="fy" aria-label="السنة الدراسية"><option value="">كل السنين</option>${D.years.map(y=>`<option>${esc(y)}</option>`).join('')}</select></div><p id="cn" class="sub" style="margin-bottom:12px" aria-live="polite"></p><div class="grid" id="kg"></div>`},
class(a){const c=sw(a),o=D.cls[c];return`<h2>${ic('users')}كروت أطفال ملاك ${esc(c)}</h2>${chips(c,'class')}<div class="card info rv"><span>العدد: <b>${o.n}</b></span><span>${o.boys} ولد · ${o.girls} بنت</span><span>متوسط الحضور: <b>${pc(o.avg)}</b></span><span>جمع مسجلة: <b>${o.meet}</b></span><span>محتاجين افتقاد: <b>${o.need}</b></span><span>أماكن فاضية: <b>${30-o.ks.length}</b></span></div><div class="grid">${o.ks.map(card).join('')||`<div class="empty">${ic('users')}مفيش أطفال في الفصل ده</div>`}</div>`},
att(a){const c=sw(a),o=D.cls[c],dt=D.dates;return`<h2>${ic('calendar')}قايمة حضور ملاك ${esc(c)}</h2>${chips(c,'att')}<div class="card info rv"><span>جمع مسجلة: <b>${o.meet}</b></span><span>المتوسط: <b>${pc(o.avg)}</b></span><span class="lgd"><i class="lg-ok">ح</i>حاضر</span><span class="lgd"><i class="lg-no">غ</i>غائب</span><span class="lgd"><i class="lg-ex">ع</i>معتذر</span></div>
 <div class="card tw rv"><table><thead><tr><th>الاسم</th><th>حضور</th><th>غياب</th><th>بعذر</th><th>النسبة</th><th>غياب متتالي</th><th>آخر حضور</th><th>الحالة</th>${dt.map((x,i)=>`<th class="${i==D.now?'now':''}" ${i==D.now?'id="nowc"':''}>${fd(x)}</th>`).join('')}</tr></thead><tbody>${o.ks.map(k=>`<tr><td><a href="javascript:void 0" data-open="${k.row}">${esc(k.name)}</a></td><td>${k.P}</td><td>${k.A}</td><td>${k.E}</td><td>${pc(k.pct)}</td><td>${k.streak}</td><td>${k.last?fd(k.last):'—'}</td><td style="color:${TIER[k.tier][2]}">${TIER[k.tier][1]}</td>${k.att.map((v,i)=>`<td class="${v==='ح'?'ok':v==='غ'?'no':v==='ع'?'ex':''} ${i==D.now?'now':''} ${k.pd[i]?'pd':''}">${v}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr><td>الحاضرين</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.any?x.p:''}</td>`).join('')}</tr><tr><td>النسبة</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.t?Math.round(x.p/x.t*100)+'%':''}</td>`).join('')}</tr></tfoot></table></div>`},
follow(){const l=D.kids.filter(k=>k.tier==3||k.tier==4).sort((a,b)=>b.streak-a.streak||(a.pct??0)-(b.pct??0));
 return`<h2>${ic('phone')}قايمة الافتقاد</h2><div class="card info rv"><span>الأطفال اللي نسبة حضورهم أقل من <b>${Math.round(D.th.ex*100)}%</b> أو غايبين <b>${D.th.streak}</b> اجتماعات ورا بعض أو أكتر — العدد: <b>${l.length}</b></span></div><div class="fl">${l.map((k,i)=>`<div class="card fi rv" style="--c:${TIER[k.tier][2]}"><div class="rank">${i+1}</div><div><h3>${esc(k.name)}</h3><p class="meta">${esc(k.cls)} · نسبة ${pc(k.pct)} · غياب متتالي ${k.streak} · آخر حضور ${k.last?fdl(k.last):'—'}<br>${ic('pin')} ${esc(k.addr||'—')}</p>${tag(TIER[k.tier])}</div><div class="acts">${btns(k)||'<span class="sub">— مفيش رقم —</span>'}<button class="b sm pri" data-open="${k.row}">التفاصيل</button></div></div>`).join('')||`<div class="card empty">${ic('checkc')}مفيش حد محتاج افتقاد دلوقتي</div>`}</div>`},
settings(){const k=(a,b)=>`<div><small>${a}</small><b>${b}</b></div>`;return`<h2>${ic('settings')}الإعدادات</h2><div class="card kv rv">${k('بداية العام الخدمي',fdl(D.start))+k('نهاية العام الخدمي',fdl(D.end))+k('يوم الاجتماع',esc(D.dayName))+k('«ممتاز» من',Math.round(D.th.ex*100)+'%')+k('«جيد» من',Math.round(D.th.good*100)+'%')+k('افتقاد ضروري بعد غياب',D.th.streak+' اجتماعات')+k('الفصول (الملائكة)',D.classes.map(esc).join(' · '))+k('السنين الدراسية',D.years.map(esc).join(' · '))}</div><div class="note">${ic('alert')}<span>الإعدادات بتتغيّر من شيت «الإعدادات» في الإكسيل، والموقع بيقراها أوتوماتيك. ماتغيّرش أسماء الفصول بعد تسجيل الأطفال.</span></div>`},
guide(){const s=[['plus','إضافة طفل جديد','افتح ملف الإكسيل وأضف الطفل في أول صف فاضي في شيت «قاعدة البيانات» (الاسم، النوع، السنة، الفصل، التليفونات…). بعدها ارفع الملف على GitHub بنفس الاسم، والطفل يظهر في الموقع لوحده.'],['camera','صورة الطفل','حطّ الصورة في عمود «الصورة» بـ Insert ← Pictures ← Place in Cell. بتظهر في كارت الطفل، ولو مفيش صورة بيظهر أول حرف من الاسم.'],['calendar','تسجيل الحضور','سجّل ح / غ / ع في شيت «حضور …» الخاص بالفصل، ثم ارفع الملف. النسب والألوان بتتحسب في الموقع أوتوماتيك.'],['phone','الافتقاد','القايمة بتتجمّع أوتوماتيك حسب الحدود في «الإعدادات». اضغط على زر الاتصال أو واتساب للتواصل مع الأم أو الأب.'],['ban','طفل غير نشط','بدل ما تمسح الصف، غيّر «الحالة» لـ «غير نشط». ماتعملش Sort ولا تمسح صفوف.'],['check','تسجيل الحضور من الموقع','من تبويب «تسجيل» اختار الفصل والتاريخ وادوس ح / غ / ع جنب كل طفل. بيتحفظ فورًا على الجهاز ده حتى من غير ما تفتح الإكسيل. ولو فتحت ملف الطفل تقدر تضيف ملاحظة بتاريخها.'],['download','نقل التسجيلات للإكسيل','افتح صفحة «للإكسيل»، انسخ عمود كل تاريخ (أو الملاحظة) والصقه في الخلية المكتوبة، احفظ الملف وارفعه. أول ما الموقع يقرا الملف الجديد بيشيل التسجيلات اللي اتنقلت لوحده. التسجيلات محفوظة على نفس الجهاز والمتصفح بس.'],['refresh','تحديث','بعد رفع الملف الجديد اضغط زر التحديث في أعلى الصفحة (أو افتح الموقع من جديد) — قد يستغرق GitHub دقيقة أو اثنين.']];return`<h2>${ic('book')}دليل الاستخدام</h2><div class="steps">${s.map(x=>`<div class="card rv"><div class="e">${ic(x[0])}</div><h3>${x[1]}</h3><p>${x[2]}</p></div>`).join('')}</div>`}};

function renderKids(){const q=$('#q').value.trim().toLowerCase(),f=id=>$(id).value;
 const l=D.kids.filter(k=>(!f('#fc')||k.cls===f('#fc'))&&(!f('#fg')||k.g===f('#fg'))&&(!f('#fy')||k.year===f('#fy'))&&(!f('#fs')||k.active==(f('#fs')==='1'))&&(!q||[k.name,k.mom,k.dad,k.addr,k.id,k.notes].join(' ').toLowerCase().includes(q)));
 $('#cn').textContent='النتائج: '+l.length+' من '+D.kids.length;$('#kg').innerHTML=l.map(card).join('')||`<div class="empty">${ic('list')}مفيش نتائج — جرّب تغيّر البحث أو الفلاتر</div>`;fx($('#kg'))}

function openKid(row,again,fn){const k=D.kids.find(x=>x.row==row);if(!k)return;const t=TIER[k.tier],f=(a,b)=>b!==''&&b!=null?`<div><small>${a}</small><b>${b}</b></div>`:'';
 if(!again)lastFocus=document.activeElement;
 $('#mc').innerHTML=`<div class="mh in"><div class="av ${k.g==='بنت'?'f':''}">${av(k)}</div><div><h2>${esc(k.name)}</h2><p class="meta">${k.id} · ${esc(k.cls)} · ${esc(k.year)}</p>${tag(t)}</div></div>
 <div class="kv">${f('حضور',k.P)+f('غياب',k.A)+f('بعذر',k.E)+f('النسبة',pc(k.pct))+f('غياب متتالي',k.streak)+f('آخر حضور',k.last?fdl(k.last):'—')+f('يروح لوحده؟',esc(k.alone))+f('النوع',k.g==='بنت'?'بنت':'ولد')}</div>
 ${k.addr?`<div class="mp">${ic('pin')}<span>${esc(k.addr)}</span><a class="b sm" target="_blank" rel="noopener" href="https://www.google.com/maps/search/${encodeURIComponent(k.addr)}">${ic('map')}الخريطة</a></div>`:''}
 ${k.notes?`<p class="mp">${ic('note')}<span>${esc(k.notes)}</span></p>`:''}<div style="display:flex;gap:6px;flex-wrap:wrap">${btns(k)}</div>
 <div class="nt"><h3>${ic('note')}ملاحظات جديدة</h3>${k.pn.map(n=>`<div class="pnote"><span class="d">${esc(n.date)}</span><span>${esc(n.text)}</span><button class="b sm" data-delnote="${n.id}" data-r="${k.row}" aria-label="حذف الملاحظة">${ic('close')}</button></div>`).join('')}<div class="addn"><input type="date" id="nd" value="${dayKey()}" aria-label="تاريخ الملاحظة"><textarea id="ntx" rows="2" placeholder="اكتب ملاحظة…" aria-label="الملاحظة"></textarea><button class="b pri" data-addnote="${k.row}">${ic('plus')}إضافة ملاحظة</button></div><p class="sub" style="margin-top:8px">بتتحفظ على الجهاز ده لحد ما تنقلها للإكسيل من صفحة «للإكسيل».</p></div>
 <div class="tl">${k.att.map((v,i)=>v?`<span class="${v==='ح'?'ok':v==='غ'?'no':'ex'}${k.pd[i]?' pd':''}">${fd(D.dates[i])} ${v}</span>`:'').join('')}</div>`;
 $('#modal').hidden=false;document.body.style.overflow='hidden';if(fn){const t=$('#ntx');t.scrollIntoView({block:'center'});t.focus()}else if(!again)$('#mx').focus()}
const closeM=()=>{if($('#modal').hidden)return;$('#modal').hidden=true;document.body.style.overflow='';if($('#rl')){renderRec();return}if(lastFocus&&lastFocus.focus)lastFocus.focus()};

const io=new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}}),{threshold:.08});
function fx(r=view){$$('.rv:not(.in)',r).forEach(e=>io.observe(e));
 $$('[data-count]',r).forEach(e=>{const to=+e.dataset.count,s=e.dataset.suf||'',t0=performance.now();(function f(t){const p=Math.min(1,(t-t0)/1100);e.textContent=Math.round(to*(1-Math.pow(1-p,4)))+s;p<1&&requestAnimationFrame(f)})(t0)})}
function route(){const [r,a,b]=decodeURIComponent(location.hash.slice(1)||'home').split('/'),n=V[r]?r:'home';
 view.innerHTML=`<div class="page">${V[n](a,b)}</div>`;scrollTo(0,0);fx();
 if(n==='kids'){['#q','#fc','#fg','#fs','#fy'].forEach(i=>$(i).addEventListener('input',renderKids));renderKids()}
 if(n==='rec')$('#rd').addEventListener('change',e=>{location.hash='#rec/'+encodeURIComponent(RC.c)+'/'+e.target.value});
 if(n==='att')setTimeout(()=>{const c=$('#nowc');c&&c.scrollIntoView({inline:'center',block:'nearest',behavior:'smooth'})},500);
 const href='#'+n;$$('#nav a,#tabs a').forEach(x=>{const on=x.getAttribute('href')===href;x.classList.toggle('on',on);on?x.setAttribute('aria-current','page'):x.removeAttribute('aria-current')})}
function init(){
 $('#nav').innerHTML=NAV.map(n=>`<a href="#${n[0]}">${ic(n[1])}${n[2]}</a>`).join('');
 $('#tabs').innerHTML=NAV.filter(n=>!n[3]).map(n=>`<a href="#${n[0]}">${ic(n[1])}<span>${n[2]}</span></a>`).join('');
 $('#dl').href=FILE;addEventListener('hashchange',route)}
document.addEventListener('click',e=>{const o=e.target.closest('[data-open]');if(o)openKid(o.dataset.open);if(e.target.id==='modal'||e.target.closest('#mx'))closeM()});
addEventListener('keydown',e=>e.key==='Escape'&&closeM());
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
unlocked()?start():showGate();

document.addEventListener('click',e=>{
 const s=e.target.closest('[data-set]');
 if(s){const k=D.kids.find(x=>x.row==s.dataset.r),j=+s.dataset.j,key=k.cls+'|'+k.name+'|'+dk(D.dates[j]);let v=s.dataset.set;if(k.att[j]===v)v='';
  if(v===k.xl[j])delete PEND.att[key];else PEND.att[key]={v,t:Date.now()};savePend();build();renderRec();return}
 const an=e.target.closest('[data-addnote]');
 if(an){const k=D.kids.find(x=>x.row==an.dataset.addnote),t=$('#ntx').value.trim();if(!t){$('#ntx').focus();return}
  PEND.notes.push({id:Date.now().toString(36)+Math.random().toString(36).slice(2,5),key:k.cls+'|'+k.name,date:noteDate($('#nd').value||dayKey()),text:t,t:Date.now()});savePend();build();openKid(k.row,true);return}
 const dn=e.target.closest('[data-delnote]');
 if(dn){PEND.notes=PEND.notes.filter(n=>n.id!==dn.dataset.delnote);savePend();build();openKid(dn.dataset.r,true);return}
 const cb=e.target.closest('[data-copy]');
 if(cb){copyTxt(CP[cb.dataset.copy]).then(()=>{const o=cb.innerHTML;cb.innerHTML=ic('check')+'تم النسخ';setTimeout(()=>cb.innerHTML=o,1600)});return}
 if(e.target.closest('[data-clearall]')&&confirm('تمسح كل الحضور والملاحظات اللي لسه ما اتنقلتش للإكسيل؟')){PEND={att:{},notes:[]};savePend();build();route()}
});

/* ---------- تحميل الإكسيل وفيه التسجيلات (بيعدّل الخلايا بس ويسيب باقي الملف زي ما هو) ---------- */
const xe=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
const colNum=ref=>{let n=0;for(const ch of ref.match(/^[A-Z]+/)[0])n=n*26+ch.charCodeAt(0)-64;return n};
function setCell(xml,ref,val){
 const m=xml.match(new RegExp('<c r="'+ref+'"(?:\\s[^>]*?)?(?:/>|>[\\s\\S]*?</c>)')),sm=m&&m[0].match(/^<c[^>]*?\ss="(\d+)"/),sa=sm?' s="'+sm[1]+'"':'';
 const nc=val===''?'<c r="'+ref+'"'+sa+'/>':'<c r="'+ref+'"'+sa+' t="inlineStr"><is><t xml:space="preserve">'+xe(val)+'</t></is></c>';
 if(m)return xml.replace(m[0],()=>nc);
 const rn=ref.match(/\d+/)[0],rm=xml.match(new RegExp('<row r="'+rn+'"[^>]*?[^/]>([\\s\\S]*?)</row>'));if(!rm)return null;
 let at=-1,cn=colNum(ref),re=/<c r="([A-Z]+)\d+"/g,t;while((t=re.exec(rm[1]))){if(colNum(t[1])>cn){at=t.index;break}}
 const inner=at<0?rm[1]+nc:rm[1].slice(0,at)+nc+rm[1].slice(at);
 return xml.replace(rm[0],()=>rm[0].replace(rm[1],()=>inner));
}
async function exportXlsx(btn){
 const old=btn.innerHTML;btn.disabled=true;btn.textContent='جاري التجهيز…';
 try{
  const z=await JSZip.loadAsync(RAW.buf),rd=p=>z.file(p).async('string'),px=async p=>new DOMParser().parseFromString(await rd(p),'text/xml');
  const wbx=await px('xl/workbook.xml'),rels=await px('xl/_rels/workbook.xml.rels');
  const path=name=>{const s=[...wbx.getElementsByTagName('sheet')].find(s=>s.getAttribute('name')===name);if(!s)throw Error('مش لاقي شيت '+name);
   const r=[...rels.getElementsByTagName('Relationship')].find(r=>r.getAttribute('Id')===s.getAttribute('r:id'));return 'xl/'+r.getAttribute('Target').replace(/^\/?(xl\/)?/,'')};
  const ed={};const add=(sh,ref,v)=>(ed[sh]=ed[sh]||[]).push([ref,v]);
  D.classes.forEach(c=>D.cls[c].ks.slice(0,30).forEach((k,i)=>k.pd.forEach((p,j)=>{if(p)add('حضور '+c,XLSX.utils.encode_col(9+j)+(7+i),k.att[j])})));
  D.kids.filter(k=>k.pn.length).forEach(k=>add('قاعدة البيانات','L'+k.row,comb(k)));
  let done=0,skip=0;
  for(const sh in ed){const p=path(sh);let x=await rd(p);for(const [ref,v] of ed[sh]){const r=setCell(x,ref,v);if(r==null)skip++;else{x=r;done++}}z.file(p,x,{createFolders:false})}
  let w=await rd('xl/workbook.xml');if(!/fullCalcOnLoad/.test(w))w=w.replace('<calcPr ','<calcPr fullCalcOnLoad="1" ');z.file('xl/workbook.xml',w,{createFolders:false});
  const blob=await z.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const u=URL.createObjectURL(blob),l=document.createElement('a');l.href=u;l.download=FILE;document.body.appendChild(l);l.click();l.remove();setTimeout(()=>URL.revokeObjectURL(u),4000);
  if(skip)alert('اتحطّ '+done+' خانة، و'+skip+' خانة ما اتحطتش — استخدم النسخ اليدوي ليهم.');
  btn.innerHTML=old;
 }catch(err){alert('حصلت مشكلة في تجهيز الملف: '+err.message+'\nاستخدم النسخ اليدوي.');btn.innerHTML=old}
 btn.disabled=false}
document.addEventListener('click',e=>{const b=e.target.closest('[data-xlsx]');if(b)exportXlsx(b)});

/* ---------- ملف مستقل بالأسماء والحضور والتواريخ (على الجهاز أو GitHub) ---------- */
const GHK='b3_gh';
const ghCfg=()=>{try{return JSON.parse(localStorage.getItem(GHK))||{}}catch(e){return{}}};
const ghGuess=()=>location.hostname.endsWith('.github.io')?location.hostname.split('.')[0]+'/'+(location.pathname.split('/')[1]||location.hostname):'';
const stamp=()=>{const t=cairoNow(),h=new Intl.DateTimeFormat('en-GB',{timeZone:'Africa/Cairo',hour:'2-digit',minute:'2-digit',hour12:false}).format(new Date()).replace(':','-');return t.year+'-'+t.month+'-'+t.day+'_'+h};
function rosterData(){const today=dayKey(),js=D.dates.map((_,j)=>j).filter(j=>D.kids.some(k=>k.pd[j]));
 return{js,val:(k,j)=>k.att[j]||(k.active&&dk(D.dates[j])<=today&&D.cls[k.cls].ks.some(x=>x.pd[j])?'غ':''),hdr:j=>fdl(D.dates[j])+' · عمود '+XLSX.utils.encode_col(9+j),nt:k=>k.pn.map(n=>n.date+': '+n.text).join(' | ')}}
function recordsCSV(){const R=rosterData(),cs=v=>'"'+String(v).replace(/"/g,'""')+'"',rows=[['الفصل','الاسم',...R.js.map(R.hdr),'ملاحظات جديدة']];
 D.classes.forEach(c=>D.cls[c].ks.slice(0,30).forEach(k=>rows.push([c,k.name,...R.js.map(j=>R.val(k,j)),R.nt(k)])));
 return'﻿'+rows.map(r=>r.map(cs).join(',')).join('\r\n')}
function recordsWB(){const R=rosterData(),wb=XLSX.utils.book_new(),tm=t=>t?new Date(t).toLocaleString('en-GB',{timeZone:'Africa/Cairo',hour12:false}):'';
 const sheet=(rows,w,name)=>{const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=w.map(x=>({wch:x}));XLSX.utils.book_append_sheet(wb,ws,name.slice(0,31))};
 D.classes.forEach(c=>sheet([['الاسم',...R.js.map(R.hdr),'ملاحظات جديدة'],...D.cls[c].ks.slice(0,30).map(k=>[k.name,...R.js.map(j=>R.val(k,j)),R.nt(k)])],[34,...R.js.map(()=>24),50],'حضور '+c));
 sheet([['تاريخ الملاحظة','الفصل','الاسم','الملاحظة','وقت التسجيل'],...PEND.notes.map(n=>{const i=n.key.indexOf('|');return[n.date,n.key.slice(0,i),n.key.slice(i+1),n.text,tm(n.t)]})],[14,12,34,60,20],'الملاحظات');
 sheet([['كل شيت «حضور …» فيه كل أطفال الفصل بنفس ترتيب شيت الحضور في الإكسيل.'],['اللي ما اتسجلش حضوره اتحسب غايب (غ). ح = حاضر، ع = معتذر.'],['انسخ عمود التاريخ (من تحت العنوان) والصقه في شيت «حضور <الفصل>» في العمود المكتوب في العنوان، بادئ من الصف 7.'],['الملاحظات الجديدة في آخر عمود وفي شيت «الملاحظات» (الصقها في عمود L في قاعدة البيانات).']],[110],'تعليمات');
 wb.Workbook={Views:[{RTL:true}]};return wb}
function recCard(){const c=ghCfg(),ok=c.repo&&c.token;
 return`<div class="card sy rv"><div><h3>${ic('list')}ملف مستقل بالأسماء والحضور والتواريخ</h3><p class="sub" style="line-height:1.9">ملف فيه كل أطفال كل فصل بنفس ترتيب شيت الحضور، وعمود لكل تاريخ سجلته (اللي ما اتسجلش حاضر بيتحسب غايب)، جاهز تنسخ العمود وتلصقه في شيت الحضور في الإكسيل. الملاحظات الجديدة بتاريخها في آخر عمود. تحفظه على الجهاز أو ترفعه على GitHub في فولدر <b dir="ltr">records</b>.</p></div><div style="display:flex;gap:8px;flex-wrap:wrap"><button class="b pri" data-rec="dev">${ic('download')}حفظ على الجهاز (إكسيل)</button><button class="b" data-rec="csv">${ic('download')}ملف CSV عادي</button><button class="b" data-rec="gh">${ic('upload')}رفع على GitHub</button></div><p id="ghs" class="sub" role="status" aria-live="polite"></p>
 <details id="ghd"><summary class="sub" style="cursor:pointer">${ok?'إعداد GitHub متظبط ✓ (تعديل)':'إعداد GitHub (مرة واحدة)'}</summary><div class="addn"><input id="ghr" dir="ltr" placeholder="اسم-الحساب/اسم-الريبو" value="${esc(c.repo||ghGuess())}" aria-label="الريبو"><input id="ght" type="password" dir="ltr" autocomplete="off" placeholder="${c.token?'التوكن محفوظ — اكتب جديد لتغييره':'التوكن (Personal access token)'}" aria-label="التوكن"><button class="b" data-ghsave>حفظ الإعداد</button>${ok?'<button class="b" data-ghclear>مسح الإعداد</button>':''}<p class="sub" style="line-height:1.9">GitHub ← Settings ← Developer settings ← Fine-grained tokens ← Generate: اختار الريبو ده بس، وصلاحية <b dir="ltr">Contents: Read and write</b>. التوكن بيتحفظ في المتصفح ده بس، ماتحطوش في أي ملف على الموقع.</p></div></details></div>`}
document.addEventListener('click',async e=>{
 const r=e.target.closest('[data-rec]');
 if(r){const name='bara3em_'+stamp()+'.xlsx',msg=$('#ghs');
  if(!Object.keys(PEND.att).length&&!PEND.notes.length){msg.textContent='مفيش تسجيلات لسه.';return}
  if(r.dataset.rec==='csv'){const u=URL.createObjectURL(new Blob([recordsCSV()],{type:'text/csv;charset=utf-8'})),l=document.createElement('a');l.href=u;l.download=name.replace('.xlsx','.csv');document.body.appendChild(l);l.click();l.remove();setTimeout(()=>URL.revokeObjectURL(u),4000);msg.textContent='اتحفظ على الجهاز: '+name.replace('.xlsx','.csv');return}
  if(r.dataset.rec==='dev'){XLSX.writeFile(recordsWB(),name);msg.textContent='اتحفظ على الجهاز: '+name;return}
  const c=ghCfg();if(!c.repo||!c.token){$('#ghd').open=true;msg.textContent='ظبّط إعداد GitHub الأول (تحت).';return}
  r.disabled=true;msg.textContent='جاري الرفع…';
  try{const body=XLSX.write(recordsWB(),{bookType:'xlsx',type:'base64'});
   const res=await fetch('https://api.github.com/repos/'+c.repo+'/contents/records/'+name,{method:'PUT',headers:{'Authorization':'Bearer '+c.token,'Accept':'application/vnd.github+json','Content-Type':'application/json'},body:JSON.stringify({message:'تسجيلات '+stamp(),content:body})});
   if(res.ok)msg.textContent='اترفع على GitHub: records/'+name;
   else{const why=res.status===401?'التوكن غلط أو منتهي':res.status===403?'التوكن مالوش صلاحية كتابة':res.status===404?'اسم الريبو غلط أو التوكن مش مديله صلاحية عليه':'خطأ '+res.status;msg.textContent='ما اترفعش: '+why+'. تقدر تحفظ على الجهاز بدل كده.'}
  }catch(err){msg.textContent='ما اترفعش (النت؟). احفظ على الجهاز بدل كده.'}
  r.disabled=false;return}
 if(e.target.closest('[data-ghsave]')){const c=ghCfg(),repo=$('#ghr').value.trim(),tk=$('#ght').value.trim()||c.token||'';
  if(!/^[\w.-]+\/[\w.-]+$/.test(repo)){$('#ghs').textContent='اكتب الريبو بالشكل: اسم-الحساب/اسم-الريبو';return}
  try{localStorage.setItem(GHK,JSON.stringify({repo,token:tk}))}catch(_){}route();$('#ghs').textContent='اتحفظ الإعداد ✓';return}
 if(e.target.closest('[data-ghclear]')){try{localStorage.removeItem(GHK)}catch(_){}route()}
});

document.addEventListener('click',e=>{const n=e.target.closest('[data-note]');if(n)openKid(n.dataset.note,false,true)});
