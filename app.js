const FILE='Bara3em_System_Final.xlsx';
const $=(s,e=document)=>e.querySelector(s),$$=(s,e=document)=>[...e.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const ic=n=>`<svg class="ic" aria-hidden="true"><use href="#i-${n}"/></svg>`;
// [أيقونة، الوصف، اللون]
const TIER=[['clock','لسه مفيش حضور','var(--none)'],['star','ممتاز','var(--ok)'],['checkc','جيد','var(--good)'],['alert','محتاج متابعة','var(--warn)'],['phone','افتقاده ضروري','var(--bad)'],['ban','غير نشط','var(--off)']];
const CC=['var(--c1)','var(--c2)','var(--c3)','var(--c4)','var(--c5)','var(--c6)'];
const NAV=[['home','home','الرئيسية'],['kids','list','الأطفال'],['class','users','الفصول'],['att','calendar','الحضور'],['follow','phone','الافتقاد'],['settings','settings','الإعدادات',1],['guide','book','الدليل',1]];
const cell=(ws,r,c)=>{const x=ws&&ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:undefined};
const ds=n=>new Date(Math.round((n-25569)*864e5)),fd=d=>d.getUTCDate()+'/'+(d.getUTCMonth()+1),fdl=d=>fd(d)+'/'+d.getUTCFullYear();
const ph=v=>{if(v==null||v==='')return'';let s=String(v).replace(/\D/g,'');if(s.length===10&&s[0]!=='0')s='0'+s;return s};
const wa=p=>'https://wa.me/20'+p.slice(-10),pc=p=>p==null?'—':Math.round(p*100)+'%';
let D,view=$('#view'),LM='',lastFocus=null;

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
 const buf=await res.arrayBuffer(),wb=XLSX.read(buf,{type:'array'}),photos=await photosOf(buf);
 const S=wb.Sheets['الإعدادات'],B=wb.Sheets['قاعدة البيانات'];
 const classes=[],days=[],years=[];
 for(let r=4;r<=8;r++){const v=cell(S,r,2);if(typeof v==='string')classes.push(v)}
 for(let r=4;r<=10;r++){const v=cell(S,r,8);if(v)days.push(v);const y=cell(S,r,4);if(y)years.push(y)}
 const th={ex:cell(S,16,2),good:cell(S,17,2),streak:cell(S,18,2)},dayName=cell(S,12,2),want=days.indexOf(dayName)<0?5:days.indexOf(dayName);
 const end=ds(cell(S,11,2));let d=ds(cell(S,10,2));while(d.getUTCDay()!==want)d=new Date(+d+864e5);
 const dates=[];while(d<=end&&dates.length<53){dates.push(new Date(d));d=new Date(+d+6048e5)}
 const kids=[],cnt={};
 for(let r=6;r<=305;r++){const name=cell(B,r,3);if(!name)continue;
  const cls=String(cell(B,r,6)??''),i=cnt[cls]=(cnt[cls]||0)+1,A=wb.Sheets['حضور '+cls];
  const k={row:r,id:'B-'+String(r-5).padStart(3,'0'),name:String(name).trim(),g:cell(B,r,4)||'',year:cell(B,r,5)||'',cls,alone:cell(B,r,7)||'',mom:ph(cell(B,r,8)),dad:ph(cell(B,r,9)),addr:cell(B,r,10)||'',notes:cell(B,r,12)||'',active:cell(B,r,13)!=='غير نشط',photo:photos[r]};
  k.att=dates.map((_,c)=>(A&&i<=30&&cell(A,6+i,10+c))||'');
  let P=0,Ab=0,E=0,last=-1;k.att.forEach((v,j)=>{if(v==='ح'){P++;last=j}else if(v==='غ')Ab++;else if(v==='ع')E++});
  Object.assign(k,{P,A:Ab,E,pct:P+Ab?P/(P+Ab):null,streak:k.att.filter((v,j)=>v==='غ'&&j>last).length,last:last>=0?dates[last]:null});
  k.tier=!k.active?5:k.pct==null?0:k.streak>=th.streak?4:k.pct>=th.ex?1:k.pct>=th.good?2:3;kids.push(k)}
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
const chips=(sel,pre)=>`<div class="chips">${D.classes.map((c,i)=>`<a href="#${pre}/${encodeURIComponent(c)}" class="${c===sel?'on':''}" style="--c:${CC[i]}">${esc(c)}</a>`).join('')}</div>`;
const sw=(a,c)=>D.classes.includes(a)?a:D.classes[0];

const V={
home(){const a=D.kids.filter(k=>k.active),P=D.kids.reduce((s,k)=>s+k.P,0),Ab=D.kids.reduce((s,k)=>s+k.A,0),avg=P+Ab?P/(P+Ab):null;
 const st=[['users','المسجلين',D.kids.length],['checkc','النشطين',a.length],['user','أولاد',a.filter(k=>k.g==='ولد').length],['user','بنات',a.filter(k=>k.g==='بنت').length],['calendar','جمع مسجلة',D.meet],['alert','محتاجين افتقاد',D.kids.filter(k=>k.tier==3||k.tier==4).length,'warn']];
 let last=D.tot.map(x=>x.any).lastIndexOf(true);const mx=Math.max(1,...D.tot.map(x=>x.p));
 return`<section class="hero"><div class="card hero-main"><h1>${ic('sprout')}خدمة براعم</h1><p>نظام الافتقاد والحضور — كل أطفال الخدمة وفصولهم ومتابعتهم في مكان واحد.</p><div class="meta-row"><span class="chip">${ic('calendar')}العام الخدمي ${fdl(D.start)} ← ${fdl(D.end)}</span><span class="chip">${ic('clock')}اجتماع كل ${esc(D.dayName)}</span></div></div>
 <div class="card hero-side rv">${ring(avg,'var(--p)',132)}<div><b>متوسط الحضور</b><p class="sub">من ${D.meet} اجتماع مسجّل</p></div></div></section>
 <div class="stats">${st.map(s=>`<div class="card st ${s[3]||''}"><div class="e">${ic(s[0])}</div><div><b data-count="${s[2]}">0</b><small>${s[1]}</small></div></div>`).join('')}</div>
 <h2>${ic('users')}الفصول</h2><div class="grid">${D.classes.map((c,i)=>{const o=D.cls[c];return`<div class="card ct rv" style="--c:${CC[i]}"><h3><span>ملاك ${esc(c)}</span>${ring(o.avg,CC[i],54)}</h3><ul><li>${ic('users')}${o.n} طفل</li><li>${ic('alert')}${o.need} محتاج افتقاد</li><li>${ic('user')}${o.boys} ولد · ${o.girls} بنت</li><li>${ic('plus')}${30-o.ks.length} مكان فاضي</li></ul><div class="row"><a class="b pri" href="#class/${encodeURIComponent(c)}">${ic('users')}الكروت</a><a class="b" href="#att/${encodeURIComponent(c)}">${ic('calendar')}الحضور</a></div></div>`}).join('')}</div>
 <h2>${ic('chart')}الحضور في كل اجتماع</h2><div class="card rv">${last<0?`<div class="empty">${ic('calendar')}لسه مفيش حضور مسجّل</div>`:`<div class="bars">${D.tot.slice(0,last+1).map((x,i)=>`<div class="bar ${i==D.now?'now':''}" title="${fdl(D.dates[i])}: ${x.p}" style="--h:${x.p/mx*78}%;--i:${i}"><b>${x.p}</b><i></i><span>${fd(D.dates[i])}</span></div>`).join('')}</div>`}</div>
 <h2>${ic('map')}اختصارات</h2><div style="display:flex;gap:10px;flex-wrap:wrap"><a class="b pri" href="#kids">${ic('list')}قاعدة البيانات</a><a class="b" href="#follow">${ic('phone')}الافتقاد</a><a class="b" href="#settings">${ic('settings')}الإعدادات</a><a class="b" href="#guide">${ic('book')}الدليل</a><a class="b" href="${FILE}" download>${ic('plus')}إضافة طفل (حمّل الإكسيل)</a></div>
 ${LM?`<p class="sub" style="margin-top:26px;text-align:center">آخر تحديث للبيانات: ${LM}</p>`:''}`},
kids(){return`<h2>${ic('list')}قاعدة بيانات الأطفال</h2><div class="tb"><input id="q" type="search" aria-label="بحث" placeholder="ابحث بالاسم أو الرقم أو العنوان…"><select id="fc" aria-label="الفصل"><option value="">كل الفصول</option>${D.classes.map(c=>`<option>${esc(c)}</option>`).join('')}</select><select id="fg" aria-label="النوع"><option value="">أولاد وبنات</option><option>ولد</option><option>بنت</option></select><select id="fs" aria-label="الحالة"><option value="">كل الحالات</option><option value="1">نشط</option><option value="0">غير نشط</option></select><select id="fy" aria-label="السنة الدراسية"><option value="">كل السنين</option>${D.years.map(y=>`<option>${esc(y)}</option>`).join('')}</select></div><p id="cn" class="sub" style="margin-bottom:12px" aria-live="polite"></p><div class="grid" id="kg"></div>`},
class(a){const c=sw(a),o=D.cls[c];return`<h2>${ic('users')}كروت أطفال ملاك ${esc(c)}</h2>${chips(c,'class')}<div class="card info rv"><span>العدد: <b>${o.n}</b></span><span>${o.boys} ولد · ${o.girls} بنت</span><span>متوسط الحضور: <b>${pc(o.avg)}</b></span><span>جمع مسجلة: <b>${o.meet}</b></span><span>محتاجين افتقاد: <b>${o.need}</b></span><span>أماكن فاضية: <b>${30-o.ks.length}</b></span></div><div class="grid">${o.ks.map(card).join('')||`<div class="empty">${ic('users')}مفيش أطفال في الفصل ده</div>`}</div>`},
att(a){const c=sw(a),o=D.cls[c],dt=D.dates;return`<h2>${ic('calendar')}قايمة حضور ملاك ${esc(c)}</h2>${chips(c,'att')}<div class="card info rv"><span>جمع مسجلة: <b>${o.meet}</b></span><span>المتوسط: <b>${pc(o.avg)}</b></span><span class="lgd"><i class="lg-ok">ح</i>حاضر</span><span class="lgd"><i class="lg-no">غ</i>غائب</span><span class="lgd"><i class="lg-ex">ع</i>معتذر</span></div>
 <div class="card tw rv"><table><thead><tr><th>الاسم</th><th>حضور</th><th>غياب</th><th>بعذر</th><th>النسبة</th><th>غياب متتالي</th><th>آخر حضور</th><th>الحالة</th>${dt.map((x,i)=>`<th class="${i==D.now?'now':''}" ${i==D.now?'id="nowc"':''}>${fd(x)}</th>`).join('')}</tr></thead><tbody>${o.ks.map(k=>`<tr><td><a href="javascript:void 0" data-open="${k.row}">${esc(k.name)}</a></td><td>${k.P}</td><td>${k.A}</td><td>${k.E}</td><td>${pc(k.pct)}</td><td>${k.streak}</td><td>${k.last?fd(k.last):'—'}</td><td style="color:${TIER[k.tier][2]}">${TIER[k.tier][1]}</td>${k.att.map((v,i)=>`<td class="${v==='ح'?'ok':v==='غ'?'no':v==='ع'?'ex':''} ${i==D.now?'now':''}">${v}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr><td>الحاضرين</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.any?x.p:''}</td>`).join('')}</tr><tr><td>النسبة</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.t?Math.round(x.p/x.t*100)+'%':''}</td>`).join('')}</tr></tfoot></table></div>`},
follow(){const l=D.kids.filter(k=>k.tier==3||k.tier==4).sort((a,b)=>b.streak-a.streak||(a.pct??0)-(b.pct??0));
 return`<h2>${ic('phone')}قايمة الافتقاد</h2><div class="card info rv"><span>الأطفال اللي نسبة حضورهم أقل من <b>${Math.round(D.th.ex*100)}%</b> أو غايبين <b>${D.th.streak}</b> اجتماعات ورا بعض أو أكتر — العدد: <b>${l.length}</b></span></div><div class="fl">${l.map((k,i)=>`<div class="card fi rv" style="--c:${TIER[k.tier][2]}"><div class="rank">${i+1}</div><div><h3>${esc(k.name)}</h3><p class="meta">${esc(k.cls)} · نسبة ${pc(k.pct)} · غياب متتالي ${k.streak} · آخر حضور ${k.last?fdl(k.last):'—'}<br>${ic('pin')} ${esc(k.addr||'—')}</p>${tag(TIER[k.tier])}</div><div class="acts">${btns(k)||'<span class="sub">— مفيش رقم —</span>'}<button class="b sm pri" data-open="${k.row}">التفاصيل</button></div></div>`).join('')||`<div class="card empty">${ic('checkc')}مفيش حد محتاج افتقاد دلوقتي</div>`}</div>`},
settings(){const k=(a,b)=>`<div><small>${a}</small><b>${b}</b></div>`;return`<h2>${ic('settings')}الإعدادات</h2><div class="card kv rv">${k('بداية العام الخدمي',fdl(D.start))+k('نهاية العام الخدمي',fdl(D.end))+k('يوم الاجتماع',esc(D.dayName))+k('«ممتاز» من',Math.round(D.th.ex*100)+'%')+k('«جيد» من',Math.round(D.th.good*100)+'%')+k('افتقاد ضروري بعد غياب',D.th.streak+' اجتماعات')+k('الفصول (الملائكة)',D.classes.map(esc).join(' · '))+k('السنين الدراسية',D.years.map(esc).join(' · '))}</div><div class="note">${ic('alert')}<span>الإعدادات بتتغيّر من شيت «الإعدادات» في الإكسيل، والموقع بيقراها أوتوماتيك. ماتغيّرش أسماء الفصول بعد تسجيل الأطفال.</span></div>`},
guide(){const s=[['plus','إضافة طفل جديد','افتح ملف الإكسيل وأضف الطفل في أول صف فاضي في شيت «قاعدة البيانات» (الاسم، النوع، السنة، الفصل، التليفونات…). بعدها ارفع الملف على GitHub بنفس الاسم، والطفل يظهر في الموقع لوحده.'],['camera','صورة الطفل','حطّ الصورة في عمود «الصورة» بـ Insert ← Pictures ← Place in Cell. بتظهر في كارت الطفل، ولو مفيش صورة بيظهر أول حرف من الاسم.'],['calendar','تسجيل الحضور','سجّل ح / غ / ع في شيت «حضور …» الخاص بالفصل، ثم ارفع الملف. النسب والألوان بتتحسب في الموقع أوتوماتيك.'],['phone','الافتقاد','القايمة بتتجمّع أوتوماتيك حسب الحدود في «الإعدادات». اضغط على زر الاتصال أو واتساب للتواصل مع الأم أو الأب.'],['ban','طفل غير نشط','بدل ما تمسح الصف، غيّر «الحالة» لـ «غير نشط». ماتعملش Sort ولا تمسح صفوف.'],['refresh','تحديث','بعد رفع الملف الجديد اضغط زر التحديث في أعلى الصفحة (أو افتح الموقع من جديد) — قد يستغرق GitHub دقيقة أو اثنين.']];return`<h2>${ic('book')}دليل الاستخدام</h2><div class="steps">${s.map(x=>`<div class="card rv"><div class="e">${ic(x[0])}</div><h3>${x[1]}</h3><p>${x[2]}</p></div>`).join('')}</div>`}};

function renderKids(){const q=$('#q').value.trim().toLowerCase(),f=id=>$(id).value;
 const l=D.kids.filter(k=>(!f('#fc')||k.cls===f('#fc'))&&(!f('#fg')||k.g===f('#fg'))&&(!f('#fy')||k.year===f('#fy'))&&(!f('#fs')||k.active==(f('#fs')==='1'))&&(!q||[k.name,k.mom,k.dad,k.addr,k.id,k.notes].join(' ').toLowerCase().includes(q)));
 $('#cn').textContent='النتائج: '+l.length+' من '+D.kids.length;$('#kg').innerHTML=l.map(card).join('')||`<div class="empty">${ic('list')}مفيش نتائج — جرّب تغيّر البحث أو الفلاتر</div>`;fx($('#kg'))}

function openKid(row){const k=D.kids.find(x=>x.row==row);if(!k)return;const t=TIER[k.tier],f=(a,b)=>b!==''&&b!=null?`<div><small>${a}</small><b>${b}</b></div>`:'';
 lastFocus=document.activeElement;
 $('#mc').innerHTML=`<div class="mh in"><div class="av ${k.g==='بنت'?'f':''}">${av(k)}</div><div><h2>${esc(k.name)}</h2><p class="meta">${k.id} · ${esc(k.cls)} · ${esc(k.year)}</p>${tag(t)}</div></div>
 <div class="kv">${f('حضور',k.P)+f('غياب',k.A)+f('بعذر',k.E)+f('النسبة',pc(k.pct))+f('غياب متتالي',k.streak)+f('آخر حضور',k.last?fdl(k.last):'—')+f('يروح لوحده؟',esc(k.alone))+f('النوع',k.g==='بنت'?'بنت':'ولد')}</div>
 ${k.addr?`<div class="mp">${ic('pin')}<span>${esc(k.addr)}</span><a class="b sm" target="_blank" rel="noopener" href="https://www.google.com/maps/search/${encodeURIComponent(k.addr)}">${ic('map')}الخريطة</a></div>`:''}
 ${k.notes?`<p class="mp">${ic('note')}<span>${esc(k.notes)}</span></p>`:''}<div style="display:flex;gap:6px;flex-wrap:wrap">${btns(k)}</div>
 <div class="tl">${k.att.map((v,i)=>v?`<span class="${v==='ح'?'ok':v==='غ'?'no':'ex'}">${fd(D.dates[i])} ${v}</span>`:'').join('')}</div>`;
 $('#modal').hidden=false;document.body.style.overflow='hidden';$('#mx').focus()}
const closeM=()=>{if($('#modal').hidden)return;$('#modal').hidden=true;document.body.style.overflow='';if(lastFocus&&lastFocus.focus)lastFocus.focus()};

const io=new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}}),{threshold:.08});
function fx(r=view){$$('.rv:not(.in)',r).forEach(e=>io.observe(e));
 $$('[data-count]',r).forEach(e=>{const to=+e.dataset.count,s=e.dataset.suf||'',t0=performance.now();(function f(t){const p=Math.min(1,(t-t0)/1100);e.textContent=Math.round(to*(1-Math.pow(1-p,4)))+s;p<1&&requestAnimationFrame(f)})(t0)})}
function route(){const [r,a]=decodeURIComponent(location.hash.slice(1)||'home').split('/'),n=V[r]?r:'home';
 view.innerHTML=`<div class="page">${V[n](a)}</div>`;scrollTo(0,0);fx();
 if(n==='kids'){['#q','#fc','#fg','#fs','#fy'].forEach(i=>$(i).addEventListener('input',renderKids));renderKids()}
 if(n==='att')setTimeout(()=>{const c=$('#nowc');c&&c.scrollIntoView({inline:'center',block:'nearest',behavior:'smooth'})},500);
 const href='#'+n;$$('#nav a,#tabs a').forEach(x=>{const on=x.getAttribute('href')===href;x.classList.toggle('on',on);on?x.setAttribute('aria-current','page'):x.removeAttribute('aria-current')})}
function init(){
 $('#nav').innerHTML=NAV.map(n=>`<a href="#${n[0]}">${ic(n[1])}${n[2]}</a>`).join('');
 $('#tabs').innerHTML=NAV.filter(n=>!n[3]).map(n=>`<a href="#${n[0]}">${ic(n[1])}<span>${n[2]}</span></a>`).join('');
 $('#dl').href=FILE;addEventListener('hashchange',route)}
document.addEventListener('click',e=>{const o=e.target.closest('[data-open]');if(o)openKid(o.dataset.open);if(e.target.id==='modal'||e.target.closest('#mx'))closeM()});
addEventListener('keydown',e=>e.key==='Escape'&&closeM());
async function boot(first){try{await load();if(first)init();route()}catch(e){view.innerHTML=`<div class="card empty" style="padding:34px">${ic('alert')}<h2 style="justify-content:center;margin:0">حصلت مشكلة</h2><p class="sub" style="line-height:2">${esc(e.message)}<br>لو بتفتح الملف من الكمبيوتر مباشرة (file://) لازم يتفتح من GitHub Pages أو سيرفر محلي.</p></div>`}$('#splash').classList.add('off')}
$('#rf').onclick=()=>{$('#splash').classList.remove('off');boot(false)};boot(true);
