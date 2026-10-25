const FILE='Bara3em_System_Final.xlsx';
const $=(s,e=document)=>e.querySelector(s),$$=(s,e=document)=>[...e.querySelectorAll(s)];
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const TIER=[['⏳','لسه مفيش حضور','#8fa3c8'],['⭐','ممتاز','#3ddc97'],['👍','جيد','#5cc8ff'],['⚠️','محتاج متابعة','#ffb454'],['📞','افتقاده ضروري','#ff6b7a'],['⛔','غير نشط','#6b7894']];
const CC=['#5cc8ff','#e8c26a','#b79cff','#3ddc97','#ff8fb1','#ffb454'];
const NAV=[['home','🏠','الرئيسية'],['kids','📋','الأطفال'],['class','🪪','الفصول'],['att','📅','الحضور'],['follow','📞','الافتقاد'],['settings','⚙️','الإعدادات',1],['guide','📖','الدليل',1]];
const cell=(ws,r,c)=>{const x=ws&&ws[XLSX.utils.encode_cell({r:r-1,c:c-1})];return x?x.v:undefined};
const ds=n=>new Date(Math.round((n-25569)*864e5)),fd=d=>d.getUTCDate()+'/'+(d.getUTCMonth()+1),fdl=d=>fd(d)+'/'+d.getUTCFullYear();
const ph=v=>{if(v==null||v==='')return'';let s=String(v).replace(/\D/g,'');if(s.length===10&&s[0]!=='0')s='0'+s;return s};
const wa=p=>'https://wa.me/20'+p.slice(-10),pc=p=>p==null?'—':Math.round(p*100)+'%';
let D,view=$('#view'),LM='';

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

const av=k=>k.photo?`<img src="${k.photo}" alt="" loading="lazy">`:`<span>${esc([...k.name][0])}</span>`;
const ring=(p,c,sz=46)=>{const L=2*Math.PI*18;return`<svg class="ring" width="${sz}" height="${sz}" viewBox="0 0 44 44"><circle cx="22" cy="22" r="18" class="bg"/><circle cx="22" cy="22" r="18" class="fg" style="stroke:${c};--L:${L};--o:${L*(1-(p||0))}"/><text x="22" y="26" text-anchor="middle">${p==null?'—':Math.round(p*100)}</text></svg>`};
const btns=k=>[['أم',k.mom],['أب',k.dad]].filter(x=>x[1]).map(([l,p])=>`<a class="b" href="tel:${p}">📞 ${l}</a><a class="b wa" target="_blank" rel="noopener" href="${wa(p)}">💬 ${l}</a>`).join('');
const card=k=>{const t=TIER[k.tier];return`<article class="card kid rv tilt ${k.g==='بنت'?'f':''} ${k.active?'':'off'}"><div class="av">${av(k)}</div><div class="inf"><h3>${esc(k.name)}</h3><p class="meta">${esc(k.cls)} · ${esc(k.year)} · ${k.g==='بنت'?'👧':'👦'}</p><span class="tag" style="--c:${t[2]}">${t[0]} ${t[1]}</span></div>${ring(k.pct,t[2])}<div class="acts">${btns(k)}<button class="b gold" data-open="${k.row}">التفاصيل</button></div></article>`};
const chips=(sel,pre)=>`<div class="chips">${D.classes.map((c,i)=>`<a href="#${pre}/${encodeURIComponent(c)}" class="${c===sel?'on':''}" style="--c:${CC[i]}">${esc(c)}</a>`).join('')}</div>`;
const sw=(a,c)=>D.classes.includes(a)?a:D.classes[0];

const V={
home(){const a=D.kids.filter(k=>k.active),P=D.kids.reduce((s,k)=>s+k.P,0),Ab=D.kids.reduce((s,k)=>s+k.A,0);
 const st=[['👶','المسجلين',D.kids.length],['✅','النشطين',a.length],['👦','أولاد',a.filter(k=>k.g==='ولد').length],['👧','بنات',a.filter(k=>k.g==='بنت').length],['📈','متوسط الحضور',P+Ab?Math.round(P/(P+Ab)*100):0,'%'],['📅','جمع مسجلة',D.meet],['⚠️','محتاجين افتقاد',D.kids.filter(k=>k.tier==3||k.tier==4).length]];
 let last=D.tot.map(x=>x.any).lastIndexOf(true);const mx=Math.max(1,...D.tot.map(x=>x.p));
 return`<section class="hero"><h1>🌱 خدمة براعم</h1><p>نظام الافتقاد والحضور · العام الخدمي ${fdl(D.start)} ← ${fdl(D.end)} · اجتماع كل ${esc(D.dayName)}</p></section>
 <div class="stats">${st.map((s,i)=>`<div class="card st rv tilt"><div class="e">${s[0]}</div><b data-count="${s[2]}" data-suf="${s[3]||''}">0</b><small>${s[1]}</small></div>`).join('')}</div>
 <h2>🪪 الفصول</h2><div class="grid">${D.classes.map((c,i)=>{const o=D.cls[c];return`<div class="card ct rv tilt" style="--c:${CC[i]}"><h3>👼 ملاك ${esc(c)}${ring(o.avg,CC[i],54)}</h3><p style="color:var(--mu);font-size:14px">👶 ${o.n} طفل · 👦 ${o.boys} · 👧 ${o.girls}<br>⚠️ محتاجين افتقاد: ${o.need} · 🆓 أماكن فاضية: ${30-o.ks.length}</p><div class="row"><a class="b gold" href="#class/${encodeURIComponent(c)}">🪪 الكروت</a><a class="b" href="#att/${encodeURIComponent(c)}">📅 الحضور</a></div></div>`}).join('')}</div>
 <h2>📊 الحضور في كل اجتماع</h2><div class="card rv">${last<0?'<div class="bar2">لسه مفيش حضور مسجّل</div>':`<div class="bars">${D.tot.slice(0,last+1).map((x,i)=>`<div class="bar" title="${fdl(D.dates[i])}: ${x.p}" style="--h:${x.p/mx*78}%;--i:${i}"><b>${x.p}</b><i></i><span>${fd(D.dates[i])}</span></div>`).join('')}</div>`}</div>
 <h2>🧭 تنقّل سريع</h2><div class="row" style="display:flex;gap:10px;flex-wrap:wrap"><a class="b gold" href="#kids">📋 قاعدة البيانات</a><a class="b" href="#follow">📞 الافتقاد</a><a class="b" href="#settings">⚙️ الإعدادات</a><a class="b" href="#guide">📖 الدليل</a><a class="b" href="${FILE}" download>➕ إضافة طفل (حمّل الإكسيل)</a></div>
 ${LM?`<p style="color:var(--mu);font-size:12px;margin-top:22px;text-align:center">آخر تحديث للبيانات: ${LM}</p>`:''}`},
kids(){return`<h2>📋 قاعدة بيانات الأطفال</h2><div class="tb"><input id="q" type="search" placeholder="🔍 ابحث بالاسم أو الرقم أو العنوان…"><select id="fc"><option value="">كل الفصول</option>${D.classes.map(c=>`<option>${esc(c)}</option>`).join('')}</select><select id="fg"><option value="">أولاد وبنات</option><option>ولد</option><option>بنت</option></select><select id="fs"><option value="">كل الحالات</option><option value="1">نشط</option><option value="0">غير نشط</option></select><select id="fy"><option value="">كل السنين</option>${D.years.map(y=>`<option>${esc(y)}</option>`).join('')}</select></div><p id="cn" style="color:var(--mu);margin-bottom:12px"></p><div class="grid" id="kg"></div>`},
class(a){const c=sw(a),o=D.cls[c];return`<h2>🪪 كروت أطفال ملاك ${esc(c)}</h2>${chips(c,'class')}<div class="card info rv"><span>👶 العدد: <b>${o.n}</b></span><span>👦 ${o.boys} | 👧 ${o.girls}</span><span>📈 متوسط الحضور: <b>${pc(o.avg)}</b></span><span>📅 جمع مسجلة: <b>${o.meet}</b></span><span>⚠️ محتاجين افتقاد: <b>${o.need}</b></span><span>🆓 فاضي: <b>${30-o.ks.length}</b></span></div><div class="grid">${o.ks.map(card).join('')||'<p class="bar2">مفيش أطفال في الفصل ده</p>'}</div>`},
att(a){const c=sw(a),o=D.cls[c],dt=D.dates;return`<h2>📅 قايمة حضور ملاك ${esc(c)}</h2>${chips(c,'att')}<div class="card info rv"><span>📅 جمع مسجلة: <b>${o.meet}</b></span><span>📊 المتوسط: <b>${pc(o.avg)}</b></span><span><b style="color:#3ddc97">ح</b> حاضر · <b style="color:#ff6b7a">غ</b> غائب · <b style="color:#ffb454">ع</b> معتذر</span></div>
 <div class="card tw rv"><table><thead><tr><th>الاسم</th><th>✅</th><th>❌</th><th>🟠</th><th>النسبة</th><th>غياب متتالي</th><th>آخر حضور</th><th>الحالة</th>${dt.map((x,i)=>`<th class="${i==D.now?'now':''}" ${i==D.now?'id="nowc"':''}>${fd(x)}</th>`).join('')}</tr></thead><tbody>${o.ks.map(k=>`<tr><td><a href="javascript:void 0" data-open="${k.row}">${esc(k.name)}</a></td><td>${k.P}</td><td>${k.A}</td><td>${k.E}</td><td>${pc(k.pct)}</td><td>${k.streak}</td><td>${k.last?fd(k.last):'—'}</td><td style="color:${TIER[k.tier][2]}">${TIER[k.tier][0]} ${TIER[k.tier][1]}</td>${k.att.map((v,i)=>`<td class="${v==='ح'?'ok':v==='غ'?'no':v==='ع'?'ex':''} ${i==D.now?'now':''}">${v}</td>`).join('')}</tr>`).join('')}</tbody><tfoot><tr><td>👥 الحاضرين</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.any?x.p:''}</td>`).join('')}</tr><tr><td>📈 النسبة</td><td colspan="7"></td>${o.per.map(x=>`<td>${x.t?Math.round(x.p/x.t*100)+'%':''}</td>`).join('')}</tr></tfoot></table></div>`},
follow(){const l=D.kids.filter(k=>k.tier==3||k.tier==4).sort((a,b)=>b.streak-a.streak||(a.pct??0)-(b.pct??0));
 return`<h2>📞 قايمة الافتقاد</h2><div class="card info rv"><span>⚠️ الأطفال اللي نسبة حضورهم أقل من <b>${Math.round(D.th.ex*100)}%</b> أو غايبين <b>${D.th.streak}</b> اجتماعات ورا بعض أو أكتر — العدد: <b>${l.length}</b></span></div><div class="fl">${l.map((k,i)=>`<div class="card fi rv" style="--c:${TIER[k.tier][2]}"><div class="rank">${i+1}</div><div><h3 style="font-size:16px">${esc(k.name)}</h3><p class="meta" style="color:var(--mu);font-size:12.5px">${esc(k.cls)} · نسبة ${pc(k.pct)} · غياب متتالي ${k.streak} · آخر حضور ${k.last?fdl(k.last):'—'}<br>📍 ${esc(k.addr||'—')}</p><span class="tag" style="--c:${TIER[k.tier][2]}">${TIER[k.tier][0]} ${TIER[k.tier][1]}</span></div><div class="acts">${btns(k)||'<span style="color:var(--mu)">— مفيش رقم —</span>'}<button class="b gold" data-open="${k.row}">التفاصيل</button></div></div>`).join('')||'<p class="bar2">🎉 مفيش حد محتاج افتقاد دلوقتي</p>'}</div>`},
settings(){const k=(a,b)=>`<div><small>${a}</small><b>${b}</b></div>`;return`<h2>⚙️ الإعدادات</h2><div class="card kv rv">${k('بداية العام الخدمي',fdl(D.start))+k('نهاية العام الخدمي',fdl(D.end))+k('يوم الاجتماع',esc(D.dayName))+k('«ممتاز» من',Math.round(D.th.ex*100)+'%')+k('«جيد» من',Math.round(D.th.good*100)+'%')+k('افتقاد ضروري بعد غياب',D.th.streak+' اجتماعات')+k('الفصول (الملائكة)',D.classes.map(esc).join(' · '))+k('السنين الدراسية',D.years.map(esc).join(' · '))}</div><p style="color:var(--mu);margin-top:14px;font-size:13px">الإعدادات بتتغيّر من شيت «الإعدادات» في الإكسيل، والموقع بيقراها أوتوماتيك. ⚠️ ماتغيّرش أسماء الفصول بعد تسجيل الأطفال.</p>`},
guide(){const s=[['➕ إضافة طفل جديد','افتح ملف الإكسيل وأضف الطفل في أول صف فاضي في شيت «قاعدة البيانات» (الاسم، النوع، السنة، الفصل، التليفونات…). بعدها ارفع الملف على GitHub بنفس الاسم، والطفل يظهر في الموقع لوحده.'],['📷 صورة الطفل','حطّ الصورة في عمود «الصورة» بـ Insert ← Pictures ← Place in Cell. بتظهر في كارت الطفل، ولو مفيش صورة بيظهر أول حرف من الاسم.'],['📅 تسجيل الحضور','سجّل ح / غ / ع في شيت «حضور …» الخاص بالفصل، ثم ارفع الملف. النسب والألوان بتتحسب في الموقع أوتوماتيك.'],['📞 الافتقاد','القايمة بتتجمّع أوتوماتيك حسب الحدود في «الإعدادات». اضغط 📞 للاتصال أو 💬 لفتح واتساب.'],['⛔ طفل غير نشط','بدل ما تمسح الصف، غيّر «الحالة» لـ «غير نشط». ماتعملش Sort ولا تمسح صفوف.'],['🔄 تحديث','بعد رفع الملف الجديد اضغط 🔄 في أعلى الصفحة (أو افتح الموقع من جديد) — قد يستغرق GitHub دقيقة أو اثنين.']];return`<h2>📖 دليل الاستخدام</h2><div class="steps">${s.map(x=>`<div class="card rv"><h3>${x[0]}</h3><p>${x[1]}</p></div>`).join('')}</div>`}};

function renderKids(){const q=$('#q').value.trim().toLowerCase(),f=id=>$(id).value;
 const l=D.kids.filter(k=>(!f('#fc')||k.cls===f('#fc'))&&(!f('#fg')||k.g===f('#fg'))&&(!f('#fy')||k.year===f('#fy'))&&(!f('#fs')||k.active==(f('#fs')==='1'))&&(!q||[k.name,k.mom,k.dad,k.addr,k.id,k.notes].join(' ').toLowerCase().includes(q)));
 $('#cn').textContent='النتائج: '+l.length+' من '+D.kids.length;$('#kg').innerHTML=l.map(card).join('')||'<p class="bar2">مفيش نتائج</p>';fx($('#kg'))}

function openKid(row){const k=D.kids.find(x=>x.row==row);if(!k)return;const t=TIER[k.tier],f=(a,b)=>b?`<div><small>${a}</small><b>${b}</b></div>`:'';
 $('#mc').innerHTML=`<div class="mh in"><div class="av ${k.g==='بنت'?'f':''}" style="${k.g==='بنت'?'background:linear-gradient(135deg,#ff6fa5,#ffb3cf)':''}">${av(k)}</div><div><h2>${esc(k.name)}</h2><p class="meta" style="color:var(--mu)">${k.id} · ${esc(k.cls)} · ${esc(k.year)}</p><span class="tag" style="--c:${t[2]}">${t[0]} ${t[1]}</span></div></div>
 <div class="kv" style="padding:0 0 14px">${f('✅ حضور',k.P)+f('❌ غياب',k.A)+f('🟠 بعذر',k.E)+f('📈 النسبة',pc(k.pct))+f('غياب متتالي',k.streak)+f('آخر حضور',k.last?fdl(k.last):'—')+f('يروح لوحده؟',esc(k.alone))+f('النوع',k.g==='بنت'?'👧 بنت':'👦 ولد')}</div>
 ${k.addr?`<p style="margin:6px 0 12px">📍 ${esc(k.addr)} <a class="b" target="_blank" rel="noopener" href="https://www.google.com/maps/search/${encodeURIComponent(k.addr)}">🗺️ الخريطة</a></p>`:''}
 ${k.notes?`<p style="margin-bottom:12px">📝 ${esc(k.notes)}</p>`:''}<div class="acts" style="display:flex;gap:6px;flex-wrap:wrap">${btns(k)}</div>
 <div class="tl">${k.att.map((v,i)=>v?`<span class="${v==='ح'?'ok':v==='غ'?'no':'ex'}">${fd(D.dates[i])} ${v}</span>`:'').join('')}</div>`;$('#modal').hidden=false;document.body.style.overflow='hidden'}
const closeM=()=>{$('#modal').hidden=true;document.body.style.overflow=''};

const io=new IntersectionObserver(es=>es.forEach(e=>{if(e.isIntersecting){e.target.classList.add('in');io.unobserve(e.target)}}),{threshold:.08});
function fx(r=view){$$('.rv:not(.in)',r).forEach((e,i)=>{e.style.setProperty('--d',Math.min(i,10)*55+'ms');io.observe(e)});
 $$('[data-count]',r).forEach(e=>{const to=+e.dataset.count,s=e.dataset.suf||'',t0=performance.now();(function f(t){const p=Math.min(1,(t-t0)/1300);e.textContent=Math.round(to*(1-Math.pow(1-p,4)))+s;p<1&&requestAnimationFrame(f)})(t0)})}
function route(){const [r,a]=decodeURIComponent(location.hash.slice(1)||'home').split('/'),n=V[r]?r:'home';
 view.innerHTML=`<div class="page">${V[n](a)}</div>`;scrollTo(0,0);fx();
 if(n==='kids'){['#q','#fc','#fg','#fs','#fy'].forEach(i=>$(i).addEventListener('input',renderKids));renderKids()}
 if(n==='att')setTimeout(()=>{const c=$('#nowc');c&&c.scrollIntoView({inline:'center',block:'nearest',behavior:'smooth'})},500);
 const href=n==='class'||n==='att'?'#'+n:'#'+n;$$('#nav a,#tabs a').forEach(x=>x.classList.toggle('on',x.getAttribute('href')===href))}
function init(){const first=encodeURIComponent(D.classes[0]);
 $('#nav').innerHTML=NAV.map(n=>`<a href="#${n[0]}">${n[1]} ${n[2]}</a>`).join('');
 $('#tabs').innerHTML=NAV.filter(n=>!n[3]).map(n=>`<a href="#${n[0]}"><b>${n[1]}</b>${n[2]}</a>`).join('');
 $('#dl').href=FILE;addEventListener('hashchange',route)}
document.addEventListener('click',e=>{const o=e.target.closest('[data-open]');if(o)openKid(o.dataset.open);if(e.target.id==='modal'||e.target.id==='mx')closeM()});
addEventListener('keydown',e=>e.key==='Escape'&&closeM());
if(matchMedia('(hover:hover)').matches){addEventListener('pointermove',e=>{document.body.style.setProperty('--mx',e.clientX+'px');$('#glow').style.setProperty('--mx',e.clientX+'px');$('#glow').style.setProperty('--my',e.clientY+'px');
 const t=e.target.closest&&e.target.closest('.tilt');$$('.tilt').forEach(x=>{if(x!==t&&x.style.transform)x.style.transform=''});if(t){const r=t.getBoundingClientRect(),px=(e.clientX-r.left)/r.width-.5,py=(e.clientY-r.top)/r.height-.5;t.style.transform=`perspective(800px) rotateY(${-px*7}deg) rotateX(${py*7}deg) translateY(-3px)`}})}
async function boot(first){try{await load();if(first)init();route()}catch(e){view.innerHTML=`<div class="card" style="padding:30px;text-align:center"><h2 style="justify-content:center">⚠️ حصلت مشكلة</h2><p style="color:var(--mu);line-height:2">${esc(e.message)}<br>لو بتفتح الملف من الكمبيوتر مباشرة (file://) لازم يتفتح من GitHub Pages أو سيرفر محلي.</p></div>`}$('#splash').classList.add('off')}
$('#rf').onclick=()=>{$('#splash').classList.remove('off');boot(false)};boot(true);
