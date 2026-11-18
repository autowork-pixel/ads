const $ = id => document.getElementById(id);
const fields = ['spend','impressions','clicks','orders','revenue'];
const labels = {meta:'메타 DA',naver:'네이버 SA',coupang:'쿠팡 SA'};
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const n = v => {const x=Number(String(v ?? '').replace(/[₩,원\s]/g,''));return Number.isFinite(x)?x:0};
const norm = v => String(v ?? '').toLowerCase().replace(/\([^)]*\)/g,'').replace(/\s/g,'');
const find = (h,names,needed=true) => {for(const name of names){const i=h.findIndex(v=>norm(v)===norm(name));if(i>=0)return i}if(needed)throw Error(`필수 컬럼을 찾지 못했습니다: ${names.join(' / ')}`);return -1};
const val = (row,i) => i>=0 ? row[i] : null;
const dateValue = v => {if(v instanceof Date)return v.toISOString().slice(0,10);const s=String(v??'');const m=s.match(/(20\d\d)[./-](\d{1,2})[./-](\d{1,2})/);return m?`${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`:''};
const empty = () => Object.fromEntries(fields.map(k=>[k,0]));
const sum = list => {const out=empty();for(const r of list)for(const k of fields)out[k]+=r[k];return out};
const div = (a,b) => b ? a/b : null;
const derived = r => ({roas:div(r.revenue,r.spend),ctr:div(r.clicks,r.impressions),cpc:div(r.spend,r.clicks),cpa:div(r.spend,r.orders)});
const fmt = (v,type='int') => v==null?'—':type==='pct'?`${(v*100).toLocaleString('ko-KR',{maximumFractionDigits:1})}%`:type==='roas'?`${Math.round(v*100).toLocaleString('ko-KR')}%`:Math.round(v).toLocaleString('ko-KR');
const delta = (a,b) => b==null?'비교 없음':b===0?(a?'신규':'—'):`${a/b-1>=0?'+':''}${((a/b-1)*100).toFixed(1)}%`;
const warn = (message,error=false) => {const el=$('status');el.textContent=message;el.className=error?'error':''};

async function workbook(file){
  if(!file)throw Error('금주 원본을 선택하세요.');
  if(!window.XLSX)throw Error('Excel 읽기 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인하고 새로고침하세요.');
  const bytes=await file.arrayBuffer();
  const wb=XLSX.read(bytes,{type:'array',cellDates:true});
  let best=null;
  for(const name of wb.SheetNames){
    const grid=XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true});
    for(let i=0;i<Math.min(grid.length,15);i++){
      const h=grid[i].map(String);
      if(h.some(v=>['캠페인 이름','캠페인명','캠페인'].includes(v)) && h.some(v=>['광고비','지출 금액 (KRW)','총비용','비용'].includes(v))){best={header:h,body:grid.slice(i+1),file:file.name};break}
    }
    if(best)break;
  }
  if(!best)throw Error(`${file.name}: 지원하는 광고 원본 헤더를 찾지 못했습니다.`);
  return best;
}

function ingest(book,platform,start,end){
  const h=book.header;
  let c;
  if(platform==='meta')c={name:find(h,['캠페인 이름']),spend:find(h,['지출 금액 (KRW)','지출 금액']),impressions:find(h,['노출']),clicks:find(h,['클릭(전체)','클릭']),orders:find(h,['공유 항목이 포함된 구매','구매','결과']),revenue:find(h,['공유 항목의 구매 전환값','구매 전환값']),start:find(h,['보고 시작']),end:find(h,['보고 종료'])};
  else if(platform==='naver')c={name:find(h,['캠페인','캠페인명']),group:find(h,['캠페인유형','캠페인 유형'],false),spend:find(h,['총비용','비용','광고비']),impressions:find(h,['노출수','노출']),clicks:find(h,['클릭수','클릭']),orders:find(h,['총 전환수','전환수','전환']),revenue:find(h,['총 전환매출액(원)','전환매출액','매출액'])};
  else c={name:find(h,['캠페인명']),group:find(h,['광고그룹'],false),spend:find(h,['광고비']),impressions:find(h,['노출수']),clicks:find(h,['클릭수']),orders:find(h,['총 주문수(14일)']),revenue:find(h,['총 전환매출액(14일)'])};
  const map=new Map(), periods=new Set();let count=0;
  for(const row of book.body){
    if(platform==='meta'){
      const a=dateValue(val(row,c.start)),b=dateValue(val(row,c.end));
      if(!a||!b)continue;
      periods.add(`${a}~${b}`);
      if(start && (a<start||b>end))continue;
    }
    const name=String(val(row,c.name)??'').trim();
    if(!name||['합계','총합계','total'].includes(name.toLowerCase()))continue;
    const group=String(val(row,c.group)??'').trim();
    const id=group?`${name} · ${group}`:name;
    if(!map.has(id))map.set(id,empty());
    const target=map.get(id);
    for(const k of fields)target[k]+=n(val(row,c[k]));
    count++;
  }
  if(!count)throw Error(`${book.file}: 선택 기간에 집계할 행이 없습니다.`);
  if(platform==='meta'&&!start&&periods.size>1)throw Error(`${book.file}: 보고 기간이 ${periods.size}개입니다. 메타 기간을 지정하세요.`);
  return {map,total:sum(map.values()),count,periods,file:book.file};
}

function findings(cur,prev){
  if(!prev)return ['전주 원본이 없어 전주 대비 변동을 계산하지 않았습니다.'];
  const a=cur.total,b=prev.total,am=derived(a),bm=derived(b);
  const lines=[['spend','광고비'],['revenue','전환매출'],['orders','전환수']].map(([k,l])=>`${l} ${fmt(b[k])} → ${fmt(a[k])} (${delta(a[k],b[k])})`);
  for(const [k,l,type] of [['roas','ROAS','roas'],['ctr','CTR','pct'],['cpa','CPA','int']])if(am[k]!=null&&bm[k]!=null)lines.push(`${l} ${fmt(bm[k],type)} → ${fmt(am[k],type)}`);
  return lines;
}

function actions(cur,prev){
  const out=[];
  for(const [name,row] of cur.map){
    if(!row.spend)continue;
    const old=prev?.map.get(name),m=derived(row),om=old&&derived(old);
    if(!row.orders)out.push({weight:row.spend,name,text:'지출 대비 전환 0건. 전환 추적과 랜딩 페이지를 확인하고, 성과가 계속 없으면 입찰·예산 축소를 검토하세요.'});
    else if(old?.spend){
      if(m.roas!=null&&om.roas!=null&&m.roas<om.roas*.8&&row.spend>=old.spend*.8)out.push({weight:row.spend,name,text:`ROAS ${fmt(om.roas,'roas')} → ${fmt(m.roas,'roas')}. 소재·검색어·상품별 효율을 확인하고 저효율 구간의 예산을 조정하세요.`});
      if(m.cpc!=null&&om.cpc!=null&&m.cpc>om.cpc*1.25)out.push({weight:row.spend,name,text:`CPC ${fmt(om.cpc)} → ${fmt(m.cpc)}원. 입찰가와 타깃·키워드 경쟁도를 확인하세요.`});
    }
  }
  return out.sort((a,b)=>b.weight-a.weight).slice(0,8).length?out.sort((a,b)=>b.weight-a.weight).slice(0,8):[{name:'전체',text:prev?'급격한 악화 신호는 확인되지 않았습니다. 기존 예산을 유지하며 세부 성과를 점검하세요.':'전주 원본을 추가하면 변화에 따른 개선사항을 제안할 수 있습니다.'}];
}

function report(client,platform,cur,prev,currentLabel,previousLabel){
  const a=cur.total,b=prev?.total,am=derived(a),bm=b&&derived(b);
  const kpis=[['spend','광고비','int','원'],['revenue','전환매출','int','원'],['orders','전환수','int','건'],['roas','ROAS','roas',''],['ctr','CTR','pct',''],['cpc','CPC','int','원'],['cpa','CPA','int','원']];
  const cards=kpis.map(([k,l,t,u])=>{const now=k in a?a[k]:am[k],old=b?(k in b?b[k]:bm[k]):null;return `<div class="kpi"><span>${l}</span><strong>${fmt(now,t)}${now==null?'':u}</strong><small>전주 ${fmt(old,t)}${old==null?'':u} · ${delta(now??0,old)}</small></div>`}).join('');
  const summary=findings(cur,prev).map(s=>`<li>${esc(s)}</li>`).join('');
  const actionRows=actions(cur,prev).map(x=>`<tr><td>${esc(x.name)}</td><td class="action">${esc(x.text)}</td></tr>`).join('');
  const names=[...new Set([...cur.map.keys(),...(prev?.map.keys()??[])])].sort((x,y)=>(cur.map.get(y)?.spend??0)-(cur.map.get(x)?.spend??0));
  const detail=names.map(name=>{const r=cur.map.get(name)??empty(),old=prev?.map.get(name),m=derived(r),om=old&&derived(old),gap=m.roas!=null&&om?.roas!=null?`${(m.roas-om.roas)*100>=0?'+':''}${Math.round((m.roas-om.roas)*100)}%p`:'—';return `<tr><td>${esc(name)}</td><td>${fmt(r.spend)}</td><td>${delta(r.spend,old?.spend)}</td><td>${fmt(r.revenue)}</td><td>${fmt(m.roas,'roas')}</td><td>${gap}</td><td>${fmt(r.orders)}</td><td>${fmt(m.ctr,'pct')}</td><td>${fmt(m.cpc)}</td><td>${fmt(m.cpa)}</td></tr>`}).join('');
  const caveats=['전환매출은 매체가 보고한 귀속 매출이며 실제 정산 매출과 다를 수 있습니다.','전환 기여기간과 VAT 기준이 매체마다 다릅니다. 매체 간 ROAS를 단순 비교하지 마세요.'];
  if(platform==='meta')caveats.push('메타는 공유 항목 구매·구매 전환값 기준입니다. 계정의 전환 설정과 내보낸 컬럼을 확인하세요.');
  if(platform==='coupang')caveats.push('쿠팡은 14일 전환 기준입니다. 상품별 전환 매출 행의 중복 가능성을 원본 합계와 대조하세요.');
  if(!prev)caveats.push('전주 원본이 없어 변화 기반 조치안은 제한됩니다.');
  return `<div class="report-head"><div><div class="eyebrow">WEEKLY PERFORMANCE</div><h2>${esc(client)} · ${labels[platform]} 주간 광고 성과</h2><p>금주 ${esc(currentLabel)} · 전주 ${esc(previousLabel||'원본 없음')}<br>금주 ${esc(cur.file)} (${cur.count.toLocaleString()}행)${prev?` · 전주 ${esc(prev.file)} (${prev.count.toLocaleString()}행)`:''}</p></div><button id="download" type="button">HTML 내려받기 ↓</button></div><div class="kpis">${cards}</div><section class="panel"><h2>전주 현황</h2><ul class="summary-list">${summary}</ul></section><section class="panel"><h2>금주 개선사항</h2><div class="table-wrap"><table><thead><tr><th>대상</th><th style="text-align:left">권장 조치 · 확인할 사항</th></tr></thead><tbody>${actionRows}</tbody></table></div></section><section class="panel"><h2>캠페인별 성과</h2><div class="table-wrap"><table><thead><tr><th>캠페인 / 그룹</th><th>광고비</th><th>전주 대비</th><th>전환매출</th><th>ROAS</th><th>ROAS 차이</th><th>전환수</th><th>CTR</th><th>CPC</th><th>CPA</th></tr></thead><tbody>${detail}</tbody></table></div></section><section class="panel"><h2>집계 기준과 확인 사항</h2><ul class="summary-list">${caveats.map(s=>`<li>${esc(s)}</li>`).join('')}</ul></section>`;
}

async function generate(){
  const button=$('generate');button.disabled=true;warn('원본을 읽고 있습니다…');
  try{
    const client=$('client').value.trim(),platform=$('platform').value;
    if(!client)throw Error('고객명을 입력하세요.');
    const currentFile=$('current').files[0],previousFile=$('previous').files[0];
    if(!currentFile)throw Error('금주 원본을 선택하세요.');
    const cs=$('currentStart').value,ce=$('currentEnd').value,ps=$('previousStart').value,pe=$('previousEnd').value;
    if(Boolean(cs)!==Boolean(ce)||Boolean(ps)!==Boolean(pe))throw Error('기간의 시작일과 종료일을 함께 입력하세요.');
    if((cs&&cs>ce)||(ps&&ps>pe))throw Error('기간의 시작일이 종료일보다 늦습니다.');
    if(platform!=='meta'&&(cs||ps))throw Error('날짜 필터는 메타 원본에만 사용합니다.');
    if(previousFile && previousFile.name===currentFile.name && previousFile.size===currentFile.size && previousFile.lastModified===currentFile.lastModified && !ps)throw Error('같은 파일을 금주·전주에 넣었습니다. 메타 원본이면 두 기간을 지정하세요.');
    const currentBook=await workbook(currentFile),previousBook=previousFile?(previousFile===currentFile?currentBook:await workbook(previousFile)):null;
    const cur=ingest(currentBook,platform,cs,ce);
    const prev=previousBook?ingest(previousBook,platform,ps,pe):(ps?ingest(currentBook,platform,ps,pe):null);
    if(prev&&platform==='meta'&&[...cur.periods].join() === [...prev.periods].join() && cs===ps)throw Error('금주와 전주가 같은 보고 기간입니다.');
    const label=(d,s,e)=>s?`${s} ~ ${e}`:platform==='meta'&&d.periods.size===1?[...d.periods][0].replace('~',' ~ '):d.file;
    const content=report(client,platform,cur,prev,label(cur,cs,ce),prev?label(prev,ps,pe):'');
    const output=$('result');output.innerHTML=content;output.hidden=false;
    $('download').onclick=async()=>{
      const title=`${client}_${platform}_주간광고보고서`;
      let css='';
      try{css=await (await fetch('styles.css')).text()}catch{css='body{font:14px sans-serif;margin:24px}table{border-collapse:collapse}td,th{border:1px solid #ddd;padding:8px}'}
      const payload=`<!doctype html><html lang="ko"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>${esc(title)}</title><style>${css}</style></head><body><main class="shell">${content.replace(/<button id="download"[^>]*>.*?<\/button>/,'')}</main></body></html>`;
      const url=URL.createObjectURL(new Blob([payload],{type:'text/html;charset=utf-8'}));
      const a=document.createElement('a');a.href=url;a.download=`${title.replace(/[\\/:*?"<>|]/g,'_')}.html`;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);
    };
    output.scrollIntoView({behavior:'smooth',block:'start'});
    warn('보고서를 생성했습니다.');
  }catch(e){warn(e.message||String(e),true)}finally{button.disabled=false}
}

$('generate').addEventListener('click',generate);
$('platform').addEventListener('change',()=>{$('dates').hidden=$('platform').value!=='meta'});

