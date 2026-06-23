// 薪資月對月差異分析 — 純前端。資料不離開瀏覽器。
// 核心方法：組成拆解法（component bridge），跨月無員工編號也能精準對帳。
//   實發PHP = 實發RMB×匯率 + 小計PESO − 扣款小計PESO
//   差異 = 匯率效應 + RMB實發差(@上月匯率) + PESO補貼差 − PESO扣款差

const COL = {
  ym:'年月', group:'事业群', div:'事业部', dept:'部门名称',
  hire:'入职日期', leave:'离职日期', has:'是否享有当月薪资',
  rmbNet:'实发工资RMB', pesoAdd:'小计PESO', pesoDed:'扣款小计PESO',
  rate:'汇率PESO', php:'实发工资PHP比索',
};
// RMB 應發組成（用來把 RMB 實發差再拆細）
const RMB_PARTS = [
  ['底薪',   ['本月薪资']],
  ['獎金分紅',['提成/奖金RMB','分红奖金RMB','年终奖RMB（13/14薪）','劝退补偿金']],
  ['加班值班',['假日加班费','1.25倍加班费','1.5倍加班费','值班小时金额','存工折现金额','1倍加班费']],
  ['補發補助',['补发上月出勤工资','调薪差额补发RMB','其他补发RMB','年度机票补助金额','宿舍维修补贴RMB','年度年假补贴金额','特殊津贴_RMB']],
  ['扣款',   ['扣款小计RMB'], -1],   // 負向：增加扣款會降低實發
];

let WB = null;          // SheetJS workbook
let SHEETS = {};        // name -> rows[]

const $ = s => document.querySelector(s);
const num = v => (typeof v === 'number' && isFinite(v)) ? v : 0;
const fmt = n => Math.round(n).toLocaleString('en-US');
const signed = n => (n >= 0 ? '+' : '−') + fmt(Math.abs(n));
const cls = n => n >= 0 ? 'pos' : 'neg';

// ---- 載檔 ----
const drop = $('#drop'), fileInput = $('#file');
$('#pick').onclick = () => fileInput.click();
fileInput.onchange = e => e.target.files[0] && loadFile(e.target.files[0]);
['dragover','dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => e.dataTransfer.files[0] && loadFile(e.dataTransfer.files[0]));

function loadFile(file){
  const r = new FileReader();
  r.onload = e => {
    WB = XLSX.read(new Uint8Array(e.target.result), { type:'array', cellDates:true });
    SHEETS = {};
    WB.SheetNames.forEach(n => {
      const rows = XLSX.utils.sheet_to_json(WB.Sheets[n], { defval:null });
      SHEETS[n] = rows.filter(r => r[COL.ym] != null && r[COL.ym] !== '');
    });
    const names = WB.SheetNames.filter(n => (SHEETS[n]||[]).length).sort().reverse();
    fill('#curSheet', names, names[0]);
    fill('#prevSheet', names, names[1] || names[0]);
    syncRates();
    $('#controls').classList.remove('hidden');
  };
  r.readAsArrayBuffer(file);
}
function fill(sel, opts, val){
  const el = $(sel); el.innerHTML = '';
  opts.forEach(o => { const op = document.createElement('option'); op.value = op.textContent = o; el.appendChild(op); });
  el.value = val;
}
// 月匯率 = 該月最常見的 汇率PESO
function monthRate(rows){
  const c = {};
  rows.forEach(r => { const v = num(r[COL.rate]); if (v) c[v] = (c[v]||0)+1; });
  let best = 0, n = -1;
  for (const k in c) if (c[k] > n) { n = c[k]; best = +k; }
  return best;
}
function syncRates(){
  $('#curRate').value  = monthRate(SHEETS[$('#curSheet').value]  || []);
  $('#prevRate').value = monthRate(SHEETS[$('#prevSheet').value] || []);
}
$('#curSheet').onchange = syncRates;
$('#prevSheet').onchange = syncRates;
$('#run').onclick = run;

// ---- 計算 ----
function sum(rows, key, mult=1){ return rows.reduce((a,r)=>a + num(r[key])*mult, 0); }
function sumParts(rows, keys, mult=1){ return keys.reduce((a,k)=>a + sum(rows,k,mult), 0); }

function ymOf(v){ // 入/離職日期 -> 'YYYYMM'
  if (v == null) return '';
  if (v instanceof Date) return '' + v.getFullYear() + String(v.getMonth()+1).padStart(2,'0');
  const s = String(v).replace(/[^0-9]/g,'');
  return s.length >= 6 ? s.slice(0,6) : '';
}

function run(){
  const cur = SHEETS[$('#curSheet').value], prev = SHEETS[$('#prevSheet').value];
  const curYM = String(cur[0][COL.ym]);
  const rC = +$('#curRate').value, rP = +$('#prevRate').value;

  // 各月實發PHP（真實匯率）
  const phpPrev = sum(prev, COL.php);
  const phpCurReal = sum(cur, COL.php);

  // 瀑布橋（一律從上月真實 → 本月真實，匯率效應獨立一段）
  const rateEff = sum(cur, COL.rmbNet) * (rC - rP);
  const dRmbNet = (sum(cur, COL.rmbNet) - sum(prev, COL.rmbNet)) * rP;
  const dPesoAdd = sum(cur, COL.pesoAdd) - sum(prev, COL.pesoAdd);
  const dPesoDed = -(sum(cur, COL.pesoDed) - sum(prev, COL.pesoDed));

  const bridge = [
    { label:'上月總額', amount:phpPrev, type:'start' },
    { label:'匯率效應', amount:rateEff, type:'rate' },
    { label:'實發RMB差', amount:dRmbNet, type:'flow' },
    { label:'PESO補貼差', amount:dPesoAdd, type:'flow' },
    { label:'PESO扣款差', amount:dPesoDed, type:'flow' },
    { label:'本月總額', amount:phpCurReal, type:'end' },
  ];

  // RMB 實發差拆細（@上月匯率）。補一條殘差「其他」保證加總 = 實發RMB差。
  const rmbBreak = RMB_PARTS.map(([name, keys, mult=1]) => ({
    name,
    delta: (sumParts(cur, keys, mult) - sumParts(prev, keys, mult)) * rP,
  }));
  const breakSum = rmbBreak.reduce((a,x)=>a+x.delta, 0);
  const residual = dRmbNet - breakSum;
  if (Math.abs(residual) > 1) rmbBreak.push({ name:'其他', delta:residual });

  // 人數 / 名單
  const joins = cur.filter(r => ymOf(r[COL.hire]) === curYM);
  const leaves = cur.filter(r => ymOf(r[COL.leave]) === curYM || String(r[COL.has]||'') === '否');

  // 事業群 / 事業部
  const groups = buildGroups(prev, cur, rP, rC);

  render({
    headPrev: prev.length, headCur: cur.length,
    phpPrev, phpCurReal,
    rateEff, dPesoAdd, dPesoDed, bridge, rmbBreak, joins, leaves, groups, curYM,
  });
}

function buildGroups(prev, cur, rP, rC){
  const key = r => [r[COL.group]||'(未填)', r[COL.div]||'(未填)'];
  const acc = {};
  const add = (rows, slot) => rows.forEach(r => {
    const [g,d] = key(r);
    acc[g] = acc[g] || { name:g, head:{prev:0,cur:0}, php:{prev:0,cur:0}, rmb:{prev:0,cur:0}, subs:{} };
    acc[g].subs[d] = acc[g].subs[d] || { name:d, head:{prev:0,cur:0}, php:{prev:0,cur:0}, rmb:{prev:0,cur:0} };
    acc[g].head[slot]++; acc[g].subs[d].head[slot]++;
    const php = num(r[COL.php]), rmb = num(r[COL.rmbNet]);
    acc[g].php[slot] += php; acc[g].subs[d].php[slot] += php;
    acc[g].rmb[slot] += rmb; acc[g].subs[d].rmb[slot] += rmb;
  });
  add(prev,'prev'); add(cur,'cur');
  // 匯率效應(本月RMB×Δrate) vs 真實差(總差−匯率)
  const finish = o => {
    o.rateEff = o.rmb.cur * (rC - rP);
    o.totalDiff = o.php.cur - o.php.prev;
    o.realDiff = o.totalDiff - o.rateEff;
    o.headDiff = o.head.cur - o.head.prev;
  };
  return Object.values(acc).map(g => {
    finish(g);
    g.subs = Object.values(g.subs).map(s => (finish(s), s)).sort((a,b)=>Math.abs(b.totalDiff)-Math.abs(a.totalDiff));
    return g;
  }).sort((a,b)=>Math.abs(b.totalDiff)-Math.abs(a.totalDiff));
}

// 驅動因素白話說明
const WHY = {
  '底薪':'底薪調整（加薪・到職滿月）',
  '獎金分紅':'提成・分紅・年終獎（13/14薪）',
  '加班值班':'加班費・值班・假日加班',
  '補發補助':'補發・機票／宿舍補助・特殊津貼',
  '扣款':'出勤扣款（遲到・事假・曠工）',
  '其他':'未分類差額',
};

// ---- 畫面 ----
function render(d){
  $('#result').classList.remove('hidden');
  const headDiff = d.headCur - d.headPrev;
  const totalRealDiff = d.phpCurReal - d.phpPrev;          // 真實總差（用各月真匯率）
  const rP = $('#prevRate').value, rC = $('#curRate').value;

  // 大數字
  $('#cHead').textContent = `${d.headPrev} → ${d.headCur}`;
  $('#cHeadD').innerHTML = `<span class="${cls(headDiff)}">${headDiff>=0?'▲':'▼'} ${signed(headDiff)} 人</span>`;
  $('#cTotal').textContent = `${fmt(d.phpPrev)} → ${fmt(d.phpCurReal)}`;
  $('#cTotalD').innerHTML = `<span class="${cls(totalRealDiff)}">${totalRealDiff>=0?'▲':'▼'} ${signed(totalRealDiff)} PHP</span>`;

  // 統一的驅動因素清單（全部加起來 = 真實總差）
  const drivers = [
    { name:'匯率變動', why:`人民幣換披索匯率 ${rP} → ${rC}`, amount:d.rateEff },
    ...d.rmbBreak.map(x => ({ name:x.name, why:WHY[x.name]||'', amount:x.delta })),
    { name:'披索補貼', why:'房補・餐補・駐菲補貼等', amount:d.dPesoAdd },
    { name:'披索扣款', why:'水電・房補超額等扣款', amount:d.dPesoDed },
  ].filter(x => Math.abs(x.amount) >= 1)
   .sort((a,b) => Math.abs(b.amount) - Math.abs(a.amount));

  // 結論導言
  const upTxt = totalRealDiff>=0 ? '增加' : '減少';
  const headTxt = headDiff>0 ? `多了 ${headDiff} 人` : headDiff<0 ? `少了 ${-headDiff} 人` : '人數沒變';
  let lead = `本月跟上月比：<b>${headTxt}</b>，但總額<b class="${cls(totalRealDiff)}">${upTxt}了 ${fmt(Math.abs(totalRealDiff))} 披索</b>。`;
  if (headDiff < 0 && totalRealDiff > 0)
    lead += `<br>「人少了、錢卻變多」的主因是下面排第一的項目 —— <b>${drivers[0].name}</b>，跟人頭數無關。`;
  else if (drivers.length)
    lead += `<br>最大的原因是 <b>${drivers[0].name}</b>（${signed(drivers[0].amount)} 披索）。`;
  $('#lead').innerHTML = lead;

  renderDrivers(drivers);
  $('#driverFoot').textContent = '＊以上每一項相加，剛好等於總額差。金額已換算成披索（PHP），方便直接比大小。';

  // 部門長條 + 完整表
  renderDeptBars(d.groups);
  renderGroups(d.groups);

  // 瀑布（折疊）
  renderWaterfall(d.bridge);
  $('#bridgeNote').textContent =
    `匯率：上月 ${rP} → 本月 ${rC}。「匯率變動」= 本月實發RMB × 匯率差，單獨抽出；其餘各段一律以上月匯率計，反映真實人/薪/獎金變動。`;

  // 名單
  $('#joinTitle').textContent = `本月新進（${d.joins.length} 人，合計 ${fmt(sum(d.joins,COL.php))} 披索）`;
  $('#leaveTitle').textContent = `本月離職／不享當月薪資（${d.leaves.length} 人）`;
  renderRoster('#joinTbl', d.joins);
  renderRoster('#leaveTbl', d.leaves);
}

function renderDrivers(drivers){
  const max = Math.max(1, ...drivers.map(x => Math.abs(x.amount)));
  $('#drivers').innerHTML = drivers.map((x, i) => {
    const up = x.amount >= 0;
    const w = Math.abs(x.amount) / max * 100;
    const tag = i === 0 ? '<span class="tag">最大主因</span>' : '';
    return `<li class="${up?'up':'down'}">
      <span class="ico">${up?'🔺':'🔻'}</span>
      <span class="dmain"><div class="dname">${x.name}${tag}</div><div class="dwhy">${x.why}</div></span>
      <span class="dbar"><i style="width:${w}%;background:${up?'var(--up)':'var(--down)'}"></i></span>
      <span class="damt">${signed(x.amount)}</span>
    </li>`;
  }).join('');
}

function renderDeptBars(groups){
  const top = groups.slice(0, 8);
  const max = Math.max(1, ...top.map(g => Math.abs(g.totalDiff)));
  $('#deptBars').innerHTML = top.map(g => {
    const up = g.totalDiff >= 0;
    const w = Math.abs(g.totalDiff) / max * 48;
    const left = up ? 50 : 50 - w;
    const hd = g.headDiff === 0 ? '人數持平' : `人數 ${signed(g.headDiff)}`;
    return `<div class="db-row">
      <div class="db-name">${g.name} <small>${hd}</small></div>
      <div class="db-track"><div class="db-fill" style="left:${left}%;width:${w}%;background:${up?'var(--up)':'var(--down)'}"></div></div>
      <div class="db-amt ${cls(g.totalDiff)}">${signed(g.totalDiff)}</div>
    </div>`;
  }).join('');
}

function renderWaterfall(bridge){
  const max = Math.max(...bridge.map(b=>Math.abs(b.amount)));
  const flowMax = Math.max(1, ...bridge.filter(b=>b.type!=='start'&&b.type!=='end').map(b=>Math.abs(b.amount)));
  const html = bridge.map(b => {
    let bar;
    if (b.type==='start' || b.type==='end'){
      bar = `<div class="wf-fill" style="left:0;width:${Math.abs(b.amount)/max*100}%;background:#3b466b"></div>`;
    } else {
      const w = Math.abs(b.amount)/flowMax*48;
      const left = b.amount>=0 ? 50 : 50-w;
      const c = b.type==='rate' ? 'var(--rate)' : (b.amount>=0?'var(--up)':'var(--down)');
      bar = `<div class="wf-fill" style="left:${left}%;width:${w}%;background:${c}"></div>`;
    }
    const amt = (b.type==='start'||b.type==='end') ? fmt(b.amount) : signed(b.amount);
    const ac = (b.type==='start'||b.type==='end') ? '' : cls(b.amount);
    return `<div class="wf-row"><div class="wf-label">${b.label}</div>
      <div class="wf-bar">${bar}</div><div class="wf-amt ${ac}">${amt}</div></div>`;
  }).join('');
  $('#waterfall').innerHTML = html;
}

function renderGroups(groups){
  const head = `<tr><th>事業群／事業部</th><th>人數(上→本)</th><th>人數Δ</th>
    <th>金額Δ(PHP)</th><th>其中匯率</th><th>其中真實</th></tr>`;
  const rowCells = (o, isSub) =>
    `<td>${o.name}</td>
     <td>${o.head.prev}→${o.head.cur}</td>
     <td class="${cls(o.headDiff)}">${o.headDiff===0?'0':signed(o.headDiff)}</td>
     <td class="${cls(o.totalDiff)}">${signed(o.totalDiff)}</td>
     <td class="${cls(o.rateEff)}">${signed(o.rateEff)}</td>
     <td class="${cls(o.realDiff)}">${signed(o.realDiff)}</td>`;
  let html = head;
  groups.forEach((g, gi) => {
    html += `<tr class="grp-row" data-g="${gi}">${rowCells(g)}</tr>`;
    g.subs.forEach(s => { html += `<tr class="sub-row g${gi}">${rowCells(s, true)}</tr>`; });
  });
  const tbl = $('#groupTbl');
  tbl.innerHTML = html;
  tbl.querySelectorAll('.grp-row').forEach(r => r.onclick = () => {
    r.classList.toggle('open');
    tbl.querySelectorAll('.g'+r.dataset.g).forEach(s => s.classList.toggle('show'));
  });
}

function renderRoster(sel, rows){
  if (!rows.length){ $(sel).innerHTML = '<tr><td class="note">無</td></tr>'; return; }
  const head = '<tr><th>事業群</th><th>部門</th><th>級別</th><th>實發PHP</th></tr>';
  $(sel).innerHTML = head + rows
    .sort((a,b)=>num(b[COL.php])-num(a[COL.php]))
    .map(r => `<tr><td>${r[COL.group]||''}</td><td>${r[COL.dept]||''}</td>
      <td>${r['福利级别']||''}</td><td>${fmt(num(r[COL.php]))}</td></tr>`).join('');
}
