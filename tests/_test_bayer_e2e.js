// 拜耳真实底表端到端验证：诺倍戈/博瑞纳品种 + 销售明细 + 随访历史任务
// 用法: node tests/_test_bayer_e2e.js
const fs = require('fs');
const path = require('path');

// ---------- 最小 DOM stub ----------
function makeEl(id) {
  return {
    id, innerHTML: '', textContent: '', value: '', disabled: false,
    style: {}, dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    appendChild(){}, insertBefore(){}, removeChild(){},
    querySelectorAll(){ return []; },
    querySelector(){ return makeEl(id + '>child'); },
    parentNode: { insertBefore(){} },
    nextElementSibling: null,
    onclick: null, onchange: null, oninput: null,
  };
}
global.window = global;
global.document = {
  querySelector(sel){ return makeEl(String(sel)); },
  querySelectorAll(){ return []; },
  addEventListener(){},
  createElement(){ return makeEl('created'); },
  documentElement: { outerHTML: '<html></html>' },
  body: { appendChild(){}, remove(){}, removeChild(){} },
};
global.URL = { createObjectURL(){ return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.alert = () => {};
global.confirm = () => true;

const load = f => { const code = fs.readFileSync(path.join(__dirname, '..', f), 'utf8'); (new Function(code))(); };
load('mapping.js');
load('pipeline.js');
load('app.js');

const M = global.Mapping, P = global.Pipeline, App = global.AppCore;

function fileObj(p, name) {
  const buf = fs.readFileSync(p);
  return {
    name: name || path.basename(p),
    size: buf.length,
    async arrayBuffer(){ return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); },
  };
}

// 真实底表（拜耳诺倍戈）
const SALES = 'D:/1111卢馨兰/厂牌/拜耳/数据底表/9月数据/销售明细查询报表_oneid (1).xlsx';
const FOLLOWUP_CANDIDATES = [
  'D:/1111卢馨兰/厂牌/拜耳/DTP随访跟进/历史任务 - 2026-09-15T084659.001.xls',
  'D:/1111卢馨兰/厂牌/拜耳/DTP随访跟进/历史任务 (5)(1).xls',
  'D:/1111卢馨兰/厂牌/拜耳/历史任务.xls',
];

async function main() {
  const files = [];
  if (!fs.existsSync(SALES)) { console.log('⏭️ 跳过：销售明细不存在', SALES); process.exit(0); }
  files.push(fileObj(SALES));
  const FU = FOLLOWUP_CANDIDATES.find(p => fs.existsSync(p));
  if (FU) files.push(fileObj(FU)); else console.log('⚠ 未找到随访任务表，仅验证销售表');

  // 解析前先看随访表表头识别情况
  if (FU) {
    const buf = fs.readFileSync(FU);
    const wb = XLSX.read(new Uint8Array(buf), { type: 'array', cellDates: true, sheetRows: 20 });
    for (const name of wb.SheetNames) {
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true, defval: null, cellDates: true });
      const hr = P.detectHeaderRow(aoa.map(r => r.map(c => String(c == null ? '' : c))));
      const cols = (aoa[hr] || []).map(c => String(c == null ? '' : c).trim()).filter(Boolean);
      console.log(`[随访表 sheet=${name}] 表头行=${hr} 列数=${cols.length} 识别=${M.detectTableType(cols)}`);
      if (name === wb.SheetNames[0]) console.log('  表头前25列:', cols.slice(0, 25).join(' | '));
    }
  }

  const t0 = Date.now();
  const res = await P.processFiles(files);
  console.log('[解析] 耗时', ((Date.now() - t0) / 1000).toFixed(1) + 's');
  console.log('[解析] 销售明细:', res.sales.length, '条');
  console.log('[解析] 随访任务:', res.followups.length, '条');
  console.log('[解析] 用药周期表:', Object.keys(res.cycles).length, '人');

  const prodCount = res.sales.reduce((a, r) => { a[r.product] = (a[r.product] || 0) + 1; return a; }, {});
  console.log('[销售品种归一]', JSON.stringify(prodCount));
  const badProd = res.sales.filter(r => r.product !== '诺倍戈' && r.product !== '博瑞纳').length;
  if (badProd) console.log('⚠ 未归一到两个品种的行数:', badProd);
  const fuProd = res.followups.reduce((a, r) => { a[r.product] = (a[r.product] || 0) + 1; return a; }, {});
  console.log('[随访品种归一]', JSON.stringify(fuProd));

  console.log('[销售样例]', JSON.stringify(res.sales[0]));
  const fu0 = res.followups.find(f => f.exec_time || f.plan_time);
  console.log('[随访样例]', JSON.stringify(fu0));

  // ---------- 计算 ----------
  App.STORE.sales = res.sales;
  App.STORE.followups = res.followups;
  App.STORE.cycles = res.cycles;
  const t1 = Date.now();
  const rows = App.buildRows();
  console.log('\n[计算] 耗时', ((Date.now() - t1) / 1000).toFixed(2) + 's');
  console.log('[计算] 名单总数:', rows.length);

  const sum = App.buildSummary(rows);
  console.log('[状态分布]', JSON.stringify(sum.by_status));
  console.log('[品种分布]', JSON.stringify(sum.by_product));
  console.log('[医院分布 TOP5]', JSON.stringify(Object.entries(sum.by_hospital || {}).sort((a,b)=>b[1]-a[1]).slice(0,5)));

  const byStatus = {};
  for (const r of rows) (byStatus[r.status] = byStatus[r.status] || []).push(r);
  for (const st of Object.keys(byStatus)) {
    const s = byStatus[st][0];
    console.log(`[${st} 示例] 患者=${s.patient_name} 品种=${s.product} 最近购药=${s.last_purchase} 周期=${s.cycle_days} 应购=${s.due_date} 距今=${s.days_to_due} 信号=${s.fu_signal}`);
  }

  // 周期核验：应全部为 30（诺倍戈/博瑞纳标准周期）
  const cycleSet = [...new Set(rows.map(r => r.cycle_days))];
  console.log('[周期取值]', JSON.stringify(cycleSet), '→ 期望全部 30');

  // 脱敏验证
  const desen = P.desensitize({ patient_name: '杨绍玉', phone: '18981606138', physician: '杜勇何正宇' }, false, false, false);
  console.log('[脱敏]', JSON.stringify(desen), '→ 期望 杨*玉 / 189****6138 / 杜***宇');
  if (desen.patient_name !== '杨*玉' || desen.phone !== '189****6138' || desen.physician !== '杜***宇') {
    throw new Error('脱敏规则不符合预期');
  }

  // 筛选验证：诺倍戈 + 应回购
  const f1 = App.filterRows(rows.filter(r => r.product === '诺倍戈' && r.status === '应回购'));
  console.log('\n[筛选] 诺倍戈+应回购:', f1.length, '条');

  console.log('\n✅ 拜耳真实底表端到端验证通过');
}

main().catch(e => { console.error('FAIL', e); process.exit(1); });
