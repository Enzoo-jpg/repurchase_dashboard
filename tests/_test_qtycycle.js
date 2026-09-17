// 「购药数量 × 用药周期 → 下次应购日」验证
// 覆盖：数量解析、品种开关、三种周期算法、自动重置、同日合并、预判应购日、
//       主状态机与整体小结口径一致、行内购药数量列
const fs = require('fs');
const path = require('path');

function makeEl(id) {
  return { id, innerHTML: '', textContent: '', value: '', disabled: false, style: {}, dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; }, appendChild(){}, insertBefore(){}, removeChild(){},
    querySelectorAll(){ return []; }, click(){}, remove(){},
    querySelector(){ return makeEl(id + '>child'); }, parentNode: { insertBefore(){} },
    nextElementSibling: null, onclick: null, onchange: null, oninput: null };
}
global.window = global;
global.document = {
  querySelector(sel){ return makeEl(String(sel)); }, querySelectorAll(){ return []; },
  addEventListener(){}, createElement(){ return makeEl('created'); },
  documentElement: { outerHTML: '<html></html>' }, body: { appendChild(){}, remove(){}, removeChild(){} },
};
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.ExcelJS = require(path.join(__dirname, '..', 'vendor', 'exceljs.min.js'));
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const App = global.AppCore;
const S = App.state, ST = App.STORE;

let pass = 0, fail = 0;
const eq = (name, got, want) => {
  if (JSON.stringify(got) === JSON.stringify(want)) { pass++; console.log('  ✅ ' + name + ' → ' + JSON.stringify(got)); }
  else { fail++; console.log('  ❌ ' + name + ' → 实际 ' + JSON.stringify(got) + '，期望 ' + JSON.stringify(want)); }
};
const ok = (name, cond) => { if (cond) { pass++; console.log('  ✅ ' + name); } else { fail++; console.log('  ❌ ' + name); } };
const P = (d, q) => ({ date: d, qty: q });

console.log('\n[1] 购药数量解析（缺失/异常按 1）');
eq('null', App.qtyNum(null), 1);
eq('空串', App.qtyNum(''), 1);
eq('纯数字', App.qtyNum('3'), 3);
eq('带单位「2盒」', App.qtyNum('2盒'), 2);
eq('小数「1.5」', App.qtyNum('1.5'), 1.5);
eq('千分位「1,200」', App.qtyNum('1,200'), 1200);
eq('零 -> 1', App.qtyNum('0'), 1);
eq('非数字「未知」-> 1', App.qtyNum('未知'), 1);
eq('负数「-2」-> 1', App.qtyNum('-2'), 1);

console.log('\n[2] 品种开关：默认关闭时数量不放大（= 历史行为）');
S.qtyScale = {}; S.cycleAlgo = 'reset'; S.cycleTol = 7;
eq('未开启时 3 个周期只算 1 个周期', App.stockEndAfter([P('2026-08-01', 3)], 21, '百泽安'), '2026-08-22');
S.qtyScale = { '百泽安': true };
eq('开启后 3 个周期算 3 个周期', App.stockEndAfter([P('2026-08-01', 3)], 21, '百泽安'), '2026-10-03');
eq('另一品种不受影响', App.stockEndAfter([P('2026-08-01', 3)], 28, '百悦泽'), '2026-08-29');

console.log('\n[3] 三种算法（周期 21：8/1 买 3 个周期、8/20 买 1 个周期）');
S.qtyScale = { '百泽安': true };
const seq = [P('2026-08-01', 3), P('2026-08-20', 1)];
S.cycleAlgo = 'reset';
eq('重置式 → 8/20 + 21', App.stockEndAfter(seq, 21, '百泽安'), '2026-09-10');
S.cycleAlgo = 'accum';
eq('累计式 → 10/3 + 21（囤药顺延）', App.stockEndAfter(seq, 21, '百泽安'), '2026-10-24');
S.cycleAlgo = 'accum_tol'; S.cycleTol = 7;
eq('累计+容差(7)：剩余 44 天 > 7 → 顺延', App.stockEndAfter(seq, 21, '百泽安'), '2026-10-24');
S.cycleTol = 60;
eq('累计+容差(60)：剩余 44 天 ≤ 60 → 重置', App.stockEndAfter(seq, 21, '百泽安'), '2026-09-10');

console.log('\n[4] 药吃完了才买 → 三种算法都自动重置');
const gap = [P('2026-08-01', 1), P('2026-09-30', 1)]; // 8/22 就吃完了，9/30 才买
for (const algo of ['reset', 'accum', 'accum_tol']) {
  S.cycleAlgo = algo;
  eq(`算法 ${algo}：自 9/30 重新起算`, App.stockEndAfter(gap, 21, '百泽安'), '2026-10-21');
}
S.cycleAlgo = 'accum';

console.log('\n[5] 提前几天购药（剩余库存小于容差）');
const slight = [P('2026-08-01', 1), P('2026-08-20', 1)]; // 8/22 才该吃？不，8/1+21=8/22，8/20 提前 2 天
S.cycleAlgo = 'accum_tol'; S.cycleTol = 7;
eq('剩余 2 天 ≤ 容差 7 → 按 8/20 重新起算', App.stockEndAfter(slight, 21, '百泽安'), '2026-09-10');
S.cycleAlgo = 'accum';
eq('纯累计 → 接在 8/22 之后', App.stockEndAfter(slight, 21, '百泽安'), '2026-09-12');
S.cycleAlgo = 'reset';

console.log('\n[6] 同一天多行数量相加');
eq('3 行各买 1 支 = 一次买 3 支', App.normalizePurchases([P('2026-08-01', 1), P('2026-08-01', 1), P('2026-08-01', 1)]), [{ date: '2026-08-01', qty: 3 }]);
eq('乱序输入也能按日期归并', App.normalizePurchases([P('2026-08-20', 2), P('2026-08-01', 1), P('2026-08-20', 3)]),
  [{ date: '2026-08-01', qty: 1 }, { date: '2026-08-20', qty: 5 }]);

console.log('\n[7] 预判应购日 = 按算法推进到「除最后一次以外」');
eq('重置式：8/1 的 3 个周期覆盖到 10/3', App.stockEndAfter(seq.slice(0, -1), 21, '百泽安'), '2026-10-03');

console.log('\n[8] 主状态机端到端（固定参考日 2026-09-10，自然周）');
S.refDate = '2026-09-10'; S.weekSel = 'this'; S.weekEnd = 0;
S.stdCycle = { '百泽安': 21 }; ST.cycles = {}; ST.followups = [];
const sales = (q3, q1) => ([
  { source: 'sales', patient_name: '甲', phone: '13900000001', sales_time: '2026-08-01', product: '百泽安', qty: q3, hospital: 'H', pharmacy: 'P', physician: '医' },
  { source: 'sales', patient_name: '甲', phone: '13900000001', sales_time: '2026-08-20', product: '百泽安', qty: q1, hospital: 'H', pharmacy: 'P', physician: '医' },
]);
S.qtyScale = { '百泽安': true };
S.cycleAlgo = 'accum';
ST.sales = sales('3', '1');
let row = App.buildRows().find(r => r.product === '百泽安');
eq('累计式 下次应购日', row.due_date, '2026-10-24');
eq('累计式 预判应购日', row.expected_due, '2026-10-03');
eq('非「应回已回」行不呈现提前/延后', row.due_offset, null);
eq('最近一次购药数量', row.qty, 1);
// 把参考日放到 8/20 所在周：该患者本周有购药 → 应回已回，此时呈现提前天数
S.refDate = '2026-08-20';
row = App.buildRows().find(r => r.product === '百泽安');
eq('应回已回行 应购药日期=预判应购日', row.due_date, '2026-10-03');
eq('提前天数（负=提前 44 天）', row.due_offset, -44);
S.refDate = '2026-09-10';

S.cycleAlgo = 'reset';
row = App.buildRows().find(r => r.product === '百泽安');
eq('重置式 下次应购日', row.due_date, '2026-09-10');

S.qtyScale = {};
row = App.buildRows().find(r => r.product === '百泽安');
eq('关闭品种开关 → 回到「最近购药 + 周期」', row.due_date, '2026-09-10');

console.log('\n[9] 数量缺失时按 1（与旧行为一致）');
S.qtyScale = { '百泽安': true }; S.cycleAlgo = 'reset';
ST.sales = [
  { source: 'sales', patient_name: '乙', phone: '13900000002', sales_time: '2026-08-01', product: '百泽安', qty: null, hospital: 'H', pharmacy: 'P', physician: '医' },
  { source: 'sales', patient_name: '乙', phone: '13900000002', sales_time: '2026-08-20', product: '百泽安', qty: '未知', hospital: 'H', pharmacy: 'P', physician: '医' },
];
row = App.buildRows().find(r => r.product === '百泽安');
eq('两次购药均无数量 → 各按 1 个周期', row.due_date, '2026-09-10');
eq('数量列回退为 1', row.qty, 1);
// 与「关闭品种开关」必须完全一致
S.qtyScale = {};
eq('与关闭开关的结果一致', App.buildRows().find(r => r.product === '百泽安').due_date, '2026-09-10');

console.log('\n[10] 整体小结与主状态机口径一致');
S.qtyScale = { '百泽安': true }; S.cycleAlgo = 'accum';
ST.sales = sales('3', '1');
// 小结按「所选周」判定：把参考日放在库存耗尽周内，使该患者进入「应回未回」
S.refDate = '2026-10-24';
row = App.buildRows().find(r => r.product === '百泽安');
eq('主状态机：应购日落在所选周 → 应回购', row.status, '应回购');
const st = App.buildSummaryStats('百泽安');
eq('小结：应回未回人数', st.notBought, 1);
eq('小结：下周窗口随之平移', st.nStart + '~' + st.nEnd, '2026-10-26~2026-11-01');

S.qtyScale = {}; S.cycleAlgo = 'reset'; S.cycleTol = 7; S.refDate = ''; ST.sales = [];
console.log('\n结果: ' + pass + ' 通过 / ' + fail + ' 失败');
if (fail) process.exit(1);
