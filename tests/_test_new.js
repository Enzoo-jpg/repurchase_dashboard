// 新需求验证：① 姓名脱敏保留首尾字 ② 标准周期可配置 ③ 时间窗筛选豁免已逾期
const fs = require('fs');
const path = require('path');
function makeEl(id) {
  return { id, innerHTML: '', textContent: '', value: '', disabled: false, style: {}, dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    appendChild(){}, insertBefore(){}, removeChild(){}, querySelectorAll(){ return []; },
    querySelector(){ return makeEl(id + '>child'); }, parentNode: { insertBefore(){} },
    nextElementSibling: null, onclick: null, onchange: null, oninput: null };
}
global.window = global;
global.document = {
  querySelector(sel){ return makeEl(String(sel)); }, querySelectorAll(){ return []; },
  addEventListener(){}, createElement(){ return makeEl('created'); },
  documentElement: { outerHTML: '<html></html>' }, body: { appendChild(){}, remove(){} },
};
global.URL = { createObjectURL(){ return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const P = global.Pipeline, App = global.AppCore;

// ---------- ① 姓名脱敏 ----------
console.log('===== ① 姓名/医生脱敏（保留首尾字，中间按位数打 *） =====');
const maskCases = [
  ['郑辉', '郑*'],       // 2字
  ['王建国', '王*国'],   // 3字
  ['欧阳明德', '欧**德'], // 4字
  ['爱新觉罗启明', '爱****明'], // 6字
  ['张', '*'],           // 单字
];
let okAll = true;
for (const [src, want] of maskCases) {
  const got = P.desensitize({ patient_name: src, phone: '13800138000', physician: src }).patient_name;
  const pass = got === want;
  okAll = okAll && pass;
  console.log(`  ${src} → ${got} ${pass ? '✅' : '❌ 期望 ' + want}`);
}
// 医生同样规则
const doc = P.desensitize({ patient_name: 'x', phone: '1', physician: '张建国' }).physician;
console.log('  医生张建国 →', doc, doc === '张*国' ? '✅' : '❌');

// 三种脱敏方式（默认 edge；first=首字保留；all=全部隐藏）
console.log('\n  [脱敏方式选择]');
const m = n => P.desensitize({ patient_name: n, phone: '1', physician: '王建国' }, false, false, false, 'edge').patient_name;
const mf = n => P.desensitize({ patient_name: n, phone: '1', physician: '王建国' }, false, false, false, 'first').patient_name;
const ma = n => P.desensitize({ patient_name: n, phone: '1', physician: '王建国' }, false, false, false, 'all').patient_name;
console.log('  edge 欧阳明德 →', m('欧阳明德'), m('欧阳明德') === '欧**德' ? '✅' : '❌');
console.log('  first 欧阳明德 →', mf('欧阳明德'), mf('欧阳明德') === '欧***' ? '✅' : '❌');
console.log('  all 欧阳明德 →', ma('欧阳明德'), ma('欧阳明德') === '****' ? '✅' : '❌');
console.log('  first 王建国 →', mf('王建国'), mf('王建国') === '王**' ? '✅' : '❌');
console.log('  first 医生王建国 →', P.desensitize({ patient_name: 'x', phone: '1', physician: '王建国' }, false, false, false, 'first').physician === '王**' ? '✅' : '❌');
console.log('  all 医生王建国 →', P.desensitize({ patient_name: 'x', phone: '1', physician: '王建国' }, false, false, false, 'all').physician === '***' ? '✅' : '❌');
okAll = okAll && m('欧阳明德') === '欧**德' && mf('欧阳明德') === '欧***' && ma('欧阳明德') === '****';

// ---------- ② 标准周期可配置 ----------
console.log('\n===== ② 标准周期维护 =====');
App.state.stdCycle = { "百泽安": 21, "百悦泽": 28 };
App.state.advance = 7;
App.state.start = null; App.state.end = null;
const sales = [{
  source: 'sales', patient_name: '测试', phone: '10000000000', sales_time: (() => {
    const d = new Date(); d.setDate(d.getDate() - 10); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  })(), product: '百泽安', hospital: 'H', pharmacy: 'P', physician: '医'
}];
App.STORE.sales = sales; App.STORE.followups = []; App.STORE.cycles = {};
const dueDefault = App.buildRows()[0].due_date;
App.state.stdCycle['百泽安'] = 14; // 改短周期
const due14 = App.buildRows()[0].due_date;
console.log(`  周期21天应购=${dueDefault} | 改14天后应购=${due14} ${due14 < dueDefault ? '✅ 周期配置生效' : '❌'}`);

// ---------- ③ 状态边界：已逾期 / 应回购（按「所选周」判定；原「时间窗筛选」已被周视图取代） ----------
console.log('\n===== ③ 状态边界（周维度） =====');
App.state.stdCycle = { "百泽安": 21, "百悦泽": 28 }; // 重置周期（避免 ② 的改动污染）
const W3 = App.getWeekRange('this');
function shift(base, n){ const d = new Date(base + 'T00:00:00'); d.setDate(d.getDate() + n);
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
// 相对所选周构造（不依赖"今天"是星期几）：应购日落在本周内 → 应回购；应购日早于本周首 → 已逾期
const dueInWeek = shift(W3.start, 3);
const dueBeforeWeek = shift(W3.start, -19);
App.STORE.sales = [
  { source:'sales', patient_name:'逾期甲', phone:'10000000001', sales_time: shift(dueBeforeWeek, -21), product:'百泽安', hospital:'H', pharmacy:'P', physician:'医' },
  { source:'sales', patient_name:'应购乙', phone:'10000000002', sales_time: shift(dueInWeek, -21), product:'百泽安', hospital:'H', pharmacy:'P', physician:'医' },
];
App.STORE.followups = []; App.STORE.cycles = {};
const all = App.buildRows();
console.log('  全量:', all.map(r => r.patient_name + ':' + r.status).join(', '));
const filtered = App.filterRows(all);
console.log('  筛选后:', filtered.map(r => r.patient_name + ':' + r.status).join(', '));
const hasOverdue = all.some(r => r.status === '已逾期');
const hasDue = all.some(r => r.status === '应回购');
console.log(`  已逾期判定=${hasOverdue}（期望true） 应回购判定=${hasDue}（期望true）`);
okAll = okAll && hasOverdue && hasDue;
console.log(okAll ? '\n✅ 三项新需求验证全部通过' : '\n❌ 存在失败项');
if (!okAll) process.exit(1);
