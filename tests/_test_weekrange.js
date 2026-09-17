// 「本周」时间范围可配置（周截止日）回归验证
// 覆盖：自然周（周日截止）/ 上周六~本周五（周五截止）/ 周日~周六（周六截止）、边界日、上周下周平移、默认值
const fs = require('fs');
const path = require('path');

function makeEl(id) {
  return { id, innerHTML: '', textContent: '', value: '', disabled: false, style: {}, dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; }, appendChild(){}, insertBefore(){}, removeChild(){}, querySelectorAll(){ return []; },
    click(){}, remove(){},
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
const S = App.state;

let pass = 0, fail = 0;
const eq = (name, got, want) => {
  if (got === want) { pass++; console.log('  ✅ ' + name + ' → ' + JSON.stringify(got)); }
  else { fail++; console.log('  ❌ ' + name + ' → 实际 ' + JSON.stringify(got) + '，期望 ' + JSON.stringify(want)); }
};
const WD = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"];
const wdOf = d => WD[new Date(d + "T00:00:00").getDay()];
const rng = sel => { const w = App.getWeekRange(sel); return w.start + "~" + w.end; };

// 2026-09-10 是周四
S.refDate = '2026-09-10';

console.log('\n[1] 默认值 = 周日截止（自然周 周一~周日），保持既有行为');
eq('state.weekEnd 默认', S.weekEnd, 0);
S.weekEnd = 0;
eq('本周（周一~周日）', rng('this'), '2026-09-07~2026-09-13');
eq('上周', rng('last'), '2026-08-31~2026-09-06');
eq('下周', rng('next'), '2026-09-14~2026-09-20');
eq('截止日星期核对', wdOf(App.getWeekRange('this').end), '周日');

console.log('\n[2] 周五截止 = 上周六~本周五（用户举例）');
S.weekEnd = 5;
eq('本周（上周六~本周五）', rng('this'), '2026-09-05~2026-09-11');
eq('上周', rng('last'), '2026-08-29~2026-09-04');
eq('下周', rng('next'), '2026-09-12~2026-09-18');
eq('起始日星期核对（应为周六）', wdOf(App.getWeekRange('this').start), '周六');
eq('截止日星期核对（应为周五）', wdOf(App.getWeekRange('this').end), '周五');

console.log('\n[3] 周六截止 = 周日~周六');
S.weekEnd = 6;
eq('本周（周日~周六）', rng('this'), '2026-09-06~2026-09-12');
eq('截止日星期核对', wdOf(App.getWeekRange('this').end), '周六');

console.log('\n[4] 边界：参考日恰好是截止日当天 / 是起始日当天');
S.weekEnd = 5;
S.refDate = '2026-09-11';               // 周五 = 截止日当天 → 本周就是含它的那一周
eq('参考日=周五（截止日当天）', rng('this'), '2026-09-05~2026-09-11');
S.refDate = '2026-09-05';               // 周六 = 起始日当天
eq('参考日=周六（起始日当天，同一窗口）', rng('this'), '2026-09-05~2026-09-11');
S.refDate = '2026-09-12';               // 周六 → 上一个窗口已截止，滚入新窗口
eq('参考日=次周周六 → 窗口滚动', rng('this'), '2026-09-12~2026-09-18');
S.refDate = '2026-09-13';               // 周日
eq('参考日=周日（仍属同一窗口，至次周五）', rng('this'), '2026-09-12~2026-09-18');

console.log('\n[5] 七种截止日都自洽（窗口 7 天 / 必含参考日 / 截止日星期正确 / 上周下周各平移 7 天）');
S.refDate = '2026-09-10';
let allOk = true, detail = '';
const dayMs = 86400000;
for (let d = 0; d < 7; d++) {
  S.weekEnd = d;
  const w = App.getWeekRange('this');
  const last = App.getWeekRange('last'), next = App.getWeekRange('next');
  const span = (new Date(w.end) - new Date(w.start)) / dayMs;
  const shiftL = (new Date(w.end) - new Date(last.end)) / dayMs;
  const shiftN = (new Date(next.end) - new Date(w.end)) / dayMs;
  const contains = S.refDate >= w.start && S.refDate <= w.end;
  const ok = wdOf(w.end) === WD[d] && span === 6 && shiftL === 7 && shiftN === 7 && contains;
  if (!ok) { allOk = false; detail += ` [截止${WD[d]}: span=${span} L=${shiftL} N=${shiftN} 含参考日=${contains}]`; }
}
eq('7 种截止日全部自洽', allOk ? 'OK' : 'FAIL' + detail, 'OK');

console.log('\n[6] 小结「下周预计」窗口 = 所选周之后紧邻 7 天，随截止日平移');
const iso = d => d.toISOString().slice(0, 10);
const nextWin = () => {
  const W = App.getWeekRange('this');
  const a = new Date(W.end); a.setDate(a.getDate() + 1);
  const b = new Date(W.end); b.setDate(b.getDate() + 7);
  return iso(a) + "~" + iso(b);
};
S.weekEnd = 0;
eq('周一~周日（周日截止）下周窗口', nextWin(), '2026-09-14~2026-09-20');
S.weekEnd = 5;
eq('上周六~本周五（周五截止）下周窗口', nextWin(), '2026-09-12~2026-09-18');

console.log('\n[7] 切换截止日不应丢失「参考日」与「所选周」');
S.refDate = '2026-09-10'; S.weekEnd = 5;
const before = rng('this');
S.weekEnd = 0;   // 切到自然周
S.weekEnd = 5;   // 再切回周五截止
eq('切走再切回，范围复原', rng('this'), before);
eq('切换过程中参考日不变', S.refDate, '2026-09-10');

S.weekEnd = 0; S.refDate = '';
console.log('\n结果: ' + pass + ' 通过 / ' + fail + ' 失败');
if (fail) process.exit(1);
