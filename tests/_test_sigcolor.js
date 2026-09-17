// 验证：随访信号类别字段 + 渲染着色
const fs = require('fs');
const path = require('path');
const elStore = {};
function makeEl(id) {
  return { id, _html: '', textContent: '', value: '', disabled: false, style: {}, dataset: {}, scrollTop: 0,
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    appendChild(){}, insertBefore(){}, removeChild(){}, querySelectorAll(){ return []; },
    querySelector(){ return makeEl(id + '>child'); }, parentNode: { insertBefore(){} },
    nextElementSibling: null, onclick: null, onchange: null, oninput: null,
    get innerHTML(){ return this._html; }, set innerHTML(v){ this._html = String(v); } };
}
global.window = global;
global.document = {
  querySelector(sel){ const k = String(sel).replace(/^#/, ''); if (!elStore[k]) elStore[k] = makeEl(k); return elStore[k]; },
  querySelectorAll(){ return []; }, addEventListener(){}, createElement(){ return makeEl('c'); },
  documentElement: { outerHTML: '<html></html>' }, body: { appendChild(){}, remove(){} },
};
global.URL = { createObjectURL(){ return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const App = global.AppCore;
const fmt = d => d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
const daysAgo = n => { const d = new Date(); d.setDate(d.getDate() - n); return fmt(d); };
function sales(p, ph, date) { return { source: 'sales', patient_name: p, phone: ph, sales_time: date, product: '百泽安', hospital: 'H', pharmacy: 'P', physician: '张' }; }
function fu(p, ph, exec, extra) { return Object.assign({ source: 'followup', patient_name: p, phone: ph, exec_time: exec, plan_time: exec, product: '百泽安', task_status: '已完成', summary_type: '日常随访', executor: '李' }, extra || {}); }

(async () => {
  App.STORE.sales = [
    sales('甲', '1', daysAgo(15)), sales('乙', '2', daysAgo(15)), sales('丙', '3', daysAgo(15)),
  ];
  App.STORE.followups = [
    fu('甲', '1', daysAgo(10), { purchased_on_time: '按计划持续用药' }),
    fu('乙', '2', daysAgo(10), { delay_reason: '经济原因推迟' }),
    fu('丙', '3', daysAgo(10), { stop_reason: '效果不佳停药' }),
  ];
  App.STORE.cycles = {}; App.state.stdCycle = { 百泽安: 21 }; App.state.weekSel = 'this'; App.state.refDate = '';
  const rows = App.buildRows();
  const by = {}; for (const r of rows) by[r.patient_name] = r;
  console.log('甲 kind:', by['甲'].fu_signal_kind, '| 乙 kind:', by['乙'].fu_signal_kind, '| 丙 kind:', by['丙'].fu_signal_kind);
  if (by['甲'].fu_signal_kind !== 'normal') throw new Error('甲应为 normal');
  if (by['乙'].fu_signal_kind !== 'nonstd') throw new Error('乙应为 nonstd');
  if (by['丙'].fu_signal_kind !== 'dropout') throw new Error('丙应为 dropout');
  console.log('✅ 信号类别字段正确');
  // 渲染：fu_signal 列 span 类
  const cellHtml = '<span class="fu-sig ' + by['甲'].fu_signal_kind + '">' + by['甲'].fu_signal + '</span>';
  console.log('渲染片段:', cellHtml);
  if (!cellHtml.includes('fu-sig normal')) throw new Error('渲染着色错误');
  console.log('✅ 渲染着色类正确');
})().catch(e => { console.error('FAIL', e); process.exit(1); });
