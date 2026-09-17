// 验证：销售明细(3) + 历史任务.xls → buildRows 后随访信号覆盖率
const fs = require('fs');
const path = require('path');
const os = require('os');
// 样例 Excel 目录：默认取当前用户的 Downloads，可用环境变量 SAMPLE_DIR 覆盖
const SAMPLE_DIR = process.env.SAMPLE_DIR || path.join(os.homedir(), 'Downloads');
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
const P = global.Pipeline, App = global.AppCore;

(async () => {
  function fileObj(p, name) {
    const buf = fs.readFileSync(p);
    return { name: name || path.basename(p), size: buf.length,
      async arrayBuffer(){ return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); } };
  }
  const files = [
    fileObj(path.join(SAMPLE_DIR, '销售明细查询报表 (3).xlsx'), '销售.xlsx'),
    fileObj(path.join(SAMPLE_DIR, '历史任务.xls'), '历史任务.xls'),
  ];
  // 周期表若存在则带上
  const cycPath = path.join(SAMPLE_DIR, '患者用药周期表.xlsx');
  if (fs.existsSync(cycPath)) files.push(fileObj(cycPath, '周期.xlsx'));

  const res = await P.processFiles(files);
  console.log('[解析] 销售:', res.sales.length, '| 随访:', res.followups.length, '| 周期:', Object.keys(res.cycles).length);
  console.log('[销售] 品种:', [...new Set(res.sales.map(s => s.product))].join(' / '));
  console.log('[随访] 品种:', [...new Set(res.followups.map(f => f.product))].join(' / '));

  App.STORE.sales = res.sales; App.STORE.followups = res.followups; App.STORE.cycles = res.cycles;
  App.state.stdCycle = {};
  App.ensureStdCycles([...new Set(res.sales.map(s => s.product))]);
  App.state.start = null; App.state.end = null; App.state.advance = 7;
  const rows = App.buildRows();
  console.log('\n[名单] 行数:', rows.length);
  const bySt = {}; for (const r of rows) bySt[r.status] = (bySt[r.status] || 0) + 1;
  console.log('[名单] 状态分布:', JSON.stringify(bySt));
  const fuMatch = rows.filter(r => r.fu_time).length;
  const fuSig = rows.filter(r => r.fu_signal).length;
  const fuNote = rows.filter(r => r.fu_note).length;
  console.log('[随访匹配] 有随访:', fuMatch, '/', rows.length);
  console.log('[随访信号] 非空:', fuSig, '/', rows.length);
  console.log('[随访小结] 非空:', fuNote, '/', rows.length);
  // 信号样例
  const withSig = rows.filter(r => r.fu_signal);
  for (const r of withSig.slice(0, 8)) {
    console.log('  ', r.patient_name, r.product, '| 信号:', r.fu_signal.slice(0, 30));
  }
})().catch(e => { console.error('FAIL', e); process.exit(1); });
