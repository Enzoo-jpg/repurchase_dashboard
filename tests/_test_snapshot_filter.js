// 快照模式多次筛选不丢失数据回归测试
// 背景：修复前 refresh() 快照分支以 DATA.rows 为基准，而 DATA.rows 每次被 filterRows 覆盖为子集，
//       连续点击筛选/取消会导致基准越来越小，最终所有信息被清零。
// 修复：引入不可变基准 SNAP_BASE，快照模式筛选永远基于全量行。
const fs = require('fs');
const path = require('path');
const os = require('os');
// 样例 Excel 目录：默认取当前用户的 Downloads，可用环境变量 SAMPLE_DIR 覆盖
const SAMPLE_DIR = process.env.SAMPLE_DIR || path.join(os.homedir(), 'Downloads');
const elStore = {};
function makeEl(id) {
  return {
    id, _html: '', textContent: '', value: '', disabled: false, style: {}, dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    appendChild(){}, insertBefore(){}, removeChild(){},
    querySelectorAll(){ return []; },
    querySelector(){ return makeEl(id + '>child'); },
    parentNode: { insertBefore(){} },
    nextElementSibling: null, onclick: null, onchange: null, oninput: null,
    get innerHTML(){ return this._html; },
    set innerHTML(v){ this._html = String(v); },
  };
}
global.window = global;
global.document = {
  querySelector(sel){
    const key = String(sel).replace(/^#/, '');
    if (!elStore[key]) elStore[key] = makeEl(key);
    return elStore[key];
  },
  querySelectorAll(){ return []; },
  addEventListener(){},
  createElement(){ return makeEl('created'); },
  documentElement: { outerHTML: '<html><body></body></html>' },
  body: { appendChild(){}, remove(){} },
};
global.URL = { createObjectURL(){ return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const P = global.Pipeline, App = global.AppCore;

const fmt = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
const addDays = (s, n) => { const d = new Date(s); d.setDate(d.getDate() + n); return fmt(d); };
const today = fmt(new Date());

function assert(cond, msg) {
  if (!cond) throw new Error('❌ ' + msg);
  console.log('  ✅ ' + msg);
}

(async () => {
  // 1. 用真实数据生成全量行，并构造混合状态快照（应回购/已逾期/未到期）
  const DATA_FILES = [
    path.join(SAMPLE_DIR, '销售明细查询报表 (47).xlsx'),
    path.join(SAMPLE_DIR, '随访任务导出 (5).xlsx'),
    path.join(SAMPLE_DIR, '患者用药周期表.xlsx'),
  ];
  const missing = DATA_FILES.filter(p => !fs.existsSync(p));
  if (missing.length) {
    console.log('⏭️ 跳过快照筛选回归：样例数据文件不在 ' + SAMPLE_DIR + '（' +
      missing.map(p => path.basename(p)).join('、') + ' 缺失）');
    process.exit(0);
  }
  function fileObj(p, name) {
    const buf = fs.readFileSync(p);
    return { name: name || path.basename(p), size: buf.length,
      async arrayBuffer(){ return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); } };
  }
  const res = await P.processFiles([
    fileObj(DATA_FILES[0], '销售.xlsx'),
    fileObj(DATA_FILES[1], '随访.xlsx'),
    fileObj(DATA_FILES[2], '周期.xlsx'),
  ]);
  App.STORE.sales = res.sales; App.STORE.followups = res.followups; App.STORE.cycles = res.cycles;
  const all = App.buildRows();
  const total = all.length;
  console.log(`[准备] 真实数据全量行 = ${total}`);

  // 构造混合状态：前 10 行→应回购(正常状态,应购日=今天)，10~19→已逾期(应购日=10天前)，其余→未到期
  // （显式覆盖全部行状态，避免真实数据本身含 应回购/已回购 干扰筛选计数）
  const rows = all.map((r, i) => Object.assign({}, r));
  rows.forEach((r, i) => {
    if (i < 10) { r.status = '应回购'; r.substatus = '正常状态'; r.due_date = today; r.days_to_due = 3; }
    else if (i < 20) { r.status = '已逾期'; r.substatus = ''; r.due_date = addDays(today, -10); r.days_to_due = -10; }
    else { r.status = '未到期'; r.substatus = ''; r.due_date = addDays(today, 20); r.days_to_due = 20; }
  });
  const snap = {
    desen: true, buildAt: new Date().toISOString(), notes: {}, subOverrides: {},
    rows: rows.map(r => Object.assign({}, r, { _matched: undefined, _key: undefined })),
    summary: App.buildSummary(rows),
    state: { weekSel: 'this', refDate: '', stdCycle: { 百泽安: 21, 百悦泽: 28 }, maskMode: 'edge' },
  };

  console.log('\n===== 打开快照 =====');
  App.loadSnapshot(snap);
  assert(App.STORE.notes !== null, '快照加载完成');
  console.log(`  [初始] DATA.rows = ${App.DATA.rows.length}（期望 ${total}）`);
  assert(App.DATA.rows.length === total, `打开快照后名单 = 全量 ${total}`);

  console.log('\n===== 连续筛选/取消（核心回归） =====');
  // 单状态筛选后取消 → 必须恢复全量（修复前这里会停留在 10）
  App.state.cats.add('应回购'); await App.refresh();
  assert(App.DATA.rows.length === 10, `筛选「应回购」→ ${App.DATA.rows.length} 行（期望 10）`);
  App.state.cats.delete('应回购'); await App.refresh();
  assert(App.DATA.rows.length === total, `取消「应回购」→ 恢复全量 ${total} 行（修复前此处会丢失）`);

  // 多状态叠加再清空 → 恢复全量
  App.state.cats.add('应回购'); App.state.cats.add('已逾期'); await App.refresh();
  assert(App.DATA.rows.length === 20, `叠加「应回购+已逾期」→ 20 行`);
  App.state.cats.clear(); await App.refresh();
  assert(App.DATA.rows.length === total, `清空状态筛选 → 恢复全量 ${total} 行`);

  console.log('\n===== 下钻子状态筛选 =====');
  App.state.cats.add('应回购'); App.state.subs.add('正常状态'); await App.refresh();
  assert(App.DATA.rows.length === 10, `「应回购+下钻正常状态」→ 10 行`);
  App.state.cats.clear(); App.state.subs.clear(); await App.refresh();
  assert(App.DATA.rows.length === total, `清空下钻 → 恢复全量 ${total} 行`);

  console.log('\n===== 快照模式周控件只读（名单不随周变化） =====');
  // 快照行状态为保存时判定结果，无法按周重算 → 周选择器只读；切换 weekSel 不影响名单
  App.state.weekSel = 'last'; await App.refresh();
  assert(App.DATA.rows.length === total, `快照模式切换周视图不重算 → 仍全量 ${total} 行`);
  App.state.weekSel = 'this'; await App.refresh();
  assert(App.DATA.rows.length === total, `恢复本周 → 全量 ${total} 行`);

  console.log('\n===== 关键词搜索 =====');
  App.state.q = '不存在的关键词xyz不存在'; await App.refresh();
  assert(App.DATA.rows.length === 0, `搜索无匹配 → 0 行`);
  App.state.q = ''; await App.refresh();
  assert(App.DATA.rows.length === total, `清空搜索 → 恢复全量 ${total} 行`);

  console.log('\n✅ 快照模式多次筛选不丢失数据 —— 全部通过');
})().catch(e => { console.error('FAIL', e); process.exit(1); });
