// 分页优化测试：① 每页条数可选 50/100/200 ② 翻页只回表格滚动条、不滚动页面
const fs = require('fs');
const path = require('path');
const elStore = {};
function makeEl(id) {
  return {
    id, _html: '', textContent: '', value: '', disabled: false, style: {}, dataset: {},
    scrollTop: 0,
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
let scrollToCalls = 0;
global.scrollTo = () => { scrollToCalls++; }; // 页面滚动调用计数（应始终为 0）
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
const App = global.AppCore;

function assert(cond, msg) {
  if (!cond) throw new Error('❌ ' + msg);
  console.log('  ✅ ' + msg);
}

(async () => {
  // 构造 114 行假数据，模拟已分析完成
  App.STORE.sales = [];
  App.DATA.rows = Array.from({ length: 114 }, (_, i) => ({
    _key: 'p' + i, product: '百泽安', patient_name: '患者' + i, phone: '1380000' + String(i).padStart(4, '0'),
    physician: '张医生', hospital: '测试医院', pharmacy: '药房', last_purchase: '2026-08-01',
    due_date: '2026-08-22', cycle_days: 21, status: '未到期', substatus: '', days_to_due: 20,
    due_offset: null, purchase_days_ago: null,
  }));
  App.state.page = 1; App.state.pageSize = 50;

  console.log('===== ① 每页条数选择 =====');
  App.renderTable();
  let pg = elStore['pagination'].innerHTML;
  assert(pg.includes('id="pageSizeSel"'), '分页栏包含每页条数下拉');
  assert(pg.includes('value="50" selected') && pg.includes('value="100"') && pg.includes('value="200"'),
    '选项 50(选中)/100/200 齐全');
  assert(pg.includes('第 1/3 页'), '50条/页 → 3 页（114条）');

  // 切到 100 条/页
  elStore['pageSizeSel'].value = '100';
  elStore['pageSizeSel'].onchange({ target: elStore['pageSizeSel'] });
  assert(App.state.pageSize === 100, '切换后 pageSize=100');
  pg = elStore['pagination'].innerHTML;
  assert(pg.includes('<option value="100" selected>'), '下拉选中 100');
  assert(pg.includes('第 1/2 页'), '100条/页 → 2 页');
  const shown = elStore['tbody'].innerHTML.match(/data-row/g) || [];
  assert(shown.length === 100, '当前页渲染 100 行');

  // 切到 200 → 1 页
  elStore['pageSizeSel'].value = '200';
  elStore['pageSizeSel'].onchange({ target: elStore['pageSizeSel'] });
  assert(App.state.pageSize === 200 && elStore['pagination'].innerHTML.includes('第 1/1 页'), '200条/页 → 1 页');

  console.log('\n===== ② 翻页滚动行为 =====');
  App.state.pageSize = 50; App.state.page = 1; App.renderTable();
  // 模拟表格容器滚动到中部，然后翻页
  const wrapEl = elStore['.tbl-wrap'] || elStore['tbl-wrap'];
  wrapEl.scrollTop = 500;
  const nextBtn = elStore['nextPg'];
  nextBtn.onclick();
  assert(App.state.page === 2, '下一页 → 第 2 页');
  assert(wrapEl.scrollTop === 0, '表格容器滚动条已回到顶部');
  assert(scrollToCalls === 0, 'window.scrollTo 从未被调用（页面滚动条不动）');

  console.log('\n✅ 分页优化测试全部通过');
})().catch(e => { console.error('FAIL', e); process.exit(1); });
