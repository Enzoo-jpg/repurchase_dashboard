// 导出内容 + 浅色样式验证：真实调用 app.js 的 doExport 逻辑
// 通过 DOM stub + 捕获 Blob 回读校验：表头加粗、浅色填充、数据行非空
const fs = require('fs');
const path = require('path');
const ExcelJS = require(path.join(__dirname, '..', 'vendor', 'exceljs.min.js'));

function makeEl(id) {
  return { id, innerHTML: '', textContent: '', value: '', disabled: false, style: {}, dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    appendChild(){}, insertBefore(){}, removeChild(){}, querySelectorAll(){ return []; },
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
let captured = null;
global.URL = { createObjectURL(b){ captured = b; return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.ExcelJS = ExcelJS; // 模拟浏览器 <script> 全局
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const App = global.AppCore;

(async () => {
  // 构造数据行（对应真实 buildRows 输出结构）
  App.DATA.rows = [
    { _key: '13800001111', product: '百泽安', patient_name: '张三', phone: '13800001111', physician: '李医生',
      hospital: '人民医院', pharmacy: '药房A', last_purchase: '2026-08-01', cycle_days: 21, due_date: '2026-08-22',
      days_to_due: 2, fu_time: '2026-08-15', fu_type: '复购确认随访任务', executor: '刘倩', fu_note: '患者表示会按时购药',
      fu_signal: '', status: '应回购', repur_part: '应回未回', reason: 'delay', due_in_week: true, member_id: 'VIP88888888' },
    { _key: '13900002222', product: '百悦泽', patient_name: '李四', phone: '13900002222', physician: '王医生',
      hospital: '中医院', pharmacy: '药房B', last_purchase: '2026-07-01', cycle_days: 28, due_date: '2026-07-29',
      days_to_due: -5, fu_time: '2026-07-20', fu_type: '日常随访任务', executor: '高金敏', fu_note: '',
      fu_signal: '未按时购药', status: '已逾期', reason: '' },
  ];
  await App.doExport(true); // 脱敏导出
  if (!captured) throw new Error('未捕获到导出 Blob');
  console.log('[导出] Blob 大小:', captured.size, '字节');
  if (captured.size < 3000) throw new Error('导出内容过小，疑似空表');

  const ab = await captured.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(ab);
  const ws = wb.getWorksheet('预计购药名单');
  const lastRow = ws.rowCount;
  const lastCol = ws.columnCount;
  console.log('[导出] 行数:', lastRow, '列数:', lastCol);
  if (lastRow < 3) throw new Error('数据行缺失，导出疑似空内容');

  // 表头
  const h = ws.getRow(1);
  const h1 = h.getCell(1);
  console.log('[表头] A1=', h1.value, '| 加粗:', h1.font.bold, '| 底色:', h1.fill.fgColor && h1.fill.fgColor.argb, '期望 FFE6F1FB(浅蓝)');
  if (h1.font.bold !== true || (h1.fill.fgColor && h1.fill.fgColor.argb) !== 'FFE6F1FB') throw new Error('表头样式不符');

  // 行2：应回购 / 未购药原因=延迟用药
  const r2 = ws.getRow(2);
  const idxStatus = 16, idxReason = 17, idxDays = 12; // 与 LIST_COLS 对应（1-based；新增「购药数量」列后，days=12/status=16/reason=17）
  console.log('[行2] 患者=', r2.getCell(1).value, '| 状态列=', r2.getCell(idxStatus).value, '(期望 应回未回)',
    '| 底色:', r2.getCell(idxStatus).fill.fgColor && r2.getCell(idxStatus).fill.fgColor.argb, '期望 FFFCEBEB(浅红)');
  console.log('[行2] 未购药原因列=', r2.getCell(idxReason).value,
    '| 字色:', r2.getCell(idxReason).font.color && r2.getCell(idxReason).font.color.argb, '期望 FFE8590C(橙)');
  console.log('[行2] 随访小结=', r2.getCell(lastCol).value);

  // 行3：已逾期 + 逾期标红
  const r3 = ws.getRow(3);
  console.log('[行3] 状态列=', r3.getCell(idxStatus).value,
    '| 底色:', r3.getCell(idxStatus).fill.fgColor && r3.getCell(idxStatus).fill.fgColor.argb, '期望 FFFAEEDA(浅橙)');
  console.log('[行3] 距今天数=', r3.getCell(idxDays).value,
    '| 字色:', r3.getCell(idxDays).font.color && r3.getCell(idxDays).font.color.argb, '期望 FFE03131(红)');

  // 周内应购行：无专属样式列整行浅金底（患者列第1列）
  console.log('[行2] 周内应购底色(患者列):', r2.getCell(1).fill.fgColor && r2.getCell(1).fill.fgColor.argb, '期望 FFFFF3D6(浅金)');
  console.log('[行3] 非周内行无底色(患者列):', r3.getCell(1).fill && r3.getCell(1).fill.fgColor && r3.getCell(1).fill.fgColor.argb, '(应为空/无)');

  const ok = r2.getCell(1).value === '张*' &&
    (r2.getCell(1).fill.fgColor && r2.getCell(1).fill.fgColor.argb) === 'FFFFF3D6' &&
    !(r3.getCell(1).fill && r3.getCell(1).fill.fgColor && r3.getCell(1).fill.fgColor.argb) &&
    r2.getCell(idxStatus).value === '应回未回' &&
    r2.getCell(idxStatus).fill.fgColor.argb === 'FFFCEBEB' &&
    r2.getCell(idxReason).value === '推迟购药' &&
    r2.getCell(idxReason).font.color.argb === 'FFE8590C' &&
    r3.getCell(idxStatus).fill.fgColor.argb === 'FFFAEEDA' &&
    r3.getCell(idxDays).font.color.argb === 'FFE03131';
  console.log(ok ? '\n✅ 导出内容 + 浅色样式验证通过' : '\n❌ 校验失败');
  if (!ok) process.exit(1);

  // —— 追加：仅会员号（maskMode=id）导出：患者列=会员号、医生列空 ——
  App.state.maskMode = 'id';
  captured = null;
  await App.doExport(true);
  const ab2 = await captured.arrayBuffer();
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(ab2);
  const ws2 = wb2.getWorksheet('预计购药名单');
  const r2id = ws2.getRow(2);
  console.log('[仅会员号] 患者列=', r2id.getCell(1).value, '(期望 VIP88888888) | 医生列=', r2id.getCell(3).value, '(期望空)');
  const okId = r2id.getCell(1).value === 'VIP88888888' && (r2id.getCell(3).value === null || r2id.getCell(3).value === '');
  console.log(okId ? '✅ 仅会员号导出验证通过' : '❌ 仅会员号导出失败');
  App.state.maskMode = 'edge';
  if (!okId) process.exit(1);
})().catch(e => { console.error('FAIL', e); process.exit(1); });
