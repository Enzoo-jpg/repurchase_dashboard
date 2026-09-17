// 随访回传闭环验证：看板填写 → 导出「随访回传表」→ 重新导入 → 备注/原因自动接续
const fs = require('fs');
const path = require('path');
const ExcelJS = require(path.join(__dirname, '..', 'vendor', 'exceljs.min.js'));

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
let captured = null, alerts = [];
global.URL = { createObjectURL(b){ captured = b; return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.ExcelJS = ExcelJS;
global.alert = m => alerts.push(String(m));
global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const App = global.AppCore;
const M = global.Mapping;
const P = global.Pipeline;

let pass = 0, fail = 0;
const eq = (name, got, want) => {
  if (JSON.stringify(got) === JSON.stringify(want)) { pass++; console.log('  ✅ ' + name + ' → ' + JSON.stringify(got)); }
  else { fail++; console.log('  ❌ ' + name + ' → 实际 ' + JSON.stringify(got) + '，期望 ' + JSON.stringify(want)); }
};
const ok = (name, cond) => { if (cond) { pass++; console.log('  ✅ ' + name); } else { fail++; console.log('  ❌ ' + name); } };

const S = App.state, ST = App.STORE;

// 三条名单行：A 填了备注、B 点选了原因、C 什么都没填
const rowA = { _key: '13800001111', product: '百泽安', patient_name: '张小三', phone: '13800001111', physician: '李大夫',
  executor: '刘倩', fu_note: '原始随访小结：患者表示会按时购药', status: '应回购', repur_part: '应回未回', reason: '', days_to_due: 2 };
const rowB = { _key: '13900002222', product: '百悦泽', patient_name: '李四', phone: '13900002222', physician: '王医生',
  executor: '高金敏', fu_note: '原始随访小结：电话未接通', status: '已逾期', repur_part: '', reason: '', days_to_due: -5 };
const rowC = { _key: '13700003333', product: '百泽安', patient_name: '王五', phone: '13700003333', physician: '赵医生',
  executor: '刘倩', fu_note: '', status: '未到期', repur_part: '', reason: '', days_to_due: 9 };
App.DATA.rows = [rowA, rowB, rowC];

(async () => {
  // 用户在看板里的操作
  const keyA = rowA._key + '::' + rowA.product;
  const keyB = rowB._key + '::' + rowB.product;
  ST.notes[keyA] = '已电话联系，患者说本周五去买';
  ST.reasonOverrides[keyB] = 'dropout_econ';
  // 故意把三组开关都设为「脱敏」，验证回传表仍然强制明文
  S.plainName = false; S.plainDoctor = false; S.plainPhone = false; S.maskMode = 'edge';

  console.log('\n[1] 回传表只包含有填写的行');
  eq('待回传行数', App.callbackRows().length, 2);
  ok('未填写的行被排除', App.callbackRows().every(r => r._key !== rowC._key));

  console.log('\n[2] 导出回传表：结构、明文、内容');
  captured = null;
  await App.doExportCallback();
  ok('已产出文件', !!captured);
  const buf = await captured.arrayBuffer();
  const wb = new ExcelJS.Workbook(); await wb.xlsx.load(buf);
  const ws = wb.worksheets[0];
  eq('sheet 名', ws.name, '随访回传');
  const head = ws.getRow(1).values.filter(Boolean);
  eq('表头', head, ['任务状态', '服务摘要', '患者姓名', '联系电话', '产品', '执行时间', '执行人', '未购药原因', '跟进备注', '随访小结']);
  eq('数据行数', ws.rowCount - 1, 2);

  const idx = {}; head.forEach((h, i) => idx[h] = i + 1);
  const r1 = ws.getRow(2).values, r2 = ws.getRow(3).values;
  eq('签名-任务状态', r1[idx['任务状态']], '已完成');
  eq('签名-服务摘要', r1[idx['服务摘要']], '跟进回传');
  eq('患者姓名强制明文', r1[idx['患者姓名']], '张小三');
  eq('联系电话强制明文', r1[idx['联系电话']], '13800001111');
  eq('产品', r1[idx['产品']], '百泽安');
  eq('执行人', r1[idx['执行人']], '刘倩');
  eq('跟进备注', r1[idx['跟进备注']], '已电话联系，患者说本周五去买');
  eq('随访小结保留原记录', r1[idx['随访小结']], '原始随访小结：患者表示会按时购药');
  eq('未购药原因（人工选定的行）', r2[idx['未购药原因']], '脱落·经济');
  eq('未购药原因（未点选的行留空）', r1[idx['未购药原因']], '');
  ok('执行时间为日期', /^\d{4}-\d{2}-\d{2}$/.test(String(r1[idx['执行时间']])));

  console.log('\n[3] 该文件能被导入侧识别为「随访数据」');
  eq('表类型识别', M.detectTableType(head), 'followup');
  const fakeFile = { name: '随访回传_2026-09-10.xlsx', arrayBuffer: async () => buf };
  const res = await P.processFiles([fakeFile]);
  eq('识别为 0 条销售', res.sales.length, 0);
  eq('识别为 2 条随访', res.followups.length, 2);
  const fuA = res.followups.find(f => f.phone === '13800001111');
  const fuB = res.followups.find(f => f.phone === '13900002222');
  ok('回传记录 A 存在且字段完整', !!fuA);
  if (fuA) {
    eq('A-姓名', fuA.patient_name, '张小三');
    eq('A-产品', fuA.product, '百泽安');
    eq('A-回传备注', fuA.callback_note, '已电话联系，患者说本周五去买');
    eq('A-随访小结未被备注污染', fuA.follow_note, '原始随访小结：患者表示会按时购药');
    eq('A-执行人', fuA.executor, '刘倩');
    eq('A-服务摘要(回传标记)', fuA.summary_type, '跟进回传');
  }
  if (fuB) eq('B-未购药原因标签', fuB.reason_label, '脱落·经济');

  console.log('\n[4] 原因标签 → 分类 key 反查');
  eq('脱落·经济', App.reasonKeyFromLabel('脱落·经济'), 'dropout_econ');
  eq('推迟购药', App.reasonKeyFromLabel('推迟购药'), 'delay');
  eq('延长周期·医嘱', App.reasonKeyFromLabel('延长周期·医嘱'), 'prolong_doctor');
  eq('随访失败未探寻原因', App.reasonKeyFromLabel('随访失败未探寻原因'), 'fuFail');
  eq('未知标签返回空', App.reasonKeyFromLabel('不存在的分类'), '');

  console.log('\n[5] 新一轮分析时自动接续（模拟清空后重新上传）');
  ST.notes = {}; ST.reasonOverrides = {};
  const cb = App.applyCallbackRecords(res.followups);
  eq('接续备注数', cb.nNote, 1);
  eq('接续原因数', cb.nReason, 1);
  eq('备注回到「跟进备注」列', ST.notes[keyA], '已电话联系，患者说本周五去买');
  eq('原因等同人工点选', ST.reasonOverrides[keyB], 'dropout_econ');

  console.log('\n[6] 不覆盖本轮已编辑的内容');
  ST.notes[keyA] = '用户刚刚改过的新内容';
  ST.reasonOverrides[keyB] = 'delay';
  App.applyCallbackRecords(res.followups);
  eq('已有备注不被导入覆盖', ST.notes[keyA], '用户刚刚改过的新内容');
  eq('已有原因不被导入覆盖', ST.reasonOverrides[keyB], 'delay');

  console.log('\n[7] 无可回传内容时给出提示且不产出文件');
  ST.notes = {}; ST.reasonOverrides = {}; captured = null; alerts = [];
  await App.doExportCallback();
  eq('未产出文件', captured, null);
  ok('给出中文提示', alerts.length === 1 && alerts[0].includes('没有可回传的内容'));

  console.log('\n结果: ' + pass + ' 通过 / ' + fail + ' 失败');
  if (fail) process.exit(1);
})();
