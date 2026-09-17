// 三组脱敏开关「相互独立」回归验证（重点：医生不被患者姓名绑定）
// 覆盖：disp 展示层 4 种组合 + doExport 脱敏导出是否跟随开关
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
// 构造三组脱敏按钮桩（模拟真实点击路径，直接复现「按钮点了没反应」）
const dtBtns = [];
["name", "doctor", "phone"].forEach(field => {
  const grp = { querySelectorAll: () => dtBtns.filter(b => b.dataset.field === field) };
  ["mask", "plain"].forEach(mode => {
    const b = makeEl('dt-' + field + '-' + mode);
    b.dataset = { field, mode };
    b.parentElement = grp;
    dtBtns.push(b);
  });
});
const clickDt = (field, mode) => dtBtns.find(b => b.dataset.field === field && b.dataset.mode === mode).onclick();

global.window = global;
global.document = {
  querySelector(sel){ return makeEl(String(sel)); },
  querySelectorAll(sel){ return String(sel) === '.dt-btn' ? dtBtns : []; },
  addEventListener(){}, createElement(){ return makeEl('created'); },
  documentElement: { outerHTML: '<html></html>' }, body: { appendChild(){}, remove(){}, removeChild(){} },
};
let captured = null;
global.URL = { createObjectURL(b){ captured = b; return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.ExcelJS = ExcelJS;
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const App = global.AppCore;

let pass = 0, fail = 0;
const eq = (name, got, want) => {
  if (got === want) { pass++; console.log('  ✅ ' + name + ' → ' + JSON.stringify(got)); }
  else { fail++; console.log('  ❌ ' + name + ' → 实际 ' + JSON.stringify(got) + '，期望 ' + JSON.stringify(want)); }
};

// 姓名用 3 字、医生用 3 字，便于观察 edge（保留首+末）掩码效果
const row = { _key: '13800001111', product: '百泽安', patient_name: '张小三', phone: '13800001111',
  physician: '李大夫', member_id: 'VIP88888888', status: '应回购', repur_part: '应回未回',
  reason: 'delay', days_to_due: 2, due_in_week: true };

(async () => {
  const S = App.state;
  S.maskMode = 'edge';   // edge=保留首+末：张小三→张*三、李大夫→李*夫

  // ---- 1. 展示层：医生开关必须独立于姓名开关 ----
  console.log('\n[1] disp 展示层（maskMode=edge）');
  S.plainName = false; S.plainDoctor = false;
  eq('姓名脱敏 / 医生脱敏', [App.disp(row, 'patient_name'), App.disp(row, 'physician')].join(' | '), '张*三 | 李*夫');

  S.plainName = false; S.plainDoctor = true;
  eq('姓名脱敏 / 医生不脱敏 ← 本次修复的核心场景', [App.disp(row, 'patient_name'), App.disp(row, 'physician')].join(' | '), '张*三 | 李大夫');

  S.plainName = true; S.plainDoctor = false;
  eq('姓名不脱敏 / 医生脱敏', [App.disp(row, 'patient_name'), App.disp(row, 'physician')].join(' | '), '张小三 | 李*夫');

  S.plainName = true; S.plainDoctor = true;
  eq('姓名不脱敏 / 医生不脱敏', [App.disp(row, 'patient_name'), App.disp(row, 'physician')].join(' | '), '张小三 | 李大夫');

  // ---- 2. 电话开关独立 ----
  console.log('\n[2] disp 电话开关独立');
  S.plainName = false; S.plainDoctor = false; S.plainPhone = true;
  eq('仅电话不脱敏，姓名/医生仍脱敏', [App.disp(row, 'patient_name'), App.disp(row, 'physician'), App.disp(row, 'phone')].join(' | '), '张*三 | 李*夫 | 13800001111');
  S.plainPhone = false;

  // ---- 3. doExport 是否跟随开关（合并为单一按钮后无参） ----
  console.log('\n[3] 导出名单 跟随开关');
  App.DATA.rows = [row];
  const readback = async () => {
    const wb = new ExcelJS.Workbook(); await wb.xlsx.load(await captured.arrayBuffer());
    const ws = wb.worksheets[0];
    const head = ws.getRow(1).values.filter(Boolean);
    return { ws, v: ws.getRow(2).values, i: { name: head.indexOf('患者'), doc: head.indexOf('医生'), phone: head.indexOf('电话') } };
  };

  S.plainName = false; S.plainDoctor = true; S.plainPhone = false;
  await App.doExport();
  let { v, i } = await readback();
  eq('导出名单：姓名列仍脱敏', v[i.name + 1], '张*三');
  eq('导出名单：医生列跟随「不脱敏」输出明文', v[i.doc + 1], '李大夫');
  eq('导出名单：电话列仍脱敏', v[i.phone + 1], '138****1111');

  S.plainDoctor = false;
  await App.doExport();
  ({ v, i } = await readback());
  eq('导出名单：医生开关复位后回到脱敏', v[i.doc + 1], '李*夫');

  // ---- 4. 三组开关全切「不脱敏」→ 导出全明文（原「未脱敏名单」能力） ----
  console.log('\n[4] 三组开关全切「不脱敏」→ 全明文导出');
  S.plainName = true; S.plainDoctor = true; S.plainPhone = true;
  await App.doExport();
  ({ v, i } = await readback());
  eq('全明文导出：姓名', v[i.name + 1], '张小三');
  eq('全明文导出：医生', v[i.doc + 1], '李大夫');
  eq('全明文导出：电话', v[i.phone + 1], '13800001111');
  S.plainName = false; S.plainDoctor = false; S.plainPhone = false;

  // ---- 5. 「仅会员号」模式下医生开关仍生效 ----
  console.log('\n[5] maskMode=id（仅会员号）时医生开关独立生效');
  S.maskMode = 'id'; S.plainDoctor = false;
  eq('id 模式 + 医生脱敏 → 医生列占位', App.disp(row, 'physician'), '—');
  S.plainDoctor = true;
  eq('id 模式 + 医生不脱敏 → 医生列明文', App.disp(row, 'physician'), '李大夫');
  S.plainName = false;
  eq('id 模式 + 姓名脱敏 → 患者列显示会员号', App.disp(row, 'patient_name'), 'VIP88888888');
  S.plainName = true;
  eq('id 模式 + 姓名不脱敏 → 患者列显示姓名', App.disp(row, 'patient_name'), '张小三');
  S.maskMode = 'edge'; S.plainName = false; S.plainDoctor = false;

  // ---- 6. 点击按钮的真实路径（复现用户反馈的「按钮无效/被绑定」） ----
  console.log('\n[6] 点击按钮 → 状态是否真的改变');
  S.maskMode = 'edge';
  try {
    clickDt('doctor', 'plain');
    eq('点「医生：不脱敏」→ state.plainDoctor', S.plainDoctor, true);
    eq('点「医生：不脱敏」不应带动 state.plainName', S.plainName, false);
    eq('点后医生列取值', App.disp(row, 'physician'), '李大夫');

    clickDt('name', 'plain');
    eq('点「姓名：不脱敏」→ state.plainName', S.plainName, true);
    eq('点「姓名：不脱敏」不应复位 state.plainDoctor', S.plainDoctor, true);
    eq('此时医生列仍明文', App.disp(row, 'physician'), '李大夫');

    clickDt('doctor', 'mask');
    eq('点「医生：脱敏」→ state.plainDoctor', S.plainDoctor, false);
    eq('医生列回到掩码', App.disp(row, 'physician'), '李*夫');
    eq('姓名列仍为明文（互不影响）', App.disp(row, 'patient_name'), '张小三');

    clickDt('name', 'mask');
    eq('姓名复位后回到掩码', App.disp(row, 'patient_name'), '张*三');
  } catch (e) {
    fail++; console.log('  ❌ 点击路径抛错: ' + e.message);
  }

  console.log('\n结果: ' + pass + ' 通过 / ' + fail + ' 失败');
  if (fail) process.exit(1);
})();
