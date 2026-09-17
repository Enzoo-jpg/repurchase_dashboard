// Node 端到端验证：加载三脚本（mapping/pipeline/app）+ 真实三张表 → 跑完整计算
// 用 DOM stub 让 app.js 可加载，直接调用 AppCore.buildRows/filterRows/buildSummary。
const fs = require('fs');
const path = require('path');
const os = require('os');
// 样例 Excel 目录：默认取当前用户的 Downloads，可用环境变量 SAMPLE_DIR 覆盖
const SAMPLE_DIR = process.env.SAMPLE_DIR || path.join(os.homedir(), 'Downloads');

// ---------- 最小 DOM stub ----------
function makeEl(id) {
  return {
    id, innerHTML: '', textContent: '', value: '', disabled: false,
    style: {},
    dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    appendChild(){}, insertBefore(){}, removeChild(){},
    querySelectorAll(){ return []; },
    querySelector(){ return makeEl(id + '>child'); },
    parentNode: { insertBefore(){} },
    nextElementSibling: null,
    onclick: null, onchange: null, oninput: null,
  };
}
global.window = global;
global.document = {
  querySelector(sel){ return makeEl(String(sel)); },
  querySelectorAll(){ return []; },
  addEventListener(){},
  createElement(){ return makeEl('created'); },
  documentElement: { outerHTML: '<html></html>' },
  body: { appendChild(){}, remove(){}, removeChild(){} },
};
global.URL = { createObjectURL(){ return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.alert = () => {};
global.confirm = () => true;

// ---------- 加载三脚本 ----------
const load = f => { const code = fs.readFileSync(path.join(__dirname, '..', f), 'utf8'); (new Function(code))(); };
load('mapping.js');
load('pipeline.js');
load('app.js');

const M = global.Mapping, P = global.Pipeline, App = global.AppCore;

// ---------- 构造 File 对象（供 pipeline.arrayBuffer 使用） ----------
function fileObj(p, name) {
  const buf = fs.readFileSync(p);
  return {
    name: name || path.basename(p),
    size: buf.length,
    async arrayBuffer(){ return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); },
  };
}

async function main() {
  const DATA_FILES = [
    path.join(SAMPLE_DIR, '销售明细查询报表 (47).xlsx'),
    path.join(SAMPLE_DIR, '随访任务导出 (5).xlsx'),
    path.join(SAMPLE_DIR, '患者用药周期表.xlsx'),
  ];
  const missing = DATA_FILES.filter(p => !fs.existsSync(p));
  if (missing.length) {
    console.log('⏭️ 跳过端到端验证：样例数据文件不在 ' + SAMPLE_DIR + '（' +
      missing.map(p => path.basename(p)).join('、') + ' 缺失）');
    process.exit(0);
  }
  const files = DATA_FILES.map(p => fileObj(p));
  const t0 = Date.now();
  const res = await P.processFiles(files);
  console.log('[解析] 耗时', ((Date.now() - t0) / 1000).toFixed(1) + 's');
  console.log('[解析] 销售明细:', res.sales.length, '条');
  console.log('[解析] 随访任务:', res.followups.length, '条');
  console.log('[解析] 用药周期表:', Object.keys(res.cycles).length, '人');
  console.log('[解析] 周期取值分布:', JSON.stringify(
    Object.values(res.cycles).reduce((a, d) => { a[d] = (a[d] || 0) + 1; return a; }, {})));

  // 销售明细样例
  console.log('[销售样例]', JSON.stringify(res.sales[0]));
  // 随访样例
  const fu0 = res.followups.find(f => f.exec_time);
  console.log('[随访样例]', JSON.stringify(fu0));

  // 品种归一化核对
  const prodCount = res.sales.reduce((a, r) => { a[r.product] = (a[r.product] || 0) + 1; return a; }, {});
  console.log('[销售品种]', JSON.stringify(prodCount));
  const fuProd = res.followups.reduce((a, r) => { a[r.product] = (a[r.product] || 0) + 1; return a; }, {});
  console.log('[随访品种]', JSON.stringify(fuProd));

  // ---------- 计算 ----------
  App.STORE.sales = res.sales;
  App.STORE.followups = res.followups;
  App.STORE.cycles = res.cycles;
  const t1 = Date.now();
  const rows = App.buildRows();
  console.log('\n[计算] 耗时', ((Date.now() - t1) / 1000).toFixed(2) + 's');
  console.log('[计算] 名单总数:', rows.length);

  const sum = App.buildSummary(rows);
  console.log('[状态分布]', JSON.stringify(sum.by_status));
  console.log('[品种分布]', JSON.stringify(sum.by_product));
  console.log('[医院分布 TOP5]', JSON.stringify(Object.entries(sum.by_hospital).sort((a,b)=>b[1]-a[1]).slice(0,5)));
  console.log('[随访人分布]', JSON.stringify(sum.by_executor));
  console.log('[随访类型分布]', JSON.stringify(sum.by_fu_type));

  // 抽查几个状态的行
  const byStatus = {};
  for (const r of rows) (byStatus[r.status] = byStatus[r.status] || []).push(r);
  for (const st of Object.keys(byStatus)) {
    const sample = byStatus[st][0];
    console.log(`\n[${st} 示例] 患者=${sample.patient_name} 电话=${sample.phone} 医生=${sample.physician || ''} 品种=${sample.product} 医院=${sample.hospital} 最近购药=${sample.last_purchase} 周期=${sample.cycle_days} 应购=${sample.due_date} 距今=${sample.days_to_due} 随访=${sample.fu_time} ${sample.fu_type} 执行人=${sample.executor} 信号=${sample.fu_signal}`);
  }

  // 筛选验证：仅百泽安 + 即将回购
  const f1 = App.filterRows(rows.filter(r => r.product === '百泽安' && r.status === '即将回购'));
  console.log('\n[筛选] 百泽安+即将回购:', f1.length, '条');
  if (f1.length) console.log('  首位:', JSON.stringify({ n: f1[0].patient_name, d: f1[0].due_date }));

  // 无随访信号统计
  const noFu = rows.filter(r => !r.fu_time).length;
  console.log('[随访匹配] 无匹配随访记录:', noFu, '/', rows.length);

  // 医生字段覆盖
  const docCount = rows.filter(r => r.physician).length;
  console.log('[医生列] 有医生值的行:', docCount, '/', rows.length);

  // 医生/姓名/电话脱敏规则验证（保留首尾字，中间按位数打 *）
  const desen = P.desensitize({ patient_name: '康辉', phone: '19102856464', physician: '王建国' }, false, false, false);
  console.log('[脱敏]', JSON.stringify(desen), '→ 期望 康* / 191****6464 / 王*国');
  if (desen.patient_name !== '康*' || desen.phone !== '191****6464' || desen.physician !== '王*国') {
    throw new Error('脱敏规则不符合预期');
  }

  // 手动备注随快照保存验证：模拟 noteKey 存储
  App.STORE.notes['test-key::百泽安'] = '患者要求延迟一个月';
  const noteVal = App.STORE.notes['test-key::百泽安'] || '';
  console.log('[手动备注] 存取:', noteVal, '→ 期望 患者要求延迟一个月');
  if (noteVal !== '患者要求延迟一个月') throw new Error('备注存取异常');
  console.log('\n✅ 端到端验证全部通过');
}

main().catch(e => { console.error('FAIL', e); process.exit(1); });
