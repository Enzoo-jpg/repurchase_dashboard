// 整体小结测试（固定按「本周」统计，口径与周维度状态机一致）
//   本周 = 参考日期所在自然周（周一~周日完整周）
//   应购药 = 下次应购日（最近购药+周期）∈ 本周；实际购药 = 当周有购药记录（含提前购药）
//   未购药 = 应购日∈本周 且 当周未购；下周预计 = 下次应购日∈下周，按随访信号分正常/推迟
const fs = require('fs');
const path = require('path');
const elStore = {};
function makeEl(id) {
  return {
    id, _html: '', textContent: '', value: '', disabled: false, style: {}, dataset: {}, scrollTop: 0,
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
const App = global.AppCore, M = global.Mapping;

function assert(cond, msg) {
  if (!cond) throw new Error('❌ ' + msg);
  console.log('  ✅ ' + msg);
}

(async () => {
  // 固定参考日期 2026-08-28（周五）→ 本周 = 08-24(周一) ~ 08-30(周日)；下周 = 08-31 ~ 09-06
  App.state.refDate = '2026-08-28';
  App.state.weekSel = 'this';
  App.state.stdCycle = { 百泽安: 21 };
  App.state.selFam = '百泽安'; App.state.fuAdj = {};

  console.log('===== 准备：构造 10 名患者（周期21天，本周=08-24~08-30） =====');
  // 甲 购药08-03 → 下次应购08-24(本周一) 未购 随访延迟用药
  // 乙 购药08-03 → 下次应购08-24 未购 随访停药·效果不佳 → 脱落(效果不佳)
  // 丙 购药08-03、08-25 → 下次应购09-15(不在本周) 但 08-25 当周购药 → 实际购药（已回购）
  // 丁 购药08-03 → 应购08-24 未购 随访小结含"转渠道"
  // 戊 购药08-03 → 应购08-24 未购 随访任务状态"执行失败"
  // 己 购药08-03 → 应购08-24 未购 随访小结含"医嘱延长"
  // 庚 购药08-03 → 应购08-24 未购 随访小结含"换药"
  // 辛 购药07-30 → 下次应购08-20(<本周一) → 已逾期（不计应购/未购）
  // 壬 购药08-11 → 下次应购09-01(下周) 随访推迟(出差) → 下周预计推迟1
  // 癸 购药08-11 → 下次应购09-01(下周) 随访正常 → 下周预计正常1
  const S = [
    sales('甲', '10000000001', '2026-08-03'),
    sales('乙', '10000000002', '2026-08-03'),
    sales('丙', '10000000003', '2026-08-03'),
    sales('丙', '10000000003', '2026-08-25'),
    sales('丁', '10000000004', '2026-08-03'),
    sales('戊', '10000000005', '2026-08-03'),
    sales('己', '10000000006', '2026-08-03'),
    sales('庚', '10000000007', '2026-08-03'),
    sales('辛', '10000000008', '2026-07-30'),
    sales('壬', '10000000009', '2026-08-11'),
    sales('癸', '10000000010', '2026-08-11'),
  ];
  const F = [
    fu('甲', '10000000001', '2026-08-10', { delay_reason: '经济压力，推迟购药' }),
    fu('乙', '10000000002', '2026-08-10', { stop_reason: '效果不佳，停药' }),
    fu('丁', '10000000004', '2026-08-10', { follow_note: '患者转渠道去外地药房购药' }),
    fu('戊', '10000000005', '2026-08-10', { task_status: '执行失败', follow_note: '' }),
    fu('己', '10000000006', '2026-08-10', { follow_note: '医生建议医嘱延长用药疗程' }),
    fu('庚', '10000000007', '2026-08-10', { follow_note: '患者换药改用其他品种' }),
    fu('壬', '10000000009', '2026-08-15', { delay_reason: '出差在外，推迟购药' }),
    fu('癸', '10000000010', '2026-08-15', { follow_note: '患者规范用药' }),
  ];
  function sales(patient, phone, date, product, hosp, pharm) {
    return { source: 'sales', patient_name: patient, phone, sales_time: date, product: product || '百泽安',
      hospital: hosp || 'H', pharmacy: pharm || 'P', physician: '张' };
  }
  function fu(patient, phone, exec, extra) {
    return Object.assign({ source: 'followup', patient_name: patient, phone, exec_time: exec, plan_time: exec,
      product: '百泽安', task_status: '已完成', summary_type: '日常随访任务', executor: '李' }, extra || {});
  }
  App.STORE.sales = S; App.STORE.followups = F; App.STORE.cycles = {};
  await App.refresh(); // 触发 buildRows → 模块级 ALL_ROWS 就绪（小结原因判定依赖）

  console.log('\n===== ① 固定本周统计（口径与状态机一致） =====');
  const st = App.buildSummaryStats('百泽安');
  console.log('  本周范围:', st.start, '~', st.end);
  assert(st.start === '2026-08-24' && st.end === '2026-08-30', '本周=周一~周日完整周（与状态机一致）');
  // 应回购（总集）= 应回未回 6 + 应回已回 1 = 7
  assert(st.due === 7, `应回购 7 人（应回未回6 + 应回已回1）（实际 ${st.due}）`);
  assert(st.bought === 1, `应回已回 1 人（丙，当周购药）（实际 ${st.bought}）`);
  assert(st.notBought === 6, `应回未回 6 人（实际 ${st.notBought}）`);
  assert(st.due === st.bought + st.notBought, '应回购 = 应回已回 + 应回未回（自洽）');
  // 新分类树：推迟购药 / 延长周期(医嘱·自行) / 脱落(换药·经济·去世·其他) / 转渠道 / 随访失败
  assert(st.cnt.delay === 1 && st.cnt.prolong === 1 && st.cnt.prolong_doctor === 1 && st.cnt.prolong_self === 0 &&
    st.cnt.dropout === 2 && st.cnt.dropout_switch === 1 && st.cnt.dropout_econ === 0 && st.cnt.dropout_death === 0 && st.cnt.dropout_other === 1 &&
    st.cnt.channel === 1 && st.cnt.fuFail === 1, '原因细分：推迟1/延长周期1(医嘱1)/脱落2(换药1·其他1)/转渠道1/随访失败1');
  assert(st.dTotal === 2, '脱落合计 2 人');
  assert(st.normal === 1 && st.postpone === 1, `下周预计 正常1/推迟1（实际 正常${st.normal}/推迟${st.postpone}）`);
  assert(st.nextTotal === 2, `下周预计合计 2 人（壬/癸）（实际 ${st.nextTotal}）`);

  console.log('\n===== ② 与状态机一致性核对 =====');
  // 小结「应回未回」应等于状态机「应回购·应回未回」行数；「应回已回」=「应回购·应回已回」行数
  const rows = App.buildRows();
  const notBack = rows.filter(r => r.status === '应回购' && r.repur_part === '应回未回').length;
  const back = rows.filter(r => r.status === '应回购' && r.repur_part === '应回已回').length;
  const yingTotal = rows.filter(r => r.status === '应回购').length;
  assert(st.notBought === notBack, `应回未回 ${st.notBought} = 状态机应回未回 ${notBack}`);
  assert(st.bought === back, `应回已回 ${st.bought} = 状态机应回已回 ${back}`);
  assert(st.due === yingTotal, `应回购总集 ${st.due} = 状态机应回购 ${yingTotal}`);

  console.log('\n===== ③ 文案格式 =====');
  const { text } = App.buildSummaryText('百泽安');
  console.log(text.split('\n').map(l => '  | ' + l).join('\n'));
  assert(text.includes('本周小结（百泽安）：'), '标题按品种');
  assert(text.includes('1. 本周应回购 7 人（应回已回 1 人、应回未回 6 人）；'), '第1行数字正确');
  assert(text.includes('①推迟购药1人、②延长周期1人（医嘱1/自行0）、③脱落2人（换药1/经济0/去世0/其他1）、④转渠道1人、⑤随访失败未探寻原因1人'), '第2行原因正确');
  assert(text.includes('3. 下周预计复购 2 人，预计正常回购 1 人，推迟 1 人') && text.includes('出差'), '第3行下周预计+推迟原因');
  assert(!/\n4\. /.test(text), '小结行号连续：无跳号(1/2/3)');

  console.log('\n===== ④ 患者标注驱动统计（明细点选原因 → 小结跟随） =====');
  // 将「庚」（原自动判定=脱落·换药）在原因列覆盖为「推迟购药」→ 推迟1→2、脱落·换药1→0、脱落父类2→1
  App.STORE.reasonOverrides['10000000007::百泽安'] = 'delay';
  await App.refresh(); // 重建全量行（行 reason = 人工覆盖优先）
  const stD = App.buildSummaryStats('百泽安');
  assert(stD.cnt.delay === 2 && stD.cnt.dropout_switch === 0 && stD.cnt.dropout === 1, `覆盖庚为推迟购药 → delay2/dropout_switch0/dropout1（实际 ${stD.cnt.delay}/${stD.cnt.dropout_switch}/${stD.cnt.dropout}）`);
  const text2 = App.buildSummaryText('百泽安').text;
  assert(text2.includes('①推迟购药2人') && text2.includes('③脱落1人（换药0'), '小结文案跟随原因标注变化');
  App.STORE.reasonOverrides = {}; await App.refresh();

  console.log('\n===== ④b 小结人数人工修正（fuAdj） =====');
  // 平级项修正：推迟购药 自动1 → 人工3
  App.state.fuAdj['百泽安'] = { delay: 3 };
  const textF = App.buildSummaryText('百泽安').text;
  assert(textF.includes('①推迟购药3人'), `人工修正后推迟购药 3 人（实际: ${textF.split('\n')[2]}）`);
  // 叶子修正 → 父类=子类修正之和（换药2+经济1+去世0+其他1(乙自动) = 4）
  App.state.fuAdj['百泽安'] = { dropout_switch: 2, dropout_econ: 1 };
  const textF2 = App.buildSummaryText('百泽安').text;
  assert(textF2.includes('③脱落4人（换药2/经济1'), '父类脱落=子类修正之和（4=2+1+0+1）');
  App.state.fuAdj = {};

  console.log('\n===== ④c 下周预计人数人工修正 =====');
  // next_normal/next_postpone 修正 → 复购总数=正常+推迟
  App.state.fuAdj['百泽安'] = { next_normal: 5, next_postpone: 2 };
  const textN4 = App.buildSummaryText('百泽安').text;
  assert(textN4.includes('3. 下周预计复购 7 人，预计正常回购 5 人，推迟 2 人'), `下周预计修正生效（实际: ${textN4.split('\n')[3]}）`);
  App.state.fuAdj = {};

  console.log('\n===== ⑤ 分类维护：新增/删除分类 → 文案与标注跟随 =====');
  // 新增平级分类（默认树 5 顶项 → 新项为第⑥）
  App.addReason('新分类X');
  let textN = App.buildSummaryText('百泽安').text;
  assert(textN.includes('⑥新分类X0人'), `新增分类 → 文案出现第⑥项（实际: ${textN.split('\n')[2]}）`);
  // 给「庚」标注新分类 → 统计 1 人
  const newKey = App.state.reasonTree[App.state.reasonTree.length - 1].key;
  App.STORE.reasonOverrides['10000000007::百泽安'] = newKey;
  await App.refresh();
  const stN = App.buildSummaryStats('百泽安');
  assert(stN.cnt[newKey] === 1, `新分类标注统计 1 人（实际 ${stN.cnt[newKey]}）`);
  // 删除该分类 → 标注清理 + 文案移除
  App.removeReason(newKey); await App.refresh();
  textN = App.buildSummaryText('百泽安').text;
  assert(!textN.includes('新分类X'), '删除分类后文案移除');
  assert(App.STORE.reasonOverrides['10000000007::百泽安'] === undefined, '删除分类后患者标注清理');
  // 恢复默认树
  App.state.reasonTree = App.cloneReasonTree(App.DEFAULT_REASON_TREE); await App.refresh();
  const textR = App.buildSummaryText('百泽安').text;
  assert(textR.includes('⑤随访失败未探寻原因1人'), '恢复默认后文案还原');

  console.log('\n===== ⑥ 周期控件不影响（统计跟随周视图） =====');
  // 小结统计范围由 上周/本周/下周 视图决定；近7天/自定义遗留选项不影响
  App.state.periodType = '7d';
  const stA = App.buildSummaryStats('百泽安');
  App.state.periodType = 'custom'; App.state.periodStart = '2026-01-01'; App.state.periodEnd = '2026-12-31';
  const stB = App.buildSummaryStats('百泽安');
  assert(stA.due === 7 && stB.due === 7 && stA.bought === stB.bought, '近7天/自定义选项均不影响统计');
  App.state.periodType = '7d';

  console.log('\n===== ⑦ 小结跟随周视图（点上周/下周 → 总结对应周） =====');
  // 切「上周」（08-17~08-23）：应购日∈上周 = 辛（下次应购 08-20，08-03 购药者应购 08-24 不在上周）
  App.state.weekSel = 'last'; await App.refresh();
  const stC = App.buildSummaryStats('百泽安');
  console.log('  上周范围:', stC.start, '~', stC.end, '| 应回购:', stC.due);
  assert(stC.start === '2026-08-17' && stC.end === '2026-08-23', `切上周 → 小结统计上周（实际 ${stC.start}~${stC.end}）`);
  assert(stC.due === 1 && stC.notBought === 1, `上周应回购 1 人（辛，应购 08-20）（实际 ${stC.due}）`);
  let txt = App.buildSummaryText('百泽安').text;
  assert(txt.startsWith('上周小结（百泽安）：') && txt.includes('1. 上周应回购 1 人'), '文案标签跟随「上周」');
  // 切「下周」（08-31~09-06）：壬/癸 下次应购 09-01 ∈ 下周 → 应回未回 2
  App.state.weekSel = 'next'; await App.refresh();
  const stNext = App.buildSummaryStats('百泽安');
  assert(stNext.start === '2026-08-31' && stNext.due === 2 && stNext.notBought === 2, `下周应回购 2 人（壬/癸，应购 09-01）（实际 ${stNext.due}）`);
  txt = App.buildSummaryText('百泽安').text;
  assert(txt.startsWith('下周小结（百泽安）：') && txt.includes('1. 下周应回购 2 人'), '文案标签跟随「下周」');
  App.state.weekSel = 'this'; await App.refresh();

  console.log('\n===== ⑧ 小结跟随医院/药房筛选（需求） =====');
  // 追加 2 名患者（T1/T2 药房=P2、医院=H2），应购 08-24（本周）
  App.STORE.sales = S.concat([
    sales('双', '10000000019', '2026-08-03', '百泽安', 'H2', 'P2'),
    sales('双', '10000000020', '2026-08-03', '百泽安', 'H2', 'P2'),
  ]);
  await App.refresh();
  const stAll = App.buildSummaryStats('百泽安');
  assert(stAll.due === 9, `未筛选 应回购 9 人（原7 + P2的2）（实际 ${stAll.due}）`);
  // 筛选药房 = P → 只统计 P 药房记录（P2 的 2 人排除）
  App.state.pharmacies.add('P'); await App.refresh();
  const stP = App.buildSummaryStats('百泽安');
  assert(stP.due === 7 && stP.notBought === 6, `筛选药房=P → 应回购 7 / 应回未回 6（实际 ${stP.due}/${stP.notBought}）`);
  // 筛选药房 = P2 → 只统计 P2 的 2 人
  App.state.pharmacies.clear(); App.state.pharmacies.add('P2'); await App.refresh();
  const stP2 = App.buildSummaryStats('百泽安');
  assert(stP2.due === 2 && stP2.notBought === 2, `筛选药房=P2 → 应回购 2 / 应回未回 2（实际 ${stP2.due}/${stP2.notBought}）`);
  // 筛选医院 = H2（药房不筛）→ P2 的 2 人
  App.state.pharmacies.clear(); App.state.hospitals.add('H2'); await App.refresh();
  const stH2 = App.buildSummaryStats('百泽安');
  assert(stH2.due === 2, `筛选医院=H2 → 应回购 2（实际 ${stH2.due}）`);
  // filterRows 药房筛选
  App.state.hospitals.clear(); App.state.pharmacies.add('P2');
  const fr = App.filterRows(App.buildRows());
  assert(fr.length === 2 && fr.every(r => r.pharmacy === 'P2'), `filterRows 药房筛选 → 仅 P2 行（实际 ${fr.length}）`);
  App.state.pharmacies.clear(); App.state.hospitals.clear();
  App.STORE.sales = S; await App.refresh();

  console.log('\n===== ⑨ 小结可选呈现所选医院/药房名单（scopeNames 开关） =====');
  App.state.hospitals.add('H2'); App.state.scopeNames = true; await App.refresh();
  let tScope = App.buildSummaryText('百泽安').text;
  assert(tScope.includes('统计范围（医院：H2）'), `开启开关 → 小结列出医院名单（实际: ${tScope.split('\n').pop()}）`);
  App.state.scopeNames = false; await App.refresh();
  tScope = App.buildSummaryText('百泽安').text;
  assert(!tScope.includes('统计范围'), '关闭开关 → 小结不含范围行');
  App.state.scopeNames = true; App.state.hospitals.clear(); await App.refresh();

  console.log('\n✅ 整体小结测试全部通过');
})().catch(e => { console.error('FAIL', e); process.exit(1); });
