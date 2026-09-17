// 随访信号判定测试：覆盖历史任务.xls 新版措辞的值（用药周期状态/是否脱落/脱落原因/未按计划原因）
const fs = require('fs');
const path = require('path');
global.window = global;
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js');
const M = global.Mapping;

function assert(cond, msg) {
  if (!cond) throw new Error('❌ ' + msg);
  console.log('  ✅ ' + msg);
}
// 直接构造已归一化的随访记录（模拟历史任务.xls 的字段值）
function fu(over) {
  return Object.assign({
    source: 'followup', patient_name: '测试', phone: '13800000000', product: '某药',
    exec_time: '2026-08-01', task_status: '已完成', summary_type: '日常随访',
  }, over || {});
}

console.log('===== 用药周期状态（历史任务表） =====');
let s = M.followupSignal(fu({ purchased_on_time: '按计划持续用药' }));
assert(s.signal === 'normal' || s.signal === 'unknown', '按计划持续用药 → 正常（实际 ' + s.signal + '）');

s = M.followupSignal(fu({ purchased_on_time: '停药----脱落' }));
assert(s.signal === 'dropout', '停药----脱落 → dropout（实际 ' + s.signal + '）');

s = M.followupSignal(fu({ purchased_on_time: '推迟购药' }));
assert(s.signal === 'nonstd', '推迟购药 → nonstd（实际 ' + s.signal + '）');

s = M.followupSignal(fu({ purchased_on_time: '未按计划持续用药-不依从' }));
assert(s.signal === 'nonstd', '未按计划持续用药-不依从 → nonstd（实际 ' + s.signal + '）');

console.log('\n===== 是否脱落 / 脱落原因 =====');
s = M.followupSignal(fu({ is_dropout: '是', dropout_reason: '疗效不佳自行停药' }));
assert(s.signal === 'dropout', '是否脱落=是 + 停药原因 → dropout（实际 ' + s.signal + '）');

s = M.followupSignal(fu({ is_dropout: '否' }));
assert(s.signal === 'normal' || s.signal === 'unknown', '是否脱落=否 → 正常（实际 ' + s.signal + '）');

s = M.followupSignal(fu({ dropout_reason: '患者转渠道去其他药店' }));
assert(s.signal === 'dropout', '脱落原因含转渠道 → dropout（实际 ' + s.signal + '）');

console.log('\n===== 未按计划持续用药原因 → 延迟 =====');
s = M.followupSignal(fu({ delay_reason: '经济原因未按时购药' }));
assert(s.signal === 'nonstd', '未按计划原因 → nonstd（实际 ' + s.signal + '）');

console.log('\n===== 患者反馈（follow_note 措辞） =====');
s = M.followupSignal(fu({ follow_note: '患者自述效果不佳，考虑停药' }));
assert(s.signal === 'dropout', '反馈含停药 → dropout（实际 ' + s.signal + '）');

s = M.followupSignal(fu({ follow_note: '患者经济负担重，推迟购药' }));
assert(s.signal === 'nonstd', '反馈含推迟/经济 → nonstd（实际 ' + s.signal + '）');

console.log('\n✅ 随访信号判定（新版措辞）—— 全部通过');
