// 统一上传自动分类验证：detectFileType 对真实三张表的识别结果 + 速度
const fs = require('fs');
const path = require('path');
const os = require('os');
// 样例 Excel 目录：默认取当前用户的 Downloads，可用环境变量 SAMPLE_DIR 覆盖
const SAMPLE_DIR = process.env.SAMPLE_DIR || path.join(os.homedir(), 'Downloads');
function makeEl(id) {
  return { id, innerHTML: '', textContent: '', value: '', disabled: false, style: {}, dataset: {},
    classList: { add(){}, remove(){}, toggle(){}, contains(){ return false; } },
    addEventListener(){}, setAttribute(){}, getAttribute(){ return null; },
    appendChild(){}, insertBefore(){}, removeChild(){}, querySelectorAll(){ return []; },
    querySelector(){ return makeEl(id + '>child'); }, parentNode: { insertBefore(){} },
    nextElementSibling: null, onclick: null, onchange: null, oninput: null };
}
global.window = global;
global.document = {
  querySelector(sel){ return makeEl(String(sel)); }, querySelectorAll(){ return []; },
  addEventListener(){}, createElement(){ return makeEl('created'); },
  documentElement: { outerHTML: '<html></html>' }, body: { appendChild(){}, remove(){} },
};
global.URL = { createObjectURL(){ return 'blob:x'; }, revokeObjectURL(){} };
global.XLSX = require(path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js'));
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(__dirname, '..', f), 'utf8')))();
load('mapping.js'); load('pipeline.js'); load('app.js');
const P = global.Pipeline;

function fileObj(p) {
  const buf = fs.readFileSync(p);
  return { name: path.basename(p), size: buf.length,
    async arrayBuffer(){ return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); } };
}

(async () => {
  const cases = [
    [path.join(SAMPLE_DIR, '销售明细查询报表 (47).xlsx'), 'sales'],
    [path.join(SAMPLE_DIR, '随访任务导出 (5).xlsx'), 'followup'],
    [path.join(SAMPLE_DIR, '患者用药周期表.xlsx'), 'cycle'],
  ];
  const missing = cases.filter(([p]) => !fs.existsSync(p));
  if (missing.length) {
    console.log('⏭️ 跳过自动分类识别：样例数据文件不在 Downloads（' +
      missing.map(([p]) => path.basename(p)).join('、') + ' 缺失）');
    process.exit(0);
  }
  for (const [p, expect] of cases) {
    const t0 = Date.now();
    const kind = await P.detectFileType(fileObj(p));
    const ms = Date.now() - t0;
    const ok = kind === expect;
    console.log(`[识别] ${path.basename(p)} → ${kind} (期望 ${expect}) ${ok ? '✅' : '❌'} ${ms}ms`);
    if (!ok) throw new Error('识别结果不符');
  }
  console.log('\n✅ 自动分类识别全部通过（均为快速读取前20行，大文件也不慢）');
})().catch(e => { console.error('FAIL', e); process.exit(1); });
