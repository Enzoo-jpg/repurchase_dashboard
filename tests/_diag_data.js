// 诊断：新数据文件（销售明细3 + 历史任务.xls）的表头结构与列映射命中情况
const fs = require('fs');
const path = require('path');
const os = require('os');
// 样例 Excel 目录：默认取当前用户的 Downloads，可用环境变量 SAMPLE_DIR 覆盖
const SAMPLE_DIR = process.env.SAMPLE_DIR || path.join(os.homedir(), 'Downloads');
const ROOT = path.join(__dirname, '..');
global.window = global;
global.XLSX = require(path.join(ROOT, 'vendor', 'xlsx.full.min.js'));
global.alert = () => {}; global.confirm = () => true;
const load = f => (new Function(fs.readFileSync(path.join(ROOT, f), 'utf8')))();
load('mapping.js'); load('pipeline.js');
const M = global.Mapping, P = global.Pipeline;

function dumpSheet(path_, label, maxRows = 6) {
  console.log('\n' + '='.repeat(90));
  console.log('FILE:', label, '|', path_);
  const buf = fs.readFileSync(path_);
  const wb = XLSX.read(buf, { type: 'buffer', sheetRows: maxRows });
  for (const sn of wb.SheetNames) {
    const ws = wb.Sheets[sn];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
    console.log('  SHEET:', sn, '| 前', maxRows, '行,', '列数:', rows[0] ? rows[0].length : 0);
    rows.slice(0, 3).forEach((r, i) => {
      const cells = r.map(c => String(c).slice(0, 14));
      while (cells.length && cells[cells.length - 1] === '') cells.pop();
      console.log('   R' + i + ':', cells.join(' | '));
    });
  }
}

function fileObj(p) {
  const buf = fs.readFileSync(p);
  return { name: path.basename(p), size: buf.length,
    async arrayBuffer(){ return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength); } };
}

(async () => {
  const salesPath = path.join(SAMPLE_DIR, '销售明细查询报表 (3).xlsx');
  const fuPath = path.join(SAMPLE_DIR, '历史任务.xls');

  // 1. 表头结构
  dumpSheet(salesPath, '销售明细查询报表 (3).xlsx');
  dumpSheet(fuPath, '历史任务.xls');

  // 2. 类型识别
  console.log('\n' + '='.repeat(90));
  console.log('[识别] 销售明细 →', await P.detectFileType(fileObj(salesPath)));
  console.log('[识别] 历史任务 →', await P.detectFileType(fileObj(fuPath)));

  // 3. 完整归一化统计关键字段覆盖
  console.log('\n' + '='.repeat(90));
  console.log('[列映射命中] 历史任务.xls 表头 → 字段（含未命中关键列）');
  const wb2 = XLSX.read(fs.readFileSync(fuPath), { type: 'buffer', sheetRows: 20 });
  const ws2 = wb2.Sheets[wb2.SheetNames[0]];
  const rows2 = XLSX.utils.sheet_to_json(ws2, { header: 1, defval: '' });
  const hr = P.detectHeaderRow(rows2);
  const sheets = await P.loadWorkbook(fileObj(fuPath));
  const sh = sheets.find(s => s.ttype === 'followup');
  const fieldOf = {};
  if (sh) {
    Object.keys(sh.cols).forEach(f => { const cs = sh.cols[f]; (Array.isArray(cs) ? cs : [cs]).forEach(c => { fieldOf[c] = f; }); });
    rows2[hr].forEach((h, i) => {
      const f = fieldOf[i];
      const norm = String(h).replace(/\s+/g, '');
      if (f) console.log('  命中:', String(h).slice(0, 18).padEnd(20), '→', f);
    });
    console.log('  —— 关键列是否命中：');
    ['电话号码','计划执行日期','任务摘要','药品名称','执行时间','患者反馈内容','备注','用药周期状态','是否脱落','是否确认脱落','脱落原因','未按计划持续用药原因','其他原因'].forEach(h => {
      const idx = rows2[hr].findIndex(c => String(c).replace(/\s+/g, '') === h);
      console.log('   ', h.padEnd(14), idx < 0 ? '(无此列)' : (fieldOf[idx] ? '→ ' + fieldOf[idx] : '→ ✗ 未命中'));
    });
  }

  // 4. 归一化后字段覆盖统计
  console.log('\n' + '='.repeat(90));
  if (sh) {
    const recs = P.normalizeSheet('followup', sh.rows, sh.cols, sh.source_file, sh.sheet_name);
    console.log('[随访] 归一化条数:', recs.length);
    const has = k => recs.filter(r => r[k] && String(r[k]).trim()).length;
    const keys = ['patient_name','phone','product','plan_time','exec_time','summary_type','executor','task_status','follow_note',
      'usage_status','purchased_on_time','is_dropout','dropout_reason','stop_reason','delay_reason','still_using'];
    for (const k of keys) console.log('  字段', k.padEnd(18), '有值:', has(k), '/', recs.length);
  } else {
    console.log('[随访] 未识别为 followup 类型！sheets:', sheets.map(s => s.ttype));
  }
})().catch(e => { console.error('FAIL', e); process.exit(1); });
