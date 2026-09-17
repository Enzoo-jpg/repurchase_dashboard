// 三表解析与归一化 pipeline（浏览器端）
// 用 SheetJS 解析 Excel，按表类型（销售明细/随访任务/患者用药周期）归一化。
// 无 DOM 依赖（XLSX 为全局变量）。

(function () {
const M = (typeof window !== "undefined" ? window : globalThis).Mapping;
const XLSX = (typeof window !== "undefined" ? window : globalThis).XLSX;

function normHeader(h) { return M.normHeader(h); }

function asList(v) { return v == null ? [] : (Array.isArray(v) ? v : [v]); }

function pad2(n) { return String(n).padStart(2, "0"); }
function fmtDateTime(d) {
  const Y = d.getFullYear(), Mo = pad2(d.getMonth() + 1), D = pad2(d.getDate());
  const h = pad2(d.getHours()), mi = pad2(d.getMinutes()), s = pad2(d.getSeconds());
  const date = `${Y}-${Mo}-${D}`;
  return (h === "00" && mi === "00" && s === "00") ? date : `${date} ${h}:${mi}:${s}`;
}

// 单元格 → 字符串（Date 保留时分；Excel 日期序列号 → 日期字符串）
function cellStr(v) {
  if (v == null) return null;
  if (v instanceof Date) return fmtDateTime(v);
  if (typeof v === "number") {
    // Excel 序列号（1900 日期系统，基准 1899-12-30）≈ 20000~80000
    if (v > 20000 && v < 80000) {
      const ms = Math.round((v - 25569) * 86400000); // 25569 = 1970-01-01 的序列号
      return fmtDateTime(new Date(ms));
    }
    if (Number.isNaN(v)) return null;
  }
  const s = String(v).trim();
  if (s === "" || s === "nan" || s === "NaN" || s === "None") return null;
  return s;
}
function cellStrSafe(v) { return cellStr(v) == null ? "" : cellStr(v); }

// 日期字段统一取 YYYY-MM-DD
function datePart(v) {
  const s = cellStr(v);
  if (!s) return null;
  const m = s.match(/(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})/);
  if (!m) return null;
  return `${m[1]}-${pad2(+m[2])}-${pad2(+m[3])}`;
}

// 单值字段关键字映射（只取本表相关字段）
function mapColumns(tableType, rawCols) {
  const rules = M.KEYWORD_RULES;
  const assigned = {};
  const used = new Set();
  // 各表生效的字段前缀
  const FOLLOWUP_FIELDS = ["task_status", "summary_type", "executor", "follow_note", "cancel_reason",
    "create_time", "plan_time", "exec_time", "task_no", "patient_id", "is_key", "reason_label", "callback_note"];
  let prefix;
  if (tableType === "sales") prefix = f => !f.startsWith("f_") && !f.startsWith("c_");
  else if (tableType === "followup") prefix = f => f.startsWith("f_") || FOLLOWUP_FIELDS.includes(f);
  else prefix = f => f === "c_patient_name" || f === "cycle_days";

  const candList = rawCols.map((rc, idx) => [idx, normHeader(rc)]);
  for (const [field, kws] of rules) {
    if (!prefix(field)) continue;
    if (field === "exec_time") {
      // 执行时间 vs 计划执行时间：排除含「计划」的列
      const clean = candList.filter(([, nc]) => !used.has(nc) && nc.includes("执行时间") && !nc.includes("计划"));
      // 注意：candList 的 key 是 idx，这里需要重新按 idx 过滤
      const cand = candList.filter(([idx]) => !used.has(idx) && !normHeader(rawCols[idx]).includes("计划") && normHeader(rawCols[idx]).includes("执行时间"));
      let best = null, bestHits = 0;
      for (const [idx, nc] of cand) {
        let hits = 0;
        for (const kw of kws) if (nc.includes(kw)) hits++;
        if (hits > 0 && hits > bestHits) { bestHits = hits; best = idx; }
      }
      if (best !== null) { assigned[field] = best; used.add(best); }
      continue;
    }
    let best = null, bestHits = 0;
    for (const [idx, nc] of candList) {
      if (used.has(idx)) continue;
      let hits = 0;
      for (const kw of kws) if (nc.includes(kw)) hits++;
      if (hits > 0 && hits > bestHits) { bestHits = hits; best = idx; }
    }
    if (best !== null) { assigned[field] = best; used.add(best); }
  }
  // 多值字段（仅随访表）：收集所有命中列
  if (tableType === "followup") {
    for (const [field, kws] of M.KEYWORD_RULES_MULTI) {
      const cols = [];
      candList.forEach(([idx, nc]) => {
        if (!used.has(idx) && kws.some(kw => nc.includes(kw))) cols.push(idx);
      });
      if (cols.length) { assigned[field] = cols; cols.forEach(i => used.add(i)); }
    }
  }
  return assigned;
}

// 在原始数组的前 maxScan 行中找最像表头的一行
function detectHeaderRow(aoaStr, maxScan = 15) {
  if (!aoaStr || !aoaStr.length) return 0;
  const kwPools = [];
  for (const [, kws] of M.KEYWORD_RULES) kwPools.push(...kws);
  for (const [, kws] of M.KEYWORD_RULES_MULTI) kwPools.push(...kws);
  const kwLower = kwPools.map(k => k.toLowerCase());
  let bestRow = 0, bestScore = -1;
  const scan = Math.min(maxScan, aoaStr.length);
  for (let i = 0; i < scan; i++) {
    const row = aoaStr[i];
    const cells = row.map(c => String(c == null ? "" : c).trim())
      .filter(c => c !== "" && c !== "nan" && c !== "None");
    if (!cells.length) continue;
    let hit = 0;
    for (const c of cells) {
      const cl = c.toLowerCase();
      if (kwLower.some(kw => cl.includes(kw))) hit++;
    }
    const ratio = cells.length / Math.max(1, row.length);
    const score = hit * 3 + ratio * 2;
    if (score > bestScore) { bestScore = score; bestRow = i; }
  }
  return bestRow;
}

function rowToObj(cols, r) {
  const o = {};
  for (let i = 0; i < r.length; i++) o[i] = cellStr(r[i]);
  return o;
}

function _gtext(row, colmap, field) {
  const c = colmap[field];
  if (c == null) return "";
  const idxs = Array.isArray(c) ? c : [c];
  const vals = [], seen = new Set();
  for (const i of idxs) {
    const v = row[i];
    if (v && !seen.has(v)) { seen.add(v); vals.push(v); }
  }
  return vals.join("\n") || "";
}

// ---------- 销售明细归一化 ----------
function normalizeSales(row, colmap, sourceFile, sheetName, i) {
  const rec = {
    source: "sales", _row_id: `${sourceFile}::sales::${sheetName}::${i}`,
    sales_time: datePart(_gtext(row, colmap, "sales_time")) || _gtext(row, colmap, "sales_time"),
    order_status: _gtext(row, colmap, "order_status") || null,
    product_raw: _gtext(row, colmap, "product") || null,
    product: M.normalizeProduct(_gtext(row, colmap, "product")) || null,
    qty: _gtext(row, colmap, "qty") || null,
    amount: _gtext(row, colmap, "amount") || null,
    member_id: _gtext(row, colmap, "member_id") || null,
    patient_name: _gtext(row, colmap, "patient_name") || null,
    phone: _gtext(row, colmap, "phone") || null,
    hospital: _gtext(row, colmap, "hospital") || null,
    pharmacy: _gtext(row, colmap, "pharmacy") || null,
    physician: _gtext(row, colmap, "physician") || null,
    department: _gtext(row, colmap, "department") || null,
    indication: _gtext(row, colmap, "indication") || null,
    age: _gtext(row, colmap, "age") || null,
    gender: _gtext(row, colmap, "gender") || null,
  };
  return rec;
}

// ---------- 随访任务归一化 ----------
function normalizeFollowup(row, colmap, sourceFile, sheetName, i) {
  const multi = f => _gtext(row, colmap, f) || null;
  const rec = {
    source: "followup", _row_id: `${sourceFile}::followup::${sheetName}::${i}`,
    task_status: _gtext(row, colmap, "task_status") || null,
    patient_name: _gtext(row, colmap, "f_patient_name") || null,
    is_key: _gtext(row, colmap, "is_key") || null,
    patient_id: _gtext(row, colmap, "patient_id") || null,
    phone: _gtext(row, colmap, "f_phone") || null,
    plan_time: _gtext(row, colmap, "plan_time") || null,
    exec_time: _gtext(row, colmap, "exec_time") || null,
    product_raw: _gtext(row, colmap, "f_product") || null,
    product: M.normalizeProduct(_gtext(row, colmap, "f_product")) || null,
    indication: _gtext(row, colmap, "f_indication") || null,
    task_no: _gtext(row, colmap, "task_no") || null,
    summary_type: _gtext(row, colmap, "summary_type") || null,
    create_time: _gtext(row, colmap, "create_time") || null,
    executor: _gtext(row, colmap, "executor") || null,
    cancel_reason: _gtext(row, colmap, "cancel_reason") || null,
    follow_note: _gtext(row, colmap, "follow_note") || null,
    reason_label: _gtext(row, colmap, "reason_label") || null, // 回传表专用：人工选定的未购药原因分类标签
    callback_note: _gtext(row, colmap, "callback_note") || null, // 回传表专用：用户在「跟进备注」列填写的内容
    usage_status: multi("usage_status"),
    purchased_on_time: multi("purchased_on_time"),
    is_dropout: multi("is_dropout"),
    dropout_reason: multi("dropout_reason"),
    still_using: multi("still_using"),
    near_usage: multi("near_usage"),
    delay_reason: multi("delay_reason"),
    stop_reason: multi("stop_reason"),
    dosage: multi("dosage"),
  };
  return rec;
}

// ---------- 用药周期表归一化 ----------
function normalizeCycle(row, colmap, sourceFile, sheetName, i) {
  const name = (_gtext(row, colmap, "c_patient_name") || "").trim();
  const days = (_gtext(row, colmap, "cycle_days") || "").trim();
  const n = parseInt(days, 10);
  return {
    source: "cycle", _row_id: `${sourceFile}::cycle::${sheetName}::${i}`,
    patient_name: name || null,
    cycle_days: (Number.isFinite(n) && n > 0) ? n : null,
    cycle_raw: days || null,
  };
}

function normalizeSheet(tableType, rows, cols, sourceFile, sheetName) {
  const colmap = mapColumns(tableType, cols);
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    if (tableType === "sales") out.push(normalizeSales(row, colmap, sourceFile, sheetName, i));
    else if (tableType === "followup") out.push(normalizeFollowup(row, colmap, sourceFile, sheetName, i));
    else out.push(normalizeCycle(row, colmap, sourceFile, sheetName, i));
  }
  return out;
}

async function loadWorkbook(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true });
  const out = [];
  for (const name of wb.SheetNames) {
    const sheet = wb.Sheets[name];
    const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, cellDates: true });
    const aoaStr = aoa.map(r => r.map(cellStrSafe));
    const hr = detectHeaderRow(aoaStr);
    let cols, rows;
    if (hr === 0) {
      cols = aoa[0].map((c, i) => String(c == null ? "" : c).trim() || `col_${i}`);
      rows = aoa.slice(1).map(r => rowToObj(cols, r));
    } else {
      cols = aoa[hr].map((c, i) => String(c == null ? "" : c).trim() || `col_${i}`);
      rows = aoa.slice(hr + 1).map(r => rowToObj(cols, r));
    }
    // 跳过全空行
    rows = rows.filter(r => Object.values(r).some(v => v != null && String(v).trim() !== ""));
    const ttype = M.detectTableType(cols);
    out.push({ source_file: file.name, sheet_name: name, ttype, cols, rows });
  }
  return out;
}

async function processFiles(files) {
  const result = { sales: [], followups: [], cycles: {} };
  for (const f of files) {
    try {
      const sheets = await loadWorkbook(f);
      for (const sh of sheets) {
        if (sh.ttype === "unknown") {
          console.warn("无法识别表类型:", sh.source_file, sh.sheet_name, sh.cols.slice(0, 10));
          continue;
        }
        if (sh.ttype === "cycle") {
          // 周期表：合并成 {患者名: 周期天数}
          for (const rec of normalizeSheet("cycle", sh.rows, sh.cols, sh.source_file, sh.sheet_name)) {
            if (rec.patient_name && rec.cycle_days) result.cycles[rec.patient_name] = rec.cycle_days;
          }
        } else {
          const recs = normalizeSheet(sh.ttype, sh.rows, sh.cols, sh.source_file, sh.sheet_name);
          if (sh.ttype === "sales") result.sales.push(...recs);
          else result.followups.push(...recs);
        }
      }
    } catch (e) {
      console.warn("处理文件失败", f.name, e);
    }
  }
  return result;
}

// 快速识别单个 Excel 文件的表类型（销售/随访/周期）。
// 用 sheetRows 只读前 20 行做表头判定，避免全量解析大文件。
async function detectFileType(file) {
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(new Uint8Array(buf), { type: "array", cellDates: true, sheetRows: 20 });
    for (const name of wb.SheetNames) {
      const sheet = wb.Sheets[name];
      const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, cellDates: true });
      const aoaStr = aoa.map(r => r.map(cellStrSafe));
      const hr = detectHeaderRow(aoaStr);
      const cols = (aoa[hr] || []).map((c, i) => String(c == null ? "" : c).trim() || `col_${i}`);
      const ttype = M.detectTableType(cols);
      if (ttype !== "unknown") return ttype;
    }
    return "unknown";
  } catch (e) {
    console.warn("识别文件类型失败", file.name, e);
    return "unknown";
  }
}

// 电话归一化：仅数字（跨表匹配用）
function phoneDigits(v) {
  if (!v) return "";
  return String(v).replace(/\D/g, "");
}

// 脱敏姓名 / 电话 / 医生
// 姓名/医生脱敏方式 maskMode：edge=首尾保留（张*三/欧**德）、first=首字保留（张**/欧***）、all=全部隐藏（***）
// 替换几个字符就有几个 *
function maskPersonName(v, mode) {
  if (!v) return "";
  const s = String(v).trim();
  if (s.length <= 1) return "*";
  if (mode === "all") return "*".repeat(s.length);
  if (mode === "first") return s[0] + "*".repeat(s.length - 1);
  if (s.length === 2) return s[0] + "*";
  return s[0] + "*".repeat(s.length - 2) + s[s.length - 1];
}
function desensitize(rec, namePlain = false, phonePlain = false, doctorPlain = false, maskMode = "edge") {
  const out = Object.assign({}, rec);
  const name = out.patient_name;
  if (name && !namePlain) {
    out.patient_name = maskPersonName(name, maskMode);
  }
  const phone = out.phone;
  if (phone && !phonePlain) {
    const s = String(phone);
    const digits = s.replace(/\D/g, "");
    if (digits.length >= 7) out.phone = digits.slice(0, 3) + "****" + digits.slice(-4);
    else out.phone = s.replace(/\d/g, "*");
  }
  const doctor = out.physician;
  if (doctor && !doctorPlain) {
    out.physician = maskPersonName(doctor, maskMode);
  }
  return out;
}

if (typeof window !== "undefined") {
  window.Pipeline = { cellStr, datePart, mapColumns, detectHeaderRow, normalizeSheet,
    loadWorkbook, processFiles, detectFileType, desensitize, phoneDigits, fmtDateTime };
}
})();
