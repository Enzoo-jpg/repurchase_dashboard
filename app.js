// 预计购药分析看板：前端编排层
// 全部计算在浏览器内完成：三表解析(pipeline.js) → 归一化 → 合并计算 → 渲染。
// 无后端、无网络请求；数据从不离开本机。

(function () {
const M = window.Mapping;
const P = window.Pipeline;
const CATEGORIES = M.CATEGORIES;    // 应回购（=应回未回+应回已回）/ 未到期 / 已逾期
const REPUR_PARTS = M.REPUR_PARTS;  // 应回购细分：应回未回 / 应回已回

const CAT_COLOR = {
  "应回购": "#e03131",   // 所选周应购清单（应回未回 + 应回已回）
  "未到期": "#3b5bdb",   // 应购日尚未到所选周
  "已逾期": "#f08c00",   // 应购日已过所选周
};
const REPUR_PART_COLOR = {
  "应回未回": "#e03131", // 应购日∈所选周但未购药，需跟进
  "应回已回": "#0f9d6b", // 所选周内已购药（含提前购药）
};

// ---- 未购药原因分类树（可维护：增删改 / 加子类 / 恢复默认）----
// 层级结构：父类人数 = 其子类人数之和；默认「脱落」为父类 + 5 个子类
// 默认结构：推迟购药 / 延长周期（医嘱·自行）/ 脱落（换药·经济·去世·其他）/ 转渠道 / 随访失败未探寻原因
const DEFAULT_REASON_TREE = [
  { key: "delay", label: "推迟购药", children: [] },
  { key: "prolong", label: "延长周期", children: [
    { key: "prolong_doctor", label: "医嘱" },
    { key: "prolong_self", label: "自行" },
  ]},
  { key: "dropout", label: "脱落", children: [
    { key: "dropout_switch", label: "换药" },
    { key: "dropout_econ", label: "经济" },
    { key: "dropout_death", label: "去世" },
    { key: "dropout_other", label: "其他" },
  ]},
  { key: "channel", label: "转渠道", children: [] },
  { key: "fuFail", label: "随访失败未探寻原因", children: [] },
];
function cloneReasonTree(t) { return (t || []).map(n => ({ key: n.key, label: n.label, children: (n.children || []).map(c => ({ key: c.key, label: c.label, children: [] })) })); }
// 展平：-> [{ key, label, parent, path }]（path 用于展示，如「脱落·效果不佳」）
function flattenReasonTree(t) {
  const out = [];
  (t || []).forEach(n => {
    out.push({ key: n.key, label: n.label, parent: "", path: n.label });
    (n.children || []).forEach(c => out.push({ key: c.key, label: c.label, parent: n.key, path: n.label + "·" + c.label }));
  });
  return out;
}
function reasonLabel(tree, key) {
  const f = flattenReasonTree(tree).find(x => x.key === key);
  return f ? f.path : "";
}

// 名单表列（note = 用户手动填写的跟进备注列，随快照保存）
const LIST_COLS = [
  ["patient_name", "患者"],
  ["phone", "电话"],
  ["physician", "医生"],
  ["department", "科室"],
  ["product", "品种"],
  ["hospital", "医院"],
  ["pharmacy", "药房"],
  ["last_purchase", "最近购药"],
  ["qty", "购药数量"],
  ["cycle_days", "周期(天)"],
  ["due_date", "应购药日期"],
  ["days_to_due", "距今天数"],
  ["fu_time", "匹配随访时间"],
  ["fu_type", "随访类型"],
  ["executor", "随访人"],
  ["status", "状态"],
  ["reason", "未购药原因"],
  ["fu_signal", "随访信号"],
  ["note", "跟进备注"],
];

const STORE = { sales: [], followups: [], cycles: {}, files: [], seq: 0, notes: {}, reasonOverrides: {} };
let SNAP_MODE = false;
let SNAP_BASE = null; // 快照全量行基准（不可变，供快照模式筛选；避免 DATA.rows 被覆盖后累积丢失）

// 界面状态
const state = {
  cats: new Set(), reasons: new Set(), repurParts: new Set(), products: new Set(), hospitals: new Set(), pharmacies: new Set(), executors: new Set(),
  q: "",
  weekSel: "this", // 周视图：last=上周 / this=本周 / next=下周（单选，默认本周）
  refDate: "",     // 参考日期（空=今天）；「本周」= 该日期所在周（起止由 weekEnd 决定），可改以回看历史周
  weekEnd: 0,      // 周截止日：0=周日（自然周，周一~周日，默认）1=周一…6=周六（选周五即「上周六~本周五」）。前后周按同一规则平移
  stdCycle: {},   // 标准周期仅按「上传数据中出现品种」动态生成，不写死内置商品
  cycleAlgo: "reset", // 周期算法：reset=重置式（默认，最近一次购药 + 覆盖）/ accum=累计式 / accum_tol=累计+容差
  cycleTol: 7,        // 容差天数（仅 accum_tol 生效）：剩余库存 ≤ 容差视为小幅提前，按重置处理
  qtyScale: {},       // 每个品种是否「按数量计算」：true=覆盖天数=数量×周期（默认关闭，数量按 1 计）
  plainName: false, plainPhone: false, plainDoctor: false,  // false=脱敏显示（姓名/医生各自独立开关；脱敏掩码方式共用「姓名/医生脱敏方式」）
  maskMode: "edge",   // 姓名/医生脱敏方式：first=保留首字、edge=保留首+末字、id=仅会员号、all=全部隐藏
  hiddenCols: new Set(),
  page: 1, pageSize: 50,
  // 整体小结：固定统计「本周」+ 选中品种 + 未购药原因分类树（可维护）
  periodType: "7d", periodStart: "", periodEnd: "",
  selFam: "",
  scopeNames: true, // 小结中是否呈现所选医院/药房名单（可开关）
  reasonTree: cloneReasonTree(DEFAULT_REASON_TREE), // 用户维护后的分类树（随快照保存）
  fuAdj: {},   // { [fam]: { [分类key]: n } } 小结人数人工修正（仅存用户改过的值；key=分类树叶子 key）
};
let SUMMARY = { text: "" }; // 当前整体小结文本（快照保存用）
let ALL_ROWS = [];          // 全量计算行（小结统计用，不随筛选变化）
let DATA = { rows: [] };
let CURRENT = { summary: null };

const $ = s => document.querySelector(s);

/* ============ 工具 ============ */
function esc(s) { return String(s == null ? "" : s).replace(/[&<>]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c])); }
function fmtSize(n) {
  if (n < 1024) return n + " B";
  if (n < 1024 * 1024) return (n / 1024).toFixed(1) + " KB";
  return (n / 1024 / 1024).toFixed(2) + " MB";
}
function showLoading(txt) { $("#loadingTxt").textContent = txt || "正在处理…"; $("#loading").classList.remove("hidden"); }
function hideLoading() { $("#loading").classList.add("hidden"); }
function download(blob, name) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(a.href);
}
function todayStr() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function addDays(dateStr, days) {
  const d = new Date(dateStr + "T00:00:00");
  d.setDate(d.getDate() + days);
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}
function diffDays(a, b) { // a - b（天）
  return Math.round((new Date(a + "T00:00:00") - new Date(b + "T00:00:00")) / 86400000);
}
function phoneDigits(v) { return P.phoneDigits(v); }

// 脱敏规则（展示层按 state 开关实时切换；导出/快照与所选脱敏方式保持一致）
// maskMode（state.maskMode）：first=保留首字（张**/欧***）、edge=保留首+末字（张*三/欧**德）、id=仅会员号（患者列显示会员号）、all=全部隐藏（***）
// 替换几个字符就有几个 *
function maskPersonName(v, mode) {
  if (!v) return "";
  const s = String(v).trim();
  if (s.length <= 1) return "*";
  if (mode === "all") return "*".repeat(s.length);
  if (mode === "first") return s[0] + "*".repeat(s.length - 1);
  // edge（默认）：保留首尾
  if (s.length === 2) return s[0] + "*";
  return s[0] + "*".repeat(s.length - 2) + s[s.length - 1];
}
function maskName(v) { return maskPersonName(v, state.maskMode === "id" ? "edge" : state.maskMode); }
function maskPhone(v) {
  if (!v) return "";
  const d = String(v).replace(/\D/g, "");
  return d.length >= 7 ? d.slice(0, 3) + "****" + d.slice(-4) : String(v).replace(/\d/g, "*");
}
// 医生与姓名共用同一脱敏方式（保持一致）
function maskDoctor(v) { return maskPersonName(v, state.maskMode); }
// 按当前脱敏开关取展示值（mode=id「仅会员号」：患者列=会员号、医生列无会员身份显示 —）
function disp(r, key) {
  const idOnly = state.maskMode === "id";
  if (key === "patient_name") {
    if (state.plainName) return r.patient_name || "";
    if (idOnly) return r.member_id ? r.member_id : "—";
    return maskName(r.patient_name);
  }
  if (key === "phone") return state.plainPhone ? (r.phone || "") : maskPhone(r.phone);
  // 医生：独立开关（state.plainDoctor）；脱敏时的方式与姓名共用（state.maskMode）
  if (key === "physician") {
    if (state.plainDoctor) return r.physician || "";
    if (idOnly) return "—";
    return maskDoctor(r.physician);
  }
  return r[key] == null ? "" : r[key];
}
// 手动备注的存储键（患者 × 品种）
function noteKey(r) { return (r._key || "") + "::" + r.product; }

/* ============ 周视图（状态判定的时间基准） ============ */
// 参考日期：默认今天，可手动改（用于回看任意历史周的应购/购药情况）
function refToday() { return state.refDate || todayStr(); }
// 所选周范围：以参考日期所在「周」（起止由 state.weekEnd 决定）为「本周」，前后推 上周/下周
// weekEnd = 周截止日：0=周日（自然周 周一~周日）… 6=周六（即 上周六~本周五）
// 本周 = [截止日-6, 截止日]，其中截止日取「参考日当天或之后最近的该星期几」（保证参考日落在本周内）
function getWeekRange(sel) {
  const base = refToday();
  const endDow = (state.weekEnd == null ? 0 : Number(state.weekEnd)) % 7; // 0=周日
  const dow = new Date(base + "T00:00:00").getDay(); // 0=周日
  let end = addDays(base, (endDow - dow + 7) % 7);   // 落在参考日当天或之后最近的「截止日」
  if (sel === "last") end = addDays(end, -7);
  else if (sel === "next") end = addDays(end, 7);
  return { start: addDays(end, -6), end };
}
const WEEK_LABEL = { last: "上周", this: "本周", next: "下周" };
const WEEKDAY_CN = ["周日", "周一", "周二", "周三", "周四", "周五", "周六"]; // 索引=Date.getDay()

/* ============ 购药数量 × 用药周期 → 下次应购日 ============ */
// 患者一次可购买多个周期的药品，故「本次购药覆盖天数 = 购药数量 × 单次用药周期」。
// 数量是否参与计算按品种开关（state.qtyScale，默认关闭）；容差与算法见 state.cycleAlgo：
//   reset     = 重置式：每次购药都从本次重新起算
//   accum     = 累计式：只要还有剩余库存就往后顺延
//   accum_tol = 累计 + 容差：剩余库存 > 容差N 视为囤药（顺延），≤ N 视为小幅提前（重置）
// 「剩余库存 = 上次库存耗尽日 - 本次购药日」，故药已吃完才买时自动走重置分支。

// 购药数量解析：取数字部分（容错 "2盒"/"1.5"/"1,200"），缺失、非数字或非正数按 1
// （= 不放大，保持历史行为）。负数按 1，避免 "-2" 被正则截成 2。
function qtyNum(v) {
  if (v == null) return 1;
  const m = String(v).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);
  if (!m) return 1;
  const n = parseFloat(m[0]);
  return Number.isFinite(n) && n > 0 ? n : 1;
}
// 该品种的数量是否参与计算（默认关闭 → 数量按 1 计，行为与历史一致）
function qtyScales(fam) { return !!state.qtyScale[fam]; }
// 单次购药的覆盖天数
function coverDays(qty, cycle, fam) {
  return (qtyScales(fam) ? qtyNum(qty) : 1) * cycle;
}
// 购药记录规整：按日期升序，同一天的多行数量相加（3 行各买 1 支 = 一次买了 3 支）
function normalizePurchases(list) {
  const out = [];
  for (const p of (list || []).slice().sort((a, b) => a.date.localeCompare(b.date))) {
    const last = out[out.length - 1];
    if (last && last.date === p.date) last.qty += p.qty;
    else out.push({ date: p.date, qty: p.qty });
  }
  return out;
}
// 按所选算法逐条推进「库存耗尽日」，返回最后一条购药的库存耗尽日（= 下次应购日）
function stockEndAfter(purchases, cycle, fam) {
  const algo = state.cycleAlgo || "reset";
  const tol = Number.isFinite(Number(state.cycleTol)) ? Number(state.cycleTol) : 7;
  let end = null;
  for (const p of purchases) {
    const cover = coverDays(p.qty, cycle, fam);
    if (end == null) { end = addDays(p.date, cover); continue; }
    let base;
    if (algo === "reset") base = p.date;                              // 重置：本次重新起算
    else if (algo === "accum") base = diffDays(p.date, end) > 0 ? p.date : end; // 累计：取较晚者（逾期自动重置）
    else base = diffDays(end, p.date) <= tol ? p.date : end;          // 累计+容差
    end = addDays(base, cover);
  }
  return end;
}

/* ============ 合并计算（核心） ============ */
// 1) 按患者（电话/姓名）聚合销售记录 → 每患者每品种的购药时间序列（含购药数量）
// 2) 应购药日期 = 按「周期算法」推进库存耗尽日（数量是否放大见「按数量计算」品种开关）
// 3) 匹配随访：应购药日前最近一次（同品种优先；无则其他品种；再无则之后最早一条）
// 4) 状态判定（周维度，购药事实优先）：
//    当周有购药记录 → 应回购·应回已回（含提前购药）
//    应购日 < 周首 → 已逾期；应购日 ∈ 所选周 → 应回购·应回未回；应购日 > 周末 → 未到期
//   所有患者按下钻标注：正常状态 / 预判延期 / 已脱落（依据随访）
function buildRows() {
  const sales = STORE.sales;
  const followups = STORE.followups;
  const cycles = STORE.cycles;

  // ---- 销售按患者聚合 ----
  const patMap = {}; // key: phoneDigits || name
  const keyOf = r => phoneDigits(r.phone) || (r.patient_name || "未知");
  for (const s of sales) {
    if (!s.sales_time || !s.product) continue;
    const k = keyOf(s);
    const p = (patMap[k] = patMap[k] || { name: s.patient_name || "未知", phone: s.phone || "", byProduct: {} });
    if (s.patient_name) p.name = s.patient_name;
    if (s.phone) p.phone = s.phone;
    const fam = s.product;
    const bp = (p.byProduct[fam] = p.byProduct[fam] || { purchases: [], hospital: s.hospital || "", pharmacy: s.pharmacy || "", physician: s.physician || "", member_id: s.member_id || "", department: s.department || "" });
    bp.purchases.push({ date: s.sales_time, qty: qtyNum(s.qty) });
    if (s.hospital) bp.hospital = s.hospital;
    if (s.pharmacy) bp.pharmacy = s.pharmacy;
    if (s.member_id) bp.member_id = s.member_id; // 会员号取最近购药记录的（脱敏方式「仅会员号」用）
    if (s.department) bp.department = s.department; // 科室取最近购药记录的
    if (s.physician) bp.physician = s.physician;
  }

  // ---- 随访按患者索引 ----
  const fuByKey = {}; // key -> [fu]
  const fuByName = {};
  for (const fu of followups) {
    const k = phoneDigits(fu.phone);
    const nm = (fu.patient_name || "").trim();
    if (k) (fuByKey[k] = fuByKey[k] || []).push(fu);
    if (nm) (fuByName[nm] = fuByName[nm] || []).push(fu);
  }
  const fuIndex = (name, phone) => {
    const k = phoneDigits(phone);
    const list = (k && fuByKey[k]) || fuByName[name] || [];
    return list;
  };
  const fuDate = fu => P.datePart(fu.exec_time) || P.datePart(fu.plan_time) || null;

  // ---- 逐患者×品种计算 ----
  const rows = [];
  const today = todayStr();
  for (const k in patMap) {
    const p = patMap[k];
    const fus = fuIndex(p.name, p.phone);
    for (const fam in p.byProduct) {
      const bp = p.byProduct[fam];
      if (!bp.purchases.length) continue;
      bp.purchases = normalizePurchases(bp.purchases);
      const lastRec = bp.purchases[bp.purchases.length - 1];
      const lastPurchase = lastRec.date;
      const cycle = cycles[p.name] || state.stdCycle[fam] || 30;
      // 周视图状态判定基准：所选周 W（起止由「周截止日」决定）
      //  当周 [W.start, W.end] 内有购药记录 → 应回购·应回已回（购药事实最优先）
      //  应购日 < W.start → 已逾期；应购日 ∈ W → 应回购；应购日 > W.end → 未到期
      // 应购药日期列显示「预判应购日」= 按周期算法推进到「倒数第二次购药」的库存耗尽日（≥2条记录时），
      // 用于对比实际购药是提前还是延后（应回已回行不呈现下次应购时间）
      const W = getWeekRange(state.weekSel);
      const purchasedInWeek = bp.purchases.some(x => x.date >= W.start && x.date <= W.end);
      let expectedDue = null, dueOffset = null;
      if (bp.purchases.length >= 2) {
        expectedDue = stockEndAfter(bp.purchases.slice(0, -1), cycle, fam);
        dueOffset = diffDays(lastPurchase, expectedDue); // 正=延期，负=提前，0=按时
      }
      const dueDate = stockEndAfter(bp.purchases, cycle, fam);

      // 匹配随访：优先应购药日前最近一次（同品种；优先「已完成/有执行时间」的真实随访）
      let matched = null;
      if (fus.length) {
        const sameFam = fus.filter(f => f.product === fam);
        let pool = sameFam.length ? sameFam : fus;
        // 优先已执行的随访（有执行时间或任务已完成），避免未执行任务污染匹配
        const executed = pool.filter(f => f.exec_time || (f.task_status && f.task_status.includes("已完成")));
        if (executed.length) pool = executed;
        const withDate = pool.map(f => ({ f, d: fuDate(f) })).filter(x => x.d);
        if (withDate.length) {
          const before = withDate.filter(x => x.d <= dueDate);
          if (before.length) {
            before.sort((a, b) => b.d.localeCompare(a.d)); // 最近一次在前
            matched = before[0].f;
          } else {
            withDate.sort((a, b) => a.d.localeCompare(b.d)); // 之后最早
            matched = withDate[0].f;
          }
        }
      }

      // 状态判定（周维度，购药事实优先）：
      //  当周有购药记录 → 应回购·应回已回（最优先：本周该买也买了 / 提前买都算）
      //  应购日 < 周首 → 已逾期；应购日 ∈ 所选周 → 应回购·应回未回；应购日 > 周末 → 未到期
      // 未购药原因：自动判定（随访文本 classifyFuReason）预填，用户可在「未购药原因」列点选覆盖
      const sig = matched ? M.followupSignal(matched) : { signal: "unknown", reason: "" };
      const days = diffDays(dueDate, today); // 正=未来，负=逾期（相对真实今天，供列上标注）
      const subKey = k + "::" + fam;
      // 自动判定原因 key（dropout 细分 dropout_switch/econ/death/other；prolong 细分 prolong_doctor/self；仅保留分类树中存在的分类）
      const treeKeys = flattenReasonTree(state.reasonTree).map(f => f.key);
      let autoReason = "";
      if (matched) {
        const c = M.classifyFuReason(matched);
        if (c) {
          const k = c.key === "dropout"
            ? ("dropout_" + (c.detail === "换药" ? "switch" : c.detail === "经济" ? "econ" : c.detail === "去世" ? "death" : "other"))
            : c.key === "prolong"
              ? ("prolong_" + (c.detail === "医嘱" ? "doctor" : "self"))
              : c.key;
          if (treeKeys.includes(k)) autoReason = k;
        }
      }
      const reason = STORE.reasonOverrides[subKey] || autoReason; // 用户点选覆盖优先
      let category, repurPart = "";
      if (purchasedInWeek) {
        category = "应回购"; repurPart = "应回已回";
      } else if (dueDate < W.start) {
        category = "已逾期";
      } else if (dueDate <= W.end) {
        category = "应回购"; repurPart = "应回未回";
      } else {
        category = "未到期";
      }

      const isRepur = repurPart === "应回已回";
      // 应购药日期列展示值（应回已回行=预判应购日）；是否落在所选周 → 列表置顶 + 特殊底色
      const rowDue = isRepur ? (expectedDue || "") : dueDate;
      const dueInWeek = !!rowDue && rowDue >= W.start && rowDue <= W.end;
      rows.push({
        _key: k, product: fam,
        patient_name: p.name, phone: p.phone, physician: bp.physician,
        hospital: bp.hospital, pharmacy: bp.pharmacy,
        member_id: bp.member_id || "", // 会员号（脱敏方式「仅会员号」时患者列显示）
        department: bp.department || "", // 科室（销售底表「处方科室」）
        last_purchase: lastPurchase, qty: lastRec.qty, cycle_days: cycle,
        // 应回已回行：due_date 显示「预判应购日」（不呈现下次应购时间）；无预判基准显示空（渲染为 —）
        due_date: rowDue,
        due_in_week: dueInWeek, // 应购药日期落在所选周 → 优先展示并标底色
        expected_due: expectedDue || "",
        due_offset: isRepur ? dueOffset : null,   // 应回已回行：正=延期/负=提前/0=按时
        purchase_days_ago: isRepur ? Math.max(0, diffDays(today, lastPurchase)) : null, // 最近购药距今
        days_to_due: days,
        repurchased: purchasedInWeek,
        repur_part: repurPart, // 应回购细分：应回未回 / 应回已回（其他主状态为空）
        fu_time: matched ? (P.datePart(matched.exec_time) || P.datePart(matched.plan_time) || "") : "",
        fu_type: matched ? (matched.summary_type || "") : "",
        executor: matched ? (matched.executor || "") : "",
        fu_note: matched ? (matched.follow_note || "") : "",
        fu_signal: sig.signal === "unknown" ? "" : sig.reason,
        fu_signal_kind: sig.signal, // 信号类别：normal / nonstd / dropout / unknown（用于着色）
        reason: reason, // 未购药原因 key（自动判定预填，用户可覆盖；空=未标）
        status: category,
        _matched: matched,
      });
    }
  }

  // 排序：应购药日期落在「所选周」的患者置顶（周内应购优先），组内外均按应购药日期升序；
  // 应回已回行按预判应购日参与；空值排最后。
  rows.sort((a, b) => {
    if (!!a.due_in_week !== !!b.due_in_week) return a.due_in_week ? -1 : 1;
    const da = a.due_date || "9999-12-31", db = b.due_date || "9999-12-31";
    return da.localeCompare(db);
  });
  return rows;
}

/* ============ 筛选 ============ */
// opts.skipCat=true 时跳过状态/下钻/应回购细分筛选（用于计算「状态计数基数」：
// 状态卡片计数跟随 品种/医院/随访人/关键词 筛选，但不受状态本身筛选影响）
function filterRows(rows, opts) {
  let rs = rows;
  if (!(opts && opts.skipCat)) {
    if (state.cats.size) rs = rs.filter(r => state.cats.has(r.status));
    if (state.reasons.size) rs = rs.filter(r => r.reason && state.reasons.has(r.reason));
    if (state.repurParts.size) rs = rs.filter(r => r.repur_part && state.repurParts.has(r.repur_part));
  }
  if (state.products.size) rs = rs.filter(r => state.products.has(r.product));
  if (state.hospitals.size) rs = rs.filter(r => state.hospitals.has(r.hospital || "未知"));
  if (state.pharmacies.size) rs = rs.filter(r => state.pharmacies.has(r.pharmacy || "未知"));
  if (state.executors.size) rs = rs.filter(r => state.executors.has(r.executor || "未知"));
  if (state.q) {
    const ql = state.q.trim().toLowerCase();
    rs = rs.filter(r => {
      for (const f of ["patient_name", "phone", "physician", "hospital", "pharmacy", "product", "executor"]) {
        if (r[f] && String(r[f]).toLowerCase().includes(ql)) return true;
      }
      return false;
    });
  }
  return rs;
}

function buildSummary(rows) {
  const by_status = {}, by_reason = {}, by_repur_part = {}, by_product = {}, by_hospital = {}, by_pharmacy = {}, by_executor = {}, by_fu_type = {};
  for (const r of rows) {
    by_status[r.status] = (by_status[r.status] || 0) + 1;
    if (r.reason) by_reason[r.reason] = (by_reason[r.reason] || 0) + 1;
    if (r.repur_part) by_repur_part[r.repur_part] = (by_repur_part[r.repur_part] || 0) + 1;
    by_product[r.product || "未知"] = (by_product[r.product || "未知"] || 0) + 1;
    by_hospital[r.hospital || "未知"] = (by_hospital[r.hospital || "未知"] || 0) + 1;
    by_pharmacy[r.pharmacy || "未知"] = (by_pharmacy[r.pharmacy || "未知"] || 0) + 1;
    by_executor[r.executor || "未知"] = (by_executor[r.executor || "未知"] || 0) + 1;
    by_fu_type[r.fu_type || "无随访"] = (by_fu_type[r.fu_type || "无随访"] || 0) + 1;
  }
  return { total: rows.length, by_status, by_reason, by_repur_part, by_product, by_hospital, by_pharmacy, by_executor, by_fu_type };
}

/* ============ 渲染 ============ */
function renderSummary(d) {
  let html = `<div class="total"><div class="n">${d.total}</div><div class="l">名单总数</div></div>`;
  for (const s of CATEGORIES) {
    const n = (d.by_status && d.by_status[s]) || 0;
    const active = state.cats.has(s) ? "active" : "";
    const color = CAT_COLOR[s] || "#868e96";
    html += `<div class="card ${active}" data-status="${s}" title="点击筛选「${s}」，再次点击取消" style="border-left:4px solid ${color}">
      <span class="dot" style="background:${color}"></span><div class="n" style="color:${color}">${n}</div><div class="l">${s}</div><span class="cue">点选</span></div>`;
  }
  $("#summary").innerHTML = html;
  document.querySelectorAll("#summary .card").forEach(c => {
    c.onclick = () => {
      const s = c.dataset.status;
      if (state.cats.has(s)) state.cats.delete(s); else state.cats.add(s);
      state.page = 1;
      refresh();
    };
  });
  // 应回购细分（应回未回 / 应回已回）chips 点选筛选
  const repurBar = $("#repurPartBar");
  const rp = d.by_repur_part || {};
  if (Object.keys(rp).length) {
    repurBar.classList.remove("hidden");
    repurBar.innerHTML = '<span class="lbl">应回购细分</span>' + REPUR_PARTS.map(k => {
      const n = rp[k] || 0;
      if (!n) return "";
      const act = state.repurParts.has(k) ? "active" : "";
      return `<span class="chip ${act}" data-rp="${k}" style="border-color:${REPUR_PART_COLOR[k]}">${k} (${n})</span>`;
    }).join("");
    repurBar.querySelectorAll(".chip").forEach(c => c.onclick = () => {
      const k = c.dataset.rp;
      if (state.repurParts.has(k)) state.repurParts.delete(k); else state.repurParts.add(k);
      state.page = 1;
      refresh();
    });
  } else {
    repurBar.classList.add("hidden");
  }
  // 未购药原因（chips 点选筛选；选项=用户维护的分类树叶子项）
  const reasonBar = $("#reasonBar");
  const rc = d.by_reason || {};
  if (Object.keys(rc).length) {
    const flat = flattenReasonTree(state.reasonTree);
    reasonBar.classList.remove("hidden");
    reasonBar.innerHTML = '<span class="lbl">未购药原因</span>' + flat.map(f => {
      const n = rc[f.key] || 0;
      if (!n) return "";
      const act = state.reasons.has(f.key) ? "active" : "";
      const color = f.parent ? "#868e96" : "#e8590c";
      return `<span class="chip ${act}" data-reason="${f.key}" style="border-color:${color}">${esc(f.path)} (${n})</span>`;
    }).join("");
    reasonBar.querySelectorAll(".chip").forEach(c => c.onclick = () => {
      const k = c.dataset.reason;
      if (state.reasons.has(k)) state.reasons.delete(k); else state.reasons.add(k);
      state.page = 1;
      refresh();
    });
  } else {
    reasonBar.classList.add("hidden");
  }
  buildMs("prodMs", d.by_product, state.products, "品种");
  buildMs("hospMs", d.by_hospital, state.hospitals, "医院");
  buildMs("pharmMs", d.by_pharmacy, state.pharmacies, "药房");
  buildMs("execMs", d.by_executor, state.executors, "随访人");
  updateFilterInfo();
}

function buildMs(msId, data, stateSet, label) {
  const entries = Object.entries(data || {}).sort((a, b) => b[1] - a[1]);
  const list = $("#" + msId + "List");
  list.innerHTML = entries.map(([k, v]) =>
    `<label><input type="checkbox" value="${esc(k)}" ${stateSet.has(k) ? "checked" : ""}><span>${esc(k)}</span><span class="ms-cnt">${v}</span></label>`).join("");
  list.querySelectorAll("input").forEach(cb => {
    cb.onchange = () => { if (cb.checked) stateSet.add(cb.value); else stateSet.delete(cb.value); state.page = 1; refresh(); };
  });
  const panel = $("#" + msId + "Panel");
  let search = panel.querySelector(".ms-search");
  if (!search) {
    search = document.createElement("input");
    search.className = "ms-search";
    search.placeholder = "输入关键字检索…";
    list.parentNode.insertBefore(search, list);
  }
  const applyFilter = () => {
    const q = search.value.trim().toLowerCase();
    list.querySelectorAll("label").forEach(lb => {
      const k = lb.querySelector("input").value;
      const hit = !q || k.toLowerCase().includes(q) || stateSet.has(k);
      lb.style.display = hit ? "" : "none";
    });
  };
  search.oninput = applyFilter;
  applyFilter();
  panel.querySelectorAll(".ms-act").forEach(a => a.onclick = () => {
    if (a.dataset.act === "all") entries.forEach(([k]) => stateSet.add(k)); else stateSet.clear();
    state.page = 1; refresh();
  });
  const n = stateSet.size, tot = entries.length;
  const txt = (n === 0 || n === tot) ? "全部" : (n + "项");
  $("#" + msId + "Btn").innerHTML = `${label}：${txt} <span class="ms-caret">▾</span>`;
  $("#" + msId + "Btn").classList.toggle("has-sel", n > 0 && n < tot);
  $("#" + msId + "Count").textContent = n === 0 ? "未选（=全部）" : ("已选 " + n + "/" + tot);
}

function updateFilterInfo() {
  const parts = [];
  if (state.q) parts.push("搜索=" + state.q);
  if (state.cats.size) parts.push("状态=" + [...state.cats].join("/"));
  if (state.reasons.size) parts.push("原因=" + [...state.reasons].map(k => reasonLabel(state.reasonTree, k)).join("/"));
  if (state.products.size) parts.push("品种=" + state.products.size + "项");
  if (state.hospitals.size) parts.push("医院=" + state.hospitals.size + "项");
  if (state.pharmacies.size) parts.push("药房=" + state.pharmacies.size + "项");
  if (state.executors.size) parts.push("随访人=" + state.executors.size + "项");
  const W = getWeekRange(state.weekSel);
  parts.push("时间=" + WEEK_LABEL[state.weekSel] + " " + W.start + "~" + W.end);
  const plain = [];
  if (state.plainName) plain.push("姓名明文");
  if (state.plainDoctor) plain.push("医生明文");
  if (state.plainPhone) plain.push("电话明文");
  if (plain.length) parts.push("展示:" + plain.join("/"));
  $("#filterInfo").textContent = parts.length ? ("筛选：" + parts.join(" · ")) : "";
}

// 当前应展示列：用户隐藏列 之外；脱敏方式为「仅会员号」时强制隐藏电话列
const visibleCols = () => LIST_COLS.filter(([key]) => {
  if (state.maskMode === "id" && key === "phone") return false; // 仅会员号模式：不呈现电话列
  return !state.hiddenCols.has(key);
});

function renderTable() {
  const cols = visibleCols();
  const total = DATA.rows.length;
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  if (state.page > pages) state.page = pages;
  const start = (state.page - 1) * state.pageSize;
  const rows = DATA.rows.slice(start, start + state.pageSize);
  $("#thead").innerHTML = cols.map(([, label]) => `<th>${label}</th>`).join("");
  const tb = $("#tbody");
  if (!total) { tb.innerHTML = ""; $("#empty").classList.remove("hidden"); }
  else {
    $("#empty").classList.add("hidden");
    tb.innerHTML = rows.map((r, i) => rowHtml(r, start + i, cols)).join("");
  }
  // 手动备注列：输入即保存到 STORE.notes（快照时一并保存）
  tb.querySelectorAll("input.note-input").forEach(inp => {
    inp.addEventListener("input", () => { STORE.notes[inp.dataset.key] = inp.value; });
    inp.addEventListener("click", e => e.stopPropagation()); // 不触发行展开
  });
  // 未购药原因列：点击 → 行内展开层级 select 选择原因（快照模式只读）
  tb.querySelectorAll("span.reason-tag.clickable").forEach(tag => {
    tag.addEventListener("click", e => {
      e.stopPropagation();
      const td = tag.parentNode;
      if (!td || SNAP_MODE) return;
      const key = tag.dataset.key;
      const cur = tag.dataset.reason || "";
      const tblWrap = document.querySelector(".tbl-wrap");
      const scrollTop = tblWrap ? tblWrap.scrollTop : 0;
      // 选项严格对齐小结第 2 行结构：无子类的顶层为平级选项；有子类的顶层作为分组标题，只列其子类
      // （小结中仅叶子人数可点选，父类人数 = 子类之和，故父类本身不作为可选原因）
      const opts = state.reasonTree.map(n => {
        const kids = n.children || [];
        if (!kids.length) {
          return `<option value="${esc(n.key)}" ${cur === n.key ? "selected" : ""}>${esc(n.label)}</option>`;
        }
        return `<optgroup label="${esc(n.label)}">` +
          kids.map(c => `<option value="${esc(c.key)}" ${cur === c.key ? "selected" : ""}>${esc(c.label)}</option>`).join("") +
          "</optgroup>";
      }).join("");
      td.innerHTML = `<select class="reason-sel" data-key="${esc(key)}" data-cur="${esc(cur)}">
        <option value="">未标</option>${opts}</select>`;
      const sel = td.querySelector("select");
      // 一步点选：点击即展开；change 即保存；点其它处或 Esc 取消（不保存）
      // 不做 blur 关闭/保存——避免 focus/重排引发误关导致需多次点击
      let closed = false;
      const finish = (saveVal) => {
        if (closed) return;
        closed = true;
        document.removeEventListener("mousedown", onDoc);
        if (saveVal) STORE.reasonOverrides[sel.dataset.key] = sel.value;
        refresh(); // 保持当前页码与筛选
        // 渲染完成后恢复列表滚动位置，避免跳回顶端
        setTimeout(() => { const w2 = document.querySelector(".tbl-wrap"); if (w2) w2.scrollTop = scrollTop; }, 0);
      };
      const onDoc = (ev) => { if (!td.contains(ev.target)) finish(false); };
      document.addEventListener("mousedown", onDoc);
      sel.addEventListener("change", () => finish(true));
      sel.addEventListener("keydown", ev => {
        if (ev.key === "Escape") { ev.preventDefault(); finish(false); }
      });
      setTimeout(() => sel.focus(), 0); // 延后聚焦，避免点击事件未完成时 select 失焦
    });
  });
  tb.querySelectorAll("tr.data-row").forEach(tr => {
    tr.onclick = () => {
      const r = DATA.rows[+tr.dataset.idx];
      const next = tr.nextElementSibling;
      if (next && next.classList.contains("expand-row")) { next.remove(); return; }
      tb.querySelectorAll("tr.expand-row").forEach(x => x.remove());
      tr.insertAdjacentHTML("afterend", expandHtml(r, cols.length));
    };
  });
  renderPagination(total);
  // 表格滚动条回到顶端（不滚动页面，避免用户被带回顶部上传区）
  const tblWrap = document.querySelector(".tbl-wrap");
  if (tblWrap) tblWrap.scrollTop = 0;
}

function rowHtml(r, idx, cols) {
  const days = r.days_to_due;
  const isRepur = r.repur_part === "应回已回";
  // 应回已回行：应购药日期列显示「预判应购日」+ 提前/延后标注（无预判基准显示 —）
  let dueTxt = r.due_date || "";
  if (isRepur) {
    if (r.due_offset == null) dueTxt = "<span class=\"muted\">—</span>";
    else if (r.due_offset > 0) dueTxt += `<span class="sub over">延期${r.due_offset}天</span>`;
    else if (r.due_offset < 0) dueTxt += `<span class="sub due">提前${-r.due_offset}天</span>`;
    else dueTxt += `<span class="sub due">按时</span>`;
  } else if (days > 0) dueTxt += `<span class="sub due">还有${days}天</span>`;
  else if (days === 0) dueTxt += `<span class="sub due today">今天</span>`;
  else if (days < 0) dueTxt += `<span class="sub over">逾期${-days}天</span>`;
  // 跟进备注：仅「应回购 / 已逾期」显示填写窗口，其余大类显示 —
  const canNote = r.status === "应回购" || r.status === "已逾期";
  const cells = cols.map(([key]) => {
    if (key === "due_date") return `<td>${dueTxt}</td>`;
    if (key === "days_to_due") {
      if (isRepur) return `<td>${esc(r.purchase_days_ago != null ? r.purchase_days_ago + "天前已购药" : "—")}</td>`;
      return `<td>${esc(days === 0 ? "今天" : (days > 0 ? "+" + days : String(days)))}</td>`;
    }
    if (key === "status") {
      // 应回购行仅显示细分标签（应回未回 / 应回已回）；其他主状态照常显示
      if (r.status === "应回购" && r.repur_part) {
        return `<td><span class="tag" style="background:${REPUR_PART_COLOR[r.repur_part]}">${esc(r.repur_part)}</span></td>`;
      }
      const color = CAT_COLOR[r.status] || "#868e96";
      return `<td><span class="tag" style="background:${color}">${esc(r.status)}</span></td>`;
    }
    if (key === "reason") {
      // 未购药原因列：自动判定预填 / 人工点选覆盖（所有患者可点选标注）；无标注按状态显示占位
      let label = "", color = "#868e96";
      if (r.reason) {
        label = reasonLabel(state.reasonTree, r.reason) || r.reason;
        color = "#e8590c";
      } else if (isRepur) {
        label = "已购药"; color = "#0f9d6b";
      } else if (r.status === "未到期") {
        label = "—";
      } else {
        label = "未标";
      }
      const clickable = SNAP_MODE ? "" : " clickable";
      const title = SNAP_MODE ? "" : "title=\"点击标注未购药原因\"";
      return `<td><span class="tag reason-tag${clickable}" data-key="${esc(noteKey(r))}" data-reason="${esc(r.reason || "")}" style="background:${color};cursor:${SNAP_MODE ? "default" : "pointer"}" ${title}>${esc(label)}</span></td>`;
    }
    if (key === "product") return `<td><span class="tag-prod">${esc(disp(r, key))}</span></td>`;
    if (key === "note") {
      if (!canNote) return `<td><span class="muted">—</span></td>`;
      const k = noteKey(r);
      const v = STORE.notes[k] || "";
      return `<td><input class="note-input" data-key="${esc(k)}" value="${esc(v)}" placeholder="填写跟进备注" ${SNAP_MODE ? "disabled" : ""}></td>`;
    }
    if (key === "fu_type") return `<td><div class="cell-clip">${esc(disp(r, key))}</div></td>`;
    if (key === "fu_signal") {
      const txt = disp(r, key);
      if (!txt) return `<td><span class="muted">—</span></td>`;
      const kind = ["normal", "nonstd", "dropout"].includes(r.fu_signal_kind) ? r.fu_signal_kind : "unknown";
      return `<td><div class="cell-clip"><span class="fu-sig ${kind}">${esc(txt)}</span></div></td>`;
    }
    return `<td>${esc(disp(r, key))}</td>`;
  }).join("");
  return `<tr class="data-row${r.due_in_week ? " due-in-week" : ""}" data-idx="${idx}">${cells}</tr>`;
}

function expandHtml(r, colspan) {
  const sec = (lbl, val) => val ? `<div class="eb-row"><span class="eb-lbl">${lbl}：</span><span class="eb-val">${esc(val)}</span></div>` : "";
  const fu = r._matched;
  let fuDetail = "";
  if (fu) {
    const rows = [
      ["随访时间", P.datePart(fu.exec_time) || P.datePart(fu.plan_time) || ""],
      ["任务状态", fu.task_status],
      ["服务摘要", fu.summary_type],
      ["执行人", fu.executor],
      ["用药状态", fu.usage_status],
      ["是否按时购药", fu.purchased_on_time],
      ["是否易脱落", fu.is_dropout],
      ["易脱落原因", fu.dropout_reason],
      ["推迟/延迟原因", fu.delay_reason],
      ["停药原因", fu.stop_reason],
      ["用法用量", fu.dosage],
      ["随访小结", fu.follow_note],
    ];
    fuDetail = rows.map(([l, v]) => sec(l, v)).join("");
  }
  // 展开卡：仅随访信息 + 跟进备注（医院/药房/医生/购药等业务字段已在主表展示，不再重复）
  return `<tr class="expand-row"><td colspan="${colspan}"><div class="expand-box">
    ${sec("跟进备注", STORE.notes[noteKey(r)] || "")}
    ${fuDetail || '<div class="eb-row"><span class="eb-lbl">无匹配随访记录</span></div>'}
  </div></td></tr>`;
}

function renderPagination(total) {
  const pages = Math.max(1, Math.ceil(total / state.pageSize));
  const opts = [50, 100, 200].map(n =>
    `<option value="${n}" ${state.pageSize === n ? "selected" : ""}>${n}</option>`).join("");
  $("#pagination").innerHTML = `
    <span class="pg-info">共 ${total} 条 · 第 ${state.page}/${pages} 页</span>
    <span class="pg-right">
      <span class="pg-size-lbl">每页</span>
      <select id="pageSizeSel" ${SNAP_MODE ? "disabled" : ""}>${opts}</select>
      <span class="pg-size-lbl">条</span>
      <span class="pg-btns">
        <button class="btn ghost sm" id="prevPg" ${state.page <= 1 || SNAP_MODE ? "disabled" : ""}>‹ 上一页</button>
        <button class="btn ghost sm" id="nextPg" ${state.page >= pages || SNAP_MODE ? "disabled" : ""}>下一页 ›</button>
      </span>
    </span>`;
  $("#pageSizeSel").onchange = e => {
    state.pageSize = Math.max(1, parseInt(e.target.value, 10) || 50);
    state.page = 1;
    renderTable();
  };
  // 翻页只滚动表格容器回顶（不 window.scrollTo，避免整页被带回顶部上传区）
  $("#prevPg").onclick = () => { if (state.page > 1) { state.page--; renderTable(); } };
  $("#nextPg").onclick = () => { if (state.page < pages) { state.page++; renderTable(); } };
}

/* ============ 整体小结（按品种分别总结；统计所选周，随周视图联动） ============ */
// 计算某品种在「所选周」的小结统计（跟随 上周/本周/下周 视图，口径与周维度状态机完全一致）
//   所选周 = 参考日期所在周（起止由「周截止日」决定，默认周一~周日）前后推
//   应回购（总集）= 应回未回（应购日∈所选周未购）+ 应回已回（所选周内购药，含提前购药者）
//   未购药 = 应回未回；实际购药 = 应回已回
//   下周预计 = 下次应购日 ∈ 所选周之后的下一周（7 天窗口）的患者，按随访信号分正常/推迟（脱落不计复购）
// 不受小结周期控件（近7天/自定义）影响
function buildSummaryStats(fam) {
  const W = getWeekRange(state.weekSel);
  const nStart = addDays(W.end, 1), nEnd = addDays(W.end, 7); // 下周 = 紧随所选周的 7 天窗口
  const inP = (d, s, e) => d && d >= s && d <= e;
  // 按患者×品种聚合购药日期（从销售明细，与状态机同一数据源）
  // 跟随所选的医院 / 药房筛选（多选=或）：小结统计范围随筛选变化
  const pat = {}; // key -> { name, phone, dates: [] }
  for (const s of STORE.sales) {
    if (s.product !== fam || !s.sales_time) continue;
    if (state.hospitals.size && !state.hospitals.has(s.hospital || "未知")) continue;
    if (state.pharmacies.size && !state.pharmacies.has(s.pharmacy || "未知")) continue;
    const k = phoneDigits(s.phone) || s.patient_name || "";
    if (!k) continue;
    (pat[k] = pat[k] || { name: s.patient_name || "", phone: s.phone || "", purchases: [] }).purchases.push({ date: s.sales_time, qty: qtyNum(s.qty) });
  }
  const cycleOf = p => STORE.cycles[p.name || p.phone] || state.stdCycle[fam] || 30;
  const boughtKeys = new Set();     // 本周实际购药（当周有购药记录）
  const notBoughtKeys = [];         // 应回未回（下次应购日∈本周 且 未购）
  const overdueKeys = [];           // 已逾期（下次应购日<周首 且 未购）—— 未购药原因统计范围之二
  const nextKeys = [];              // 下周应购（下次应购日∈下周）
  for (const k in pat) {
    const p = pat[k];
    p.purchases = normalizePurchases(p.purchases);
    const cyc = cycleOf(p);
    // 与主状态机同一套口径：数量 × 周期 + 所选周期算法
    const nextDue = stockEndAfter(p.purchases, cyc, fam);
    const boughtThisWeek = p.purchases.some(x => inP(x.date, W.start, W.end));
    if (boughtThisWeek) boughtKeys.add(k);
    else if (inP(nextDue, W.start, W.end)) notBoughtKeys.push(k);
    else if (nextDue < W.start) overdueKeys.push(k);
    if (inP(nextDue, nStart, nEnd)) nextKeys.push(k);
  }
  // 未购药原因：统计「应回未回 + 已逾期」患者的未购药原因标注（原因列人工点选/自动判定预填）
  // 父类人数 = 其子类人数之和（按用户维护的分类树）
  const rowMap = {};
  for (const r of ALL_ROWS) rowMap[r._key + "|" + r.product] = r;
  const treeKeys = flattenReasonTree(state.reasonTree).map(f => f.key);
  const cnt = Object.fromEntries(treeKeys.map(k => [k, 0]));
  const reasonScope = notBoughtKeys.concat(overdueKeys);
  for (const k of reasonScope) {
    const r = rowMap[k + "|" + fam];
    if (r && r.reason && cnt[r.reason] != null) cnt[r.reason]++;
  }
  // 父类人数 = 子类之和（父类 key 本身不作为标注目标，仅汇总展示）
  for (const n of state.reasonTree) {
    if (n.children && n.children.length) {
      cnt[n.key] = n.children.reduce((s, c) => s + (cnt[c.key] || 0), 0);
    }
  }
  // 下周预计复购：按随访信号分正常/推迟（脱落不计入复购）
  let normal = 0, postpone = 0;
  const postponeReasons = new Set();
  for (const k of nextKeys) {
    const r = rowMap[k + "|" + fam];
    const sig = r && r._matched ? M.followupSignal(r._matched) : { signal: "unknown", reason: "" };
    if (sig.signal === "nonstd") {
      postpone++;
      const dr = (r && r._matched && r._matched.delay_reason) || "";
      if (dr) postponeReasons.add(String(dr).slice(0, 30));
    } else if (sig.signal !== "dropout") {
      normal++;
    }
  }
  // 脱落合计 = 父类人数（子类之和，已在上面汇总）
  const dTotal = cnt.dropout || 0;
  // 应回购（总集）= 应回未回（应购日∈本周未购）+ 应回已回（当周购药，含提前购药）
  const dueTotal = notBoughtKeys.length + boughtKeys.size;
  return { start: W.start, end: W.end, nStart, nEnd, due: dueTotal, bought: boughtKeys.size, notBought: notBoughtKeys.length,
    cnt, dTotal, nextTotal: nextKeys.length, normal, postpone, postponeReasons: [...postponeReasons] };
}

// 未购药原因各分类数据（人工修正优先；父类人数=子类之和）——小结文本与可编辑渲染共用
function reasonParts(st, fam) {
  const adj = state.fuAdj[fam] || {};
  const v = (k, auto) => (adj[k] != null ? adj[k] : auto);
  const out = [];
  let idx = 1;
  const cn = ["", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "⑭", "⑮"];
  for (const n of state.reasonTree) {
    if (n.children && n.children.length) {
      const subs = n.children.map(c => ({ key: c.key, label: c.label, n: v(c.key, st.cnt[c.key] || 0) }));
      out.push({ cn: cn[idx] || (idx + "."), label: n.label, subs, total: subs.reduce((s, x) => s + x.n, 0) });
    } else {
      out.push({ cn: cn[idx] || (idx + "."), label: n.label, leaf: n.key, n: v(n.key, st.cnt[n.key] || 0) });
    }
    idx++;
  }
  return out;
}

// 可选范围行：小结中呈现所选医院/药房名单（state.scopeNames 开关；未筛选时为空）
function summaryScopeLine() {
  if (!state.scopeNames || !(state.hospitals.size || state.pharmacies.size)) return "";
  const hs = state.hospitals.size ? "医院：" + [...state.hospitals].join("、") : "";
  const ps = state.pharmacies.size ? "药房：" + [...state.pharmacies].join("、") : "";
  return "统计范围（" + [hs, ps].filter(Boolean).join("；") + "）";
}

// 生成小结文案（未购药原因按用户维护的分类树动态生成；父类人数=子类之和）
function buildSummaryText(fam) {
  const st = buildSummaryStats(fam);
  const parts = reasonParts(st, fam).map(p =>
    p.subs
      ? `${p.cn}${p.label}${p.total}人（${p.subs.map(x => `${x.label}${x.n}`).join("/")}）`
      : `${p.cn}${p.label}${p.n}人`);
  const wkLabel = WEEK_LABEL[state.weekSel] || "本周";
  const lines = [];
  lines.push(`${wkLabel}小结（${fam}）：`);
  lines.push(`1. ${wkLabel}应回购 ${st.due} 人（应回已回 ${st.bought} 人、应回未回 ${st.notBought} 人）；`);
  lines.push(`2. 未购药原因：${parts.join("、")}；`);
  const pReasons = st.postponeReasons.length ? `（${st.postponeReasons.slice(0, 5).join("、")}）` : "";
  // 下周预计：人工修正优先（fuAdj 键 next_normal / next_postpone）；复购总数 = 正常 + 推迟
  const adj4 = state.fuAdj[fam] || {};
  const v4 = (k, auto) => (adj4[k] != null ? adj4[k] : auto);
  const nNormal = v4("next_normal", st.normal);
  const nPostpone = v4("next_postpone", st.postpone);
  lines.push(`3. 下周预计复购 ${nNormal + nPostpone} 人，预计正常回购 ${nNormal} 人，推迟 ${nPostpone} 人${pReasons}`);
  const scopeLine = summaryScopeLine();
  if (scopeLine) lines.push(scopeLine);
  SUMMARY.text = lines.join("\n");
  return { text: SUMMARY.text, st };
}

// 渲染整体小结面板（周期/品种/文案/原因修正/复制）
function renderSummaryPanel() {
  const panel = $("#summaryPanel");
  if (!panel) return;
  // 快照模式：只读展示保存的小结文本（快照行不含随访对象，不重算）
  if (SNAP_MODE) {
    panel.classList.remove("hidden");
    $("#summaryText").textContent = SUMMARY.text || "（快照未包含小结文本）";
    $("#summaryFamChips").innerHTML = "";
    $("#summaryAdjust").innerHTML = "";
    $("#copySummaryBtn").classList.add("hidden");
    return;
  }
  const fams = [...new Set(ALL_ROWS.map(r => r.product).filter(Boolean))].sort((a, b) => a.localeCompare(b, "zh"));
  if (!fams.length) { panel.classList.add("hidden"); return; }
  panel.classList.remove("hidden");
  if (!fams.includes(state.selFam)) state.selFam = fams[0];
  // 品种 chips
  $("#summaryFamChips").innerHTML = fams.map(f =>
    `<span class="chip ${f === state.selFam ? "active" : ""}" data-fam="${esc(f)}">${esc(f)}</span>`).join("");
  document.querySelectorAll("#summaryFamChips .chip").forEach(ch => {
    ch.onclick = () => { state.selFam = ch.dataset.fam; renderSummaryPanel(); };
  });
  // 小结统计「所选周」（跟随 上周/本周/下周 视图；完整 7 天窗口，起止由「周截止日」决定，与状态机一致）
  // 统计范围跟随 医院/药房 筛选（多选=或）
  const W0 = getWeekRange(state.weekSel);
  const scope = [];
  if (state.hospitals.size) scope.push("医院" + state.hospitals.size + "家");
  if (state.pharmacies.size) scope.push("药房" + state.pharmacies.size + "家");
  $("#periodFixedLabel").textContent = `${WEEK_LABEL[state.weekSel] || "本周"}：${W0.start} ~ ${W0.end}` + (scope.length ? `（已按 ${scope.join("、")} 统计）` : "");
  // 文案：小结文本内直接编辑（第2行原因数字=可点击的橙色数字，点击原地变输入框）
  const { text, st } = buildSummaryText(state.selFam);
  const wkLabel = WEEK_LABEL[state.weekSel] || "本周";
  const fam = state.selFam;
  const line1 = `${wkLabel}小结（${fam}）：\n1. ${wkLabel}应回购 ${st.due} 人（应回已回 ${st.bought} 人、应回未回 ${st.notBought} 人）；`;
  const line2 = reasonParts(st, fam).map(p =>
    p.subs
      ? `${p.cn}${esc(p.label)}<span class="sn-total" title="由子类人数之和决定">${p.total}</span>人（${p.subs.map(x => `${esc(x.label)}<span class="sn-num" data-k="${esc(x.key)}" title="点击修改人数">${x.n}</span>`).join("/")}）`
      : `${p.cn}${esc(p.label)}<span class="sn-num" data-k="${esc(p.leaf)}" title="点击修改人数">${p.n}</span>人`).join("、");
  const pReasons = st.postponeReasons.length ? `（${st.postponeReasons.slice(0, 5).map(esc).join("、")}）` : "";
  // 下周预计：复购总数（=正常+推迟）不可编辑，正常/推迟数字可点击修改
  const adj4 = state.fuAdj[fam] || {};
  const v4 = (k, auto) => (adj4[k] != null ? adj4[k] : auto);
  const nNormal = v4("next_normal", st.normal);
  const nPostpone = v4("next_postpone", st.postpone);
  const line4 = `3. 下周预计复购 <span class="sn-total" title="由正常+推迟之和决定">${nNormal + nPostpone}</span> 人，预计正常回购 <span class="sn-num" data-k="next_normal" title="点击修改人数">${nNormal}</span> 人，推迟 <span class="sn-num" data-k="next_postpone" title="点击修改人数">${nPostpone}</span> 人${pReasons}`;
  const line5 = summaryScopeLine();
  $("#summaryText").innerHTML =
    esc(line1) + "\n2. 未购药原因：" + line2 + "；\n" + line4 + (line5 ? "\n" + esc(line5) : "");
  // 小结是否呈现所选医院/药房名单开关
  const scopeCb = $("#scopeNamesCb");
  if (scopeCb) {
    scopeCb.checked = !!state.scopeNames;
    scopeCb.onchange = () => { state.scopeNames = !!scopeCb.checked; renderSummaryPanel(); };
  }
  // 点击数字 → 原地变输入框，回车/失焦保存（Esc 取消）
  $("#summaryText").onclick = (e) => {
    if (SNAP_MODE) return;
    const t = e.target.closest(".sn-num");
    if (!t) return;
    const k = t.dataset.k;
    const input = document.createElement("input");
    input.type = "number"; input.min = "0"; input.className = "sn-input";
    input.value = t.textContent.trim();
    t.replaceWith(input);
    input.focus(); input.select();
    let done = false;
    const commit = () => {
      if (done) return; done = true;
      const n = Math.max(0, parseInt(input.value, 10) || 0);
      (state.fuAdj[fam] = state.fuAdj[fam] || {})[k] = n;
      renderSummaryPanel();
    };
    input.addEventListener("change", commit);
    input.addEventListener("blur", commit);
    input.addEventListener("keydown", ev => {
      if (ev.key === "Enter") input.blur();
      else if (ev.key === "Escape") { done = true; renderSummaryPanel(); }
    });
  };
  // 提示条
  $("#summaryAdjust").innerHTML =
    '<span class="sa-lbl">💡 小结中的橙色数字可直接点击修改（人工修正优先于自动统计，随快照保存）</span>';
  $("#summaryAdjust").classList.toggle("hidden", SNAP_MODE);
  // 管理分类（维护小结分类树；未购药原因列选项跟随）
  const mgrBtn = $("#manageReasonBtn");
  if (SNAP_MODE) mgrBtn.classList.add("hidden");
  else {
    mgrBtn.classList.remove("hidden");
    mgrBtn.onclick = () => { $("#reasonManager").classList.toggle("hidden"); renderReasonManager(); };
  }
  // 复制
  const copyBtn = $("#copySummaryBtn");
  if (SNAP_MODE) copyBtn.classList.add("hidden");
  else copyBtn.classList.remove("hidden");
  copyBtn.onclick = async () => {
    try {
      await navigator.clipboard.writeText(SUMMARY.text);
      copyBtn.textContent = "已复制 ✓";
      setTimeout(() => { copyBtn.textContent = "复制小结"; }, 1500);
    } catch (e) {
      prompt("复制失败，请手动复制以下内容：", SUMMARY.text);
    }
  };
  // 周期固定为「本周」，无切换控件
}

/* ============ 未购药原因分类维护（小结分类树；原因列选项跟随） ============ */
function findReason(tree, key) {
  for (const n of tree) {
    if (n.key === key) return n;
    const c = (n.children || []).find(x => x.key === key);
    if (c) return c;
  }
  return null;
}
function genReasonKey() { return "r" + Date.now().toString(36) + Math.floor(Math.random() * 1e4).toString(36); }
function addReason(label) { state.reasonTree.push({ key: genReasonKey(), label, children: [] }); }
function addReasonChild(parentKey, label) {
  const p = state.reasonTree.find(n => n.key === parentKey);
  if (p) (p.children = p.children || []).push({ key: genReasonKey(), label, children: [] });
}
function renameReason(key, label) { const n = findReason(state.reasonTree, key); if (n) n.label = label; }
function removeReason(key) {
  // 删除节点（连带其子类），并清理已标注该分类的患者标注
  state.reasonTree = state.reasonTree.filter(n => n.key !== key)
    .map(n => Object.assign({}, n, { children: (n.children || []).filter(c => c.key !== key) }));
  const flat = flattenReasonTree(state.reasonTree).map(f => f.key);
  for (const rk in STORE.reasonOverrides) {
    if (!flat.includes(STORE.reasonOverrides[rk])) delete STORE.reasonOverrides[rk];
  }
}
function renderReasonManager() {
  const box = $("#reasonManagerList");
  const rows = [];
  state.reasonTree.forEach(n => {
    rows.push(`<div class="rm-row">
      <span class="rm-name">${esc(n.label)}</span>
      <span class="rm-acts">
        <a data-act="addchild" data-key="${esc(n.key)}">加子类</a>
        <a data-act="rename" data-key="${esc(n.key)}">改名</a>
        <a data-act="del" data-key="${esc(n.key)}">删除</a>
      </span></div>`);
    (n.children || []).forEach(c => {
      rows.push(`<div class="rm-row rm-child">
        <span class="rm-name">${esc(c.label)}</span>
        <span class="rm-acts">
          <a data-act="rename" data-key="${esc(c.key)}">改名</a>
          <a data-act="del" data-key="${esc(c.key)}">删除</a>
        </span></div>`);
    });
  });
  box.innerHTML = rows.join("");
  box.querySelectorAll("a[data-act]").forEach(a => a.onclick = () => {
    const act = a.dataset.act, key = a.dataset.key;
    if (act === "addchild") {
      const label = prompt("新子类名称：");
      if (!label) return;
      addReasonChild(key, label.trim());
    } else if (act === "rename") {
      const cur = findReason(state.reasonTree, key);
      const label = prompt("新的分类名称：", cur ? cur.label : "");
      if (!label) return;
      renameReason(key, label.trim());
    } else if (act === "del") {
      const cur = findReason(state.reasonTree, key);
      if (!confirm(`删除分类「${cur ? cur.label : key}」？（已标注该分类的患者将回到未标）`)) return;
      removeReason(key);
    }
    renderReasonManager();
    refresh();
  });
  $("#reasonManagerAdd").onclick = () => {
    const label = prompt("新增分类名称：");
    if (!label) return;
    addReason(label.trim());
    renderReasonManager();
    refresh();
  };
  $("#reasonManagerReset").onclick = () => {
    if (!confirm("恢复默认分类？自定义修改将丢失。")) return;
    state.reasonTree = cloneReasonTree(DEFAULT_REASON_TREE);
    renderReasonManager();
    refresh();
  };
}

async function refresh() {
  // 快照模式：数据已内嵌在快照行中（STORE 为空），基于不可变的 SNAP_BASE 筛选，不重建、不累积丢失
  const all = SNAP_MODE ? (SNAP_BASE || DATA.rows) : buildRows();
  ALL_ROWS = all;
  // 状态计数基数：跳过「状态/下钻/应回购细分」维度，跟随 品种/医院/随访人/关键词 筛选
  //（需求：选择品种、医院、时间等筛选后，总的患者状态计数栏跟随变化）
  const baseRows = filterRows(all, { skipCat: true });
  CURRENT.summary = buildSummary(baseRows);
  DATA.rows = filterRows(all);
  // 多选面板用「全量」计数（不随筛选遮蔽）
  const full = buildSummary(all);
  $("#summary").innerHTML = "";
  renderSummary(Object.assign({}, CURRENT.summary, {
    by_product: full.by_product, by_hospital: full.by_hospital, by_executor: full.by_executor,
  }));
  renderTable();
  renderSummaryPanel();
  renderCycleCfg();
  renderWeekBar();
}

/* ============ 上传 / 分析（统一上传 → 列表确认/删除 → 开始分析后保留） ============ */
const KIND_LABEL = { sales: "销售", followup: "随访", cycle: "周期", unknown: "未识别" };
// 待分析文件列表：每项 { file, kind }
let pendingFiles = [];

function addPendingFiles(files, kind) {
  let added = 0;
  for (const f of files) {
    if (!pendingFiles.some(p => p.file.name === f.name && p.file.size === f.size)) {
      pendingFiles.push({ file: f, kind });
      added++;
    }
  }
  if (added) renderPendingList();
  return added;
}
function renderPendingList() {
  const list = $("#pendingList");
  if (!pendingFiles.length) {
    list.innerHTML = '<span class="empty-tip">尚未选择文件 —— 点击上方「选择文件」或拖入表格</span>';
  } else {
    list.innerHTML = pendingFiles.map((p, i) =>
      `<span class="file-chip ${p.kind === "unknown" ? "k-unknown" : ""}">${esc(p.file.name)} (${fmtSize(p.file.size)})` +
      `<span class="k-type">${KIND_LABEL[p.kind] || "未识别"}</span>` +
      `<span class="fx" data-idx="${i}" title="移除">×</span></span>`).join("");
  }
  const unknown = pendingFiles.filter(p => p.kind === "unknown").length;
  $("#pendingTotalCount").textContent = pendingFiles.length
    ? `共 ${pendingFiles.length} 个文件${unknown ? ` · ${unknown} 个未识别` : ""}`
    : "";
  $("#startBtn").disabled = pendingFiles.length === 0;
}

// 统一上传：一次选多个文件 → 逐个识别类型 → 加入列表（识别失败标「未识别」）
async function classifyFiles(files) {
  const fileList = Array.from(files || []);
  if (!fileList.length) return;
  showLoading("正在识别文件类型…");
  for (const f of fileList) {
    let kind = "unknown";
    try {
      kind = await P.detectFileType(f);
    } catch (e) {
      console.warn("识别失败", f.name, e);
      kind = "unknown";
    }
    addPendingFiles([f], kind);
  }
  hideLoading();
}

// 绑定统一上传区：整条可点击选择文件（无独立按钮），拖拽亦可
const dropAll = $("#dropAll");
dropAll.onclick = () => $("#allInput").click();
$("#allInput").onchange = e => { if (e.target.files.length) { classifyFiles(e.target.files); e.target.value = ""; } };
["dragover", "dragenter"].forEach(ev => dropAll.addEventListener(ev, e => { e.preventDefault(); dropAll.classList.add("drag"); }));
["dragleave", "drop"].forEach(ev => dropAll.addEventListener(ev, e => { e.preventDefault(); dropAll.classList.remove("drag"); }));
dropAll.addEventListener("drop", e => { const f = e.dataTransfer.files; if (f.length) classifyFiles(f); });

// 列表项移除（事件委托）
document.addEventListener("click", e => {
  const fx = e.target.closest ? e.target.closest("#pendingList .fx") : null;
  if (!fx) return;
  pendingFiles.splice(+fx.dataset.idx, 1);
  renderPendingList();
});

$("#clearPendingBtn").onclick = () => {
  pendingFiles = [];
  renderPendingList();
};

// 开始分析：用当前列表全量解析；分析后保留列表与上传窗口，可增删后再次分析
// 回传数据回流：识别回传表记录（服务摘要=「跟进回传」），把上一轮填写的内容接回本轮
//  · 「跟进备注」列 → 回填 STORE.notes（该列仍可继续编辑）
//  · 「未购药原因」标签 → 反查为分类 key 写入 STORE.reasonOverrides（等同人工点选，优先于自动判定）
// 已有本轮编辑的字段不覆盖，避免刚填的内容被导入数据抹掉
function applyCallbackRecords(followups) {
  let nNote = 0, nReason = 0;
  for (const fu of (followups || [])) {
    if (String(fu.summary_type || "").trim() !== CB_SUMMARY_TYPE) continue;
    const key = phoneDigits(fu.phone) || String(fu.patient_name || "").trim();
    if (!key || !fu.product) continue;
    const nk = key + "::" + fu.product;
    const note = String(fu.callback_note || "").trim();
    if (note && !String(STORE.notes[nk] || "").trim()) { STORE.notes[nk] = note; nNote++; }
    const rk = reasonKeyFromLabel(fu.reason_label);
    if (rk && !STORE.reasonOverrides[nk]) { STORE.reasonOverrides[nk] = rk; nReason++; }
  }
  return { nNote, nReason };
}

$("#startBtn").onclick = async () => {
  const unknown = pendingFiles.filter(p => p.kind === "unknown");
  if (unknown.length) {
    alert("以下文件未能识别为三类数据（销售/随访/周期），请移除后重试：\n" +
      unknown.map(p => p.file.name).join("\n"));
    return;
  }
  const files = pendingFiles.map(p => p.file);
  if (!files.length) return;
  showLoading("正在解析三张表并归一化计算…");
  $("#startBtn").disabled = true;
  try {
    const res = await P.processFiles(files);
    if (!res.sales.length) { alert("未识别到销售明细数据：请确认上传了含「销售时间/商品名称/会员姓名」列的销售报表。"); return; }
    STORE.sales = res.sales;
    STORE.followups = res.followups;
    STORE.cycles = res.cycles;
    // 回传数据回流（上一轮导出的「随访回传表」在本轮自动接续）
    const cb = applyCallbackRecords(res.followups);
    $("#board").classList.remove("hidden");
    $("#dataInfo").textContent =
      `销售明细 ${res.sales.length} 条 · 随访任务 ${res.followups.length} 条 · 用药周期 ${Object.keys(res.cycles).length} 人` +
      (cb.nNote || cb.nReason ? ` · 已接续回传 ${cb.nNote} 条备注 / ${cb.nReason} 条原因` : "");
    // 按数据实际出现的品种动态维护标准周期（内置值优先，新品种默认 30 天）
    const products = [...new Set(res.sales.map(s => s.product).filter(Boolean))];
    ensureStdCycles(products);
    renderCycleInputs();
    // 默认不设应购日期范围（=全部显示）；用户选择范围后表单按范围裁剪
    state.page = 1;
    await refresh();
  } catch (err) {
    alert("分析失败：" + (err && err.message ? err.message : err));
  } finally {
    hideLoading();
    renderPendingList();
  }
};
renderPendingList();

$("#clearAllBtn").onclick = () => {
  if (!confirm("确定清空全部已加载数据？")) return;
  STORE.sales = []; STORE.followups = []; STORE.cycles = {}; STORE.files = []; STORE.notes = {}; STORE.reasonOverrides = {};
  CURRENT = { summary: null }; DATA = { rows: [] };
  $("#board").classList.add("hidden");
};

/* ============ 筛选交互 ============ */
$("#clearBtn").onclick = () => {
  state.cats.clear(); state.reasons.clear(); state.repurParts.clear(); state.products.clear(); state.hospitals.clear(); state.pharmacies.clear(); state.executors.clear();
  state.weekSel = "this"; state.refDate = "";
  state.q = "";
  $("#searchInput").value = "";
  state.page = 1; refresh();
};
// 周视图：上周/本周/下周（单选）——所选周 = 状态判定的时间基准
// 周截止日：可选「本周」的定义（周五=上周六~本周五 / 周日=自然周 周一~周日 …），前后周按同一规则平移
// 参考日期：默认今天；可手动改，以回看任意历史周的应购/购药情况（周范围随其与周截止日变化）
function renderWeekBar() {
  document.querySelectorAll("#weekTabs .wk-btn").forEach(b => {
    b.classList.toggle("active", state.weekSel === b.dataset.wk);
    b.disabled = SNAP_MODE;
  });
  const W = getWeekRange(state.weekSel);
  const extra = state.refDate ? `（参考日 ${state.refDate}）` : "";
  const endName = WEEKDAY_CN[new Date(W.end + "T00:00:00").getDay()];
  $("#weekRangeLabel").textContent = `${WEEK_LABEL[state.weekSel]}：${W.start} ~ ${W.end}（${endName}止）${extra}`;
  $("#refDateInput").value = state.refDate || todayStr();
  const wsEl = $("#weekEndInput");
  if (wsEl) {
    wsEl.value = String(state.weekEnd == null ? 0 : state.weekEnd);
    wsEl.disabled = SNAP_MODE;
  }
}
document.querySelectorAll("#weekTabs .wk-btn").forEach(b => {
  b.onclick = () => {
    if (SNAP_MODE) return;
    state.weekSel = b.dataset.wk;
    state.page = 1; refresh();
  };
});
$("#refDateInput").onchange = e => {
  if (SNAP_MODE) return;
  state.refDate = e.target.value || "";
  state.page = 1; refresh();
};
// 周截止日：「本周」的可选定义（自然周 周一~周日 / 上周六~本周五 / …）
const wsEl = $("#weekEndInput");
if (wsEl) wsEl.onchange = e => {
  if (SNAP_MODE) return;
  const v = parseInt(e.target.value, 10);
  state.weekEnd = isNaN(v) ? 0 : v;
  state.page = 1; refresh();
};
// 标准周期维护（周期表未覆盖患者时使用）
// 标准周期按「数据中出现品种」动态维护：
// 只对上传数据实际出现的品种补标准周期 —— 内置品种族用内置值，新品种默认 30 天（用户可在界面修改）
// 数据中未出现的品种一律不显示、不写入
function ensureStdCycles(products) {
  for (const fam of products) {
    if (state.stdCycle[fam] == null) {
      const builtin = M.PRODUCT_FAMILIES.find(p => p.family === fam);
      state.stdCycle[fam] = builtin ? builtin.stdCycle : 30;
    }
  }
}
// 按 state.stdCycle 动态渲染每个品种的周期输入框（可单独修改，即时重算）
function renderCycleInputs() {
  const box = $("#cycleInputs");
  const fams = Object.keys(state.stdCycle).sort((a, b) => a.localeCompare(b, "zh"));
  if (!fams.length) {
    box.innerHTML = '<span class="hint">开始分析后按数据中的品种自动生成</span>';
    return;
  }
  box.innerHTML = fams.map(fam => {
    const on = !!state.qtyScale[fam];
    return `<span class="cycle-item" data-fam="${esc(fam)}"><span class="ci-name">${esc(fam)}</span>` +
      `<input type="number" class="ci-num" min="1" max="365" value="${state.stdCycle[fam]}">` +
      `<span class="ci-unit">天</span>` +
      `<label class="ci-qty${on ? " on" : ""}" title="勾选后：该品种的购药覆盖天数 = 购药数量 × 周期（患者一次买几个周期就按几个周期算）；不勾则数量按 1 计">` +
      `<input type="checkbox" class="ci-qty-cb"${on ? " checked" : ""}>按数量</label></span>`;
  }).join("");
  box.querySelectorAll(".ci-num").forEach(inp => {
    inp.addEventListener("change", () => {
      const item = inp.closest ? inp.closest(".cycle-item") : null;
      if (!item) return;
      const fam = item.dataset.fam;
      state.stdCycle[fam] = Math.max(1, parseInt(inp.value, 10) || 30);
      refresh();
    });
  });
  box.querySelectorAll(".ci-qty-cb").forEach(cb => {
    cb.addEventListener("change", () => {
      const label = cb.closest ? cb.closest(".ci-qty") : null;
      const item = cb.closest ? cb.closest(".cycle-item") : null;
      if (!item) return;
      const fam = item.dataset.fam;
      if (cb.checked) state.qtyScale[fam] = true; else delete state.qtyScale[fam];
      if (label) label.classList.toggle("on", !!cb.checked);
      refresh();
    });
  });
}
renderCycleInputs();

// 周期算法选择器 + 容差天数（容差仅在「累计 + 容差」下可改）
function renderCycleCfg() {
  const algoEl = $("#cycleAlgoInput");
  const algo = state.cycleAlgo || "reset";
  if (algoEl) { algoEl.value = algo; algoEl.disabled = SNAP_MODE; }
  const on = algo === "accum_tol";
  const tolEl = $("#cycleTolInput");
  if (tolEl) {
    tolEl.value = String(state.cycleTol == null ? 7 : state.cycleTol);
    tolEl.disabled = SNAP_MODE || !on;
  }
  ["#cycleTolLbl", "#cycleTolSuffix"].forEach(sel => {
    const el = $(sel);
    if (el) el.style.opacity = on ? "1" : ".45";
  });
}
{
  const algoEl = $("#cycleAlgoInput");
  if (algoEl) algoEl.onchange = e => {
    if (SNAP_MODE) return;
    state.cycleAlgo = e.target.value || "reset";
    renderCycleCfg();
    state.page = 1; refresh();
  };
  const tolEl = $("#cycleTolInput");
  if (tolEl) tolEl.onchange = e => {
    if (SNAP_MODE) return;
    const v = parseInt(e.target.value, 10);
    state.cycleTol = (Number.isFinite(v) && v >= 0) ? v : 7;
    renderCycleCfg();
    state.page = 1; refresh();
  };
}
let searchTimer = null;
$("#searchInput").addEventListener("input", e => {
  clearTimeout(searchTimer);
  searchTimer = setTimeout(() => { state.q = e.target.value.trim(); state.page = 1; refresh(); }, 350);
});

function closeAllPopovers() {
  document.querySelectorAll(".ms-panel:not(.hidden)").forEach(p => p.classList.add("hidden"));
  $("#colPanel").classList.add("hidden");
}
function togglePanel(panel) {
  const open = !panel.classList.contains("hidden");
  closeAllPopovers();
  if (!open) panel.classList.remove("hidden");
}
$("#prodMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#prodMsPanel")); };
$("#hospMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#hospMsPanel")); };
$("#pharmMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#pharmMsPanel")); };
$("#execMsBtn").onclick = e => { e.stopPropagation(); togglePanel($("#execMsPanel")); };
$("#colBtn").onclick = e => {
  e.stopPropagation();
  const p = $("#colPanel");
  const open = !p.classList.contains("hidden");
  closeAllPopovers();
  if (open) return;
  p.innerHTML = LIST_COLS.map(([key, label]) => {
    const phoneIdDisabled = state.maskMode === "id" && key === "phone";
    const checked = !state.hiddenCols.has(key) && !phoneIdDisabled;
    return `<label${phoneIdDisabled ? ' class="muted" title="脱敏方式为「仅会员号」时电话列自动隐藏"' : ""}><input type="checkbox" data-col="${key}" ${phoneIdDisabled ? "disabled " : ""}${checked ? "checked" : ""}> ${label}${phoneIdDisabled ? "（仅会员号隐藏）" : ""}</label>`;
  }).join("");
  p.querySelectorAll("input").forEach(cb => cb.onchange = () => {
    if (cb.checked) state.hiddenCols.delete(cb.dataset.col); else state.hiddenCols.add(cb.dataset.col);
    renderTable();
  });
  p.classList.remove("hidden");
};
document.addEventListener("click", e => {
  if (e.target.closest(".ms") || e.target.closest("#colPanel") || e.target.closest("#colBtn")) return;
  closeAllPopovers();
});

// 脱敏开关：姓名 / 电话 / 医生（脱敏=默认，不脱敏=明文）
function bindDesenToggle() {
  document.querySelectorAll(".dt-btn").forEach(b => b.onclick = () => {
    const field = b.dataset.field, mode = b.dataset.mode;
    const plain = (mode === "plain");
    if (field === "name") state.plainName = plain;
    else if (field === "phone") state.plainPhone = plain;
    else if (field === "doctor") state.plainDoctor = plain;
    b.parentElement.querySelectorAll(".dt-btn").forEach(x => x.classList.toggle("active", x === b));
    state.page = 1;
    refresh();
  });
}
bindDesenToggle();
function syncDesenBtns() {
  const map = { name: state.plainName, phone: state.plainPhone, doctor: state.plainDoctor };
  document.querySelectorAll(".dt-btn").forEach(b => {
    const active = (b.dataset.mode === "plain") === !!map[b.dataset.field];
    b.classList.toggle("active", active);
  });
}

// 脱敏方式选择：edge=首尾保留 / first=首字保留 / all=全部隐藏（展示层/导出/快照共用）
function bindMaskMode() {
  document.querySelectorAll(".mm-btn").forEach(b => b.onclick = () => {
    state.maskMode = b.dataset.mm || "edge";
    syncMaskModeBtns();
    state.page = 1;
    refresh();
  });
}
bindMaskMode();
function syncMaskModeBtns() {
  document.querySelectorAll(".mm-btn").forEach(b =>
    b.classList.toggle("active", b.dataset.mm === state.maskMode));
}

/* ============ 导出 Excel（当前筛选，带样式：表头加粗、状态/下钻浅色着色、逾期标红） ============ */
// 浅色系导出配色：浅底 + 深字
const EXPORT_CAT_STYLE = {
  "应回购": { fill: "FFFCEBEB", font: "FF791F1F" },
  "未到期": { fill: "FFE6F1FB", font: "FF0C447C" },
  "已逾期": { fill: "FFFAEEDA", font: "FF633806" },
};
const EXPORT_REPUR_STYLE = {
  "应回未回": { fill: "FFFCEBEB", font: "FF791F1F" },
  "应回已回": { fill: "FFEAF3DE", font: "FF27500A" },
};

// 导出名单：内容与界面「所见即所得」——姓名/医生/电话按各自脱敏开关取值，
// 三个开关都切「不脱敏」即得到全明文名单（原「导出未脱敏名单」的能力）
async function doExport() {
  const rows = DATA.rows;
  if (!rows.length) { alert("当前筛选无数据可导出"); return; }
  // 导出列 = 当前列表列（「仅会员号」时不含电话列，与展示一致）
  const effCols = state.maskMode === "id" ? LIST_COLS.filter(([k]) => k !== "phone") : LIST_COLS;
  const header = effCols.map(([, l]) => l).concat(["随访小结"]);
  // 列索引（用于着色）：状态=status、未购药原因=reason、距今天数=days_to_due、随访信号=fu_signal
  const idxStatus = effCols.findIndex(([k]) => k === "status");
  const idxReason = effCols.findIndex(([k]) => k === "reason");
  const idxDays = effCols.findIndex(([k]) => k === "days_to_due");
  const idxSig = effCols.findIndex(([k]) => k === "fu_signal");
  const EXPORT_SIG_FONT = { normal: "FF0F9D6B", nonstd: "FFE8590C", dropout: "FFE03131" };

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("预计购药名单");
  // 表头：浅蓝底 + 深蓝加粗字
  const hRow = ws.addRow(header);
  hRow.eachCell(c => {
    c.font = { bold: true, color: { argb: "FF0C447C" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6F1FB" } };
    c.alignment = { vertical: "middle", horizontal: "center" };
    c.border = { bottom: { style: "thin", color: { argb: "FFB5D4F4" } } };
  });
  // 数据行
  for (const r of rows) {
    const vals = effCols.map(([k]) => {
      if (k === "note") return STORE.notes[noteKey(r)] || "";
      // 应回购行仅导出细分（应回未回 / 应回已回，与列表显示一致）
      if (k === "status") return (r.status === "应回购" && r.repur_part) ? r.repur_part : r.status;
      if (k === "reason") {
        // 未购药原因：标注分类 / 无标注按状态占位（与展示一致）
        if (r.reason) return reasonLabel(state.reasonTree, r.reason) || r.reason;
        if (r.repur_part === "应回已回") return "已购药";
        return "";
      }
      // 仅会员号模式（id）：患者列导出会员号、医生列空（与列表展示一致）
      // 与界面三组开关保持一致：该字段开关为「不脱敏」时导出明文（所见即所得）
      if (k === "patient_name") {
        if (state.plainName) return r.patient_name || "";
        if (state.maskMode === "id") return r.member_id || "";
        return maskName(r.patient_name);
      }
      if (k === "phone") {
        if (state.plainPhone) return r.phone || "";
        return maskPhone(r.phone);
      }
      if (k === "physician") {
        if (state.plainDoctor) return r.physician || "";
        if (state.maskMode === "id") return "";
        return maskDoctor(r.physician);
      }
      // 应回已回行：距今天数列导出「N天前已购药」（与展示一致），应购药日期列已是预判应购日
      if (k === "days_to_due" && r.repur_part === "应回已回") {
        return r.purchase_days_ago != null ? r.purchase_days_ago + "天前已购药" : "";
      }
      return r[k] == null ? "" : r[k];
    });
    vals.push(r.fu_note || "");
    const row = ws.addRow(vals);
    // 应购药日期落在所选周：整行浅金底色（与列表特殊底色一致；列专属样式随后覆盖）
    if (r.due_in_week) {
      row.eachCell(c => {
        c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF3D6" } };
      });
    }
    // 状态列：浅色填充 + 深色字（应回购细分按 应回未回=浅红 / 应回已回=浅绿 区分着色）
    const st = (r.status === "应回购" && r.repur_part) ? EXPORT_REPUR_STYLE[r.repur_part] : EXPORT_CAT_STYLE[r.status];
    if (idxStatus >= 0 && st) {
      row.getCell(idxStatus + 1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: st.fill } };
      row.getCell(idxStatus + 1).font = { bold: true, color: { argb: st.font } };
    }
    // 未购药原因列：有标注橙色着色
    if (idxReason >= 0 && r.reason) {
      row.getCell(idxReason + 1).font = { bold: true, color: { argb: "FFE8590C" } };
    }
    // 距今天数：逾期标红加粗
    if (idxDays >= 0 && typeof r.days_to_due === "number" && r.days_to_due < 0) {
      row.getCell(idxDays + 1).font = { bold: true, color: { argb: "FFE03131" } };
    }
    // 随访信号：按类别着色（正常=绿 / 可能延期=橙 / 脱落=红）
    if (idxSig >= 0 && r.fu_signal && EXPORT_SIG_FONT[r.fu_signal_kind]) {
      row.getCell(idxSig + 1).font = { color: { argb: EXPORT_SIG_FONT[r.fu_signal_kind] } };
    }
  }
  // 列宽（19 个 LIST_COLS 列 + 随访小结）
  const widths = [10, 14, 10, 10, 22, 18, 12, 10, 10, 12, 10, 14, 18, 10, 10, 10, 24, 22, 40];
  ws.columns.forEach((c, i) => { c.width = widths[i] || 12; });
  // 冻结首行
  ws.views = [{ state: "frozen", ySplit: 1 }];
  const buf = await wb.xlsx.writeBuffer();
  download(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    "预计购药名单.xlsx");
}
// 导出名单（单一按钮）：按当前三组开关导出，所见即所得；含明文字段时先提示隐私风险
$("#exportBtn").onclick = () => {
  if (!DATA.rows.length) { alert("当前筛选无数据可导出"); return; }
  const plain = [];
  if (state.plainName) plain.push("姓名");
  if (state.plainDoctor) plain.push("医生");
  if (state.plainPhone) plain.push("电话");
  if (plain.length) {
    const ok = confirm(
      "⚠️ 隐私风险提示\n\n当前「" + plain.join("、") + "」开关为「不脱敏」，\n本次导出将包含这些字段的明文。\n\n确定导出？"
    );
    if (!ok) return;
  }
  doExport();
};

/* ============ 随访回传表导出（供下一轮作为随访数据回传） ============ */
// 回传标记：识别签名用固定值，导入侧据此把数据回流到备注与未购药原因
const CB_TASK_STATUS = "已完成";
const CB_SUMMARY_TYPE = "跟进回传";
// 回传表列（顺序即导出顺序）：既满足导入侧的类型识别签名（任务状态+服务摘要+联系电话），
// 又能被逐字段映射回随访记录。强制明文，否则回传后无法匹配到患者。
const CB_COLS = [
  ["任务状态", () => CB_TASK_STATUS],
  ["服务摘要", () => CB_SUMMARY_TYPE],
  ["患者姓名", r => r.patient_name || ""],
  ["联系电话", r => r.phone || ""],
  ["产品", r => r.product || ""],
  ["执行时间", () => refToday()],
  ["执行人", r => r.executor || ""],
  ["未购药原因", r => STORE.reasonOverrides[noteKey(r)] ? reasonLabel(state.reasonTree, STORE.reasonOverrides[noteKey(r)]) : ""],
  ["跟进备注", r => STORE.notes[noteKey(r)] || ""],
  ["随访小结", r => r.fu_note || ""],
];

// 原因标签 → 分类 key 反查（用于回传数据回流；path 优先，其次节点名）
function reasonKeyFromLabel(label) {
  const s = String(label == null ? "" : label).trim();
  if (!s) return "";
  const flat = flattenReasonTree(state.reasonTree);
  const byPath = flat.find(x => x.path === s);
  if (byPath) return byPath.key;
  const byLabel = flat.find(x => x.label === s);
  return byLabel ? byLabel.key : "";
}

function callbackRows() {
  // 只回传真正有填写的行：填了跟进备注，或人工点选过未购药原因
  return DATA.rows.filter(r => {
    const k = noteKey(r);
    return (STORE.notes[k] && String(STORE.notes[k]).trim()) || STORE.reasonOverrides[k];
  });
}

async function doExportCallback() {
  const rows = callbackRows();
  if (!rows.length) {
    alert("当前筛选内没有可回传的内容。\n\n请先在名单中填写「跟进备注」或点选「未购药原因」，再导出回传表。");
    return;
  }
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("随访回传");
  const header = CB_COLS.map(([l]) => l);
  const hRow = ws.addRow(header);
  hRow.eachCell(c => {
    c.font = { bold: true, color: { argb: "FF0C447C" } };
    c.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE6F1FB" } };
    c.alignment = { vertical: "middle", horizontal: "center" };
    c.border = { bottom: { style: "thin", color: { argb: "FFB5D4F4" } } };
  });
  for (const r of rows) ws.addRow(CB_COLS.map(([, get]) => get(r)));
  // 列宽
  const widths = [12, 12, 14, 16, 12, 14, 12, 20, 40, 40];
  ws.columns.forEach((c, i) => { c.width = widths[i] || 14; });
  ws.views = [{ state: "frozen", ySplit: 1 }];
  const buf = await wb.xlsx.writeBuffer();
  const d = refToday();
  download(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
    `随访回传_${d}.xlsx`);
  alert(
    `已导出 ${rows.length} 条跟进回传记录。\n\n` +
    "下次分析时，把该文件与销售/随访/周期表一起上传即可：\n" +
    "· 跟进备注会回填到「跟进备注」列（仍可编辑）\n" +
    "· 未购药原因等同人工点选\n\n" +
    "⚠️ 该表含明文姓名与电话，请仅在团队内部流转。"
  );
}
$("#exportCallbackBtn").onclick = () => doExportCallback();

/* ============ 快照生成（自包含 HTML） ============ */
async function doSnapshot(desen) {
  if (!DATA.rows.length) { alert("当前筛选无数据可生成快照"); return; }
  showLoading(desen ? "正在生成脱敏快照…" : "正在生成不脱敏快照…");
  try {
    const rows = DATA.rows.map(r => {
      const c = Object.assign({}, r);
      delete c._matched; delete c._row_id;
      if (desen) {
        // 与界面三组开关保持一致：该字段开关为「不脱敏」时快照保留明文
        if (!state.plainName) c.patient_name = maskName(c.patient_name);
        if (!state.plainPhone) c.phone = maskPhone(c.phone);
        if (!state.plainDoctor) c.physician = maskDoctor(c.physician);
      }
      return c;
    });
    const snap = {
      desen, rows, notes: STORE.notes, reasonOverrides: STORE.reasonOverrides,
      summary: CURRENT.summary, state: {
        weekSel: state.weekSel, refDate: state.refDate, weekEnd: state.weekEnd,
        stdCycle: state.stdCycle, pageSize: state.pageSize,
        cycleAlgo: state.cycleAlgo, cycleTol: state.cycleTol, qtyScale: state.qtyScale,
        q: state.q,
        cats: [...state.cats], reasons: [...state.reasons], repurParts: [...state.repurParts],
        products: [...state.products],
        hospitals: [...state.hospitals], pharmacies: [...state.pharmacies], executors: [...state.executors],
        plainName: state.plainName, plainPhone: state.plainPhone, plainDoctor: state.plainDoctor, scopeNames: state.scopeNames,
        maskMode: state.maskMode,
        hiddenCols: [...state.hiddenCols],
        periodType: state.periodType, periodStart: state.periodStart, periodEnd: state.periodEnd,
        selFam: state.selFam, reasonTree: state.reasonTree, fuAdj: state.fuAdj,
      },
      summaryText: SUMMARY.text,
      buildAt: new Date().toISOString(),
      note: "预计购药分析快照",
    };
    const dataJson = JSON.stringify(snap).replace(/</g, "\\u003c");
    // 启动脚本（快照只读模式加载数据）
    const bootCode = `(function(){
      if(window.__SNAP__&&window.AppCore){window.AppCore.loadSnapshot(window.__SNAP__);}
      else{var ld=document.getElementById('loading');if(ld){ld.querySelector('.txt').textContent='快照加载失败：应用脚本未加载。';ld.classList.remove('hidden');}}
    })();`;
    // 剥离库代码 + 注入数据/启动脚本：全部用真实 DOM 操作（快照为只读视图，去掉 Excel 解析库减小体积）。
    // 不能用文本正则：应用代码自身内联时也含 <script>/</body> 等字面，正则会误删应用代码或注入错位。
    const isLib = s => {
      const src = s.getAttribute && s.getAttribute("src") || "";
      const id = s.id || "";
      return /vendor\/(xlsx|exceljs)\.min\.js$/.test(src) || id === "__vnd_xlsx__" || id === "__vnd_exceljs__";
    };
    const libScripts = Array.from(document.querySelectorAll("script")).filter(isLib);
    libScripts.forEach(s => s.remove());
    const ds = document.createElement("script");
    ds.textContent = "window.__SNAP__=" + dataJson;
    const bs = document.createElement("script");
    bs.textContent = bootCode;
    document.body.appendChild(ds);
    document.body.appendChild(bs);
    const html = "<!DOCTYPE html>\n" + document.documentElement.outerHTML;
    ds.remove(); bs.remove();
    libScripts.forEach(s => document.head.appendChild(s)); // 放回页面（Excel 导出仍可用）
    download(new Blob([html], { type: "text/html" }),
      desen ? "预计购药名单_脱敏快照.html" : "预计购药名单_不脱敏快照.html");
  } catch (err) {
    alert("快照生成失败：" + (err && err.message ? err.message : err));
  } finally {
    hideLoading();
  }
}
$("#snapshotDesenBtn").onclick = () => doSnapshot(true);
$("#snapshotPlainBtn").onclick = () => {
  const ok = confirm(
    "⚠️ 隐私风险提示\n\n「未脱敏快照」包含明文姓名、电话、医生等个人敏感信息。\n仅可分享给可信的团队成员，切勿通过外部渠道转发。\n\n如仅需对外分享，请改用「脱敏快照」。\n\n确定生成未脱敏快照吗？"
  );
  if (ok) doSnapshot(false);
};

/* ============ 快照打开模式 ============ */
function loadSnapshot(snap) {
  hideLoading();
  SNAP_MODE = true;
  SNAP_BASE = (snap.rows || []).map(r => Object.assign({}, r));
  DATA.rows = SNAP_BASE.slice();
  CURRENT.summary = snap.summary || buildSummary(DATA.rows);
  STORE.notes = snap.notes || {};
  STORE.reasonOverrides = snap.reasonOverrides || {};
  const s = snap.state || {};
  state.weekSel = s.weekSel || "this";
  state.refDate = s.refDate || "";
  // 周截止日（兼容旧快照的 weekStart：起始日 +6 天即为截止日）
  state.weekEnd = (() => {
    let v = s.weekEnd;
    if (v == null && s.weekStart != null) v = parseInt(s.weekStart, 10) + 6;
    if (v == null) return 0;
    const n = parseInt(v, 10);
    return isNaN(n) ? 0 : ((n % 7) + 7) % 7;
  })();
  state.pageSize = s.pageSize || 50;
  state.page = 1;
  if (s.stdCycle) state.stdCycle = s.stdCycle;
  state.cycleAlgo = s.cycleAlgo || "reset";
  state.cycleTol = (s.cycleTol == null || !Number.isFinite(Number(s.cycleTol))) ? 7 : Number(s.cycleTol);
  state.qtyScale = s.qtyScale || {};
  renderCycleInputs();
  state.q = s.q || "";
  state.cats = new Set(s.cats || []);
  state.reasons = new Set(s.reasons || []);
  state.repurParts = new Set(s.repurParts || []);
  state.products = new Set(s.products || []);
  state.hospitals = new Set(s.hospitals || []);
  state.pharmacies = new Set(s.pharmacies || []);
  state.executors = new Set(s.executors || []);
  state.plainName = !!s.plainName; state.plainPhone = !!s.plainPhone; state.plainDoctor = !!s.plainDoctor;
  state.maskMode = s.maskMode || "edge";
  state.scopeNames = s.scopeNames !== false;
  if (s.hiddenCols) state.hiddenCols = new Set(s.hiddenCols);
  // 整体小结状态恢复（快照中只读展示保存的小结文本）
  state.periodType = s.periodType || "7d";
  state.periodStart = s.periodStart || ""; state.periodEnd = s.periodEnd || "";
  state.selFam = s.selFam || "";
  if (s.reasonTree) state.reasonTree = cloneReasonTree(s.reasonTree);
  state.fuAdj = s.fuAdj || {};
  SUMMARY.text = snap.summaryText || "";
  // 上传区隐藏
  ["#dropAll", "#pendingPanel"].forEach(id => $(id).classList.add("hidden"));
  $("#snapBanner").classList.remove("hidden");
  $("#snapBanner").textContent = snap.desen
    ? "📄 这是一份脱敏快照：姓名、电话、医生已按所选方式脱敏，文件中不含任何明文个人信息，可安全分享。"
    : "⚠️ 这是一份「不脱敏」快照：包含明文姓名、电话、医生等个人信息，仅可分享给可信接收方。";
  const fb = document.querySelector(".filebar"); if (fb) fb.classList.add("hidden");
  ["#exportBtn", "#exportCallbackBtn", "#snapshotDesenBtn", "#snapshotPlainBtn", "#clearAllBtn"].forEach(sl => { const el = $(sl); if (el) el.classList.add("hidden"); });
  ["#refDateInput", "#clearBtn"].forEach(id => $(id).disabled = true);
  document.querySelectorAll("#cycleInputs .ci-num").forEach(inp => inp.disabled = true);
  document.querySelectorAll(".mm-btn").forEach(b => b.disabled = true);
  $("#dataInfo").textContent = `快照 · 生成于 ${new Date(snap.buildAt).toLocaleString("zh-CN")}`;
  $("#board").classList.remove("hidden");
  syncDesenBtns();
  syncMaskModeBtns();
  const full = buildSummary(DATA.rows);
  renderSummary(Object.assign({}, CURRENT.summary, {
    by_product: full.by_product, by_hospital: full.by_hospital, by_executor: full.by_executor,
  }));
  renderTable();
  // 快照模式：小结只读展示保存的文本（渲染面板会自动隐藏控制/修正区）
  renderSummaryPanel();
}

window.AppCore = { loadSnapshot, buildRows, filterRows, buildSummary, doExport, doExportCallback, callbackRows, applyCallbackRecords, reasonKeyFromLabel, disp, qtyNum, qtyScales, coverDays, normalizePurchases, stockEndAfter, renderCycleCfg, LIST_COLS, STORE, state, DATA, refresh, ensureStdCycles, renderCycleInputs, renderTable, renderPagination, buildSummaryStats, buildSummaryText, renderSummaryPanel, getWeekRange, renderWeekBar, refToday, WEEK_LABEL, classifyFuReason: M.classifyFuReason, DEFAULT_REASON_TREE, cloneReasonTree, flattenReasonTree, reasonLabel, renderReasonManager, addReason, addReasonChild, renameReason, removeReason };
})();
