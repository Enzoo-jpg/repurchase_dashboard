// 预计购药分析：三表映射 / 品种归一化 / 状态判定配置（浏览器端）
// 纯函数 + 常量，无 DOM 依赖，可在浏览器与 Node 共用。

(function () {
// ---------------------------------------------------------------------------
// 品种家族与标准周期
// 商品名称/产品 列的取值可能是：达罗他胺片(诺倍戈)、洛拉替尼片(博瑞纳)、
// 达罗他胺片、洛拉替尼片…… 用关键词归一到家族名。
// ---------------------------------------------------------------------------
const PRODUCT_FAMILIES = [
  { family: "诺倍戈", kws: ["达罗他胺", "诺倍戈", "darolutamide"], stdCycle: 30 },
  { family: "博瑞纳", kws: ["洛拉替尼", "博瑞纳", "lorlatinib"], stdCycle: 30 },
];

function normalizeProduct(v) {
  if (!v) return null;
  const s = String(v);
  for (const p of PRODUCT_FAMILIES) {
    if (p.kws.some(k => s.includes(k))) return p.family;
  }
  return s.length > 12 ? s.slice(0, 12) : s; // 未知品种原样截断展示
}

// 回购分析的状态分类（名单主维度）。
// 大类：应回购（= 应回未回 + 应回已回，所选周应购清单）/ 未到期 / 已逾期。
// 「应回购」细分：应回未回（应购日∈所选周未购药）/ 应回已回（所选周内有购药记录）。
// 所有患者按下钻子状态标注：正常状态 / 预判延期 / 已脱落（依据最近随访内容判定，用户可点选手动修正）。
const CATEGORIES = ["应回购", "未到期", "已逾期"];
const REPUR_PARTS = ["应回未回", "应回已回"];
const SUBSTATUS = ["正常状态", "预判延期", "已脱落"];

// ---------------------------------------------------------------------------
// 三张表的来源识别签名（表头关键字子串，抗尾标）
// ---------------------------------------------------------------------------
const TABLE_SIGNATURES = [
  ["sales", ["销售时间", "商品名称", "会员姓名", "药房名称"]],
  ["followup", ["任务状态", "服务摘要", "联系电话", "计划执行时间"]],
  ["cycle", ["用药周期"]],
];

// 单值字段关键字映射：列标题包含任一关键字即命中，取命中数最多的一列。
// 随访表是多模板合并大表（同一逻辑字段可能出现多次），故 followup 的
// 状态相关字段全部走 MULTI（收集所有命中列，行内跨列读取，取有值者）。
const KEYWORD_RULES = [
  // ---------- 销售明细 ----------
  ["sales_time",    ["销售时间", "购药时间", "成交时间"]],
  ["order_status",  ["订单状态"]],
  ["product",       ["商品名称", "药品名称", "产品名称", "产品", "药品"]],
  ["qty",           ["销售数量", "数量"]],
  ["amount",        ["含税金额", "无税金额", "实收金额", "金额"]],
  ["member_id",     ["会员号", "会员编号", "会员卡号"]],
  ["patient_name",  ["会员姓名", "客户姓名", "患者姓名", "姓名"]],
  ["phone",         ["会员电话", "手机号", "联系电话", "手机", "电话"]],
  ["hospital",      ["医疗单位", "首诊医院", "就诊医院", "医院"]],
  ["pharmacy",      ["药房名称", "门店", "药店名称", "药店", "药房"]],
  ["physician",     ["处方医生", "开单医生", "医生"]],
  ["department",    ["处方科室", "科室", "就诊科室"]],
  ["indication",    ["适应症", "临床诊断", "诊断"]],
  ["age",           ["年龄"]],
  ["gender",        ["性别"]],
  // ---------- 随访任务 ----------
  ["task_status",   ["任务状态"]],
  ["f_patient_name",["患者", "患者姓名", "姓名"]],
  ["is_key",        ["是否重点患者", "重点患者"]],
  ["patient_id",    ["患者编号", "患者编码", "编号"]],
  ["f_phone",       ["联系电话", "手机号", "电话", "电话号码"]],
  ["plan_time",     ["计划执行时间", "计划执行日期"]],
  ["f_product",     ["产品", "药品", "商品"]],
  ["f_indication",  ["适应症"]],
  ["task_no",       ["任务编号", "任务ID"]],
  ["summary_type",  ["服务摘要", "任务摘要", "随访类型", "任务类型"]],
  ["create_time",   ["创建时间"]],
  ["executor",      ["执行人", "随访人", "经办人", "负责人", "坐席", "话务员", "客服"]],
  ["exec_time",     ["执行时间"]],
  ["cancel_reason", ["取消原因"]],
  // 回传表专用：用户在看板里人工选定的未购药原因（写中文分类标签，导入后直接采用，等同人工点选）
  ["reason_label",  ["未购药原因", "未购药原因分类"]],
  // 回传表专用：用户在「跟进备注」列填写的内容。
  // 单列独立映射，避免与「随访小结」一起被 MULTI 的 follow_note 合并成一列。
  ["callback_note", ["跟进备注"]],
  // ---------- 患者用药周期表 ----------
  ["c_patient_name",["患者", "患者姓名", "姓名"]],
  ["cycle_days",    ["用药周期", "周期天数", "购药周期", "周期"]],
];

// 多值字段：随访表状态相关列（收集所有命中列；每行仅其所属模板那一列有值）
const KEYWORD_RULES_MULTI = [
  // 「患者目前的用药状态是？」在多个模板中出现；历史任务表为「用药周期状态」
  ["usage_status", ["患者目前的用药状态"]],
  // 「据上次随访后，是否已按时购药？」；历史任务表为「用药周期状态」（按计划持续用药）
  ["purchased_on_time", ["是否已按时购药", "已按时购药", "用药周期状态"]],
  // 「是否判定患者属于易脱落（失访/停药）高风险？」；历史任务表为「是否脱落 / 是否确认脱落」
  ["is_dropout", ["是否判定患者属于易脱落", "是否判定患者属于易脱落（失访/停药）高风险", "是否脱落", "是否确认脱落"]],
  ["dropout_reason", ["判断为易脱落的原因", "脱落原因", "脱落/流失原因"]],
  // 「还在按照处方继续使用诺倍戈/博瑞纳」
  ["still_using", ["还在按照处方继续使用"]],
  ["near_usage", ["患者近一次用药情况"]],
  ["delay_reason", ["推迟用药的原因", "延迟用药的具体原因", "未按计划持续用药原因"]],
  ["stop_reason", ["停药的原因是什么", "停药的具体原因", "停止用药，主要是谁的建议"]],
  ["dosage", ["的用法用量", "当前用法用量"]],
  // 随访小结 / 患者反馈（多模板措辞不同，跨列取有值者）
  ["follow_note", ["随访小结", "小结", "随访记录", "患者反馈内容", "患者反馈", "备注"]],
];

function normHeader(h) {
  return String(h == null ? "" : h).replace(/\s+/g, "").toLowerCase();
}

function _cell(row, col) {
  if (col == null) return null;
  const v = row[col];
  if (v == null) return null;
  if (typeof v === "number" && Number.isNaN(v)) return null;
  const s = String(v).trim();
  return (s === "" || s === "nan" || s === "NaN" || s === "None") ? null : s;
}

function _gtext(row, colmap, field) {
  let cols = colmap[field];
  if (cols == null) return "";
  if (typeof cols === "number" || typeof cols === "string") cols = [cols];
  const vals = [], seen = new Set();
  for (const c of cols) {
    const v = _cell(row, c);
    if (v && !seen.has(v)) { seen.add(v); vals.push(v); }
  }
  return vals.join("\n") || "";
}

const RE_PLACEHOLDER = /^(无|无。|没有|暂无|不详|未知|不清楚|不知道|无特殊|无异常|n\/?a|null|-+|\/|\.|。)$/i;
function isPlaceholder(v) {
  return !v || RE_PLACEHOLDER.test(String(v).trim());
}

// ---------------------------------------------------------------------------
// 表类型识别：按表头签名（关键字命中数打分）
// ---------------------------------------------------------------------------
function detectTableType(cols) {
  const normCols = cols.map(normHeader);
  let best = "unknown", bestScore = -1;
  for (const [ttype, sigs] of TABLE_SIGNATURES) {
    let score = 0;
    for (const sig of sigs) {
      const ns = normHeader(sig);
      if (normCols.some(c => c.includes(ns))) score++;
    }
    if (score > bestScore) { bestScore = score; best = ttype; }
  }
  // 销售表与随访表都可能含「产品」列，需用更强特征区分
  if (best === "followup" && normCols.some(c => c.includes("任务状态")) && normCols.some(c => c.includes("服务摘要"))) return "followup";
  if (best === "sales" && normCols.some(c => c.includes("销售时间")) && normCols.some(c => c.includes("会员姓名"))) return "sales";
  if (best === "cycle" && normCols.some(c => c.includes("用药周期"))) return "cycle";
  return bestScore > 0 ? best : "unknown";
}

// ---------------------------------------------------------------------------
// 随访修正信号：从匹配到的随访记录提取业务信号
// 返回 { signal: "dropout" | "nonstd" | "normal" | "unknown", reason, fu_time, task_type, executor }
//  dropout = 明确停药/脱落/转渠道/失访/去世/不再治疗
//  nonstd  = 减量/延迟/推迟/未按时购药/易脱落标记等（可能延期信号）
//  normal  = 规范/按时/足量等正向信号
// ---------------------------------------------------------------------------
function RE_DROPOUT_TXT() { return /停药|停用|停[／/]换|转渠道|转院|失访|去世|拒绝随访|脱落|不再治疗|放弃治疗|不治疗/; }
function RE_NONSTD_TXT() { return /减量|减药|延迟|推迟|延后|未按时|漏服|漏药|自行|不规范|不依从|经济|负担|买不起|吃不起|未按计划|未持续|中断/; }
function RE_NORMAL_TXT() { return /规范用药|正常用药|按时|足量|按医嘱|规律用药|继续使用|正在使用|按计划|持续用药/; }

function followupSignal(fu) {
  if (!fu) return { signal: "unknown", reason: "" };
  const parts = [];
  // 1. 患者目前的用药状态
  const usage = fu.usage_status || "";
  if (usage) {
    if (/停药|停用/.test(usage)) return { signal: "dropout", reason: "用药状态:" + usage };
    if (RE_NONSTD_TXT().test(usage)) return { signal: "nonstd", reason: "用药状态:" + usage };
    if (RE_NORMAL_TXT().test(usage)) parts.push("用药状态:" + usage);
  }
  // 2. 是否已按时购药（注意：该列部分模板行会混入「复查周期」类噪声值；
  //    历史任务表为「用药周期状态」，值为 按计划持续用药 / 停药----脱落 / 推迟购药 / 未按计划持续用药-不依从）
  const ontime = fu.purchased_on_time || "";
  if (ontime) {
    if (/停药|停用|脱落/.test(ontime)) return { signal: "dropout", reason: "用药周期状态:" + ontime };
    if (/^(否|未)/.test(ontime) || /未按计划|未持续|中断|推迟|延迟|暂缓/.test(ontime)) return { signal: "nonstd", reason: "未按时购药" };
    if (/^(是|已)/.test(ontime) || /按计划|持续用药|按时/.test(ontime)) parts.push("已按时购药");
    else if (/不清楚|未知/.test(ontime)) parts.push("按时购药情况:" + ontime);
    // 复查周期类噪声（月度/季度复查等）不参与判定
  }
  // 3. 是否判定易脱落 + 易脱落原因
  const drop = fu.is_dropout || "";
  if (drop && /^(是|有)/.test(drop)) {
    const dr = fu.dropout_reason || "";
    if (RE_DROPOUT_TXT().test(dr)) return { signal: "dropout", reason: "易脱落:" + dr };
    return { signal: "nonstd", reason: "易脱落判定:" + (dr || drop) };
  }
  if (fu.dropout_reason && RE_DROPOUT_TXT().test(fu.dropout_reason)) {
    return { signal: "dropout", reason: "易脱落原因:" + fu.dropout_reason };
  }
  // 4. 仍在按处方使用（诺倍戈/博瑞纳）
  const still = fu.still_using || "";
  if (still) {
    if (/是的|正在|在吃|在服|继续|使用/.test(still) && !/没有|停|否/.test(still)) parts.push(still);
    if (/停药|没有使用|没有在吃|停止|停用|否/.test(still) && !/正在|继续|在吃|在服/.test(still)) {
      return { signal: "dropout", reason: "未继续使用:" + still };
    }
    if (/不确定|没怎么用|漏/.test(still)) return { signal: "nonstd", reason: "未继续使用:" + still };
  }
  // 5. 停药原因列有值 → 脱落；推迟/延迟原因列有值 → 非规范
  if (fu.stop_reason) return { signal: "dropout", reason: "停药原因:" + fu.stop_reason.slice(0, 40) };
  if (fu.delay_reason) return { signal: "nonstd", reason: "推迟原因:" + fu.delay_reason.slice(0, 40) };
  // 6. 随访小结文本（自由文本，仅作弱信号）
  const note = fu.follow_note || "";
  if (note) {
    if (RE_DROPOUT_TXT().test(note)) return { signal: "dropout", reason: note.slice(0, 50) };
    if (RE_NONSTD_TXT().test(note)) return { signal: "nonstd", reason: note.slice(0, 50) };
    if (RE_NORMAL_TXT().test(note)) parts.push("小结正常");
  }
  // 7. 任务状态修正：超期未购药随访任务 + 已按时购药=否 已在上文处理；执行失败/未执行不参与
  if (fu.task_status && /执行失败|超期未完成/.test(fu.task_status) && /超期未购药/.test(fu.summary_type || "")) {
    return { signal: "nonstd", reason: "超期未购药随访:" + fu.task_status };
  }
  return { signal: parts.length ? "normal" : "unknown", reason: parts.join("；") };
}

// ---- 未购药原因细分（整体小结用，按模板顺序判定，每人不重复计数）----
// 分类（可维护树）：推迟购药 / 延长周期(医嘱·自行) / 脱落(换药·经济·去世·其他) / 转渠道 / 随访失败未探寻原因
function RE_DELAY_TXT() { return /延迟|推迟|延后|未按时|过几天|晚几天|暂缓/; }
function RE_ECON_TXT() { return /经济|费用|太贵|负担|没钱|买不起|吃不起|贵/; }
function RE_CHANNEL_TXT() { return /转渠道|转院|转店|换店|其他药房|异地|外地|其他医院/; }
function RE_FUFAIL_TXT() { return /随访失败|联系不上|拒访|拒绝随访|无法联系|关机|无人接听/; }
function RE_PROLONG_TXT() { return /医嘱延长|遵医嘱延长|医生建议延长|延长用药|延长疗程/; }
function RE_PROLONG_DOCTOR_TXT() { return /医嘱|遵医嘱|遵医|医生建议|按医嘱/; }
function RE_SWITCH_TXT() { return /换药|换品种|更换药品|换用|改用|换成/; }
function RE_DEATH_TXT() { return /去世|死亡|过世|逝世/; }
// 脱落判定专用（排除「转渠道/转院」，避免与③转渠道冲突）
function RE_DROPOUT_ONLY() { return /停药|停用|脱落|失访|去世|拒绝随访|不再治疗|放弃治疗|不治疗/; }

// 返回 { key: 'delay'|'dropout'|'channel'|'fuFail'|'prolong'|null,
//         detail: dropout→'换药'|'经济'|'去世'|'其他'；prolong→'医嘱'|'自行'；其余 '' }
function classifyFuReason(fu) {
  if (!fu) return null;
  const note = [fu.follow_note, fu.delay_reason, fu.stop_reason, fu.dropout_reason].filter(Boolean).join(" ");
  const ts = fu.task_status || "";
  // ① 推迟购药
  if (fu.delay_reason || RE_DELAY_TXT().test(note)) return { key: "delay", detail: "" };
  // ② 脱落（细分：换药/经济/去世/其他；去世最明确优先）
  if (fu.stop_reason || (fu.is_dropout && /^(是|有)/.test(fu.is_dropout)) || RE_DROPOUT_ONLY().test(note)) {
    let d = "其他";
    if (RE_DEATH_TXT().test(note)) d = "去世";
    else if (RE_SWITCH_TXT().test(note)) d = "换药";
    else if (RE_ECON_TXT().test(note)) d = "经济";
    return { key: "dropout", detail: d };
  }
  // ③ 转渠道
  if (RE_CHANNEL_TXT().test(note)) return { key: "channel", detail: "" };
  // ④ 随访失败未探寻原因
  if (RE_FUFAIL_TXT().test(note) || (/执行失败|超期未完成/.test(ts) && /随访|购药/.test(fu.summary_type || ""))) {
    return { key: "fuFail", detail: "" };
  }
  // ⑤ 延长周期（医嘱 / 自行）
  if (RE_PROLONG_TXT().test(note)) {
    return { key: "prolong", detail: RE_PROLONG_DOCTOR_TXT().test(note) ? "医嘱" : "自行" };
  }
  // ⑥ 换药（未命中停药等脱落信号时仍归「脱落·换药」）
  if (RE_SWITCH_TXT().test(note)) return { key: "dropout", detail: "换药" };
  return null;
}

if (typeof window !== "undefined") {
  window.Mapping = {
    PRODUCT_FAMILIES, normalizeProduct, CATEGORIES, REPUR_PARTS, SUBSTATUS,
    TABLE_SIGNATURES, detectTableType, normHeader, _cell, _gtext,
    isPlaceholder, followupSignal, classifyFuReason, KEYWORD_RULES, KEYWORD_RULES_MULTI,
  };
}
})();
