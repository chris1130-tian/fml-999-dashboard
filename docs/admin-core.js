const REQUIRED_COLUMNS = new Map([
  [3, "业务子类型"], [4, "订单类型"], [5, "用户ID"], [7, "归属城市"],
  [8, "归属门店ID"], [9, "归属门店名称"], [16, "商品编号"], [17, "商品名称"],
  [19, "支付时间"], [20, "支付状态"], [21, "子订单实际支付金额"],
  [22, "是否线下主营业务付费订单"], [27, "时长类型"], [28, "权益类型"],
  [31, "权益天数"], [32, "权益次数"], [33, "拉新/复购"],
]);

const cleanText = (value) => {
  if (value === null || value === undefined) return "";
  const result = String(value).trim();
  return ["(null)", "(空字符串)", "None"].includes(result) ? "" : result;
};

const asNumber = (value) => {
  const result = Number(value ?? 0);
  return Number.isFinite(result) ? result : 0;
};

const pad = (value) => String(value).padStart(2, "0");

const asDate = (value) => {
  if (value instanceof Date && Number.isFinite(value.getTime())) return value;
  if (typeof value === "number" && globalThis.XLSX?.SSF) {
    const parsed = globalThis.XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, Math.floor(parsed.S));
  }
  const text = cleanText(value).replace(/\//g, "-");
  if (!text) return null;
  const normalized = text.includes("T") ? text : text.replace(" ", "T");
  const result = new Date(normalized);
  return Number.isFinite(result.getTime()) ? result : null;
};

const dateTimeText = (value) => `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())} ${pad(value.getHours())}:${pad(value.getMinutes())}:${pad(value.getSeconds())}`;
const monthText = (value) => `${value.getFullYear()}-${pad(value.getMonth() + 1)}`;
const dateOnlyText = (value) => `${value.getFullYear()}-${pad(value.getMonth() + 1)}-${pad(value.getDate())}`;
const displayNumber = (value) => Number.isInteger(value) ? String(value) : String(value);

function isLimited(row) {
  return [3, 4, 28].map(index => cleanText(row[index]).toLowerCase()).join("|").includes("限次");
}

function isStoreMonthCard(row) {
  const product = cleanText(row[17]).toLowerCase().replaceAll("ｖ", "v").replaceAll("－", "-");
  if (product.includes("预售") || product.includes("通卡") || cleanText(row[27]) !== "月卡") return false;
  return /包月私教(?:1v[12])?会员卡月卡/.test(product) || /私教会员-?单店月卡/.test(product);
}

function automaticTag(row) {
  const product = cleanText(row[17]);
  const productLower = product.toLowerCase().replaceAll("ｖ", "v");
  const amount = asNumber(row[21]);
  const days = asNumber(row[31]);
  const times = asNumber(row[32]);
  const paidAt = asDate(row[19]);

  if (product.includes("预售")) return ["预售", "商品名称含“预售”，独立核算"];
  if (cleanText(row[16]) === "011109621" && product.includes("91天卡")) return ["1V1", "业务确认：F2包月私教通卡91天卡按1V1"];
  if (productLower.includes("1v1") || productLower.includes("1对1") || product.includes("一对一")) return ["1V1", "商品名称明确1V1"];
  if (productLower.includes("1v2") || productLower.includes("1对2") || product.includes("一对二")) return ["1V2", "商品名称明确1V2"];
  if (days === 21 && times === 9) return ["1V1", "21天9节课按1V1"];
  if (product.includes("体验") || product.includes("赠送") || amount === 0) return ["体验/赠送", "商品名称明确为体验/赠送，或实付为0"];
  if (isLimited(row) && paidAt && paidAt < new Date(2026, 6, 21)) return ["1V2", "2026-07-21前限次课程（21天9节除外）按1V2"];

  const monthlyPrice = days > 0 ? amount * 30 / days : null;
  if (monthlyPrice !== null && monthlyPrice <= 2000) return ["1V2", `折算月均价${Math.round(monthlyPrice)}元，落入1V2近似区间`];
  if (monthlyPrice !== null && monthlyPrice >= 2200) return ["1V1", `折算月均价${Math.round(monthlyPrice)}元，落入1V1近似区间`];
  if (monthlyPrice !== null) return ["待确认", `折算月均价${Math.round(monthlyPrice)}元，位于待确认区间`];
  return ["待确认", "商品名称与价格均无法直接判定"];
}

function ruleKey(row) {
  return [cleanText(row[16]), cleanText(row[17]), displayNumber(asNumber(row[21])), cleanText(row[3]), cleanText(row[27]), cleanText(row[28]), displayNumber(asNumber(row[31])), displayNumber(asNumber(row[32]))].join("|");
}

export async function parseOrderWorkbook(file) {
  if (!globalThis.XLSX) throw new Error("Excel解析组件尚未加载，请刷新页面后重试。");
  const buffer = await file.arrayBuffer();
  const workbook = globalThis.XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = globalThis.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null });
  if (!rows.length) throw new Error("Excel中没有可读取的数据。");
  const headers = rows[0].map(cleanText);
  const mismatches = [];
  for (const [index, expected] of REQUIRED_COLUMNS) if (headers[index] !== expected) mismatches.push(`第${index + 1}列应为“${expected}”`);
  if (mismatches.length) throw new Error(`文件格式与“截止26年8月数据表”不一致：${mismatches.join("；")}`);
  return { fileName: file.name, headers, rows: rows.slice(1), sheetName: workbook.SheetNames[0] };
}

export function buildDashboardData(parsed, storeStatus, overrides = {}) {
  const records = [];
  let dateMin = null;
  let dateMax = null;

  parsed.rows.forEach((row, offset) => {
    const paidAt = asDate(row[19]);
    if (!paidAt) return;
    if (!dateMin || paidAt < dateMin) dateMin = paidAt;
    if (!dateMax || paidAt > dateMax) dateMax = paidAt;
    const amount = asNumber(row[21]);
    const [autoTag, basis] = automaticTag(row);
    const key = ruleKey(row);
    const tag = overrides[key] || autoTag;
    const isPaid = cleanText(row[20]) === "已支付";
    const isMain = asNumber(row[22]) === 1;
    const isPresale = tag === "预售" || cleanText(row[17]).includes("预售");
    const isExperience = cleanText(row[17]).includes("体验") || cleanText(row[17]).includes("赠送");
    const valid = Boolean(isPaid && isMain && amount > 0 && !isPresale && !isExperience);
    const isNew = Boolean(valid && cleanText(row[33]) === "同商户拉新订单");
    records.push({
      sourceRow: offset + 2, paidAt, paidText: dateTimeText(paidAt), month: monthText(paidAt),
      user: cleanText(row[5]), city: cleanText(row[7]), storeId: cleanText(row[8]), store: cleanText(row[9]),
      productId: cleanText(row[16]), product: cleanText(row[17]), amount, subtype: cleanText(row[3]),
      duration: cleanText(row[27]), rights: cleanText(row[28]), days: asNumber(row[31]), times: asNumber(row[32]),
      key, autoTag, tag, basis, valid, isNew, nature: tag === "预售" ? "预售" : (isStoreMonthCard(row) ? "常规999" : "其他常规"),
      second: null,
    });
  });

  const validByUser = new Map();
  records.filter(row => row.valid).forEach(row => {
    if (!validByUser.has(row.user)) validByUser.set(row.user, []);
    validByUser.get(row.user).push(row);
  });
  for (const rows of validByUser.values()) rows.sort((a, b) => a.paidAt - b.paidAt || a.sourceRow - b.sourceRow);
  records.filter(row => row.isNew).forEach(row => {
    const next = validByUser.get(row.user)?.find(candidate => candidate.paidAt > row.paidAt || (candidate.paidAt.getTime() === row.paidAt.getTime() && candidate.sourceRow > row.sourceRow));
    if (next) row.second = { paidText: next.paidText, days: Math.floor((new Date(next.paidAt.getFullYear(), next.paidAt.getMonth(), next.paidAt.getDate()) - new Date(row.paidAt.getFullYear(), row.paidAt.getMonth(), row.paidAt.getDate())) / 86400000), tag: next.tag, product: next.product, path: `${row.tag}→${next.tag}` };
  });

  const pendingMap = new Map();
  records.filter(row => row.valid && row.month === "2026-08" && row.tag === "待确认").forEach(row => {
    const item = pendingMap.get(row.key) || { key: row.key, productId: row.productId, product: row.product, amount: row.amount, subtype: row.subtype, duration: row.duration, rights: row.rights, days: row.days, times: row.times, basis: row.basis, orders: 0 };
    item.orders += 1;
    pendingMap.set(row.key, item);
  });

  const rawUsers = Array.from(new Set(records.filter(row => row.valid).map(row => row.user))).sort((a, b) => a.localeCompare(b, "zh-CN", { numeric: true }));
  const anonymousUsers = new Map(rawUsers.map((user, index) => [user, `U${String(index + 1).padStart(7, "0")}`]));
  const orders = records.filter(row => row.valid).map(row => ({ m: row.month, c: row.city, sid: row.storeId, s: row.store, u: anonymousUsers.get(row.user), a: row.amount, t: row.tag, n: row.nature, new: row.isNew }));

  const firstByUser = new Map();
  records.filter(row => row.isNew).forEach(row => {
    const current = firstByUser.get(row.user);
    if (!current || row.paidAt < current.paidAt || (row.paidAt.getTime() === current.paidAt.getTime() && row.sourceRow < current.sourceRow)) firstByUser.set(row.user, row);
  });
  const cohorts = Array.from(firstByUser.values()).map(row => ({
    m: row.month, c: row.city, sid: row.storeId, s: row.store, u: anonymousUsers.get(row.user), a: row.amount,
    t: row.tag, n: row.nature, dt: row.paidText, d: row.second?.days ?? null, st: row.second?.tag ?? null, p: row.second?.path ?? null,
  }));

  const latest = dateMax ? dateOnlyText(dateMax) : "—";
  const normalizedStoreStatus = { ...storeStatus, source: "营业门店表" };
  const payload = {
    meta: {
      dateMin: dateMin ? dateTimeText(dateMin) : "", dateMax: dateMax ? dateTimeText(dateMax) : "",
      source: "订单数据（浏览器本地处理并匿名化）", generatedAt: dateTimeText(new Date()).replace(" ", "T"),
      warning: `当前源文件覆盖${dateMin ? dateOnlyText(dateMin) : "—"}至${latest}；经营指标仅统计非预售、主营、已支付且非体验订单，2026年8月为截至${latest}的月内数据。`,
    },
    orders, cohorts, storeStatus: normalizedStoreStatus,
  };

  const valid = records.filter(row => row.valid);
  const newUsers = new Set(records.filter(row => row.isNew).map(row => row.user));
  const summary = {
    sourceRows: parsed.rows.length, processedRows: records.length, validOrders: valid.length,
    gmv: valid.reduce((sum, row) => sum + row.amount, 0), newUsers: newUsers.size,
    pendingOrders: valid.filter(row => row.month === "2026-08" && row.tag === "待确认").length,
    pendingRules: pendingMap.size, dateMin: payload.meta.dateMin, dateMax: payload.meta.dateMax,
    oneVOne: valid.filter(row => row.tag === "1V1").length, oneVTwo: valid.filter(row => row.tag === "1V2").length,
    special999: cohorts.filter(row => row.m === "2026-08" && row.n === "常规999").length,
  };
  return { payload, summary, pending: Array.from(pendingMap.values()) };
}
