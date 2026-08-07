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

function stableHash(value) {
  const bytes = new TextEncoder().encode(String(value));
  let hash = 1469598103934665603n;
  for (const byte of bytes) {
    hash ^= BigInt(byte);
    hash = BigInt.asUintN(64, hash * 1099511628211n);
  }
  return hash.toString(36).padStart(13, "0");
}

function stableUserId(user) {
  return `U${stableHash(`fml-dashboard-user-v2|${user}`)}`;
}

function orderIdentity(row, paidText) {
  const identityIndexes = [5, 8, 16, 19, 21, 30, 31, 32];
  return `O${stableHash(identityIndexes.map(index => index === 19 ? paidText : cleanText(row[index])).join("|"))}`;
}

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

function rebuildPayload(orders, storeStatus, mergeInfo = {}) {
  const sortedOrders = [...orders].sort((left, right) => left.dt.localeCompare(right.dt) || left.oid.localeCompare(right.oid));
  const byUser = new Map();
  sortedOrders.forEach(order => {
    if (!byUser.has(order.u)) byUser.set(order.u, []);
    byUser.get(order.u).push(order);
  });

  const cohorts = [];
  for (const userOrders of byUser.values()) {
    const first = userOrders.find(order => order.new);
    if (!first) continue;
    const next = userOrders.find(order => order.oid !== first.oid && (order.dt > first.dt || (order.dt === first.dt && order.oid > first.oid)));
    let days = null;
    if (next) {
      const firstDay = new Date(first.dt.slice(0, 10));
      const nextDay = new Date(next.dt.slice(0, 10));
      days = Math.floor((nextDay.getTime() - firstDay.getTime()) / 86400000);
    }
    cohorts.push({
      m: first.m, c: first.c, sid: first.sid, s: first.s, u: first.u, a: first.a,
      t: first.t, n: first.n, dt: first.dt, d: days, st: next?.t ?? null,
      p: next ? `${first.t}→${next.t}` : null,
    });
  }

  const dateMin = sortedOrders[0]?.dt ?? "";
  const dateMax = sortedOrders.at(-1)?.dt ?? "";
  const latest = dateMax ? dateMax.slice(0, 10) : "—";
  return {
    meta: {
      schemaVersion: 2,
      dateMin,
      dateMax,
      source: "线上历史数据与新增订单合并（匿名化、加密存储）",
      generatedAt: dateTimeText(new Date()).replace(" ", "T"),
      warning: `当前线上数据覆盖${dateMin ? dateMin.slice(0, 10) : "—"}至${latest}；经营指标仅统计非预售、主营、已支付且非体验订单。`,
      mergeInfo,
    },
    orders: sortedOrders,
    cohorts,
    storeStatus: { ...storeStatus, source: "营业门店表" },
  };
}

export function buildDashboardData(parsed, storeStatus, overrides = {}, currentData = null) {
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
    const paidText = dateTimeText(paidAt);
    records.push({
      sourceRow: offset + 2, paidAt, paidText, month: monthText(paidAt), oid: orderIdentity(row, paidText),
      user: cleanText(row[5]), city: cleanText(row[7]), storeId: cleanText(row[8]), store: cleanText(row[9]),
      productId: cleanText(row[16]), product: cleanText(row[17]), amount, subtype: cleanText(row[3]),
      duration: cleanText(row[27]), rights: cleanText(row[28]), days: asNumber(row[31]), times: asNumber(row[32]),
      key, autoTag, tag, basis, valid, isNew, nature: tag === "预售" ? "预售" : (isStoreMonthCard(row) ? "常规999" : "其他常规"),
    });
  });

  const pendingMap = new Map();
  const uploadLatestMonth = dateMax ? monthText(dateMax) : "";
  records.filter(row => row.valid && row.month === uploadLatestMonth && row.tag === "待确认").forEach(row => {
    const item = pendingMap.get(row.key) || { key: row.key, productId: row.productId, product: row.product, amount: row.amount, subtype: row.subtype, duration: row.duration, rights: row.rights, days: row.days, times: row.times, basis: row.basis, orders: 0 };
    item.orders += 1;
    pendingMap.set(row.key, item);
  });

  const incomingOrders = records.filter(row => row.valid).map(row => ({
    oid: row.oid, dt: row.paidText, m: row.month, c: row.city, sid: row.storeId, s: row.store,
    u: stableUserId(row.user), a: row.amount, t: row.tag, n: row.nature, new: row.isNew,
  }));
  const observedOrderIds = new Set(records.map(row => row.oid));
  const currentOrders = Array.isArray(currentData?.orders) ? currentData.orders : [];
  const currentIsIncremental = currentData?.meta?.schemaVersion === 2 && currentOrders.every(row => row.oid && row.dt);
  const currentDateMin = currentData?.meta?.dateMin || "";
  const uploadDateMin = dateMin ? dateTimeText(dateMin) : "";
  if (currentOrders.length && !currentIsIncremental) {
    const looksComplete = uploadDateMin && (!currentDateMin || uploadDateMin.slice(0, 10) <= currentDateMin.slice(0, 10)) && incomingOrders.length >= currentOrders.length * 0.8;
    if (!looksComplete) throw new Error("首次启用增量更新需要上传一份完整历史订单文件，完成数据结构迁移后即可只上传新增数据。");
  }

  const mergedMap = new Map();
  if (currentIsIncremental) currentOrders.forEach(order => mergedMap.set(order.oid, order));
  const currentIds = new Set(currentOrders.map(order => order.oid));
  const incomingIds = new Set(incomingOrders.map(order => order.oid));
  const replacedOrders = currentIsIncremental ? incomingOrders.filter(order => currentIds.has(order.oid)).length : 0;
  const addedOrders = currentIsIncremental ? incomingOrders.filter(order => !currentIds.has(order.oid)).length : incomingOrders.length;
  const removedOrders = currentIsIncremental ? Array.from(observedOrderIds).filter(oid => currentIds.has(oid) && !incomingIds.has(oid)).length : 0;
  observedOrderIds.forEach(oid => mergedMap.delete(oid));
  incomingOrders.forEach(order => mergedMap.set(order.oid, order));
  const mergedOrders = Array.from(mergedMap.values());
  const mergeMode = currentOrders.length && !currentIsIncremental ? "首次完整迁移" : currentIsIncremental ? "线上增量合并" : "首次建立";
  const payload = rebuildPayload(mergedOrders, storeStatus, {
    mode: mergeMode,
    uploadRows: parsed.rows.length,
    uploadValidOrders: incomingOrders.length,
    previousOrders: currentIsIncremental ? currentOrders.length : 0,
    addedOrders,
    replacedOrders,
    removedOrders,
  });

  const valid = payload.orders;
  const newUsers = new Set(payload.cohorts.map(row => row.u));
  const latestMonth = payload.meta.dateMax.slice(0, 7);
  const summary = {
    sourceRows: parsed.rows.length, processedRows: records.length, validOrders: valid.length,
    uploadValidOrders: incomingOrders.length, addedOrders, replacedOrders, removedOrders, mergeMode,
    gmv: valid.reduce((sum, row) => sum + row.a, 0), newUsers: newUsers.size,
    pendingOrders: records.filter(row => row.valid && row.month === uploadLatestMonth && row.tag === "待确认").length,
    pendingRules: pendingMap.size, dateMin: payload.meta.dateMin, dateMax: payload.meta.dateMax,
    oneVOne: valid.filter(row => row.t === "1V1").length, oneVTwo: valid.filter(row => row.t === "1V2").length,
    special999: payload.cohorts.filter(row => row.m === latestMonth && row.n === "常规999").length,
  };
  return { payload, summary, pending: Array.from(pendingMap.values()) };
}

function bytesToBase64(bytes) {
  let binary = "";
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  return btoa(binary);
}

function base64ToBytes(value) {
  const binary = atob(value);
  return Uint8Array.from(binary, character => character.charCodeAt(0));
}

export async function encryptDashboardData(payload, password) {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const iv = crypto.getRandomValues(new Uint8Array(12));
  const iterations = 250000;
  const compressed = await new Response(new Blob([JSON.stringify(payload)]).stream().pipeThrough(new CompressionStream("gzip"))).arrayBuffer();
  const material = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
  const key = await crypto.subtle.deriveKey({ name: "PBKDF2", salt, iterations, hash: "SHA-256" }, material, { name: "AES-GCM", length: 256 }, false, ["encrypt"]);
  const encrypted = await crypto.subtle.encrypt({ name: "AES-GCM", iv }, key, compressed);
  return { v: 1, iterations, salt: bytesToBase64(salt), iv: bytesToBase64(iv), compression: "gzip", data: bytesToBase64(new Uint8Array(encrypted)) };
}

export async function decryptDashboardData(envelope, password) {
  const material = await crypto.subtle.importKey("raw", new TextEncoder().encode(password), "PBKDF2", false, ["deriveKey"]);
  const key = await crypto.subtle.deriveKey({ name: "PBKDF2", salt: base64ToBytes(envelope.salt), iterations: envelope.iterations, hash: "SHA-256" }, material, { name: "AES-GCM", length: 256 }, false, ["decrypt"]);
  const decrypted = await crypto.subtle.decrypt({ name: "AES-GCM", iv: base64ToBytes(envelope.iv) }, key, base64ToBytes(envelope.data));
  return JSON.parse(await new Response(new Blob([decrypted]).stream().pipeThrough(new DecompressionStream("gzip"))).text());
}
