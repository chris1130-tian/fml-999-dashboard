import { parseOrderWorkbook, buildDashboardData, encryptDashboardData, decryptDashboardData } from "./admin-core.js";

const $ = (id) => document.getElementById(id);
const number = new Intl.NumberFormat("zh-CN");
const money = new Intl.NumberFormat("zh-CN", { style: "currency", currency: "CNY", maximumFractionDigits: 0 });
const state = { parsed: null, result: null, storeStatus: null, overrides: JSON.parse(localStorage.getItem("fml-product-overrides") || "{}") };

function setStatus(element, message, type = "") {
  element.textContent = message;
  element.className = `status ${type}`.trim();
}

function summaryCard(label, value, warn = false) {
  return `<article class="summary-card ${warn ? "warn" : ""}"><span>${label}</span><strong>${value}</strong></article>`;
}

function renderResult() {
  const { summary, pending } = state.result;
  $("resultSection").classList.remove("hidden");
  $("publishSection").classList.remove("hidden");
  $("summaryCards").innerHTML = [
    summaryCard("数据范围", `${summary.dateMin.slice(0, 10)} 至 ${summary.dateMax.slice(0, 10)}`),
    summaryCard("有效私教订单", number.format(summary.validOrders)),
    summaryCard("有效私教GMV", money.format(summary.gmv)),
    summaryCard("拉新用户", number.format(summary.newUsers)),
    summaryCard("1V1订单", number.format(summary.oneVOne)),
    summaryCard("1V2订单", number.format(summary.oneVTwo)),
    summaryCard("8月999专项拉新", number.format(summary.special999)),
    summaryCard("8月待确认订单", number.format(summary.pendingOrders), summary.pendingOrders > 0),
  ].join("");

  const hasPending = pending.length > 0;
  $("pendingBlock").classList.toggle("hidden", !hasPending);
  $("readyBlock").classList.toggle("hidden", hasPending);
  $("pendingCount").textContent = `${summary.pendingRules}种商品 · ${summary.pendingOrders}笔订单`;
  $("pendingRows").innerHTML = pending.map((row, index) => `<tr>
    <td><strong title="${escapeHtml(row.product)}">${escapeHtml(row.product)}</strong><small>${escapeHtml(row.productId)}</small></td>
    <td>${money.format(row.amount)}</td><td>${escapeHtml(`${row.duration || "—"} · ${row.days || "—"}天 · ${row.times || "—"}次`)}</td>
    <td>${number.format(row.orders)}</td><td>${escapeHtml(row.basis)}</td>
    <td><select data-pending-index="${index}"><option value="">请选择</option><option value="1V1">1V1</option><option value="1V2">1V2</option></select></td>
  </tr>`).join("");
  document.querySelectorAll("[data-pending-index]").forEach(select => select.addEventListener("change", () => {
    if (!select.value) return;
    state.overrides[pending[Number(select.dataset.pendingIndex)].key] = select.value;
    localStorage.setItem("fml-product-overrides", JSON.stringify(state.overrides));
    recalculate();
  }));
  $("downloadButton").disabled = hasPending;
  $("publishButton").disabled = hasPending;
}

function escapeHtml(value) {
  return String(value ?? "").replace(/[&<>'"]/g, character => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;" })[character]);
}

function recalculate() {
  state.result = buildDashboardData(state.parsed, state.storeStatus, state.overrides);
  renderResult();
}

async function handleFile(file) {
  if (!file) return;
  if (!state.storeStatus) {
    setStatus($("parseStatus"), "请先输入当前看板访问密码并完成验证。", "error");
    return;
  }
  $("fileLabel").textContent = file.name;
  setStatus($("parseStatus"), "正在读取并校验Excel，请稍候…");
  $("resultSection").classList.add("hidden");
  $("publishSection").classList.add("hidden");
  try {
    state.parsed = await parseOrderWorkbook(file);
    recalculate();
    setStatus($("parseStatus"), `已读取 ${number.format(state.parsed.rows.length)} 行数据，列结构校验通过。`, "success");
  } catch (error) {
    state.parsed = null;
    state.result = null;
    setStatus($("parseStatus"), error.message || "文件读取失败，请确认格式后重试。", "error");
  }
}

function utf8Base64(text) {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  const chunk = 0x8000;
  for (let index = 0; index < bytes.length; index += chunk) binary += String.fromCharCode(...bytes.subarray(index, index + chunk));
  return btoa(binary);
}

async function githubRequest(url, token, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: { Accept: "application/vnd.github+json", Authorization: `Bearer ${token}`, "X-GitHub-Api-Version": "2022-11-28", "Content-Type": "application/json", ...(options.headers || {}) },
  });
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.message || `GitHub请求失败（${response.status}）`);
  }
  return response.json();
}

async function publish() {
  const owner = $("ownerInput").value.trim();
  const repo = $("repoInput").value.trim();
  const branch = $("branchInput").value.trim() || "main";
  const token = $("tokenInput").value.trim();
  const dataPassword = $("dataPasswordInput").value;
  if (!owner || !repo || !token) return setStatus($("publishStatus"), "请填写GitHub用户名、仓库名称和令牌。", "error");
  if (!dataPassword) return setStatus($("publishStatus"), "请填写当前看板访问密码。", "error");
  if (!state.result || state.result.pending.length) return setStatus($("publishStatus"), "请先上传数据并完成全部待确认订单。", "error");
  localStorage.setItem("fml-github-owner", owner);
  localStorage.setItem("fml-github-repo", repo);
  localStorage.setItem("fml-github-branch", branch);
  $("publishButton").disabled = true;
  setStatus($("publishStatus"), "正在连接GitHub并更新看板数据…");
  try {
    const path = "docs/dashboard-data.enc.json";
    const base = `https://api.github.com/repos/${encodeURIComponent(owner)}/${encodeURIComponent(repo)}/contents/${path}`;
    let sha = null;
    const existingResponse = await fetch(`${base}?ref=${encodeURIComponent(branch)}`, { headers: { Accept: "application/vnd.github+json", Authorization: `Bearer ${token}`, "X-GitHub-Api-Version": "2022-11-28" } });
    if (existingResponse.ok) sha = (await existingResponse.json()).sha;
    else if (existingResponse.status !== 404) throw new Error((await existingResponse.json().catch(() => ({}))).message || `无法读取现有数据文件（${existingResponse.status}）`);
    const encryptedPayload = await encryptDashboardData(state.result.payload, dataPassword);
    const body = { message: `更新加密经营看板数据至 ${state.result.summary.dateMax.slice(0, 10)}`, content: utf8Base64(JSON.stringify(encryptedPayload)), branch };
    if (sha) body.sha = sha;
    await githubRequest(base, token, { method: "PUT", body: JSON.stringify(body) });
    $("tokenInput").value = "";
    setStatus($("publishStatus"), "发布成功。GitHub Pages通常会在1—3分钟内完成更新，请稍后刷新公开看板。", "success");
  } catch (error) {
    setStatus($("publishStatus"), `${error.message || "发布失败"}。请检查令牌是否仅授权当前仓库，并拥有Contents读写权限。`, "error");
  } finally {
    $("publishButton").disabled = Boolean(state.result?.pending.length);
  }
}

function downloadData() {
  if (!state.result || state.result.pending.length) return;
  const password = $("dataPasswordInput").value;
  if (!password) return setStatus($("parseStatus"), "请先输入当前看板访问密码。", "error");
  encryptDashboardData(state.result.payload, password).then(envelope => {
    const blob = new Blob([JSON.stringify(envelope)], { type: "application/json" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `dashboard-data-${state.result.summary.dateMax.slice(0, 10)}.enc.json`;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 1000);
  });
}

async function unlockCurrentData() {
  const password = $("dataPasswordInput").value;
  if (!password) return setStatus($("unlockStatus"), "请输入当前看板访问密码。", "error");
  $("unlockDataButton").disabled = true;
  setStatus($("unlockStatus"), "正在验证密码并读取营业门店表…");
  try {
    const envelope = await fetch("./dashboard-data.enc.json", { cache: "no-store" }).then(response => {
      if (!response.ok) throw new Error("加密数据文件加载失败");
      return response.json();
    });
    const current = await decryptDashboardData(envelope, password);
    state.storeStatus = current.storeStatus;
    setStatus($("unlockStatus"), `密码验证成功，已读取${current.storeStatus.stores.filter(row => row.operating).length}家营业中门店。`, "success");
  } catch {
    state.storeStatus = null;
    setStatus($("unlockStatus"), "密码不正确，请重新输入。", "error");
  } finally {
    $("unlockDataButton").disabled = false;
  }
}

async function initialize() {
  $("ownerInput").value = localStorage.getItem("fml-github-owner") || "";
  $("repoInput").value = localStorage.getItem("fml-github-repo") || "fml-999-dashboard";
  $("branchInput").value = localStorage.getItem("fml-github-branch") || "main";
}

$("fileInput").addEventListener("change", event => handleFile(event.target.files[0]));
$("unlockDataButton").addEventListener("click", unlockCurrentData);
$("downloadButton").addEventListener("click", downloadData);
$("publishButton").addEventListener("click", publish);
const dropZone = $("dropZone");
["dragenter", "dragover"].forEach(name => dropZone.addEventListener(name, event => { event.preventDefault(); dropZone.classList.add("dragging"); }));
["dragleave", "drop"].forEach(name => dropZone.addEventListener(name, event => { event.preventDefault(); dropZone.classList.remove("dragging"); }));
dropZone.addEventListener("drop", event => handleFile(event.dataTransfer.files[0]));
initialize();
