 const templateSelectEl = document.getElementById("templateSelect");
const purchaseSelectEl = document.getElementById("purchaseSelect");
const templateSheetEl = document.getElementById("templateSheet");
const purchaseSheetEl = document.getElementById("purchaseSheet");
const templateInfoEl = document.getElementById("templateInfo");
const purchaseInfoEl = document.getElementById("purchaseInfo");
const mergeBtnEl = document.getElementById("mergeBtn");
const downloadBtnEl = document.getElementById("downloadBtn");
const saveBtnEl = document.getElementById("saveBtn");
const saveNameInputEl = document.getElementById("saveName");
const statusEl = document.getElementById("status");
const previewEl = document.getElementById("preview");
const contractsTbodyEl = document.getElementById("contractsTbody");
const previewModalEl = document.getElementById("previewModal");
const modalTitleEl = document.getElementById("modalTitle");
const modalSheetSelectEl = document.getElementById("modalSheetSelect");
const modalCloseEl = document.getElementById("modalClose");
const modalPreviewEl = document.getElementById("modalPreview");

const state = {
  template: null,
  purchase: null,
  merged: null,
  mergedSheetName: null,
  hf: null, // HyperFormula 实例
  hfSheetId: null,
};

function normalizeApiBase(base) {
  const s = String(base ?? "").trim().replace(/\/+$/, "");
  return s;
}

function getApiBase() {
  const configured = normalizeApiBase(globalThis.API_BASE);
  if (configured) return configured;
  if (location.protocol === "file:") return "http://localhost:3000";
  return "";
}

function buildApiUrl(pathname) {
  const base = getApiBase();
  if (!base) return pathname;
  return `${base}${pathname}`;
}

async function postBinaryWithFallback(pathname, buffer, contentType = "application/octet-stream") {
  const primary = buildApiUrl(pathname);
  try {
    const res = await fetch(primary, {
      method: "POST",
      headers: { "Content-Type": contentType },
      body: buffer,
    });
    if (res.ok && !isHtmlLikeResponse(res)) return res;
    if (shouldTryLocalApiFallback()) {
      const fb = await fetch(`http://localhost:3000${pathname}`, {
        method: "POST",
        headers: { "Content-Type": contentType },
        body: buffer,
      });
      return fb;
    }
    return res;
  } catch (e) {
    if (!shouldTryLocalApiFallback()) throw e;
    return await fetch(`http://localhost:3000${pathname}`, {
      method: "POST",
      headers: { "Content-Type": contentType },
      body: buffer,
    });
  }
}

function shouldTryLocalApiFallback() {
  return !globalThis.API_BASE && location.origin !== "http://localhost:3000";
}

function isHtmlLikeResponse(res) {
  const ct = String(res?.headers?.get?.("content-type") ?? "").toLowerCase();
  return ct.includes("text/html");
}

async function fetchJsonWithFallback(pathname) {
  const primary = buildApiUrl(pathname);
  let res = null;
  let json = null;

  const tryParseJson = async (response) => {
    if (!response) return null;
    const ct = String(response.headers?.get?.("content-type") ?? "").toLowerCase();
    if (!ct.includes("application/json")) return null;
    return response.json().catch(() => null);
  };

  try {
    res = await fetch(primary);
    json = await tryParseJson(res);
    if (shouldTryLocalApiFallback() && (!res.ok || json === null || isHtmlLikeResponse(res))) {
      const fallback = `http://localhost:3000${pathname}`;
      const fbRes = await fetch(fallback);
      const fbJson = await tryParseJson(fbRes);
      if (fbRes.ok && fbJson !== null && !isHtmlLikeResponse(fbRes)) {
        return { res: fbRes, json: fbJson, url: fallback };
      }
    }
    return { res, json, url: primary };
  } catch (e) {
    if (!shouldTryLocalApiFallback()) throw e;
    const fallback = `http://localhost:3000${pathname}`;
    res = await fetch(fallback);
    json = await tryParseJson(res);
    return { res, json, url: fallback };
  }
}

async function fetchWithLocalFallback(pathname) {
  let response = await fetch(buildApiUrl(pathname));
  if (!shouldTryLocalApiFallback()) return response;

  if (!response.ok || isHtmlLikeResponse(response)) {
    try {
      const fallback = await fetch(`http://localhost:3000${pathname}`);
      if (fallback.ok && !isHtmlLikeResponse(fallback)) return fallback;
      if (!response.ok) return fallback;
    } catch (e) {}
  }

  return response;
}

async function fetchWithLocalFallbackRequest(pathname, options) {
  let response = await fetch(buildApiUrl(pathname), options);
  if (!shouldTryLocalApiFallback()) return response;

  if (!response.ok || isHtmlLikeResponse(response)) {
    try {
      const fallback = await fetch(`http://localhost:3000${pathname}`, options);
      if (fallback.ok && !isHtmlLikeResponse(fallback)) return fallback;
      if (!response.ok) return fallback;
    } catch (e) {}
  }

  return response;
}

function guessExtFromFileType(fileType) {
  const t = String(fileType || "").trim().toLowerCase();
  if (!t) return ".xlsx";
  if (t.includes("xlsx")) return ".xlsx";
  if (t.includes("xls")) return ".xls";
  if (t.includes("docx")) return ".docx";
  if (t.includes("doc")) return ".doc";
  if (t.includes("pdf")) return ".pdf";
  return ".bin";
}

async function fetchTemplates() {
  try {
    setStatus("正在加载模板列表…");
    let res = null;
    let json = null;
    try {
      const out = await fetchJsonWithFallback("/api/templates");
      res = out.res;
      json = out.json;
    } catch (e) {
      throw e;
    }
    if (!res.ok || !json?.success) {
      const msg =
        json?.error ||
        json?.message ||
        `${res.status} ${res.statusText}` ||
        "加载模板失败";
      throw new Error(msg);
    }
    const templates = Array.isArray(json.data) ? json.data : [];

    templateSelectEl.innerHTML = '<option value="">-- 请选择合同模板 --</option>';
    templates.forEach((t) => {
      const opt = document.createElement("option");
      opt.value = String(t.id ?? "");
      opt.textContent = t.name;
      if (t.file_type) opt.dataset.fileType = String(t.file_type);
      templateSelectEl.appendChild(opt);
    });
    setStatus("");
    templateInfoEl.textContent = "";
  } catch (e) {
    const msg = e?.message ?? String(e);
    console.error("加载模板失败：", e);
    templateSelectEl.innerHTML = '<option value="">加载模板失败</option>';
    templateInfoEl.textContent = `加载失败：${msg}`;
    setStatus("模板列表加载失败：请确认已启动后端服务（默认 http://localhost:3000）");
  }
}

async function fetchPurchaseOrders() {
  if (!purchaseSelectEl) return;
  try {
    setStatus("正在加载订购单列表…");
    const out = await fetchJsonWithFallback("/api/purchase-orders");
    const res = out.res;
    const json = out.json;
    if (!res.ok || !json?.success) {
      const msg =
        json?.error ||
        json?.message ||
        `${res.status} ${res.statusText}` ||
        "加载订购单失败";
      throw new Error(msg);
    }

    const list = Array.isArray(json.data) ? json.data : [];
    purchaseSelectEl.innerHTML = '<option value="">-- 请选择订购单 --</option>';
    list.forEach((p) => {
      const opt = document.createElement("option");
      opt.value = String(p.id ?? "");
      opt.textContent = p.name ?? `订购单_${p.id ?? ""}`;
      if (p.file_path) opt.dataset.filePath = String(p.file_path);
      if (p.type_id) opt.dataset.typeId = String(p.type_id);
      purchaseSelectEl.appendChild(opt);
    });
    setStatus("");
    purchaseInfoEl.textContent = "";
  } catch (e) {
    const msg = e?.message ?? String(e);
    console.error("加载订购单失败：", e);
    purchaseSelectEl.innerHTML = '<option value="">加载订购单失败</option>';
    purchaseInfoEl.textContent = `加载失败：${msg}`;
    setStatus("订购单列表加载失败：请确认后端与数据库已就绪");
  }
}

async function handlePurchaseSelectChange() {
  if (!purchaseSelectEl) return;
  const purchaseId = purchaseSelectEl.value;
  state.purchase = null;
  state.merged = state.template?.workbook || null;
  state.mergedSheetName = templateSheetEl.value;
  downloadBtnEl.disabled = true;
  if (saveBtnEl) saveBtnEl.disabled = true;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = true;
    saveNameInputEl.value = "";
  }
  if (state.merged && state.mergedSheetName) {
    try {
      await ensureHyperFormula();
      initFormulaEngine(state.merged.getWorksheet(state.mergedSheetName));
    } catch (e) {}
    renderPreview(state.merged, state.mergedSheetName);
  } else {
    previewEl.innerHTML = "";
  }

  if (!purchaseId) {
    purchaseInfoEl.textContent = "";
    purchaseSheetEl.innerHTML = "";
    purchaseSheetEl.disabled = true;
    enableActionsIfReady();
    return;
  }

  try {
    setStatus("正在从服务器下载订购单…");

    const pathname = `/api/purchase-orders/${encodeURIComponent(purchaseId)}/download`;
    const response = await fetchWithLocalFallback(pathname);

    if (!response.ok) {
      const text = await response.text().catch(() => "");
      const parsed = (() => {
        try {
          return JSON.parse(text);
        } catch (e) {
          return null;
        }
      })();
      const msg = parsed?.error || parsed?.message || text || `${response.status} ${response.statusText}`;
      throw new Error(`无法从服务器获取订购单文件：${msg}`);
    }

    const buffer = await response.arrayBuffer();
    const loaded = await loadWorkbookFromArrayBuffer(buffer);
    const opt = purchaseSelectEl.options[purchaseSelectEl.selectedIndex];
    const fileName = `${opt?.text || "订购单"}.xlsx`;
    const typeId = String(opt?.dataset?.typeId ?? "");

    state.purchase = { id: purchaseId, typeId, file: { name: fileName }, workbook: loaded, buffer };
    fillSheetSelect(purchaseSheetEl, loaded);
    purchaseSheetEl.value = getFirstSheetName(loaded);
    updateFileMeta(purchaseInfoEl, { name: fileName }, loaded);
    setStatus("订购单已就绪");

    if (state.merged && state.mergedSheetName) {
      try {
        await ensureHyperFormula();
        initFormulaEngine(state.merged.getWorksheet(state.mergedSheetName));
      } catch (e) {}
      renderPreview(state.merged, state.mergedSheetName);
    }

    enableActionsIfReady();
  } catch (e) {
    setStatus("");
    purchaseInfoEl.textContent = `读取失败：${e?.message ?? e}`;
    purchaseSheetEl.disabled = true;
    enableActionsIfReady();
  }
}

async function handleTemplateSelectChange() {
  const templateId = templateSelectEl.value;
  state.template = null;
  state.merged = null;
  downloadBtnEl.disabled = true;
  if (saveBtnEl) saveBtnEl.disabled = true;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = true;
    saveNameInputEl.value = "";
  }
  previewEl.innerHTML = "";

  if (!templateId) {
    templateInfoEl.textContent = "";
    templateSheetEl.innerHTML = "";
    templateSheetEl.disabled = true;
    enableActionsIfReady();
    return;
  }

  try {
    setStatus("正在从服务器下载合同模板…");

    const pathname = `/api/templates/${encodeURIComponent(templateId)}/download`;
    const response = await fetchWithLocalFallback(pathname);
    if (!response.ok) {
      const text = await response.text().catch(() => "");
      const parsed = (() => {
        try {
          return JSON.parse(text);
        } catch (e) {
          return null;
        }
      })();
      const msg = parsed?.error || parsed?.message || text || `${response.status} ${response.statusText}`;
      throw new Error(`无法从服务器获取模板文件：${msg}`);
    }
    const buffer = await response.arrayBuffer();
    
    const loaded = await loadWorkbookFromArrayBuffer(buffer);
    const opt = templateSelectEl.options[templateSelectEl.selectedIndex];
    const ext = guessExtFromFileType(opt?.dataset?.fileType);
    const fileName = `${opt?.text || "合同模板"}${ext}`;
    
    state.template = { 
      id: templateId,
      file: { name: fileName }, 
      workbook: loaded, 
      buffer 
    };
    
    fillSheetSelect(templateSheetEl, loaded);
    templateSheetEl.value = getFirstSheetName(loaded);
    updateFileMeta(templateInfoEl, { name: fileName }, loaded);
    setStatus("合同模板已就绪");
    
    // 立即预览模板
    state.merged = loaded;
    state.mergedSheetName = templateSheetEl.value;
    try {
      await ensureHyperFormula();
      initFormulaEngine(state.merged.getWorksheet(state.mergedSheetName));
    } catch (e) {}
    renderPreview(state.merged, state.mergedSheetName);

    enableActionsIfReady();
  } catch (e) {
    setStatus("");
    templateInfoEl.textContent = `读取失败：${e?.message ?? e}`;
    templateSheetEl.disabled = true;
    enableActionsIfReady();
  }
}

function setStatus(message) {
  statusEl.textContent = message ?? "";
}

function normalizeText(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim().replace(/\s+/g, " ");
}

function getFirstSheetName(workbook) {
  return workbook.worksheets?.[0]?.name ?? "";
}

function encodeCol(colIndex1) {
  let n = colIndex1;
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function decodeCol(colLetters) {
  const s = String(colLetters || "").toUpperCase();
  let n = 0;
  for (let i = 0; i < s.length; i += 1) {
    const code = s.charCodeAt(i);
    if (code < 65 || code > 90) break;
    n = n * 26 + (code - 64);
  }
  return n;
}

function encodeCellAddr(rowIndex1, colIndex1) {
  return `${encodeCol(colIndex1)}${rowIndex1}`;
}

function decodeCellAddr(addr) {
  const s = String(addr || "").replace(/\$/g, "");
  const m = s.match(/^([A-Za-z]+)(\d+)$/);
  if (!m) return null;
  return { c: decodeCol(m[1]), r: Number(m[2]) };
}

function decodeRangeAddr(range) {
  const s = String(range || "").replace(/\$/g, "");
  const parts = s.split(":");
  const start = decodeCellAddr(parts[0]);
  const end = decodeCellAddr(parts[1] || parts[0]);
  if (!start || !end) return null;
  return {
    s: { r: start.r, c: start.c },
    e: { r: end.r, c: end.c },
  };
}

function encodeRangeAddr(sRow, sCol, eRow, eCol) {
  return `${encodeCellAddr(sRow, sCol)}:${encodeCellAddr(eRow, eCol)}`;
}

function adjustRangeByInsertRows(rangeStr, insertAtRow1, deltaRows) {
  const decoded = decodeRangeAddr(rangeStr);
  if (!decoded) return null;

  let { r: sRow, c: sCol } = decoded.s;
  let { r: eRow, c: eCol } = decoded.e;

  if (eRow < insertAtRow1) {
    return encodeRangeAddr(sRow, sCol, eRow, eCol);
  }

  if (sRow >= insertAtRow1) {
    sRow += deltaRows;
    eRow += deltaRows;
    return encodeRangeAddr(sRow, sCol, eRow, eCol);
  }

  eRow += deltaRows;
  return encodeRangeAddr(sRow, sCol, eRow, eCol);
}

function getWorksheetMergeRanges(worksheet) {
  const merges = worksheet?.model?.merges;
  if (Array.isArray(merges)) return merges.slice();
  return [];
}

function clearNonMasterCellsForMerge(worksheet, rangeStr) {
  const decoded = decodeRangeAddr(rangeStr);
  if (!decoded) return;
  for (let r = decoded.s.r; r <= decoded.e.r; r += 1) {
    const row = worksheet.getRow(r);
    for (let c = decoded.s.c; c <= decoded.e.c; c += 1) {
      if (r === decoded.s.r && c === decoded.s.c) continue;
      row.getCell(c).value = null;
    }
  }
}

function restoreMergesAfterInsert(worksheet, originalMerges, insertAtRow1, deltaRows) {
  if (!originalMerges || !deltaRows) return;

  // 1. 获取当前工作表所有的合并区域（包括 spliceRows 自动产生或错位的）
  const currentMerges = getWorksheetMergeRanges(worksheet);
  
  // 2. 彻底解除所有受影响的合并：
  //    任何结束行在插入点之后（>= insertAtRow1）的合并都需要重新处理
  //    因为它们要么被 spliceRows 错误复制了，要么位置已经错位了
  for (const rngStr of currentMerges) {
    const decoded = decodeRangeAddr(rngStr);
    if (decoded && decoded.e.r >= insertAtRow1) {
      try {
        worksheet.unMergeCells(rngStr);
      } catch (e) {}
    }
  }

  // 3. 根据原始模板的合并规则，重新应用调整后的合并
  for (const orig of originalMerges) {
    const adjusted = adjustRangeByInsertRows(orig, insertAtRow1, deltaRows);
    if (!adjusted) continue;
    
    const adjDecoded = decodeRangeAddr(adjusted);
    // 如果调整后的合并区域完全在插入点上方，说明它刚才没被解除，不需要重做
    if (adjDecoded && adjDecoded.e.r < insertAtRow1) continue;

    // 清除非左上角单元格的值（避免合并后 WPS 显示重复）
    clearNonMasterCellsForMerge(worksheet, adjusted);
    
    try {
      worksheet.mergeCells(adjusted);
    } catch (e) {}
  }
}

function deepClone(value) {
  if (value === null || value === undefined) return value;
  if (typeof structuredClone === "function") return structuredClone(value);
  if (value instanceof Date) return new Date(value.getTime());
  if (Array.isArray(value)) return value.map((v) => deepClone(v));
  if (typeof value === "object") {
    const out = {};
    for (const [k, v] of Object.entries(value)) out[k] = deepClone(v);
    return out;
  }
  return value;
}

function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error ?? new Error("读取文件失败"));
    reader.readAsArrayBuffer(file);
  });
}

function loadScriptOnce(url) {
  return new Promise((resolve, reject) => {
    const existing = document.querySelector(`script[data-exceljs-src="${url}"]`);
    if (existing) {
      if (globalThis.ExcelJS?.Workbook) resolve();
      else existing.addEventListener("load", () => resolve(), { once: true });
      return;
    }

    const script = document.createElement("script");
    script.src = url;
    script.async = true;
    script.dataset.exceljsSrc = url;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`加载脚本失败：${url}`));
    document.head.appendChild(script);
  });
}

async function ensureHyperFormula() {
  if (globalThis.HyperFormula) return;
  const urls = [
    "https://cdn.jsdelivr.net/npm/hyperformula/dist/hyperformula.full.min.js",
  ];

  let lastError = null;
  for (const url of urls) {
    try {
      await loadScriptOnce(url);
      if (globalThis.HyperFormula) return;
    } catch (e) {
      lastError = e;
    }
  }
  throw lastError ?? new Error("公式引擎加载失败（HyperFormula 未就绪）");
}

function sanitizeFormulaForHF(formula, r) {
  if (!formula || typeof formula !== "string") return formula;
  
  // 针对 SUM(J$9:INDEX(J:J, ROW()-1)) 这种动态区域公式做兼容性处理
  // 因为 HyperFormula 的 INDEX 不支持作为 Range 的端点返回引用，只返回具体值
  // 我们将其在预览阶段临时替换为固定的 A1:A{r-1} 格式
  let sanitized = formula;
  
  // 匹配模式：([列号]$[起始行]):INDEX([列号]:[列号], ROW()-1)
  // 例如：J$9:INDEX(J:J, ROW()-1)
  const dynamicRangeRegex = /([A-Z]+)\$?(\d+):INDEX\(\s*(\1):\1\s*,\s*ROW\(\)\s*-\s*1\s*\)/gi;
  sanitized = sanitized.replace(dynamicRangeRegex, (match, col, startRow) => {
    return `${col}$${startRow}:${col}${r - 1}`;
  });

  return sanitized;
}

function initFormulaEngine(ws) {
  if (!globalThis.HyperFormula) return;

  const aoa = [];
  const rowCount = Math.max(ws.rowCount || 0, ws.actualRowCount || 0);
  const colCount = Math.max(ws.columnCount || 0, ws.actualColumnCount || 0);

  for (let r = 1; r <= rowCount; r++) {
    const rowValues = [];
    const wsRow = ws.getRow(r);
    for (let c = 1; c <= colCount; c++) {
      const cell = wsRow.getCell(c);
      const val = cell.value;
      if (val && typeof val === "object" && val.formula) {
        // 在存入公式引擎前进行清洗
        const sanitized = sanitizeFormulaForHF(val.formula, r);
        rowValues.push("=" + sanitized);
      } else {
        rowValues.push(val);
      }
    }
    aoa.push(rowValues);
  }

  // 创建 HyperFormula 实例
  if (state.hf) state.hf.destroy();
  state.hf = HyperFormula.buildFromArray(aoa, {
    licenseKey: "gpl-v3",
  });
  state.hfSheetId = state.hf.getSheetId(state.hf.getSheetNames()[0]);
}

function getCellValueWithFormula(r, c) {
  if (state.hf && state.hfSheetId !== null) {
    // HyperFormula 使用 0-based index
    const val = state.hf.getCellValue({
      sheet: state.hfSheetId,
      row: r - 1,
      col: c - 1,
    });
    
    if (val !== null && typeof val === "object" && val.error) {
      return "#ERROR!";
    }
    return val;
  }
  return null;
}

async function ensureExcelJS() {
  if (globalThis.ExcelJS?.Workbook) return;
  const urls = [
    "https://cdn.jsdelivr.net/npm/exceljs@4.3.0/dist/exceljs.min.js",
    "https://unpkg.com/exceljs@4.3.0/dist/exceljs.min.js",
  ];

  let lastError = null;
  for (const url of urls) {
    try {
      await loadScriptOnce(url);
      if (globalThis.ExcelJS?.Workbook) return;
    } catch (e) {
      lastError = e;
    }
  }

  throw lastError ?? new Error("Excel 引擎加载失败（ExcelJS 未就绪）");
}

async function loadWorkbookFromFile(file) {
  const buffer = await readFileAsArrayBuffer(file);
  const workbook = await loadWorkbookFromArrayBuffer(buffer);
  return { workbook, buffer };
}

async function loadWorkbookFromArrayBuffer(buffer) {
  await ensureExcelJS();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  return workbook;
}

function fillSheetSelect(selectEl, workbook) {
  selectEl.innerHTML = "";
  for (const ws of workbook.worksheets) {
    const opt = document.createElement("option");
    opt.value = ws.name;
    opt.textContent = ws.name;
    selectEl.appendChild(opt);
  }
  selectEl.disabled = workbook.worksheets.length === 0;
}

function getWorksheet(workbook, sheetName) {
  return workbook.getWorksheet(sheetName) ?? null;
}

function getCellText(cell) {
  if (!cell) return "";
  const v = cell.value;
  if (v === null || v === undefined) return "";
  if (typeof v === "object") {
    if (v.text) return normalizeText(v.text);
    if (v.richText && Array.isArray(v.richText)) return normalizeText(v.richText.map((x) => x.text || "").join(""));
    if (v.formula !== undefined) return normalizeText(String(v.result ?? ""));
    if (v.sharedFormula !== undefined) return normalizeText(String(v.result ?? ""));
    if (v.hyperlink) return normalizeText(String(v.text ?? v.hyperlink));
    if (v instanceof Date) return normalizeText(v.toISOString());
  }
  try {
    return normalizeText(v);
  } catch {
    return "";
  }
}

function findHeaderRowInWorksheet(worksheet, purchaseHeaders) {
  const normalizedHeaders = purchaseHeaders.map(normalizeText).filter(Boolean);
  if (normalizedHeaders.length === 0) return null;
  if (!worksheet) return null;

  let best = null;
  const maxRow = Math.max(worksheet.rowCount || 0, worksheet.actualRowCount || 0);
  const maxCol = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0);

  for (let r = 1; r <= maxRow; r += 1) {
    const row = worksheet.getRow(r);
    const headerToCol = new Map();
    for (let c = 1; c <= maxCol; c += 1) {
      const v = getCellText(row.getCell(c));
      if (v && !headerToCol.has(v)) headerToCol.set(v, c);
    }

    let hitCount = 0;
    let minCol = Infinity;
    let maxHitCol = -Infinity;
    for (const header of normalizedHeaders) {
      const col = headerToCol.get(header);
      if (col === undefined) continue;
      hitCount += 1;
      minCol = Math.min(minCol, col);
      maxHitCol = Math.max(maxHitCol, col);
    }

    if (hitCount === 0) continue;
    const span = maxHitCol - minCol;

    if (
      !best ||
      hitCount > best.hitCount ||
      (hitCount === best.hitCount && span < best.span) ||
      (hitCount === best.hitCount && span === best.span && r < best.rowIndex) ||
      (hitCount === best.hitCount && span === best.span && r === best.rowIndex && minCol < best.minCol)
    ) {
      best = {
        rowIndex: r,
        hitCount,
        headerToCol,
        minCol,
        maxCol: maxHitCol,
        span,
      };
    }
  }

  return best;
}

function getRowNonEmptyMinMax(worksheet, rowIndex1) {
  const row = worksheet.getRow(rowIndex1);
  let minCol = Infinity;
  let maxCol = -Infinity;
  row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const v = getCellText(cell);
    if (!v) return;
    minCol = Math.min(minCol, colNumber);
    maxCol = Math.max(maxCol, colNumber);
  });
  if (!Number.isFinite(minCol) || !Number.isFinite(maxCol)) return null;
  return { minCol, maxCol };
}

function adjustFormula(formula, oldRow, newRow) {
  if (!formula || typeof formula !== "string") return formula;
  // 简单的正则替换：寻找紧跟在字母列号后的行号，并进行偏移
  // 例如将 G9 * H9 中的 9 替换为 10 -> G10 * H10
  // 注意：这只适用于简单的 A1 引用格式
  const diff = newRow - oldRow;
  if (diff === 0) return formula;

  return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
    const rowNum = parseInt(row, 10);
    // 只有当公式里的行号正好是模板里的示例行号时，才进行偏移
    // 这样可以避免误伤类似 SUM(G$9:G$100) 这种绝对引用或固定区域
    if (rowNum === oldRow) {
      return col + newRow;
    }
    return match;
  });
}

async function insertPurchaseRowsIntoTemplate({
  templateWb,
  templateSheetName,
  purchaseWb,
  purchaseSheetName,
  mapping,
}) {
  const templateWs = getWorksheet(templateWb, templateSheetName);
  const purchaseWs = getWorksheet(purchaseWb, purchaseSheetName);
  if (!templateWs) throw new Error("合同模板工作表不存在");
  if (!purchaseWs) throw new Error("订购单工作表不存在");

  const contractHeaders = Array.isArray(mapping?.config?.contract_headers)
    ? mapping.config.contract_headers
    : Array.isArray(mapping?.contract_headers)
      ? mapping.contract_headers
      : null;
  const purchaseHeaders = Array.isArray(mapping?.config?.purchase_headers)
    ? mapping.config.purchase_headers
    : Array.isArray(mapping?.purchase_headers)
      ? mapping.purchase_headers
      : null;
  const rules = Array.isArray(mapping?.rules) ? mapping.rules : null;

  if (!contractHeaders || contractHeaders.length === 0) throw new Error("映射配置缺少 contract_headers");
  if (!purchaseHeaders || purchaseHeaders.length === 0) throw new Error("映射配置缺少 purchase_headers");
  if (!rules || rules.length === 0) throw new Error("映射规则为空");

  const purchaseHeaderRow = findHeaderRowInWorksheet(purchaseWs, purchaseHeaders);
  if (!purchaseHeaderRow) throw new Error("未在订购单中找到映射配置指定的标题行");

  const contractHeaderRow = findHeaderRowInWorksheet(templateWs, contractHeaders);
  if (!contractHeaderRow) throw new Error("未在合同模板中找到映射配置指定的标题行");

  const mappingPairs = [];
  for (const r of rules) {
    const contractHeader = normalizeText(r?.contract_header);
    const purchaseHeader = normalizeText(r?.purchase_header);
    if (!contractHeader || !purchaseHeader) continue;
    const contractCol = contractHeaderRow.headerToCol.get(contractHeader);
    const purchaseCol = purchaseHeaderRow.headerToCol.get(purchaseHeader);
    if (!contractCol || !purchaseCol) continue;
    mappingPairs.push({ contractCol, purchaseCol });
  }
  if (mappingPairs.length === 0) throw new Error("映射规则未能匹配到任何列");

  const headerRowMinMax =
    getRowNonEmptyMinMax(templateWs, contractHeaderRow.rowIndex) ?? {
      minCol: contractHeaderRow.minCol,
      maxCol: contractHeaderRow.maxCol,
    };

  const originCol = headerRowMinMax.minCol;
  let maxAssignedCol = headerRowMinMax.maxCol;
  for (const p of mappingPairs) maxAssignedCol = Math.max(maxAssignedCol, p.contractCol);

  const insertAtRow1 = contractHeaderRow.rowIndex + 1;

  const styleSourceRow = templateWs.getRow(insertAtRow1);
  const styleByCol = new Map();
  const formulaByCol = new Map();
  let fallbackStyle = null;
  
  for (let c = originCol; c <= maxAssignedCol; c += 1) {
    const cell = styleSourceRow.getCell(c);
    if (cell?.style && Object.keys(cell.style).length > 0) {
      const s = deepClone(cell.style);
      styleByCol.set(c, s);
      if (!fallbackStyle) fallbackStyle = s;
    }
    // 检查并记录公式
    if (cell?.value && typeof cell.value === "object" && cell.value.formula) {
      formulaByCol.set(c, cell.value.formula);
    }
  }
  const rowHeight = styleSourceRow.height;

  const insertRowValuesList = [];
  const maxPurchaseRow = Math.max(purchaseWs.rowCount || 0, purchaseWs.actualRowCount || 0);
  for (let r = purchaseHeaderRow.rowIndex + 1; r <= maxPurchaseRow; r += 1) {
    const row = purchaseWs.getRow(r);
    const values = new Array(maxAssignedCol).fill(null);
    let any = false;
    for (const pair of mappingPairs) {
      const cell = row.getCell(pair.purchaseCol);
      const text = getCellText(cell);
      if (normalizeText(text)) any = true;
      values[pair.contractCol - 1] = cell.value ?? null;
    }
    if (any) insertRowValuesList.push(values);
  }
  if (insertRowValuesList.length === 0) throw new Error("订购单没有可插入的数据行");

  const originalMerges = getWorksheetMergeRanges(templateWs);
  const deltaRows = insertRowValuesList.length;

  // 在插入前，探测标题行下方有多少行是“空白模板行”
  // 逻辑：从标题行下一行开始，如果某行所有单元格要么为空，要么只有公式但没有手动输入的值，则视为模板行
  let blankTemplateRowsCount = 0;
  const maxSearchRows = 50; // 最多向上探测50行，防止意外删除过多
  for (let i = 0; i < maxSearchRows; i++) {
    const checkRow = templateWs.getRow(insertAtRow1 + i);
    let isBlank = true;
    checkRow.eachCell({ includeEmpty: false }, (cell) => {
      const val = cell.value;
      // 如果单元格有值且不是公式对象（或者公式对象没有结果），则认为不是空白模板行
      if (val !== null && val !== undefined) {
        if (typeof val === 'object' && val.formula) {
          // 仅有公式，继续检查其他格
        } else if (String(val).trim() !== "") {
          isBlank = false;
        }
      }
    });
    if (isBlank && checkRow.actualCellCount >= 0) {
      blankTemplateRowsCount++;
    } else {
      break;
    }
  }

  // 执行插入
  templateWs.spliceRows(insertAtRow1, 0, ...insertRowValuesList);

  // 插入后，原来的空白模板行被挤到了下方位置：insertAtRow1 + deltaRows
  // 删除这些多余的空白模板行
  if (blankTemplateRowsCount > 0) {
    templateWs.spliceRows(insertAtRow1 + deltaRows, blankTemplateRowsCount);
  }

  // 修正：计算净行数变化。
  // 合并单元格的偏移应该基于 (插入行数 - 删除行数)
  const netDelta = deltaRows - blankTemplateRowsCount;

  for (let i = 0; i < insertRowValuesList.length; i += 1) {
    const rowNum = insertAtRow1 + i;
    const row = templateWs.getRow(rowNum);
    if (rowHeight !== undefined) row.height = rowHeight;
    
    for (let c = originCol; c <= maxAssignedCol; c += 1) {
      const cell = row.getCell(c);
      
      // 应用样式
      const style = styleByCol.get(c) ?? fallbackStyle;
      if (style) cell.style = deepClone(style);
      
      // 处理公式：如果模板该列原本有公式，则为新行生成对应的偏移公式
      if (formulaByCol.has(c)) {
        const baseFormula = formulaByCol.get(c);
        // 这里 oldRow 使用 insertAtRow1，因为 baseFormula 是在 spliceRows 之前从该行提取的
        const newFormula = adjustFormula(baseFormula, insertAtRow1, rowNum);
        cell.value = { formula: newFormula };
      }
    }
  }

  // 使用净变化 netDelta 来恢复合并单元格位置
  restoreMergesAfterInsert(templateWs, originalMerges, insertAtRow1, netDelta);

  return {
    insertedRows: insertRowValuesList.length,
    headerRowIndex: contractHeaderRow.rowIndex,
    matchedHeaders: mappingPairs.length,
    purchaseHeaderCount: purchaseHeaders.length,
  };
}

function refreshPreviewValues(container = previewEl) {
  if (!state.hf) return;
  if (!container) return;
  const cells = container.querySelectorAll("td[data-row][data-col]");
  cells.forEach((td) => {
    const r = parseInt(td.dataset.row);
    const c = parseInt(td.dataset.col);
    const val = getCellValueWithFormula(r, c);
    const displayValue = val !== null && val !== undefined ? String(val) : "";
    td.innerText = displayValue;
    if (displayValue) td.classList.remove("empty");
    else td.classList.add("empty");
  });
}

function renderPreviewInto(container, workbook, sheetName, options = {}) {
  const readOnly = !!options.readOnly;
  const useFormulaEngine = options.useFormulaEngine !== false;

  if (!container) return;
  const ws = getWorksheet(workbook, sheetName);
  if (!ws) {
    container.textContent = "未找到工作表";
    return;
  }

  const maxRows = 60;
  const maxCols = 20;
  const endRow = Math.min(Math.max(ws.rowCount || 0, ws.actualRowCount || 0) || 1, maxRows);
  const endCol = Math.min(Math.max(ws.columnCount || 0, ws.actualColumnCount || 0) || 1, maxCols);

  const keyOf = (r, c) => `${r},${c}`;
  const mergeMaster = new Map();
  const mergeSkip = new Set();
  const merges = getWorksheetMergeRanges(ws);
  for (const rng of merges) {
    const decoded = decodeRangeAddr(rng);
    if (!decoded) continue;
    const s = decoded.s;
    const e = decoded.e;
    if (s.r < 1 || s.c < 1 || e.r < 1 || e.c < 1) continue;
    if (s.r > endRow || s.c > endCol) continue;
    const rr = Math.min(e.r, endRow) - s.r + 1;
    const cc = Math.min(e.c, endCol) - s.c + 1;
    if (rr <= 1 && cc <= 1) continue;
    mergeMaster.set(keyOf(s.r, s.c), { rowSpan: rr, colSpan: cc });
    for (let r = s.r; r <= Math.min(e.r, endRow); r += 1) {
      for (let c = s.c; c <= Math.min(e.c, endCol); c += 1) {
        if (r === s.r && c === s.c) continue;
        mergeSkip.add(keyOf(r, c));
      }
    }
  }

  const table = document.createElement("table");
  table.className = "table";

  const thead = document.createElement("thead");
  const headRow = document.createElement("tr");
  for (let c = 1; c <= endCol; c += 1) {
    const th = document.createElement("th");
    th.textContent = encodeCol(c);
    headRow.appendChild(th);
  }
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  for (let r = 1; r <= endRow; r += 1) {
    const tr = document.createElement("tr");
    const row = ws.getRow(r);
    for (let c = 1; c <= endCol; c += 1) {
      if (mergeSkip.has(keyOf(r, c))) continue;
      const td = document.createElement("td");
      const master = mergeMaster.get(keyOf(r, c));
      if (master) {
        if (master.rowSpan > 1) td.rowSpan = master.rowSpan;
        if (master.colSpan > 1) td.colSpan = master.colSpan;
      }
      
      // 优先从公式引擎获取计算结果
      let displayValue = "";
      if (useFormulaEngine && state.hf) {
        const val = getCellValueWithFormula(r, c);
        displayValue = val !== null && val !== undefined ? String(val) : "";
      } else {
        displayValue = getCellText(row.getCell(c));
      }
      
      const text = displayValue;
      td.innerText = text;
      
      if (!readOnly) {
        td.dataset.row = r;
        td.dataset.col = c;

        const cell = row.getCell(c);
        const isFormula = cell?.value && typeof cell.value === "object" && cell.value.formula;
        
        if (isFormula) {
          td.contentEditable = "false";
          td.classList.add("formula-cell");
          td.title = "公式单元格，无法手动修改";
        } else {
          td.contentEditable = "true";
          
          td.addEventListener("blur", (e) => {
            const rowNum = parseInt(e.target.dataset.row);
            const colNum = parseInt(e.target.dataset.col);
            const newValue = e.target.innerText.trim();
            
            const targetSheet = state.merged.getWorksheet(state.mergedSheetName);
            if (targetSheet) {
              const cell = targetSheet.getRow(rowNum).getCell(colNum);
              
              const num = Number(newValue);
              if (newValue !== "" && !isNaN(num)) {
                cell.value = num;
              } else {
                cell.value = newValue;
                if (newValue.includes("\n")) {
                  cell.alignment = { wrapText: true };
                }
              }

              if (useFormulaEngine && state.hf && state.hfSheetId !== null) {
                state.hf.setCellContents({
                  sheet: state.hfSheetId,
                  row: rowNum - 1,
                  col: colNum - 1
                }, [[cell.value]]);
                
                refreshPreviewValues(container);
              }
            }
          });

          td.addEventListener("keydown", (e) => {
            if (e.key === "Enter") {
              if (e.shiftKey) {
                return;
              } else {
                e.preventDefault();
                const rowNum = parseInt(e.target.dataset.row);
                const colNum = parseInt(e.target.dataset.col);
                
                let nextRow = rowNum + 1;
                let found = false;
                while (nextRow <= endRow) {
                  const nextCell = container.querySelector(`td[data-row="${nextRow}"][data-col="${colNum}"][contenteditable="true"]`);
                  if (nextCell) {
                    nextCell.focus();
                    found = true;
                    break;
                  }
                  nextRow++;
                }
                if (!found) {
                  e.target.blur();
                }
              }
            }
          });
        }
      } else {
        td.contentEditable = "false";
      }

      if (!displayValue) td.className = "empty";
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);

  container.innerHTML = "";
  container.appendChild(table);
}

function renderPreview(workbook, sheetName) {
  return renderPreviewInto(previewEl, workbook, sheetName, { readOnly: false, useFormulaEngine: true });
}

function enableActionsIfReady() {
  const ready = !!(state.template?.workbook && state.purchase?.workbook);
  mergeBtnEl.disabled = !ready;
}

async function downloadWorkbook(workbook, filename) {
  const out = await workbook.xlsx.writeBuffer();
  const blob = new Blob([out], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function updateFileMeta(targetEl, file, workbook) {
  const sheetCount = workbook.worksheets.length;
  targetEl.textContent = `文件：${file.name}；工作表：${sheetCount} 个`;
}

function sanitizeFilenamePart(input) {
  const s = String(input ?? "").trim();
  if (!s) return "";
  return s
    .replace(/[\\/:*?"<>|：]/g, "_")
    .replace(/\s+/g, "")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function computeDefaultOutputBaseName() {
  const templateNameRaw = state.template?.file?.name?.replace(/\.(xlsx|xls)$/i, "") || "合同模板";
  const templateName = sanitizeFilenamePart(templateNameRaw) || "合同模板";
  const purchaseNoRaw = state.merged ? getPurchaseOrderNoFromTemplate(state.merged, state.mergedSheetName) : "";
  const purchaseNo = sanitizeFilenamePart(purchaseNoRaw);
  return purchaseNo ? `${templateName}_${purchaseNo}` : `${templateName}_已插入订购单`;
}

function ensureXlsxFilename(name) {
  const s = String(name ?? "").trim();
  if (!s) return "合同.xlsx";
  if (/\.xlsx$/i.test(s)) return s;
  return `${s}.xlsx`;
}

function getEffectiveOutputFilename() {
  const typed = sanitizeFilenamePart(saveNameInputEl?.value);
  const base = typed ? typed.replace(/\.xlsx$/i, "") : computeDefaultOutputBaseName();
  return ensureXlsxFilename(base);
}

function formatDateTime(value) {
  if (!value) return "";
  const d = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}/${pad(d.getMonth() + 1)}/${pad(d.getDate())} ${pad(d.getHours())}:${pad(
    d.getMinutes(),
  )}:${pad(d.getSeconds())}`;
}

function downloadArrayBuffer(arrayBuffer, filename) {
  const blob = new Blob([arrayBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function renderContractsTable(list) {
  if (!contractsTbodyEl) return;
  const rows = Array.isArray(list) ? list : [];
  contractsTbodyEl.innerHTML = "";
  if (rows.length === 0) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="4" style="color:#64748b;">暂无数据</td>`;
    contractsTbodyEl.appendChild(tr);
    return;
  }

  rows.forEach((row, idx) => {
    const tr = document.createElement("tr");
    const name = row?.name ?? "";
    const createdAt = formatDateTime(row?.created_at);
    const id = String(row?.id ?? "");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${String(name)}</td>
      <td>${createdAt}</td>
      <td>
        <div class="opBtns">
          <button class="miniBtn" data-action="preview" data-id="${id}" data-name="${String(name)}">预览</button>
          <button class="miniBtn primary" data-action="download" data-id="${id}" data-name="${String(name)}">下载</button>
          <button class="miniBtn danger" data-action="delete" data-id="${id}" data-name="${String(name)}">删除</button>
        </div>
      </td>
    `;
    contractsTbodyEl.appendChild(tr);
  });
}

async function fetchContractsList() {
  if (!contractsTbodyEl) return;
  try {
    const out = await fetchJsonWithFallback("/api/contracts");
    const res = out.res;
    const json = out.json;
    if (!res.ok || !json?.success) {
      const msg = json?.error || `${res.status} ${res.statusText}`;
      throw new Error(msg);
    }
    const list = Array.isArray(json.data) ? json.data : [];
    renderContractsTable(list);
  } catch (e) {
    renderContractsTable([]);
  }
}

const modalState = {
  workbook: null,
  sheetName: null,
  title: "",
};

function closePreviewModal() {
  if (!previewModalEl) return;
  previewModalEl.classList.remove("open");
  previewModalEl.setAttribute("aria-hidden", "true");
  modalState.workbook = null;
  modalState.sheetName = null;
  if (modalTitleEl) modalTitleEl.textContent = "预览";
  if (modalSheetSelectEl) modalSheetSelectEl.innerHTML = "";
  if (modalPreviewEl) modalPreviewEl.innerHTML = "";
}

function openPreviewModal(title, workbook) {
  if (!previewModalEl || !modalPreviewEl) return;
  modalState.workbook = workbook || null;
  modalState.title = String(title || "预览");
  if (modalTitleEl) modalTitleEl.textContent = modalState.title;

  if (modalSheetSelectEl) {
    modalSheetSelectEl.innerHTML = "";
    const sheets = Array.isArray(workbook?.worksheets) ? workbook.worksheets : [];
    for (const ws of sheets) {
      const opt = document.createElement("option");
      opt.value = ws.name;
      opt.textContent = ws.name;
      modalSheetSelectEl.appendChild(opt);
    }
    modalState.sheetName = sheets[0]?.name || "";
    if (modalState.sheetName) modalSheetSelectEl.value = modalState.sheetName;
  } else {
    modalState.sheetName = workbook?.worksheets?.[0]?.name || "";
  }

  if (modalState.sheetName) {
    renderPreviewInto(modalPreviewEl, workbook, modalState.sheetName, { readOnly: true, useFormulaEngine: false });
  } else {
    modalPreviewEl.textContent = "未找到工作表";
  }

  previewModalEl.classList.add("open");
  previewModalEl.setAttribute("aria-hidden", "false");
}

if (modalCloseEl) modalCloseEl.addEventListener("click", closePreviewModal);
if (previewModalEl) {
  previewModalEl.addEventListener("click", (e) => {
    if (e.target === previewModalEl) closePreviewModal();
  });
}
if (modalSheetSelectEl) {
  modalSheetSelectEl.addEventListener("change", () => {
    if (!modalState.workbook || !modalPreviewEl) return;
    const sheetName = String(modalSheetSelectEl.value || "");
    modalState.sheetName = sheetName;
    if (!sheetName) return;
    renderPreviewInto(modalPreviewEl, modalState.workbook, sheetName, { readOnly: true, useFormulaEngine: false });
  });
}
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") closePreviewModal();
});

function extractPurchaseNoFromCellLabel(text) {
  const s = normalizeText(text);
  const m = s.match(/(?:采购单号|采购单编号|采购订单号|订购单号|订单号)\s*[:：]\s*(.+)$/);
  if (!m) return "";
  return normalizeText(m[1] || "");
}

function findPurchaseOrderNoInWorksheet(worksheet) {
  if (!worksheet) return "";

  const labels = [
    "采购单号",
    "采购单编号",
    "采购订单号",
    "订购单号",
    "订单号",
    "PO号",
    "PO编号",
    "PO",
  ];

  const maxRow = Math.min(
    80,
    Math.max(worksheet.rowCount || 0, worksheet.actualRowCount || 0) || 0,
  );
  const maxCol = Math.min(
    30,
    Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0) || 0,
  );

  for (let r = 1; r <= maxRow; r += 1) {
    const row = worksheet.getRow(r);
    for (let c = 1; c <= maxCol; c += 1) {
      const cell = row.getCell(c);
      const text = getCellText(cell);
      if (!text) continue;

      const isPurchaseNoLabelWithColon = /^采购单号\s*[:：]\s*$/.test(normalizeText(text));
      if (isPurchaseNoLabelWithColon) {
        const labelText = normalizeText(text);
        for (let k = 1; k <= 10; k += 1) {
          const candidate = getCellText(row.getCell(c + k));
          if (!candidate) continue;
          const candidateText = normalizeText(candidate);
          if (candidateText === labelText) continue;
          const cleanedCandidate = candidateText.replace(/[：:]/g, "");
          const isLabelAgain = labels.some(
            (l) => cleanedCandidate === l || cleanedCandidate.toLowerCase() === l.toLowerCase(),
          );
          if (isLabelAgain) continue;
          return candidateText;
        }
        continue;
      }

      const fromInline = extractPurchaseNoFromCellLabel(text);
      if (fromInline) return fromInline;

      const cleaned = normalizeText(text).replace(/[：:]/g, "");
      const hit = labels.some((l) => cleaned === l || cleaned.toLowerCase() === l.toLowerCase());
      if (!hit) continue;

      const right = getCellText(row.getCell(c + 1));
      if (right) {
        const rightText = normalizeText(right);
        if (rightText !== normalizeText(text)) return rightText;
      }
    }
  }

  return "";
}

function getPurchaseOrderNoFromTemplate(workbook, preferredSheetName) {
  if (!workbook) return "";
  const preferredWs = preferredSheetName ? workbook.getWorksheet(preferredSheetName) : null;
  const foundPreferred = findPurchaseOrderNoInWorksheet(preferredWs);
  if (foundPreferred) return foundPreferred;

  for (const ws of workbook.worksheets || []) {
    const found = findPurchaseOrderNoInWorksheet(ws);
    if (found) return found;
  }

  return "";
}

async function cloneWorkbook(workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  const clone = new ExcelJS.Workbook();
  await clone.xlsx.load(buffer);
  return clone;
}

async function mergeToNewWorkbook() {
  if (!state.template?.workbook || !state.purchase?.workbook) return;
  const templateSheetName = templateSheetEl.value;
  const purchaseSheetName = purchaseSheetEl.value;
  const templateId = String(state.template?.id ?? templateSelectEl?.value ?? "");
  const purchaseTypeId = String(state.purchase?.typeId ?? "");
  if (!templateId) throw new Error("缺少模板 ID");
  if (!purchaseTypeId) throw new Error("缺少订购单类型标识（type_ID）");

  // 使用当前在预览中（可能已编辑）的模板作为源进行克隆
  setStatus("正在准备合并环境…");
  const templateClone = await cloneWorkbook(state.template.workbook);

  setStatus("正在读取映射关系…");
  const mappingOut = await fetchJsonWithFallback(
    `/api/mapping?template_id=${encodeURIComponent(templateId)}&purchase_type_id=${encodeURIComponent(purchaseTypeId)}`,
  );
  if (!mappingOut.res.ok || !mappingOut.json?.success) {
    const msg = mappingOut.json?.error || `${mappingOut.res.status} ${mappingOut.res.statusText}`;
    throw new Error(`映射关系读取失败：${msg}`);
  }
  const mapping = mappingOut.json.data;

  setStatus("正在生成…");
  const result = await insertPurchaseRowsIntoTemplate({
    templateWb: templateClone,
    templateSheetName,
    purchaseWb: state.purchase.workbook,
    purchaseSheetName,
    mapping,
  });

  state.merged = templateClone;
  state.mergedSheetName = templateSheetName;

  // 初始化公式引擎
  try {
    await ensureHyperFormula();
    const targetWs = state.merged.getWorksheet(state.mergedSheetName);
    initFormulaEngine(targetWs);
  } catch (e) {
    console.warn("公式引擎初始化失败，预览可能无法显示实时计算结果", e);
  }

  downloadBtnEl.disabled = false;
  if (saveBtnEl) saveBtnEl.disabled = false;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = false;
    saveNameInputEl.value = ensureXlsxFilename(computeDefaultOutputBaseName());
  }
  renderPreview(templateClone, templateSheetName);
  setStatus(
    `已插入 ${result.insertedRows} 行（插入起始行：第 ${result.headerRowIndex + 2} 行；标题匹配 ${result.matchedHeaders}/${result.purchaseHeaderCount}）`,
  );
}

async function handleDownload() {
  if (!state.merged) return;
  const filename = getEffectiveOutputFilename();
  await downloadWorkbook(state.merged, filename);
}

async function handleSaveToServer() {
  if (!state.merged) return;
  const filename = getEffectiveOutputFilename();
  const baseName = filename.replace(/\.xlsx$/i, "");

  try {
    setStatus("正在保存到服务器…");
    const out = await state.merged.xlsx.writeBuffer();
    const pathname = `/api/contracts/save?name=${encodeURIComponent(baseName)}`;
    const res = await postBinaryWithFallback(pathname, out);
    let json = null;
    try {
      json = await res.json();
    } catch {}
    if (!res.ok || !json?.success) {
      const msg = json?.error || `${res.status} ${res.statusText}`;
      throw new Error(msg);
    }
    setStatus("已保存到合同列表");
    fetchContractsList();
  } catch (e) {
    setStatus(`保存失败：${e?.message ?? e}`);
  }
}

templateSelectEl.addEventListener("change", handleTemplateSelectChange);
if (purchaseSelectEl) purchaseSelectEl.addEventListener("change", handlePurchaseSelectChange);
mergeBtnEl.addEventListener("click", () => {
  mergeToNewWorkbook().catch((e) => {
    state.merged = null;
    state.mergedSheetName = null;
    downloadBtnEl.disabled = true;
    if (saveBtnEl) saveBtnEl.disabled = true;
    if (saveNameInputEl) {
      saveNameInputEl.disabled = true;
      saveNameInputEl.value = "";
    }
    setStatus(`生成失败：${e?.message ?? e}`);
  });
});
downloadBtnEl.addEventListener("click", handleDownload);
if (saveBtnEl) saveBtnEl.addEventListener("click", handleSaveToServer);
if (contractsTbodyEl) {
  contractsTbodyEl.addEventListener("click", (ev) => {
    const btn = ev.target?.closest?.("button[data-action]");
    if (!btn) return;
    const action = btn.dataset.action;
    const id = btn.dataset.id;
    const name = btn.dataset.name || "";
    if (!id) return;

    if (action === "download") {
      (async () => {
        try {
          const response = await fetchWithLocalFallback(`/api/contracts/${encodeURIComponent(id)}/download`);
          if (!response.ok) {
            const text = await response.text().catch(() => "");
            const parsed = (() => {
              try {
                return JSON.parse(text);
              } catch {
                return null;
              }
            })();
            const msg = parsed?.error || parsed?.message || text || `${response.status} ${response.statusText}`;
            throw new Error(msg);
          }
          const buffer = await response.arrayBuffer();
          const filename = ensureXlsxFilename(sanitizeFilenamePart(name) || `contract_${id}`);
          downloadArrayBuffer(buffer, filename);
        } catch (e) {
          setStatus(`下载失败：${e?.message ?? e}`);
        }
      })();
      return;
    }

    if (action === "preview") {
      (async () => {
        try {
          setStatus("正在加载预览…");
          const response = await fetchWithLocalFallback(`/api/contracts/${encodeURIComponent(id)}/download`);
          if (!response.ok) {
            const text = await response.text().catch(() => "");
            const parsed = (() => {
              try {
                return JSON.parse(text);
              } catch {
                return null;
              }
            })();
            const msg = parsed?.error || parsed?.message || text || `${response.status} ${response.statusText}`;
            throw new Error(msg);
          }
          const buffer = await response.arrayBuffer();
          const wb = await loadWorkbookFromArrayBuffer(buffer);
          openPreviewModal(`预览 - ${name || id}`, wb);
          setStatus("");
        } catch (e) {
          setStatus(`预览失败：${e?.message ?? e}`);
        }
      })();
      return;
    }

    if (action === "delete") {
      (async () => {
        const ok = confirm(`确定删除合同文件：${name || id} ？`);
        if (!ok) return;
        try {
          setStatus("正在删除…");
          const res = await fetchWithLocalFallbackRequest(`/api/contracts/${encodeURIComponent(id)}`, { method: "DELETE" });
          let json = null;
          try {
            json = await res.json();
          } catch {}
          if (!res.ok || !json?.success) {
            const msg = json?.error || `${res.status} ${res.statusText}`;
            throw new Error(msg);
          }
          setStatus("已删除");
          fetchContractsList();
        } catch (e) {
          setStatus(`删除失败：${e?.message ?? e}`);
        }
      })();
    }
  });
}

templateSheetEl.addEventListener("change", async () => {
  if (!state.template?.workbook) return;
  state.merged = state.template.workbook;
  state.mergedSheetName = templateSheetEl.value;
  downloadBtnEl.disabled = true;
  if (saveBtnEl) saveBtnEl.disabled = true;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = true;
    saveNameInputEl.value = "";
  }
  setStatus("");
  
  try {
    await ensureHyperFormula();
    initFormulaEngine(state.merged.getWorksheet(state.mergedSheetName));
  } catch (e) {}
  renderPreview(state.merged, state.mergedSheetName);
});

purchaseSheetEl.addEventListener("change", () => {
  // 修改采购单工作表仅清除合并后的下载状态，但不清空模板预览
  state.merged = state.template?.workbook || null;
  state.mergedSheetName = templateSheetEl.value;
  downloadBtnEl.disabled = true;
  if (saveBtnEl) saveBtnEl.disabled = true;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = true;
    saveNameInputEl.value = "";
  }
  setStatus("");
  
  if (state.merged && state.mergedSheetName) {
    (async () => {
      try {
        await ensureHyperFormula();
        initFormulaEngine(state.merged.getWorksheet(state.mergedSheetName));
      } catch (e) {}
      renderPreview(state.merged, state.mergedSheetName);
    })();
  }
});

// 初始化：获取模板列表
fetchTemplates();
fetchPurchaseOrders();
fetchContractsList();
 
setStatus("");
