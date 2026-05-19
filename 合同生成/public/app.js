 const templateSelectEl = document.getElementById("templateSelect");
const purchaseFileEl = document.getElementById("purchaseFile");
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
  mergedDocxBuffer: null,
  mergedDocxFilename: null,
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

function isWordLikeTemplate(fileType) {
  const t = String(fileType || "").trim().toLowerCase();
  return t.includes("docx") || (t.includes("doc") && !t.includes("docx"));
}

function ensureDocxPreviewCss() {
  const id = "docx-preview-inline-style";
  if (document.getElementById(id)) return;
  const style = document.createElement("style");
  style.id = id;
  style.textContent = `
.docx-wrapper { background: #f1f5f9; padding: 16px; overflow: auto; }
.docx-wrapper, .docx-wrapper * { box-sizing: border-box; }
.docx-wrapper > section.docx { margin: 0 auto 16px; box-shadow: 0 2px 10px rgba(15, 23, 42, 0.08); overflow: visible; }
.docx { max-width: 100%; }
.docx table { border-collapse: collapse; max-width: 100%; }
.docx * { max-width: 100%; }
`;
  document.head.appendChild(style);
}

function stripDocxImages(container) {
  if (!container) return;
  container.querySelectorAll("img").forEach((el) => el.remove());
  container.querySelectorAll("svg image").forEach((el) => el.remove());
  container.querySelectorAll("[style]").forEach((el) => {
    try {
      const bg = String(el.style.backgroundImage || "");
      if (bg.includes("blob:") || bg.includes("data:")) el.style.backgroundImage = "none";
    } catch {}
  });
}

function enableDocxInlineEditing(container) {
  if (!container) return;
  container.dataset.docxEditable = "1";
  const scope = container.querySelector(".docx-wrapper") || container;
  scope.querySelectorAll(".docx p, .docx td, .docx th").forEach((el) => {
    el.setAttribute("contenteditable", "true");
    el.dataset.docxEditableNode = "1";
  });
}

function disableDocxInlineEditing(container) {
  if (!container) return;
  try {
    delete container.dataset.docxEditable;
  } catch {
    container.dataset.docxEditable = "";
  }
  const scope = container.querySelector(".docx-wrapper") || container;
  scope.querySelectorAll('[data-docx-editable-node="1"]').forEach((el) => {
    el.removeAttribute("contenteditable");
    delete el.dataset.docxEditableNode;
  });
}

function escapeXml(text) {
  return String(text ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function updateDocxTcText(tcXml, text) {
  const value = escapeXml(String(text ?? "").replace(/\r\n/g, "\n").replace(/\n/g, " "));
  const tRe = /<w:t(\s+[^>]*)?>([\s\S]*?)<\/w:t>/g;
  const matches = [...String(tcXml ?? "").matchAll(tRe)];
  if (matches.length > 0) {
    let out = String(tcXml ?? "");
    const first = matches[0];
    const full = first[0];
    const attrs = first[1] ?? "";
    const attrFixed = /xml:space=/.test(attrs) ? attrs : `${attrs} xml:space="preserve"`;
    out = out.replace(full, `<w:t${attrFixed}>${value}</w:t>`);
    for (let i = 1; i < matches.length; i += 1) {
      out = out.replace(matches[i][0], `<w:t${matches[i][1] ?? ""}></w:t>`);
    }
    return out;
  }

  if (/<w:p[\s>]/.test(tcXml)) {
    return String(tcXml ?? "").replace(
      /<w:p[\s>]/,
      (m) => `${m}<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`,
    );
  }
  return String(tcXml ?? "").replace(
    /<\/w:tc>/,
    `<w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p></w:tc>`,
  );
}

function updateDocxPText(pXml, text) {
  const value = escapeXml(String(text ?? "").replace(/\r\n/g, "\n").replace(/\n/g, " "));
  const tRe = /<w:t(\s+[^>]*)?>([\s\S]*?)<\/w:t>/g;
  const matches = [...String(pXml ?? "").matchAll(tRe)];
  if (matches.length > 0) {
    let out = String(pXml ?? "");
    const first = matches[0];
    const full = first[0];
    const attrs = first[1] ?? "";
    const attrFixed = /xml:space=/.test(attrs) ? attrs : `${attrs} xml:space="preserve"`;
    out = out.replace(full, `<w:t${attrFixed}>${value}</w:t>`);
    for (let i = 1; i < matches.length; i += 1) {
      out = out.replace(matches[i][0], `<w:t${matches[i][1] ?? ""}></w:t>`);
    }
    return out;
  }
  return String(pXml ?? "").replace(
    /<\/w:p>/,
    `<w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p>`,
  );
}

async function applyDocxEditsToArrayBuffer(originalArrayBuffer, container) {
  const base = originalArrayBuffer;
  if (!base) return base;
  if (!container?.dataset?.docxEditable) return base;
  await ensureDocxPreview();
  if (!globalThis.JSZip) return base;

  const tds = Array.from(container.querySelectorAll(".docx td, .docx th"));
  const ps = Array.from(container.querySelectorAll(".docx p")).filter((p) => !p.closest("td,th"));
  if (tds.length === 0 && ps.length === 0) return base;

  const zip = await globalThis.JSZip.loadAsync(base);
  const docEntry = zip.file("word/document.xml");
  if (!docEntry) return base;
  let documentXml = await docEntry.async("string");

  if (tds.length > 0) {
    let idx = 0;
    const tcRe = /<w:tc[\s\S]*?<\/w:tc>/g;
    documentXml = documentXml.replace(tcRe, (tcXml) => {
      const td = tds[idx++];
      if (!td) return tcXml;
      const text = String(td.innerText ?? "").replace(/\u00A0/g, " ").trim();
      return updateDocxTcText(tcXml, text);
    });
  }

  if (ps.length > 0) {
    const tblRanges = [];
    const tblRe = /<w:tbl[\s\S]*?<\/w:tbl>/g;
    let tm = null;
    while ((tm = tblRe.exec(documentXml))) {
      tblRanges.push({ s: tm.index, e: tm.index + tm[0].length });
    }

    let pIdx = 0;
    let rangeIdx = 0;
    const isInTable = (offset) => {
      while (rangeIdx < tblRanges.length && offset >= tblRanges[rangeIdx].e) rangeIdx += 1;
      if (rangeIdx >= tblRanges.length) return false;
      return offset >= tblRanges[rangeIdx].s && offset < tblRanges[rangeIdx].e;
    };

    const pRe = /<w:p[\s\S]*?<\/w:p>/g;
    documentXml = documentXml.replace(pRe, (pXml, offset) => {
      if (isInTable(Number(offset) || 0)) return pXml;
      const p = ps[pIdx++];
      if (!p) return pXml;
      const text = String(p.innerText ?? "").replace(/\u00A0/g, " ").trim();
      return updateDocxPText(pXml, text);
    });
  }

  zip.file("word/document.xml", documentXml);
  const out = await zip.generateAsync({ type: "arraybuffer" });
  return out;
}

function commitActivePreviewCellEdits() {
  const active = document.activeElement;
  if (!active) return;
  const td = active.closest?.('td[contenteditable="true"][data-row][data-col]');
  if (td && typeof td.blur === "function") td.blur();
}

async function renderDocxWithoutImages(arrayBuffer, container, options) {
  ensureDocxPreviewCss();
  await ensureDocxPreview();
  if (!globalThis.docx?.renderAsync) throw new Error("docx-preview 未就绪");

  const editable = !!(options && options.editable);
  const safeOptions = { ...(options || {}) };
  delete safeOptions.editable;

  const renderOptions = {
    ignoreWidth: true,
    ignoreHeight: true,
    breakPages: true,
    renderHeaders: true,
    renderFooters: true,
    ...safeOptions,
  };

  const urlObj = globalThis.URL;
  const originalCreateObjectURL = urlObj?.createObjectURL ? urlObj.createObjectURL.bind(urlObj) : null;
  if (urlObj && originalCreateObjectURL) {
    urlObj.createObjectURL = (blob) => {
      try {
        const type = String(blob?.type || "").toLowerCase();
        if (type.startsWith("image/")) return "data:,";
      } catch {}
      return originalCreateObjectURL(blob);
    };
  }

  try {
    await globalThis.docx.renderAsync(arrayBuffer, container, null, renderOptions);
  } finally {
    if (urlObj && originalCreateObjectURL) urlObj.createObjectURL = originalCreateObjectURL;
  }

  stripDocxImages(container);
  if (editable) enableDocxInlineEditing(container);
  else disableDocxInlineEditing(container);
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

function fetchPurchaseOrders() {
  if (!purchaseFileEl) return;
  if (purchaseSheetEl) {
    purchaseSheetEl.innerHTML = "";
    purchaseSheetEl.disabled = true;
  }
}

function guessPurchaseTypeIdFromFilename(filename) {
  const base = String(filename ?? "").trim().replace(/\.(xlsx|xls)$/i, "");
  if (!base) return "";
  const m = base.match(/(\d{3})\s*$/);
  return m ? m[1] : "";
}

async function handlePurchaseFileChange() {
  if (!purchaseFileEl) return;
  const file = purchaseFileEl.files?.[0] ?? null;

  state.purchase = null;
  state.mergedDocxBuffer = null;
  state.mergedDocxFilename = null;
  downloadBtnEl.disabled = true;
  if (saveBtnEl) saveBtnEl.disabled = true;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = true;
    saveNameInputEl.value = "";
  }

  if (!file) {
    if (purchaseInfoEl) purchaseInfoEl.textContent = "";
    purchaseSheetEl.innerHTML = "";
    purchaseSheetEl.disabled = true;
    try {
      const fileType = String(state.template?.fileType ?? "");
      if (isWordLikeTemplate(fileType) && state.template?.buffer) {
        previewEl.innerHTML = '<div class="meta">正在渲染 DOCX，请稍候…</div>';
        await renderDocxWithoutImages(state.template.buffer, previewEl, {});
      } else if (state.template?.workbook) {
        state.merged = state.template.workbook;
        state.mergedSheetName = templateSheetEl.value || getFirstSheetName(state.template.workbook);
        if (state.mergedSheetName) {
          try {
            await ensureHyperFormula();
            initFormulaEngine(state.merged.getWorksheet(state.mergedSheetName));
          } catch (e) {
            try {
              if (state.hf) state.hf.destroy();
            } catch {}
            state.hf = null;
            state.hfSheetId = null;
          }
          renderPreview(state.merged, state.mergedSheetName);
        } else {
          previewEl.innerHTML = "";
        }
      }
    } catch {}
    enableActionsIfReady();
    return;
  }

  try {
    setStatus("正在读取订购单…");
    const loadedOut = await loadWorkbookFromFile(file);
    const loaded = loadedOut.workbook;
    const buffer = loadedOut.buffer;

    const typeId = guessPurchaseTypeIdFromFilename(file.name);
    state.purchase = { id: "local", typeId, file, workbook: loaded, buffer };
    fillSheetSelect(purchaseSheetEl, loaded);
    purchaseSheetEl.value = getFirstSheetName(loaded);
    updateFileMeta(purchaseInfoEl, file, loaded);
    if (purchaseInfoEl) {
      const typeHint = typeId ? `；类型ID：${typeId}` : "；类型ID：未识别（文件名末尾需为3位数字）";
      purchaseInfoEl.textContent = `${purchaseInfoEl.textContent}${typeHint}`;
    }
    setStatus(typeId ? "订购单已就绪" : "订购单已就绪：文件名末尾需为3位数字以匹配映射配置");

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
  state.mergedDocxBuffer = null;
  state.mergedDocxFilename = null;
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
    
    const opt = templateSelectEl.options[templateSelectEl.selectedIndex];
    const fileType = String(opt?.dataset?.fileType ?? "");
    const ext = guessExtFromFileType(fileType);
    const fileName = `${opt?.text || "合同模板"}${ext}`;
    
    if (isWordLikeTemplate(fileType)) {
      state.template = {
        id: templateId,
        fileType,
        file: { name: fileName },
        workbook: null,
        buffer,
      };

      templateSheetEl.innerHTML = "";
      templateSheetEl.disabled = true;
      templateInfoEl.textContent = `文件：${fileName}；类型：${ext.replace(/^\./, "").toUpperCase()}`;
      if (String(fileType).toLowerCase().includes("doc") && !String(fileType).toLowerCase().includes("docx")) {
        setStatus("合同模板已就绪（DOC）");
        previewEl.innerHTML = '<div class="meta">DOC 格式暂不支持预览，请另存为 DOCX。</div>';
        enableActionsIfReady();
        return;
      }

      setStatus("正在渲染 Word 模板预览…");
      previewEl.innerHTML = '<div class="meta">正在渲染 DOCX，请稍候…</div>';
      try {
        await renderDocxWithoutImages(buffer, previewEl, { editable: false });
        setStatus("合同模板已就绪（DOCX）");
      } catch (err) {
        previewEl.innerHTML = `<div class="meta">DOCX 预览失败：${String(err?.message ?? err)}</div>`;
        setStatus("合同模板已就绪（DOCX）");
      }
      enableActionsIfReady();
      return;
    }

    const loaded = await loadWorkbookFromArrayBuffer(buffer);
    state.template = { 
      id: templateId,
      fileType,
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
    } catch (e) {
      try {
        if (state.hf) state.hf.destroy();
      } catch {}
      state.hf = null;
      state.hfSheetId = null;
    }
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

function getSingleRowHorizontalMergeSpansAtRow(mergeRanges, row1, minCol, maxCol) {
  if (!Array.isArray(mergeRanges) || !row1) return [];
  const spans = [];
  for (const rngStr of mergeRanges) {
    const decoded = decodeRangeAddr(rngStr);
    if (!decoded) continue;
    if (decoded.s.r !== row1 || decoded.e.r !== row1) continue;
    const sC = decoded.s.c;
    const eC = decoded.e.c;
    if (eC < minCol || sC > maxCol) continue;
    spans.push({ sC, eC });
  }
  spans.sort((a, b) => a.sC - b.sC || a.eC - b.eC);
  return spans;
}

function applyHorizontalMergeSpansToRows(worksheet, spans, startRow1, rowCount) {
  if (!worksheet || !Array.isArray(spans) || spans.length === 0) return;
  if (!startRow1 || !rowCount) return;
  for (let i = 0; i < rowCount; i += 1) {
    const r = startRow1 + i;
    for (const sp of spans) {
      if (!sp || !sp.sC || !sp.eC || sp.eC <= sp.sC) continue;
      const addr = encodeRangeAddr(r, sp.sC, r, sp.eC);
      try {
        worksheet.mergeCells(addr);
      } catch (e) {}
    }
  }
}

function restoreMergesAfterInsert(worksheet, originalMerges, thresholdRow1, deltaRows) {
  if (!originalMerges || !Array.isArray(originalMerges)) return;

  // 1. 获取当前工作表所有的合并区域并备份
  const currentMerges = getWorksheetMergeRanges(worksheet);
  
  // 2. 解除所有可能受影响的合并区域
  // 只要合并区域的结束行在阈值之后，或者与插入区域有交集，就解除
  for (const rngStr of currentMerges) {
    const decoded = decodeRangeAddr(rngStr);
    if (decoded && decoded.e.r >= thresholdRow1) {
      try {
        worksheet.unMergeCells(rngStr);
      } catch (e) {}
    }
  }

  // 3. 根据原始模板的合并规则，重新计算并应用合并
  for (const orig of originalMerges) {
    const adjusted = adjustRangeByInsertRows(orig, thresholdRow1, deltaRows);
    if (!adjusted) continue;
    
    const adjDecoded = decodeRangeAddr(adjusted);
    if (!adjDecoded) continue;

    // 如果该合并区域完全在受影响区域之前，说明刚才没被解除，跳过
    if (adjDecoded.e.r < thresholdRow1) continue;

    try {
      // 尝试重新合并
      worksheet.mergeCells(adjusted);
      // 注意：这里不再调用 clearNonMasterCellsForMerge
      // 因为在恢复过程中，非主单元格可能已经包含了刚刚填入的数据或原有内容
      // 强制清空会导致内容丢失。Excel 渲染时会自动处理合并单元格的显示。
    } catch (e) {
      console.warn("恢复合并失败:", adjusted, e);
    }
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

function adjustFormulaByOffset(formula, deltaRows, deltaCols) {
  if (!formula || typeof formula !== "string") return formula;
  if (!deltaRows && !deltaCols) return formula;

  const shiftSegment = (segment) => {
    return segment.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, colAbs, col, rowAbs, rowDigits, offset) => {
      const before = offset > 0 ? segment[offset - 1] : "";
      const after = offset + m.length < segment.length ? segment[offset + m.length] : "";
      if (before && /[A-Z0-9_.]/i.test(before)) return m;
      if (after && /[A-Z0-9_]/i.test(after)) return m;

      let colPart = `${colAbs}${col}`;
      let rowPart = `${rowAbs}${rowDigits}`;

      if (deltaCols && colAbs !== "$") {
        const colNum = decodeCol(col);
        const nextCol = colNum + deltaCols;
        if (nextCol >= 1) colPart = `${colAbs}${encodeCol(nextCol)}`;
      }

      if (deltaRows && rowAbs !== "$") {
        const rowNum = Number(rowDigits);
        const nextRow = rowNum + deltaRows;
        if (Number.isFinite(nextRow) && nextRow >= 1) rowPart = `${rowAbs}${nextRow}`;
      }

      return `${colPart}${rowPart}`;
    });
  };

  let out = "";
  let inString = false;
  let start = 0;
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch !== '"') {
      i += 1;
      continue;
    }

    if (!inString) {
      out += shiftSegment(formula.slice(start, i));
      inString = true;
      start = i;
      i += 1;
      continue;
    }

    if (i + 1 < formula.length && formula[i + 1] === '"') {
      i += 2;
      continue;
    }

    i += 1;
    out += formula.slice(start, i);
    inString = false;
    start = i;
  }

  const tail = formula.slice(start);
  out += inString ? tail : shiftSegment(tail);
  return out;
}

function initFormulaEngine(ws) {
  if (!globalThis.HyperFormula) return;
  if (!ws) {
    try {
      if (state.hf) state.hf.destroy();
    } catch (e) {}
    state.hf = null;
    state.hfSheetId = null;
    return;
  }

  const aoa = [];
  const rowCount = Math.max(1, Math.max(ws.rowCount || 0, ws.actualRowCount || 0));
  const colCount = Math.max(1, Math.max(ws.columnCount || 0, ws.actualColumnCount || 0));

  for (let r = 1; r <= rowCount; r++) {
    const rowValues = [];
    const wsRow = ws.getRow(r);
    for (let c = 1; c <= colCount; c++) {
      const cell = wsRow.getCell(c);
      const val = cell.value;
      if (val && typeof val === "object" && (val.formula || val.sharedFormula)) {
        // 在存入公式引擎前进行清洗
        let formula = val.formula;
        if (!formula && val.sharedFormula) {
          const masterAddr = String(val.sharedFormula);
          const masterPos = decodeCellAddr(masterAddr);
          const masterCell = ws.getCell(masterAddr);
          const masterVal = masterCell?.value;
          const masterFormula =
            masterVal && typeof masterVal === "object" && masterVal.formula ? String(masterVal.formula) : "";
          if (masterFormula && masterPos) {
            const deltaR = r - masterPos.r;
            const deltaC = c - masterPos.c;
            formula = adjustFormulaByOffset(masterFormula, deltaR, deltaC);
          } else if (masterFormula) {
            formula = masterFormula;
          }
        }
        const sanitized = sanitizeFormulaForHF(formula, r);
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
  const sheetNames = state.hf.getSheetNames();
  const firstSheetName = Array.isArray(sheetNames) && sheetNames.length > 0 ? sheetNames[0] : null;
  state.hfSheetId = firstSheetName ? state.hf.getSheetId(firstSheetName) : null;
}

function getCellValueWithFormula(r, c) {
  if (state.hf && state.hfSheetId !== null && state.hfSheetId !== undefined) {
    // HyperFormula 使用 0-based index
    let val = null;
    try {
      val = state.hf.getCellValue({
        sheet: state.hfSheetId,
        row: r - 1,
        col: c - 1,
      });
    } catch (e) {
      return null;
    }
    
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
  const diff = newRow - oldRow;
  if (diff === 0) return formula;

  const shiftSegment = (segment) => {
    return segment.replace(/(\$?)([A-Z]{1,3})(\$?)(\d+)/g, (m, colAbs, col, rowAbs, rowDigits, offset) => {
      const before = offset > 0 ? segment[offset - 1] : "";
      const after = offset + m.length < segment.length ? segment[offset + m.length] : "";
      if (before && /[A-Z0-9_.]/i.test(before)) return m;
      if (after && /[A-Z0-9_]/i.test(after)) return m;

      if (rowAbs === "$") return m;
      const rowNum = Number(rowDigits);
      if (!Number.isFinite(rowNum)) return m;
      const nextRow = rowNum + diff;
      if (nextRow < 1) return m;
      return `${colAbs}${col}${rowAbs}${nextRow}`;
    });
  };

  let out = "";
  let inString = false;
  let start = 0;
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch !== '"') {
      i += 1;
      continue;
    }

    if (!inString) {
      out += shiftSegment(formula.slice(start, i));
      inString = true;
      start = i;
      i += 1;
      continue;
    }

    if (i + 1 < formula.length && formula[i + 1] === '"') {
      i += 2;
      continue;
    }

    i += 1;
    out += formula.slice(start, i);
    inString = false;
    start = i;
  }

  const tail = formula.slice(start);
  out += inString ? tail : shiftSegment(tail);
  return out;
}

function adjustFormulaByRowThreshold(formula, thresholdRow1, deltaRows) {
  if (!formula || typeof formula !== "string") return formula;
  if (!deltaRows) return formula;
  const threshold = Number(thresholdRow1);
  if (!Number.isFinite(threshold) || threshold < 1) return formula;

  const parseCellRef = (ref) => {
    const m = String(ref || "").match(/^(\$?)([A-Z]{1,3})(\$?)(\d+)$/);
    if (!m) return null;
    return {
      colAbs: m[1] || "",
      col: m[2] || "",
      rowAbs: m[3] || "",
      row: Number(m[4]),
    };
  };

  const encodeCellRef = (p) => `${p.colAbs}${p.col}${p.rowAbs}${p.row}`;

  const adjustRowForInsert = (row, rowAbs, insertAtRow1, delta, mode) => {
    if (rowAbs === "$") return row;
    const r = Number(row);
    if (!Number.isFinite(r)) return row;
    if (mode === "shift") {
      if (r < insertAtRow1) return r;
      return Math.max(1, r + delta);
    }
    return r;
  };

  const adjustRangeForInsert = (a, b) => {
    const s = parseCellRef(a);
    const e = parseCellRef(b);
    if (!s || !e) return `${a}:${b}`;

    const sRow = s.rowAbs === "$" ? s.row : s.row;
    const eRow = e.rowAbs === "$" ? e.row : e.row;
    const insertAt = threshold;
    const delta = deltaRows;

    if (s.rowAbs !== "$" && sRow >= insertAt) s.row = Math.max(1, sRow + delta);
    if (e.rowAbs !== "$") {
      if (eRow >= insertAt) e.row = Math.max(1, eRow + delta);
      else if (eRow === insertAt - 1 && sRow < insertAt) e.row = Math.max(1, eRow + delta);
    }

    return `${encodeCellRef(s)}:${encodeCellRef(e)}`;
  };

  const shiftSegment = (segment) => {
    const reCell = /(\$?[A-Z]{1,3}\$?\d+)/g;
    let out = "";
    let i = 0;
    while (i < segment.length) {
      const m = reCell.exec(segment);
      if (!m) {
        out += segment.slice(i);
        break;
      }
      const start = m.index;
      const ref = m[1];
      out += segment.slice(i, start);

      const before = start > 0 ? segment[start - 1] : "";
      const after = start + ref.length < segment.length ? segment[start + ref.length] : "";
      if (before && /[A-Z0-9_.]/i.test(before)) {
        out += ref;
        i = start + ref.length;
        continue;
      }

      if (after === ":") {
        const m2 = reCell.exec(segment);
        if (m2 && m2.index === start + ref.length + 1) {
          const ref2 = m2[1];
          out += adjustRangeForInsert(ref, ref2);
          i = m2.index + ref2.length;
          continue;
        }
        reCell.lastIndex = start + ref.length;
      }

      if (after && /[A-Z0-9_]/i.test(after)) {
        out += ref;
        i = start + ref.length;
        continue;
      }

      const p = parseCellRef(ref);
      if (!p) {
        out += ref;
        i = start + ref.length;
        continue;
      }
      p.row = adjustRowForInsert(p.row, p.rowAbs, threshold, deltaRows, "shift");
      out += encodeCellRef(p);
      i = start + ref.length;
    }
    return out;
  };

  let out = "";
  let inString = false;
  let start = 0;
  let i = 0;
  while (i < formula.length) {
    const ch = formula[i];
    if (ch !== '"') {
      i += 1;
      continue;
    }

    if (!inString) {
      out += shiftSegment(formula.slice(start, i));
      inString = true;
      start = i;
      i += 1;
      continue;
    }

    if (i + 1 < formula.length && formula[i + 1] === '"') {
      i += 2;
      continue;
    }

    i += 1;
    out += formula.slice(start, i);
    inString = false;
    start = i;
  }

  const tail = formula.slice(start);
  out += inString ? tail : shiftSegment(tail);
  return out;
}

function updateWorksheetFormulasAfterRowChange(worksheet, thresholdRow1, deltaRows, skipStartRow1, skipRowCount) {
  if (!worksheet) return;
  if (!deltaRows) return;

  const maxRow = Math.max(worksheet.rowCount || 0, worksheet.actualRowCount || 0);
  const maxCol = Math.max(worksheet.columnCount || 0, worksheet.actualColumnCount || 0);

  const skipStart = Number(skipStartRow1);
  const skipEnd = Number.isFinite(skipStart) && skipRowCount ? skipStart + skipRowCount - 1 : null;

  const getCellFormulaText = (ws, r1, c1, v) => {
    if (!v || typeof v !== "object") return "";
    if (v.formula) return String(v.formula);
    if (!v.sharedFormula) return "";
    const masterAddr = String(v.sharedFormula);
    const masterPos = decodeCellAddr(masterAddr);
    const masterCell = ws.getCell(masterAddr);
    const masterVal = masterCell?.value;
    const masterFormula = masterVal && typeof masterVal === "object" && masterVal.formula ? String(masterVal.formula) : "";
    if (!masterFormula) return "";
    if (!masterPos) return masterFormula;
    const deltaR = r1 - masterPos.r;
    const deltaC = c1 - masterPos.c;
    return adjustFormulaByOffset(masterFormula, deltaR, deltaC);
  };

  for (let r = 1; r <= maxRow; r += 1) {
    if (skipEnd !== null && r >= skipStart && r <= skipEnd) continue;
    const row = worksheet.getRow(r);
    for (let c = 1; c <= maxCol; c += 1) {
      const cell = row.getCell(c);
      const v = cell?.value;
      if (!v || typeof v !== "object" || (!v.formula && !v.sharedFormula)) continue;
      const baseFormula = getCellFormulaText(worksheet, r, c, v);
      if (!baseFormula) continue;
      const nextFormula = adjustFormulaByRowThreshold(baseFormula, thresholdRow1, deltaRows);
      if (nextFormula === baseFormula) continue;
      const nextValue = { ...v, formula: nextFormula };
      if (nextValue.sharedFormula !== undefined) delete nextValue.sharedFormula;
      cell.value = nextValue;
    }
  }
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

  const baseDetailRow1 = contractHeaderRow.rowIndex + 1;

  const isManualContentCell = (cell) => {
    if (!cell) return false;
    const effectiveCell = cell.isMerged && cell.master ? cell.master : cell;
    const v = effectiveCell.value;
    if (v === null || v === undefined) return false;
    if (typeof v === "object") {
      if (v.formula !== undefined || v.sharedFormula !== undefined) return false;
      return true;
    }
    return String(v).trim() !== "";
  };

  const hasManualContentInRow = (rowIndex1) => {
    const row = templateWs.getRow(rowIndex1);
    for (let c = originCol; c <= maxAssignedCol; c += 1) {
      if (isManualContentCell(row.getCell(c))) return true;
    }
    return false;
  };

  const getInsertAtRow1 = () => {
    if (!hasManualContentInRow(baseDetailRow1)) return baseDetailRow1;

    const maxRow = Math.max(templateWs.rowCount || 0, templateWs.actualRowCount || 0) || baseDetailRow1;
    const maxScan = Math.min(maxRow, baseDetailRow1 + 5000);

    let lastContentRow1 = baseDetailRow1;
    for (let r = baseDetailRow1; r <= maxScan; r += 1) {
      const has = hasManualContentInRow(r);
      if (has) {
        lastContentRow1 = r;
        continue;
      }
      return r;
    }
    return lastContentRow1 + 1;
  };

  const insertAtRow1 = getInsertAtRow1();

  const styleSourceRowAtInsert = templateWs.getRow(insertAtRow1);
  const styleSourceRowAtBase = templateWs.getRow(baseDetailRow1);
  const styleByCol = new Map();
  const formulaByCol = new Map();
  let fallbackStyle = null;
  
  for (let c = originCol; c <= maxAssignedCol; c += 1) {
    const cellAtInsert = styleSourceRowAtInsert.getCell(c);
    const cellAtBase = styleSourceRowAtBase.getCell(c);
    const pickedStyleCell =
      cellAtInsert?.style && Object.keys(cellAtInsert.style).length > 0 ? cellAtInsert : cellAtBase;
    if (pickedStyleCell?.style && Object.keys(pickedStyleCell.style).length > 0) {
      const s = deepClone(pickedStyleCell.style);
      styleByCol.set(c, s);
      if (!fallbackStyle) fallbackStyle = s;
    }
    // 检查并记录公式
    const findFormulaTemplateInCol = () => {
      const maxRow = Math.max(templateWs.rowCount || 0, templateWs.actualRowCount || 0) || baseDetailRow1;
      const end = Math.min(maxRow, baseDetailRow1 + 60);
      const preferred = [
        { row1: insertAtRow1, cell: cellAtInsert },
        { row1: baseDetailRow1, cell: cellAtBase },
      ];
      const tryCell = (row1, cell) => {
        const v = cell?.value;
        if (!v || typeof v !== "object") return null;
        if (v.formula) return { formula: String(v.formula), sourceRow1: row1 };
        if (!v.sharedFormula) return null;
        const masterAddr = String(v.sharedFormula);
        const masterPos = decodeCellAddr(masterAddr);
        const masterCell = templateWs.getCell(masterAddr);
        const masterVal = masterCell?.value;
        const masterFormula =
          masterVal && typeof masterVal === "object" && masterVal.formula ? String(masterVal.formula) : "";
        if (!masterFormula) return null;
        if (!masterPos) return { formula: masterFormula, sourceRow1: row1 };
        const deltaR = row1 - masterPos.r;
        const deltaC = c - masterPos.c;
        return { formula: adjustFormulaByOffset(masterFormula, deltaR, deltaC), sourceRow1: row1 };
      };

      for (const p of preferred) {
        const hit = tryCell(p.row1, p.cell);
        if (hit) return hit;
      }
      for (let r = baseDetailRow1; r <= end; r += 1) {
        const cell = templateWs.getRow(r).getCell(c);
        const hit = tryCell(r, cell);
        if (hit) return hit;
      }
      return null;
    };

    const tpl = findFormulaTemplateInCol();
    if (tpl) formulaByCol.set(c, tpl);
  }
  const rowHeight = styleSourceRowAtInsert.height ?? styleSourceRowAtBase.height;

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

  let blankTemplateRowsCount = 0;
  const maxSearchRows = 200;
  const footerTextRe = /(合计|小计|总计|金额合计|总价|备注|说明|条款|结算|税)/;
  const rowLooksLikePlaceholder = (rowIndex1) => {
    const row = templateWs.getRow(rowIndex1);
    let hasFooterText = false;
    let hasNonFormulaData = false;

    for (let c = originCol; c <= maxAssignedCol; c += 1) {
      const cell = row.getCell(c);
      if (!cell) continue;
      const effectiveCell = cell.isMerged && cell.master ? cell.master : cell;
      const val = effectiveCell.value;
      if (val === null || val === undefined) continue;

      if (typeof val === "object") {
        if (val.formula !== undefined || val.sharedFormula !== undefined) continue;
        const t = normalizeText(getCellText(effectiveCell));
        if (t) {
          if (footerTextRe.test(t)) hasFooterText = true;
          else hasNonFormulaData = true;
        }
        continue;
      }

      if (typeof val === "number") {
        if (val !== 0) hasNonFormulaData = true;
        continue;
      }

      const t = normalizeText(String(val));
      if (t) {
        if (footerTextRe.test(t)) hasFooterText = true;
        else hasNonFormulaData = true;
      }
    }

    if (hasFooterText) return false;
    if (hasNonFormulaData) return false;
    return true;
  };

  for (let i = 0; i < maxSearchRows; i += 1) {
    const rowIndex1 = insertAtRow1 + i;
    if (!rowLooksLikePlaceholder(rowIndex1)) break;
    blankTemplateRowsCount += 1;
  }

  let netDelta = 0;
  let thresholdRow1 = insertAtRow1;
  let skipStartRow1 = insertAtRow1;
  let skipRowCount = 0;
  const newRowsNeedingFormula = new Set();
  for (let i = 0; i < deltaRows; i += 1) newRowsNeedingFormula.add(insertAtRow1 + i);

  if (blankTemplateRowsCount > 0) {
    const extraCount = Math.max(0, deltaRows - blankTemplateRowsCount);
    if (extraCount > 0) {
      const extraInsertAtRow1 = insertAtRow1 + blankTemplateRowsCount;
      const blank = new Array(maxAssignedCol).fill(null);
      const blanks = [];
      for (let i = 0; i < extraCount; i += 1) blanks.push(blank.slice());
      templateWs.spliceRows(extraInsertAtRow1, 0, ...blanks);
      netDelta = extraCount;
      thresholdRow1 = extraInsertAtRow1;
      skipStartRow1 = extraInsertAtRow1;
      skipRowCount = extraCount;
    } else {
      netDelta = 0;
      skipRowCount = 0;
    }

    for (let i = 0; i < deltaRows; i += 1) {
      const rowNum = insertAtRow1 + i;
      const row = templateWs.getRow(rowNum);
      const values = insertRowValuesList[i];
      if (!Array.isArray(values)) continue;
      for (let c = originCol; c <= maxAssignedCol; c += 1) {
        const v = values[c - 1];
        if (v !== null && v !== undefined) row.getCell(c).value = v;
      }
    }
  } else {
    const blank = new Array(maxAssignedCol).fill(null);
    const blanks = [];
    for (let i = 0; i < deltaRows; i += 1) blanks.push(blank.slice());
    templateWs.spliceRows(insertAtRow1, 0, ...blanks);
    netDelta = deltaRows;
    thresholdRow1 = insertAtRow1;
    skipStartRow1 = insertAtRow1;
    skipRowCount = deltaRows;

    for (let i = 0; i < deltaRows; i += 1) {
      const rowNum = insertAtRow1 + i;
      const row = templateWs.getRow(rowNum);
      const values = insertRowValuesList[i];
      if (!Array.isArray(values)) continue;
      for (let c = originCol; c <= maxAssignedCol; c += 1) {
        const v = values[c - 1];
        if (v !== null && v !== undefined) row.getCell(c).value = v;
      }
    }
  }

  for (let i = 0; i < insertRowValuesList.length; i += 1) {
    const rowNum = insertAtRow1 + i;
    const row = templateWs.getRow(rowNum);
    if (rowHeight !== undefined) row.height = rowHeight;

    for (let c = originCol; c <= maxAssignedCol; c += 1) {
      const cell = row.getCell(c);

      const style = styleByCol.get(c) ?? fallbackStyle;
      if (style) cell.style = deepClone(style);

      const formulaEntry = formulaByCol.get(c);
      if (!formulaEntry) continue;
      const existing = cell?.value;
      const hasExistingFormula =
        existing && typeof existing === "object" && (existing.formula !== undefined || existing.sharedFormula !== undefined);
      if (hasExistingFormula) continue;
      if (newRowsNeedingFormula.size > 0 && !newRowsNeedingFormula.has(rowNum)) continue;
      const newFormula = adjustFormula(formulaEntry.formula, formulaEntry.sourceRow1, rowNum);
      cell.value = { formula: newFormula };
    }
  }

  if (netDelta) {
    updateWorksheetFormulasAfterRowChange(templateWs, thresholdRow1, netDelta, skipStartRow1, skipRowCount);
    restoreMergesAfterInsert(templateWs, originalMerges, thresholdRow1, netDelta);
  }

  const detailRowMergeSpans = getSingleRowHorizontalMergeSpansAtRow(
    originalMerges,
    baseDetailRow1,
    originCol,
    maxAssignedCol,
  );
  applyHorizontalMergeSpansToRows(templateWs, detailRowMergeSpans, insertAtRow1, insertRowValuesList.length);

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
  refreshFormulaCells(container);
}

function refreshFormulaCells(container = previewEl) {
  if (!state.hf) return;
  if (!container) return;
  const cells = container.querySelectorAll('td.formula-cell[data-row][data-col]');
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

function buildMergeMapsByScan(ws, endRow, endCol) {
  const keyOf = (r, c) => `${r},${c}`;
  const mergeBounds = new Map();

  for (let r = 1; r <= endRow; r += 1) {
    const row = ws.getRow(r);
    for (let c = 1; c <= endCol; c += 1) {
      const cell = row.getCell(c);
      if (!cell || !cell.isMerged) continue;
      const master = cell.master;
      if (!master || !master.address) continue;
      const masterPos = decodeCellAddr(master.address);
      if (!masterPos) continue;

      const mk = keyOf(masterPos.r, masterPos.c);
      const b = mergeBounds.get(mk) ?? { sR: masterPos.r, sC: masterPos.c, eR: masterPos.r, eC: masterPos.c };
      if (r < b.sR) b.sR = r;
      if (c < b.sC) b.sC = c;
      if (r > b.eR) b.eR = r;
      if (c > b.eC) b.eC = c;
      mergeBounds.set(mk, b);
    }
  }

  const mergeMaster = new Map();
  const mergeSkip = new Set();

  for (const b of mergeBounds.values()) {
    const sR = Math.max(1, b.sR);
    const sC = Math.max(1, b.sC);
    const eR = Math.min(endRow, b.eR);
    const eC = Math.min(endCol, b.eC);
    const rr = eR - sR + 1;
    const cc = eC - sC + 1;
    if (rr <= 1 && cc <= 1) continue;

    mergeMaster.set(keyOf(sR, sC), { rowSpan: rr, colSpan: cc });
    for (let r = sR; r <= eR; r += 1) {
      for (let c = sC; c <= eC; c += 1) {
        if (r === sR && c === sC) continue;
        mergeSkip.add(keyOf(r, c));
      }
    }
  }

  return { keyOf, mergeMaster, mergeSkip };
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
  if (useFormulaEngine && !state.hf && globalThis.HyperFormula) {
    try {
      initFormulaEngine(ws);
    } catch {
      state.hf = null;
      state.hfSheetId = null;
    }
  }

  const maxRows = 60;
  const maxCols = 20;
  const endRow = Math.min(Math.max(ws.rowCount || 0, ws.actualRowCount || 0) || 1, maxRows);
  const endCol = Math.min(Math.max(ws.columnCount || 0, ws.actualColumnCount || 0) || 1, maxCols);

  const { keyOf, mergeMaster, mergeSkip } = buildMergeMapsByScan(ws, endRow, endCol);

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
        const isFormula =
          cell?.value &&
          typeof cell.value === "object" &&
          (cell.value.formula !== undefined || cell.value.sharedFormula !== undefined);
        
        if (isFormula) {
          td.contentEditable = "false";
          td.classList.add("formula-cell");
          td.title = "公式单元格，无法手动修改";
        } else {
          td.contentEditable = "true";

          let inputTimer = null;
          const scheduleRecalc = (rowNum, colNum, value) => {
            if (!useFormulaEngine) return;
            if ((!state.hf || state.hfSheetId === null) && globalThis.HyperFormula) {
              try {
                initFormulaEngine(state.merged?.getWorksheet?.(state.mergedSheetName) ?? ws);
              } catch {
                state.hf = null;
                state.hfSheetId = null;
              }
            }
            if (!state.hf || state.hfSheetId === null) return;
            try {
              state.hf.setCellContents(
                { sheet: state.hfSheetId, row: rowNum - 1, col: colNum - 1 },
                [[value]],
              );
            } catch (e) {
              return;
            }
            if (inputTimer) clearTimeout(inputTimer);
            inputTimer = setTimeout(() => refreshFormulaCells(container), 80);
          };

          td.addEventListener("input", (e) => {
            const rowNum = parseInt(e.target.dataset.row);
            const colNum = parseInt(e.target.dataset.col);
            const raw = String(e.target.innerText ?? "").replace(/\r\n/g, "\n");
            const trimmed = raw.trim();
            let v = trimmed;
            const num = Number(trimmed);
            if (trimmed !== "" && !Number.isNaN(num)) v = num;
            scheduleRecalc(rowNum, colNum, v);
          });
          
          td.addEventListener("blur", (e) => {
            const rowNum = parseInt(e.target.dataset.row);
            const colNum = parseInt(e.target.dataset.col);
            const raw = String(e.target.innerText ?? "").replace(/\r\n/g, "\n");
            const newValue = raw.trim();
            
            const targetSheet = state.merged.getWorksheet(state.mergedSheetName);
            if (targetSheet) {
              const cell = targetSheet.getRow(rowNum).getCell(colNum);
              
              const num = Number(newValue);
              if (newValue !== "" && !isNaN(num)) {
                cell.value = num;
              } else {
                cell.value = raw;
                if (raw.includes("\n")) {
                  cell.alignment = { wrapText: true };
                }
              }

              if (useFormulaEngine && state.hf && state.hfSheetId !== null) {
                state.hf.setCellContents({
                  sheet: state.hfSheetId,
                  row: rowNum - 1,
                  col: colNum - 1
                }, [[cell.value]]);
                
                refreshFormulaCells(container);
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

      if (!displayValue) td.classList.add("empty");
      else td.classList.remove("empty");
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
  const purchaseTypeId = String(state.purchase?.typeId ?? "").trim();
  const templateReady = !!(state.template?.workbook || isWordLikeTemplate(state.template?.fileType));
  const ready = !!(templateReady && state.purchase?.workbook && purchaseTypeId);
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
  const templateNameRaw = state.template?.file?.name?.replace(/\.[A-Za-z0-9]+$/i, "") || "合同模板";
  const templateName = sanitizeFilenamePart(templateNameRaw) || "合同模板";
  const purchaseNameRaw = state.purchase?.file?.name?.replace(/\.(xlsx|xls)$/i, "") || "";
  const purchaseName = sanitizeFilenamePart(purchaseNameRaw);
  return purchaseName ? `${templateName}-${purchaseName}` : `${templateName}-已插入订购单`;
}

function ensureFilenameWithExt(name, extWithDot) {
  const ext = String(extWithDot ?? "").trim();
  const s = String(name ?? "").trim();
  if (!s) return `合同${ext || ".xlsx"}`;
  if (ext && new RegExp(`\\${ext}$`, "i").test(s)) return s;
  const stripped = s.replace(/\.[A-Za-z0-9]+$/i, "");
  return `${stripped}${ext || ".xlsx"}`;
}

function getActiveOutputExtWithDot() {
  if (state.mergedDocxBuffer) return ".docx";
  const t = String(state.template?.fileType ?? "").toLowerCase();
  if (t.includes("docx") || (t.includes("doc") && !t.includes("docx"))) return ".docx";
  return ".xlsx";
}

function getEffectiveOutputFilename() {
  const typed = sanitizeFilenamePart(saveNameInputEl?.value);
  const base = typed ? typed.replace(/\.[A-Za-z0-9]+$/i, "") : computeDefaultOutputBaseName();
  return ensureFilenameWithExt(base, getActiveOutputExtWithDot());
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
  const ext = String(filename ?? "").trim().toLowerCase().match(/\.([a-z0-9]+)$/)?.[1] ?? "";
  const type =
    ext === "docx"
      ? "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const blob = new Blob([arrayBuffer], {
    type,
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
    const filePath = row?.file_path ?? "";
    const createdAt = formatDateTime(row?.created_at);
    const id = String(row?.id ?? "");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${String(name)}</td>
      <td>${createdAt}</td>
      <td>
        <div class="opBtns">
          <button class="miniBtn" data-action="preview" data-id="${id}" data-name="${String(name)}" data-file-path="${String(
            filePath,
          )}">预览</button>
          <button class="miniBtn primary" data-action="download" data-id="${id}" data-name="${String(name)}" data-file-path="${String(
            filePath,
          )}">下载</button>
          <button class="miniBtn danger" data-action="delete" data-id="${id}" data-name="${String(name)}" data-file-path="${String(
            filePath,
          )}">删除</button>
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
  if (modalSheetSelectEl) {
    modalSheetSelectEl.innerHTML = "";
    modalSheetSelectEl.style.display = "";
  }
  if (modalPreviewEl) modalPreviewEl.innerHTML = "";
}

function openPreviewModal(title, workbook) {
  if (!previewModalEl || !modalPreviewEl) return;
  modalState.workbook = workbook || null;
  modalState.title = String(title || "预览");
  if (modalTitleEl) modalTitleEl.textContent = modalState.title;

  if (modalSheetSelectEl) {
    modalSheetSelectEl.style.display = "";
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

function ensureDocxPreview() {
  const loadScript = (url, dataAttrKey, timeoutMs = 12000) =>
    new Promise((resolve, reject) => {
      const existing = document.querySelector(`script[${dataAttrKey}="${url}"]`);
      if (existing) {
        existing.addEventListener("load", () => resolve(), { once: true });
        existing.addEventListener("error", () => reject(new Error(`加载失败：${url}`)), { once: true });
        return;
      }
      const script = document.createElement("script");
      script.src = url;
      script.async = true;
      script.setAttribute(dataAttrKey, url);

      const timer = setTimeout(() => {
        try {
          script.remove();
        } catch {}
        reject(new Error(`加载超时：${url}`));
      }, timeoutMs);

      script.onload = () => {
        clearTimeout(timer);
        resolve();
      };
      script.onerror = () => {
        clearTimeout(timer);
        reject(new Error(`加载失败：${url}`));
      };
      document.head.appendChild(script);
    });

  const ensureJsZip = async () => {
    if (globalThis.JSZip) return;
    const urls = [
      "https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js",
      "https://unpkg.com/jszip@3.10.1/dist/jszip.min.js",
    ];
    let lastErr = null;
    for (const url of urls) {
      try {
        await loadScript(url, "data-jszip-src");
        if (globalThis.JSZip) return;
      } catch (e) {
        lastErr = e;
      }
    }
    throw lastErr || new Error("JSZip 加载失败");
  };

  return (async () => {
    if (globalThis.docx?.renderAsync) return;
    await ensureJsZip();

    const urls = [
      "https://cdn.jsdelivr.net/npm/docx-preview-lib@0.1.14-fix-3/dist/docx-preview.min.js",
      "https://unpkg.com/docx-preview-lib@0.1.14-fix-3/dist/docx-preview.min.js",
    ];
    let lastErr = null;
    for (const url of urls) {
      try {
        await loadScript(url, "data-docx-preview-src");
        if (globalThis.docx?.renderAsync) return;
      } catch (e) {
        lastErr = e;
      }
    }
    throw lastErr || new Error("docx-preview 加载失败");
  })();
}

async function openDocxPreviewModal(title, arrayBuffer) {
  if (!previewModalEl || !modalPreviewEl) return;
  modalState.workbook = null;
  modalState.sheetName = null;
  modalState.title = String(title || "预览");
  if (modalTitleEl) modalTitleEl.textContent = modalState.title;
  if (modalSheetSelectEl) {
    modalSheetSelectEl.innerHTML = "";
    modalSheetSelectEl.style.display = "none";
  }
  modalPreviewEl.innerHTML = "";
  previewModalEl.classList.add("open");
  previewModalEl.setAttribute("aria-hidden", "false");

  await renderDocxWithoutImages(arrayBuffer, modalPreviewEl, { editable: false });
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

function coerceToArrayMaybeJson(value) {
  if (Array.isArray(value)) return value;
  if (typeof value === "string") {
    const s = value.trim();
    if (!s) return null;
    try {
      const parsed = JSON.parse(s);
      if (Array.isArray(parsed)) return parsed;
    } catch {}
    const parts = s.split(/[,，]/).map((x) => x.trim()).filter(Boolean);
    return parts.length > 0 ? parts : null;
  }
  return null;
}

async function mergeToDocxContract() {
  const templateId = String(state.template?.id ?? templateSelectEl?.value ?? "");
  const purchaseTypeId = String(state.purchase?.typeId ?? "");
  if (!templateId) throw new Error("缺少模板 ID");
  if (!purchaseTypeId) throw new Error("缺少订购单类型标识（type_ID）");
  if (!state.purchase?.workbook) throw new Error("订购单未就绪");

  const purchaseSheetName = purchaseSheetEl.value;
  const purchaseWs = getWorksheet(state.purchase.workbook, purchaseSheetName);
  if (!purchaseWs) throw new Error("订购单工作表不存在");

  setStatus("正在读取映射关系…");
  const mappingOut = await fetchJsonWithFallback(
    `/api/mapping?template_id=${encodeURIComponent(templateId)}&purchase_type_id=${encodeURIComponent(purchaseTypeId)}`,
  );
  if (!mappingOut.res.ok || !mappingOut.json?.success) {
    const msg = mappingOut.json?.error || `${mappingOut.res.status} ${mappingOut.res.statusText}`;
    throw new Error(`映射关系读取失败：${msg}`);
  }
  const mapping = mappingOut.json.data;

  const contractHeaders =
    coerceToArrayMaybeJson(mapping?.config?.contract_headers) ??
    coerceToArrayMaybeJson(mapping?.contract_headers);
  const purchaseHeaders =
    coerceToArrayMaybeJson(mapping?.config?.purchase_headers) ??
    coerceToArrayMaybeJson(mapping?.purchase_headers);
  const rules = Array.isArray(mapping?.rules) ? mapping.rules : null;

  if (!contractHeaders || contractHeaders.length === 0) throw new Error("映射配置缺少 contract_headers");
  if (!purchaseHeaders || purchaseHeaders.length === 0) throw new Error("映射配置缺少 purchase_headers");
  if (!rules || rules.length === 0) throw new Error("映射规则为空");

  const purchaseHeaderRow = findHeaderRowInWorksheet(purchaseWs, purchaseHeaders);
  if (!purchaseHeaderRow) throw new Error("未在订购单中找到映射配置指定的标题行");

  const contractHeaderIndex = new Map();
  contractHeaders.forEach((h, idx) => {
    const k = normalizeText(h);
    if (k) contractHeaderIndex.set(k, idx);
  });

  const mappingPairs = [];
  for (const r of rules) {
    const contractHeader = normalizeText(r?.contract_header);
    const purchaseHeader = normalizeText(r?.purchase_header);
    if (!contractHeader || !purchaseHeader) continue;
    const purchaseCol = purchaseHeaderRow.headerToCol.get(purchaseHeader);
    const contractIdx = contractHeaderIndex.get(contractHeader);
    if (!purchaseCol && purchaseCol !== 0) continue;
    if (contractIdx === undefined) continue;
    mappingPairs.push({ contractIdx, purchaseCol });
  }
  if (mappingPairs.length === 0) throw new Error("映射规则未能匹配到任何字段");

  const maxPurchaseRow = Math.max(purchaseWs.rowCount || 0, purchaseWs.actualRowCount || 0);
  const rows = [];
  for (let r = purchaseHeaderRow.rowIndex + 1; r <= maxPurchaseRow; r += 1) {
    const row = purchaseWs.getRow(r);
    const arr = new Array(contractHeaders.length).fill("");
    let any = false;
    for (const p of mappingPairs) {
      const cell = row.getCell(p.purchaseCol);
      const text = normalizeText(getCellText(cell));
      if (text) any = true;
      arr[p.contractIdx] = text;
    }
    if (any) rows.push(arr);
  }
  if (rows.length === 0) throw new Error("订购单没有可插入的数据行");

  setStatus("正在生成 Word 合同…");
  const outputBaseName = computeDefaultOutputBaseName();
  let templateDocxBase64 = "";
  try {
    if (state.template?.buffer && previewEl?.dataset?.docxEditable) {
      const editedTpl = await applyDocxEditsToArrayBuffer(state.template.buffer, previewEl);
      const bytes = new Uint8Array(editedTpl);
      let binary = "";
      const chunkSize = 0x2000;
      for (let i = 0; i < bytes.length; i += chunkSize) {
        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunkSize));
      }
      templateDocxBase64 = btoa(binary);
    }
  } catch {}
  const res = await fetchWithLocalFallbackRequest(`/api/templates/${encodeURIComponent(templateId)}/merge-docx`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      output_base_name: outputBaseName,
      contract_headers: contractHeaders,
      rows,
      template_docx_base64: templateDocxBase64 || undefined,
    }),
  });

  if (!res.ok) {
    const text = await res.text().catch(() => "");
    const parsed = (() => {
      try {
        return JSON.parse(text);
      } catch {
        return null;
      }
    })();
    const msg = parsed?.error || parsed?.message || text || `${res.status} ${res.statusText}`;
    throw new Error(`生成失败：${msg}`);
  }

  const buffer = await res.arrayBuffer();
  state.merged = null;
  state.mergedSheetName = null;
  state.mergedDocxBuffer = buffer;
  state.mergedDocxFilename = ensureFilenameWithExt(outputBaseName, ".docx");

  downloadBtnEl.disabled = false;
  if (saveBtnEl) saveBtnEl.disabled = false;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = false;
    saveNameInputEl.value = state.mergedDocxFilename;
  }
  previewEl.innerHTML = '<div class="meta">正在渲染生成后的 DOCX，请稍候…</div>';
  try {
    await renderDocxWithoutImages(buffer, previewEl, { editable: false });
  } catch (err) {
    previewEl.innerHTML = `<div class="meta">DOCX 预览失败：${String(
      err?.message ?? err,
    )}。请点击“下载修改后文件”查看结果。</div>`;
  }
  setStatus(`已生成 Word 合同：共插入 ${rows.length} 行`);
}

async function mergeToNewWorkbook() {
  const templateId = String(state.template?.id ?? templateSelectEl?.value ?? "");
  const purchaseTypeId = String(state.purchase?.typeId ?? "");
  if (!templateId) throw new Error("缺少模板 ID");
  if (!purchaseTypeId) throw new Error("缺少订购单类型标识（type_ID）");
  if (!state.purchase?.workbook) return;

  const isWord = isWordLikeTemplate(state.template?.fileType);
  if (isWord) {
    await mergeToDocxContract();
    return;
  }

  if (!state.template?.workbook) return;
  commitActivePreviewCellEdits();
  const templateSheetName = templateSheetEl.value;
  const purchaseSheetName = purchaseSheetEl.value;

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
    saveNameInputEl.value = ensureFilenameWithExt(computeDefaultOutputBaseName(), ".xlsx");
  }
  renderPreview(templateClone, templateSheetName);
  setStatus(
    `已插入 ${result.insertedRows} 行（插入起始行：第 ${result.headerRowIndex + 2} 行；标题匹配 ${result.matchedHeaders}/${result.purchaseHeaderCount}）`,
  );
}

async function handleDownload() {
  commitActivePreviewCellEdits();
  const filename = getEffectiveOutputFilename();
  if (state.mergedDocxBuffer) {
    const out = await applyDocxEditsToArrayBuffer(state.mergedDocxBuffer, previewEl).catch(() => state.mergedDocxBuffer);
    state.mergedDocxBuffer = out;
    downloadArrayBuffer(out, filename);
    return;
  }
  if (!state.merged) return;
  await downloadWorkbook(state.merged, filename);
}

async function handleSaveToServer() {
  commitActivePreviewCellEdits();
  const filename = getEffectiveOutputFilename();
  const baseName = filename.replace(/\.[A-Za-z0-9]+$/i, "");

  try {
    setStatus("正在保存到服务器…");
    let res = null;
    if (state.mergedDocxBuffer) {
      const out = await applyDocxEditsToArrayBuffer(state.mergedDocxBuffer, previewEl).catch(() => state.mergedDocxBuffer);
      state.mergedDocxBuffer = out;
      const pathname = `/api/contracts/save?name=${encodeURIComponent(baseName)}&ext=docx`;
      res = await postBinaryWithFallback(pathname, out);
    } else {
      if (!state.merged) return;
      const out = await state.merged.xlsx.writeBuffer();
      const pathname = `/api/contracts/save?name=${encodeURIComponent(baseName)}&ext=xlsx`;
      res = await postBinaryWithFallback(pathname, out);
    }
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
if (purchaseFileEl) purchaseFileEl.addEventListener("change", handlePurchaseFileChange);
mergeBtnEl.addEventListener("click", () => {
  mergeToNewWorkbook().catch((e) => {
    state.merged = null;
    state.mergedSheetName = null;
    state.mergedDocxBuffer = null;
    state.mergedDocxFilename = null;
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
    const filePath = btn.dataset.filePath || "";
    if (!id) return;
    const ext = String(filePath).toLowerCase().match(/\.([a-z0-9]+)$/)?.[1] ?? "";
    const extWithDot = ext ? `.${ext}` : ".xlsx";

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
          const filename = ensureFilenameWithExt(sanitizeFilenamePart(name) || `contract_${id}`, extWithDot);
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
          if (extWithDot.toLowerCase() === ".docx") {
            await openDocxPreviewModal(`预览 - ${name || id}`, buffer);
            setStatus("");
            return;
          }
          if (extWithDot.toLowerCase() !== ".xlsx") {
            const filename = ensureFilenameWithExt(sanitizeFilenamePart(name) || `contract_${id}`, extWithDot);
            downloadArrayBuffer(buffer, filename);
            setStatus("");
            return;
          }
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
  } catch (e) {
    try {
      if (state.hf) state.hf.destroy();
    } catch {}
    state.hf = null;
    state.hfSheetId = null;
  }
  renderPreview(state.merged, state.mergedSheetName);
});

purchaseSheetEl.addEventListener("change", () => {
  state.mergedDocxBuffer = null;
  state.mergedDocxFilename = null;
  downloadBtnEl.disabled = true;
  if (saveBtnEl) saveBtnEl.disabled = true;
  if (saveNameInputEl) {
    saveNameInputEl.disabled = true;
    saveNameInputEl.value = "";
  }
  setStatus("");
});

// 初始化：获取模板列表
fetchTemplates();
fetchPurchaseOrders();
fetchContractsList();
 
setStatus("");
