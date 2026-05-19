function $(id) {
  return document.getElementById(id);
}

function escapeHtml(s) {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function setStatus(el, text, kind) {
  el.textContent = text || '';
  el.style.color = kind === 'error' ? '#b91c1c' : '#334155';
}

async function apiJson(url, options) {
  const res = await fetch(url, options);
  const data = await res.json().catch(() => ({}));
  if (!res.ok || data?.success === false) {
    const message = data?.error || `请求失败：${res.status}`;
    throw new Error(message);
  }
  return data;
}

function formatDate(value) {
  if (!value) return '';
  const d = new Date(value);
  if (Number.isNaN(d.getTime())) return String(value);
  return d.toLocaleString();
}

function parseContentDispositionFilename(headerValue) {
  const s = String(headerValue ?? '');
  const m5987 = s.match(/filename\*\s*=\s*UTF-8''([^;]+)/i);
  if (m5987 && m5987[1]) {
    try {
      return decodeURIComponent(m5987[1]);
    } catch {
      return m5987[1];
    }
  }
  const m = s.match(/filename\s*=\s*\"([^\"]+)\"/i) || s.match(/filename\s*=\s*([^;]+)/i);
  if (m && m[1]) return m[1].trim();
  return '';
}

function extFromName(name) {
  const s = String(name ?? '');
  const m = s.match(/\.([A-Za-z0-9]+)$/);
  return m ? String(m[1]).toLowerCase() : '';
}

function stripFileExtension(filename) {
  const s = String(filename ?? '');
  const i = s.lastIndexOf('.');
  if (i > 0) return s.slice(0, i);
  return s;
}

function normalizeType(typeOrExt) {
  const t = String(typeOrExt ?? '').trim().toLowerCase();
  if (!t) return '';
  if (t.includes('xlsx')) return 'xlsx';
  if (t.includes('xls')) return 'xls';
  if (t.includes('pdf')) return 'pdf';
  if (t.includes('docx')) return 'docx';
  if (t.includes('doc')) return 'doc';
  return t.replace(/^\./, '');
}

let currentPreviewUrl = null;

function closePreview() {
  const overlay = $('previewOverlay');
  overlay.classList.add('hidden');
  $('previewTitle').textContent = '预览';
  $('previewMeta').textContent = '';
  $('previewBody').innerHTML = '';
  if (currentPreviewUrl) {
    URL.revokeObjectURL(currentPreviewUrl);
    currentPreviewUrl = null;
  }
}

function openPreviewShell(title, metaText) {
  $('previewTitle').textContent = title || '预览';
  $('previewMeta').textContent = metaText || '';
  $('previewBody').innerHTML = '<div class="meta">正在加载预览...</div>';
  $('previewOverlay').classList.remove('hidden');
}

function openMapping() {
  $('mappingOverlay').classList.remove('hidden');
}

function closeMapping() {
  $('mappingOverlay').classList.add('hidden');
}

async function fetchFileForPreview(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`预览加载失败：${res.status}`);
  const disp = res.headers.get('content-disposition') || '';
  const contentType = res.headers.get('content-type') || '';
  const filename = parseContentDispositionFilename(disp);
  const blob = await res.blob();
  return { blob, filename, contentType };
}

async function renderExcelPreview(arrayBuffer) {
  if (!window.ExcelJS) {
    return '<div class="meta">Excel 预览组件未加载，请刷新页面重试。</div>';
  }

  const columnToLetters = (n) => {
    let x = Number(n) || 0;
    if (x < 1) return '';
    let s = '';
    while (x > 0) {
      x -= 1;
      s = String.fromCharCode(65 + (x % 26)) + s;
      x = Math.floor(x / 26);
    }
    return s;
  };

  const workbook = new window.ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(new Uint8Array(arrayBuffer));
  } catch {
    await workbook.xlsx.load(arrayBuffer);
  }
  const sheet = workbook.worksheets[0];
  if (!sheet) return '<div class="meta">未找到工作表</div>';

  const maxRows = Math.min(60, sheet.rowCount || 60);
  const maxCols = Math.min(20, sheet.columnCount || 20);

  const cellToString = (cell) => {
    if (!cell) return '';
    if (cell.isMerged) {
      const master = cell.master;
      if (master && master.address && cell.address && master.address !== cell.address) return '';
    }
    const v = cell.value;
    if (v == null) return '';
    if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v);
    if (v instanceof Date) return v.toLocaleString();
    if (typeof v === 'object') {
      if (Array.isArray(v.richText)) return v.richText.map((p) => (p && p.text != null ? String(p.text) : '')).join('');
      if (v.text != null) return String(v.text);
      if (v.hyperlink) return String(v.text ?? v.hyperlink);
      if (v.formula) return v.result != null ? String(v.result) : String(v.formula);
      if (v.sharedFormula) return v.result != null ? String(v.result) : '';
      if (v.error) return String(v.error);
      if (v.result != null) return String(v.result);
    }
    try {
      return String(v);
    } catch {
      return '';
    }
  };

  let html = '<div class="previewBox"><table class="previewTable"><thead><tr><th></th>';
  for (let c = 1; c <= maxCols; c++) {
    html += `<th>${escapeHtml(columnToLetters(c))}</th>`;
  }
  html += '</tr></thead><tbody>';
  for (let r = 1; r <= maxRows; r++) {
    html += `<tr><th>${escapeHtml(String(r))}</th>`;
    for (let c = 1; c <= maxCols; c++) {
      const cell = sheet.getCell(r, c);
      const text = cellToString(cell);
      html += `<td>${escapeHtml(text)}</td>`;
    }
    html += '</tr>';
  }
  html += '</tbody></table></div>';
  return html;
}

function cellToPlainString(cell) {
  if (!cell) return '';
  if (cell.isMerged) {
    const master = cell.master;
    if (master && master.address && cell.address && master.address !== cell.address) return '';
  }
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toLocaleString();
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((p) => (p && p.text != null ? String(p.text) : '')).join('');
    if (v.text != null) return String(v.text);
    if (v.hyperlink) return String(v.text ?? v.hyperlink);
    if (v.formula) return v.result != null ? String(v.result) : String(v.formula);
    if (v.sharedFormula) return v.result != null ? String(v.result) : '';
    if (v.error) return String(v.error);
    if (v.result != null) return String(v.result);
  }
  try {
    return String(v);
  } catch {
    return '';
  }
}

function uniqHeaders(values) {
  const out = [];
  const seen = new Set();
  for (const v of values) {
    const s = String(v ?? '').trim();
    if (!s) continue;
    if (seen.has(s)) continue;
    seen.add(s);
    out.push(s);
  }
  return out;
}

async function loadXlsxSheetFromDownload(url) {
  const { blob, filename, contentType } = await fetchFileForPreview(url);
  const ext = normalizeType(extFromName(filename) || contentType);
  if (ext !== 'xlsx') {
    throw new Error('仅支持 xlsx 文件解析标题');
  }
  if (!window.ExcelJS) throw new Error('Excel 组件未加载');
  const workbook = new window.ExcelJS.Workbook();
  const arrayBuffer = await blob.arrayBuffer();
  try {
    await workbook.xlsx.load(new Uint8Array(arrayBuffer));
  } catch {
    await workbook.xlsx.load(arrayBuffer);
  }
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new Error('未找到工作表');
  return sheet;
}

function extractHeadersFromRow(sheet, rowIndex) {
  const maxCols = Math.min(80, sheet.columnCount || 80);
  const values = [];
  for (let c = 1; c <= maxCols; c++) {
    const cell = sheet.getCell(rowIndex, c);
    values.push(cellToPlainString(cell));
  }
  return uniqHeaders(values);
}

function detectTemplateHeaderRow(sheet) {
  const maxRows = Math.min(60, sheet.rowCount || 60);
  const maxCols = Math.min(80, sheet.columnCount || 80);
  let bestRow = 1;
  let bestScore = -Infinity;

  for (let r = 1; r <= maxRows; r++) {
    const values = [];
    for (let c = 1; c <= maxCols; c++) {
      const cell = sheet.getCell(r, c);
      values.push(cellToPlainString(cell));
    }
    const trimmed = values.map((s) => String(s ?? '').trim()).filter(Boolean);
    if (trimmed.length < 2) continue;
    const unique = new Set(trimmed);
    const avgLen = trimmed.reduce((a, s) => a + s.length, 0) / trimmed.length;
    const maxLen = trimmed.reduce((m, s) => Math.max(m, s.length), 0);
    const hasVeryLong = maxLen > 80;

    const score =
      trimmed.length * 12 +
      unique.size * 3 -
      avgLen * 0.6 -
      (hasVeryLong ? 30 : 0) -
      r * 0.25;

    if (score > bestScore) {
      bestScore = score;
      bestRow = r;
    }
  }

  return bestRow;
}

function renderChips(container, headers) {
  container.innerHTML = headers
    .map((h) => `<span class="chip" title="${escapeHtml(h)}">${escapeHtml(h)}</span>`)
    .join('');
}

let templatesCache = [];
let purchaseOrdersCache = [];

let mappingTemplateHeaders = [];
let mappingPurchaseHeaders = [];
let mappingTemplateHeaderRow = null;
let mappingPurchaseHeaderRow = null;
let mappingRules = [];
let mappingTemplateSheet = null;
let mappingPurchaseSheet = null;
let mappingNameIsAuto = true;
let lastAutoMappingName = '';

function getRecordNameById(cache, id) {
  const sid = String(id ?? '');
  const hit = (cache || []).find((r) => String(r?.id ?? '') === sid);
  return String(hit?.name || hit?.file_name || '').trim();
}

function computeAutoMappingName() {
  const templateId = $('mappingTemplateSelect')?.value;
  const purchaseId = $('mappingPurchaseSelect')?.value;
  const templateName = getRecordNameById(templatesCache, templateId);
  const purchaseName = getRecordNameById(purchaseOrdersCache, purchaseId);
  if (!templateName || !purchaseName) return '';
  return `${templateName}-${purchaseName}`;
}

function getPurchaseTypeIdById(purchaseId) {
  const sid = String(purchaseId ?? '');
  const hit = (purchaseOrdersCache || []).find((r) => String(r?.id ?? '') === sid);
  const fromApi = String(hit?.type_id ?? hit?.type_ID ?? '').trim();
  if (fromApi) return fromApi;
  const name = String(hit?.name ?? '').trim();
  const m = name.match(/(\d{3})$/);
  return m ? m[1] : '';
}

function updateDefaultMappingName(force) {
  const input = $('mappingName');
  if (!input) return;
  const next = computeAutoMappingName();
  if (!next) return;

  const current = String(input.value || '').trim();
  const shouldUpdate = force || mappingNameIsAuto || current === lastAutoMappingName || !current;
  if (!shouldUpdate) return;

  input.value = next;
  mappingNameIsAuto = true;
  lastAutoMappingName = next;
}

function readPositiveInt(value) {
  const n = Number.parseInt(String(value ?? '').trim(), 10);
  if (!Number.isFinite(n) || n < 1) return null;
  return n;
}

function setHeaderRowUi(templateRow, purchaseRow) {
  const templateInput = $('mappingTemplateHeaderRowInput');
  const purchaseInput = $('mappingPurchaseHeaderRowInput');
  if (templateInput) templateInput.value = templateRow ? String(templateRow) : '';
  if (purchaseInput) purchaseInput.value = purchaseRow ? String(purchaseRow) : '';
}

function fillRowSelectOptions(selectEl, maxRows, selectedRow) {
  if (!selectEl) return;
  const n = Math.max(1, Number(maxRows || 1));
  const options = [];
  for (let i = 1; i <= n; i++) {
    options.push(`<option value="${i}">${i}</option>`);
  }
  selectEl.innerHTML = options.join('');
  if (selectedRow) selectEl.value = String(selectedRow);
}

function mergeAutoMatchedRules(existingRules, templateHeaders, purchaseHeaders) {
  const templateSet = new Set((templateHeaders || []).map((h) => String(h).trim()).filter(Boolean));
  const purchaseSet = new Set((purchaseHeaders || []).map((h) => String(h).trim()).filter(Boolean));

  const out = [];
  const seen = new Set();

  for (const h of templateHeaders || []) {
    const s = String(h).trim();
    if (!s) continue;
    if (!purchaseSet.has(s)) continue;
    const key = `${s}|||${s}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ templateHeader: s, purchaseHeader: s });
  }

  for (const r of existingRules || []) {
    const t = String(r?.templateHeader ?? '').trim();
    const p = String(r?.purchaseHeader ?? '').trim();
    if (!t || !p) continue;
    if (!templateSet.has(t) || !purchaseSet.has(p)) continue;
    const key = `${t}|||${p}`;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ templateHeader: t, purchaseHeader: p });
  }

  out.push({ templateHeader: '', purchaseHeader: '' });
  return out;
}

function setMappingHeaders({ templateHeaders, purchaseHeaders, templateRow, purchaseRow, keepExistingRules }) {
  const t = Array.isArray(templateHeaders) ? templateHeaders : [];
  const p = Array.isArray(purchaseHeaders) ? purchaseHeaders : [];
  if (!t.length) throw new Error('合同模板标题行为空');
  if (!p.length) throw new Error('订购单标题行为空');

  mappingTemplateHeaderRow = templateRow ?? 1;
  mappingPurchaseHeaderRow = purchaseRow ?? 1;
  mappingTemplateHeaders = t;
  mappingPurchaseHeaders = p;

  setHeaderRowUi(mappingTemplateHeaderRow, mappingPurchaseHeaderRow);
  renderChips($('templateHeadersBox'), mappingTemplateHeaders);
  renderChips($('purchaseHeadersBox'), mappingPurchaseHeaders);

  mappingRules = mergeAutoMatchedRules(keepExistingRules ? mappingRules : [], mappingTemplateHeaders, mappingPurchaseHeaders);
  renderMappingRulesTable();
  return mappingRules.length ? mappingRules.length - 1 : 0;
}

function applyHeaderRowSelection({ keepExistingRules }) {
  const templateInput = $('mappingTemplateHeaderRowInput');
  const purchaseInput = $('mappingPurchaseHeaderRowInput');

  const templateRow = readPositiveInt(templateInput?.value) ?? mappingTemplateHeaderRow ?? 1;
  const purchaseRow = readPositiveInt(purchaseInput?.value) ?? mappingPurchaseHeaderRow ?? 1;

  let newTemplateHeaders = mappingTemplateHeaders;
  let newPurchaseHeaders = mappingPurchaseHeaders;

  if (mappingTemplateSheet) {
    const templateMax = mappingTemplateSheet.rowCount || templateRow;
    if (templateRow > templateMax) throw new Error(`合同模板标题行超出范围（最大 ${templateMax}）`);
    newTemplateHeaders = extractHeadersFromRow(mappingTemplateSheet, templateRow);
  }
  if (mappingPurchaseSheet) {
    const purchaseMax = mappingPurchaseSheet.rowCount || purchaseRow;
    if (purchaseRow > purchaseMax) throw new Error(`订购单标题行超出范围（最大 ${purchaseMax}）`);
    newPurchaseHeaders = extractHeadersFromRow(mappingPurchaseSheet, purchaseRow);
  }

  if (!newTemplateHeaders || !newTemplateHeaders.length) throw new Error('请先解析标题');
  if (!newPurchaseHeaders || !newPurchaseHeaders.length) throw new Error('请先解析标题');

  return setMappingHeaders({
    templateHeaders: newTemplateHeaders,
    purchaseHeaders: newPurchaseHeaders,
    templateRow,
    purchaseRow,
    keepExistingRules
  });
}

function resetMappingParsedState() {
  mappingTemplateHeaders = [];
  mappingPurchaseHeaders = [];
  mappingTemplateHeaderRow = null;
  mappingPurchaseHeaderRow = null;
  mappingRules = [];
  mappingTemplateSheet = null;
  mappingPurchaseSheet = null;
  setHeaderRowUi(null, null);
  const templateSelect = $('mappingTemplateHeaderRowInput');
  const purchaseSelect = $('mappingPurchaseHeaderRowInput');
  if (templateSelect) {
    templateSelect.innerHTML = '';
    templateSelect.disabled = false;
  }
  if (purchaseSelect) {
    purchaseSelect.innerHTML = '';
    purchaseSelect.disabled = false;
  }
  const tbox = $('templateHeadersBox');
  const pbox = $('purchaseHeadersBox');
  if (tbox) tbox.innerHTML = '';
  if (pbox) pbox.innerHTML = '';
  renderMappingRulesTable();
}

async function previewRecord(kind, row) {
  const id = row.id;
  const name = row.name || row.file_name || `ID ${id}`;
  const hinted = normalizeType(row.file_type || extFromName(row.file_name) || extFromName(row.file_path) || extFromName(row.name));

  openPreviewShell(`${kind === 'templates' ? '合同模板' : '采购单'}预览`, `${escapeHtml(String(name))}`);
  try {
    const { blob, filename, contentType } = await fetchFileForPreview(`/api/${kind}/${encodeURIComponent(id)}/download`);
    const effectiveName = filename || row.file_name || row.name || '';
    const ext = normalizeType(extFromName(effectiveName) || hinted || contentType);
    $('previewMeta').textContent = `${effectiveName ? effectiveName : '文件'}${ext ? `（${ext}）` : ''}`;

    if (ext === 'xlsx') {
      const arrayBuffer = await blob.arrayBuffer();
      const html = await renderExcelPreview(arrayBuffer);
      $('previewBody').innerHTML = html;
      return;
    }

    if (ext === 'docx') {
      const { data } = await apiJson(`/api/${kind}/${encodeURIComponent(id)}/preview`);
      const blocks = data?.blocks || [];
      const parts = [];
      for (const b of blocks) {
        if (b?.type === 'p') {
          const t = String(b.text ?? '').trim();
          if (!t) continue;
          parts.push(`<div class="docxText">${escapeHtml(t)}</div>`);
          continue;
        }
        if (b?.type === 'table' && Array.isArray(b.rows) && b.rows.length) {
          const rows = b.rows;
          const head = rows[0] || [];
          const bodyRows = rows.slice(1);
          const thead =
            head.length > 0
              ? `<thead><tr>${head.map((c) => `<th>${escapeHtml(String(c ?? ''))}</th>`).join('')}</tr></thead>`
              : '';
          const tbody = `<tbody>${bodyRows
            .map((r) => `<tr>${(r || []).map((c) => `<td>${escapeHtml(String(c ?? ''))}</td>`).join('')}</tr>`)
            .join('')}</tbody>`;
          parts.push(`<div class="previewBox"><table class="previewTable">${thead}${tbody}</table></div>`);
        }
      }
      $('previewBody').innerHTML = parts.join('') || '<div class="meta">未解析到可预览内容</div>';
      return;
    }

    currentPreviewUrl = URL.createObjectURL(blob);
    if (ext === 'pdf') {
      $('previewBody').innerHTML = `<iframe class="previewFrame" src="${currentPreviewUrl}"></iframe>`;
      return;
    }

    $('previewBody').innerHTML = `<div class="meta">当前格式暂不支持页面内预览，可点击“下载”查看。</div>
      <div class="actions">
        <a class="link" href="/api/${kind}/${encodeURIComponent(id)}/download">下载</a>
      </div>`;
  } catch (e) {
    $('previewBody').innerHTML = `<div class="meta" style="color:#b91c1c">${escapeHtml(e.message || '预览失败')}</div>`;
  }
}

async function refreshTemplates() {
  const tbody = $('templatesTbody');
  tbody.innerHTML = '<tr><td colspan="4">正在加载...</td></tr>';
  const { data } = await apiJson('/api/templates');
  templatesCache = data ?? [];
  if (!data || data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4">暂无数据</td></tr>';
    return;
  }
  tbody.innerHTML = data
    .map((r, idx) => {
      const id = escapeHtml(r.id);
      const name = escapeHtml(r.name || r.file_name || '');
      const created = escapeHtml(formatDate(r.created_at));
      const displayId = escapeHtml(idx + 1);
      return `<tr>
        <td>${displayId}</td>
        <td title="${name}">${name}</td>
        <td>${created}</td>
        <td class="opCell">
          <div class="opActions">
            <button class="btn small" data-action="preview-template" data-id="${id}">预览</button>
            <a class="link" href="/api/templates/${id}/download">下载</a>
            <button class="btn small danger" data-action="delete-template" data-id="${id}">删除</button>
          </div>
        </td>
      </tr>`;
    })
    .join('');
}

async function refreshPurchaseOrders() {
  const tbody = $('purchaseTbody');
  tbody.innerHTML = '<tr><td colspan="4">正在加载...</td></tr>';
  const { data } = await apiJson('/api/purchase-orders');
  purchaseOrdersCache = data ?? [];
  if (!data || data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="4">暂无数据</td></tr>';
    return;
  }
  tbody.innerHTML = data
    .map((r, idx) => {
      const id = escapeHtml(r.id);
      const name = escapeHtml(r.name || '');
      const created = escapeHtml(formatDate(r.created_at));
      const displayId = escapeHtml(idx + 1);
      return `<tr>
        <td>${displayId}</td>
        <td title="${name}">${name}</td>
        <td>${created}</td>
        <td class="opCell">
          <div class="opActions">
            <button class="btn small" data-action="preview-purchase" data-id="${id}">预览</button>
            <a class="link" href="/api/purchase-orders/${id}/download">下载</a>
            <button class="btn small danger" data-action="delete-purchase" data-id="${id}">删除</button>
          </div>
        </td>
      </tr>`;
    })
    .join('');
}

function fillSelectOptions(selectEl, items, placeholder) {
  const options = [`<option value="">${escapeHtml(placeholder || '请选择')}</option>`];
  for (const it of items) {
    const id = String(it.id ?? '');
    const label = `${it.name || it.file_name || '未命名'}（ID ${id}）`;
    options.push(`<option value="${escapeHtml(id)}">${escapeHtml(label)}</option>`);
  }
  selectEl.innerHTML = options.join('');
}

function renderMappingRulesTable() {
  const tbody = $('mappingRulesTbody');
  if (!mappingTemplateHeaders.length || !mappingPurchaseHeaders.length) {
    tbody.innerHTML = '<tr><td colspan="4">请先解析标题</td></tr>';
    return;
  }
  if (!mappingRules.length) {
    tbody.innerHTML = '<tr><td colspan="4">暂无规则，点击“新增规则”开始配置</td></tr>';
    return;
  }

  const templateOptions = mappingTemplateHeaders
    .map((h) => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`)
    .join('');
  const purchaseOptions = mappingPurchaseHeaders
    .map((h) => `<option value="${escapeHtml(h)}">${escapeHtml(h)}</option>`)
    .join('');

  tbody.innerHTML = mappingRules
    .map((r, idx) => {
      const seq = escapeHtml(idx + 1);
      const t = String(r.templateHeader ?? '');
      const p = String(r.purchaseHeader ?? '');
      return `<tr>
        <td>${seq}</td>
        <td>
          <select class="selectInline" data-rule-idx="${idx}" data-field="template">
            <option value="">请选择</option>
            ${templateOptions}
          </select>
        </td>
        <td>
          <select class="selectInline" data-rule-idx="${idx}" data-field="purchase">
            <option value="">请选择</option>
            ${purchaseOptions}
          </select>
        </td>
        <td class="opCell">
          <div class="opActions">
            <button class="btn small danger" data-action="remove-rule" data-rule-idx="${idx}">删除</button>
          </div>
        </td>
      </tr>`;
    })
    .join('');

  const selects = tbody.querySelectorAll('select[data-rule-idx]');
  selects.forEach((sel) => {
    const idx = Number(sel.getAttribute('data-rule-idx'));
    const field = sel.getAttribute('data-field');
    const value = field === 'purchase' ? mappingRules[idx]?.purchaseHeader : mappingRules[idx]?.templateHeader;
    if (value != null) sel.value = String(value);
  });
}

async function refreshMappings() {
  const tbody = $('mappingsTbody');
  tbody.innerHTML = '<tr><td colspan="5">正在加载...</td></tr>';
  const { data } = await apiJson('/api/mappings');
  if (!data || data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5">暂无映射</td></tr>';
    return;
  }
  tbody.innerHTML = data
    .map((r, idx) => {
      const seq = escapeHtml(idx + 1);
      const id = escapeHtml(r.id);
      const name = escapeHtml(r.name || '');
      const count = escapeHtml(r.rule_count ?? 0);
      const created = escapeHtml(formatDate(r.created_at));
      return `<tr>
        <td>${seq}</td>
        <td title="${name}">${name}</td>
        <td>${count}</td>
        <td>${created}</td>
        <td class="opCell">
          <div class="opActions">
            <button class="btn small" data-action="view-mapping" data-id="${id}">查看</button>
            <button class="btn small danger" data-action="delete-mapping" data-id="${id}">删除</button>
          </div>
        </td>
      </tr>`;
    })
    .join('');
}

async function parseMappingHeaders() {
  const status = $('mappingStatus');
  setStatus(status, '解析中...');
  try {
    const templateId = $('mappingTemplateSelect').value;
    const purchaseId = $('mappingPurchaseSelect').value;
    if (!templateId) throw new Error('请选择合同模板');
    if (!purchaseId) throw new Error('请选择订购单');

    const purchaseTypeId = getPurchaseTypeIdById(purchaseId);
    if (!purchaseTypeId) throw new Error('无法从订购单名称解析 type_ID（需以 001/002 等三位数字结尾）');

    const check = await apiJson(
      `/api/mappings/check?templateId=${encodeURIComponent(templateId)}&purchaseTypeId=${encodeURIComponent(
        purchaseTypeId
      )}`
    );
    if (check?.data?.exists) {
      await refreshMappings();
      setStatus(status, '已有映射关系，请删除后再建立映射关系', 'error');
      return;
    }

    const templateRecord = (templatesCache || []).find((r) => String(r?.id ?? '') === String(templateId));
    const purchaseRecord = (purchaseOrdersCache || []).find((r) => String(r?.id ?? '') === String(purchaseId));
    const detectTypeFromRecord = (r) =>
      normalizeType(r?.file_type || extFromName(r?.file_name) || extFromName(r?.file_path) || extFromName(r?.name));

    const templateType = detectTypeFromRecord(templateRecord);
    const purchaseType = detectTypeFromRecord(purchaseRecord);

    const templateRowSelect = $('mappingTemplateHeaderRowInput');
    const purchaseRowSelect = $('mappingPurchaseHeaderRowInput');

    let templateHeaders = [];
    let purchaseHeaders = [];
    let detectedTemplateRow = 1;
    let detectedPurchaseRow = 1;

    mappingTemplateSheet = null;
    mappingPurchaseSheet = null;

    if (templateType === 'xlsx') {
      const templateSheet = await loadXlsxSheetFromDownload(`/api/templates/${encodeURIComponent(templateId)}/download`);
      mappingTemplateSheet = templateSheet;
      detectedTemplateRow = detectTemplateHeaderRow(templateSheet);
      templateHeaders = extractHeadersFromRow(templateSheet, detectedTemplateRow);
      fillRowSelectOptions(templateRowSelect, templateSheet.rowCount || 1, detectedTemplateRow);
      if (templateRowSelect) templateRowSelect.disabled = false;
    } else if (templateType === 'doc' || templateType === 'docx') {
      const { data } = await apiJson(`/api/templates/${encodeURIComponent(templateId)}/headers`);
      templateHeaders = data?.headers ?? [];
      detectedTemplateRow = 1;
      fillRowSelectOptions(templateRowSelect, 1, 1);
      if (templateRowSelect) templateRowSelect.disabled = true;
    } else {
      throw new Error('合同模板仅支持 xlsx/doc/docx 建立映射');
    }

    if (purchaseType === 'xlsx') {
      const purchaseSheet = await loadXlsxSheetFromDownload(`/api/purchase-orders/${encodeURIComponent(purchaseId)}/download`);
      mappingPurchaseSheet = purchaseSheet;
      detectedPurchaseRow = 1;
      purchaseHeaders = extractHeadersFromRow(purchaseSheet, 1);
      fillRowSelectOptions(purchaseRowSelect, purchaseSheet.rowCount || 1, 1);
      if (purchaseRowSelect) purchaseRowSelect.disabled = false;
    } else if (purchaseType === 'doc' || purchaseType === 'docx') {
      const { data } = await apiJson(`/api/purchase-orders/${encodeURIComponent(purchaseId)}/headers`);
      purchaseHeaders = data?.headers ?? [];
      detectedPurchaseRow = 1;
      fillRowSelectOptions(purchaseRowSelect, 1, 1);
      if (purchaseRowSelect) purchaseRowSelect.disabled = true;
    } else {
      throw new Error('订购单仅支持 xlsx/doc/docx 建立映射');
    }

    updateDefaultMappingName(false);
    mappingRules = [];
    const autoCount = setMappingHeaders({
      templateHeaders,
      purchaseHeaders,
      templateRow: detectedTemplateRow,
      purchaseRow: detectedPurchaseRow,
      keepExistingRules: false
    });
    setStatus(status, `解析完成，已自动匹配 ${autoCount} 条规则`);
  } catch (e) {
    resetMappingParsedState();
    setStatus(status, e.message || '解析失败', 'error');
  }
}

async function saveMapping() {
  const status = $('mappingStatus');
  setStatus(status, '保存中...');
  try {
    const name = String($('mappingName').value || '').trim();
    const templateId = $('mappingTemplateSelect').value;
    const purchaseId = $('mappingPurchaseSelect').value;
    if (!name) throw new Error('请输入映射名称');
    if (!templateId) throw new Error('请选择合同模板');
    if (!purchaseId) throw new Error('请选择订购单');
    if (!mappingTemplateHeaders.length || !mappingPurchaseHeaders.length) throw new Error('请先解析标题');

    const purchaseTypeId = getPurchaseTypeIdById(purchaseId);
    if (!purchaseTypeId) throw new Error('无法从订购单名称解析 type_ID（需以 001/002 等三位数字结尾）');

    const rules = mappingRules
      .map((r) => ({
        templateHeader: String(r.templateHeader || '').trim(),
        purchaseHeader: String(r.purchaseHeader || '').trim()
      }))
      .filter((r) => r.templateHeader && r.purchaseHeader);
    if (!rules.length) throw new Error('请至少配置一条映射规则');

    await apiJson('/api/mappings', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        name,
        templateId: Number(templateId),
        purchaseTypeId,
        templateHeaders: mappingTemplateHeaders,
        purchaseHeaders: mappingPurchaseHeaders,
        rules
      })
    });

    setStatus(status, '保存成功');
    await refreshMappings();
  } catch (e) {
    setStatus(status, e.message || '保存失败', 'error');
  }
}

async function handleUploadTemplate() {
  const btn = $('uploadTemplateBtn');
  const status = $('templateStatus');
  const name = $('templateName').value;
  const fileInput = $('templateFile');
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    setStatus(status, '请选择文件', 'error');
    return;
  }

  btn.disabled = true;
  setStatus(status, '上传中...');
  try {
    const fd = new FormData();
    fd.append('name', name);
    fd.append('file', file);
    await apiJson('/api/templates/upload', { method: 'POST', body: fd });
    setStatus(status, '上传成功');
    fileInput.value = '';
    await refreshTemplates();
    fillSelectOptions($('mappingTemplateSelect'), templatesCache, '请选择合同模板');
  } catch (e) {
    setStatus(status, e.message || '上传失败', 'error');
  } finally {
    btn.disabled = false;
  }
}

async function handleUploadPurchase() {
  const btn = $('uploadPurchaseBtn');
  const status = $('purchaseStatus');
  const name = $('purchaseName').value;
  const fileInput = $('purchaseFile');
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    setStatus(status, '请选择文件', 'error');
    return;
  }

  btn.disabled = true;
  setStatus(status, '上传中...');
  try {
    const fd = new FormData();
    fd.append('name', name);
    fd.append('file', file);
    await apiJson('/api/purchase-orders/upload', { method: 'POST', body: fd });
    setStatus(status, '上传成功');
    fileInput.value = '';
    await refreshPurchaseOrders();
    fillSelectOptions($('mappingPurchaseSelect'), purchaseOrdersCache, '请选择订购单');
  } catch (e) {
    setStatus(status, e.message || '上传失败', 'error');
  } finally {
    btn.disabled = false;
  }
}

async function init() {
  $('uploadTemplateBtn').addEventListener('click', handleUploadTemplate);
  $('uploadPurchaseBtn').addEventListener('click', handleUploadPurchase);
  $('mappingName').addEventListener('input', () => {
    const v = String($('mappingName').value || '').trim();
    if (!v) {
      mappingNameIsAuto = true;
      lastAutoMappingName = '';
      updateDefaultMappingName(true);
      return;
    }
    mappingNameIsAuto = false;
  });
  $('templateFile').addEventListener('change', () => {
    const input = $('templateName');
    const fileInput = $('templateFile');
    const file = fileInput.files && fileInput.files[0];
    if (!file) return;
    input.value = stripFileExtension(file.name);
    input.dataset.autoFilled = '1';
  });
  $('purchaseFile').addEventListener('change', () => {
    const input = $('purchaseName');
    const fileInput = $('purchaseFile');
    const file = fileInput.files && fileInput.files[0];
    if (!file) return;
    input.value = stripFileExtension(file.name);
    input.dataset.autoFilled = '1';
  });

  $('templatesTbody').addEventListener('click', async (e) => {
    const el = e.target;
    if (!(el instanceof HTMLElement)) return;
    const action = el.getAttribute('data-action');
    const id = el.getAttribute('data-id');
    if (!action || !id) return;

    if (action === 'preview-template') {
      await previewRecord('templates', { id });
      return;
    }
    if (action === 'delete-template') {
      if (!confirm(`确认删除模板 ID=${id} 吗？此操作会删除数据库记录，并尝试删除已上传文件。`)) return;
      try {
        await apiJson(`/api/templates/${encodeURIComponent(id)}`, { method: 'DELETE' });
        await refreshTemplates();
        await refreshMappings();
      } catch (err) {
        setStatus($('templateStatus'), err.message || '删除失败', 'error');
      }
    }
  });

  $('purchaseTbody').addEventListener('click', async (e) => {
    const el = e.target;
    if (!(el instanceof HTMLElement)) return;
    const action = el.getAttribute('data-action');
    const id = el.getAttribute('data-id');
    if (!action || !id) return;

    if (action === 'preview-purchase') {
      await previewRecord('purchase-orders', { id });
      return;
    }
    if (action === 'delete-purchase') {
      if (!confirm(`确认删除采购单 ID=${id} 吗？此操作会删除数据库记录，并尝试删除已上传文件。`)) return;
      try {
        await apiJson(`/api/purchase-orders/${encodeURIComponent(id)}`, { method: 'DELETE' });
        await refreshPurchaseOrders();
        fillSelectOptions($('mappingPurchaseSelect'), purchaseOrdersCache, '请选择订购单');
        await refreshMappings();
      } catch (err) {
        setStatus($('purchaseStatus'), err.message || '删除失败', 'error');
      }
    }
  });

  $('previewClose').addEventListener('click', closePreview);
  $('previewOverlay').addEventListener('click', (e) => {
    if (e.target === $('previewOverlay')) closePreview();
  });

  $('openMappingBtn').addEventListener('click', openMapping);
  $('mappingClose').addEventListener('click', closeMapping);
  $('mappingOverlay').addEventListener('click', (e) => {
    if (e.target === $('mappingOverlay')) closeMapping();
  });

  $('mappingParseBtn').addEventListener('click', parseMappingHeaders);
  $('mappingAddRuleBtn').addEventListener('click', () => {
    if (!mappingTemplateHeaders.length || !mappingPurchaseHeaders.length) {
      setStatus($('mappingStatus'), '请先解析标题', 'error');
      return;
    }
    mappingRules.push({ templateHeader: '', purchaseHeader: '' });
    renderMappingRulesTable();
  });
  $('mappingSaveBtn').addEventListener('click', saveMapping);

  $('mappingTemplateSelect').addEventListener('change', () => {
    updateDefaultMappingName(false);
    resetMappingParsedState();
  });
  $('mappingPurchaseSelect').addEventListener('change', () => {
    updateDefaultMappingName(false);
    resetMappingParsedState();
  });

  $('mappingTemplateHeaderRowInput').addEventListener('change', () => {
    try {
      const autoCount = applyHeaderRowSelection({ keepExistingRules: true });
      setStatus($('mappingStatus'), `已应用标题行，自动匹配 ${autoCount} 条规则`);
    } catch (e) {
      setStatus($('mappingStatus'), e.message || '应用失败', 'error');
    }
  });
  $('mappingPurchaseHeaderRowInput').addEventListener('change', () => {
    try {
      const autoCount = applyHeaderRowSelection({ keepExistingRules: true });
      setStatus($('mappingStatus'), `已应用标题行，自动匹配 ${autoCount} 条规则`);
    } catch (e) {
      setStatus($('mappingStatus'), e.message || '应用失败', 'error');
    }
  });

  $('mappingRulesTbody').addEventListener('change', (e) => {
    const el = e.target;
    if (!(el instanceof HTMLSelectElement)) return;
    const idx = Number(el.getAttribute('data-rule-idx'));
    const field = el.getAttribute('data-field');
    if (!Number.isFinite(idx) || idx < 0 || idx >= mappingRules.length) return;
    if (field === 'template') mappingRules[idx].templateHeader = el.value;
    if (field === 'purchase') mappingRules[idx].purchaseHeader = el.value;
  });

  $('mappingRulesTbody').addEventListener('click', async (e) => {
    const el = e.target;
    if (!(el instanceof HTMLElement)) return;
    const action = el.getAttribute('data-action');
    if (action === 'remove-rule') {
      const idx = Number(el.getAttribute('data-rule-idx'));
      if (!Number.isFinite(idx) || idx < 0 || idx >= mappingRules.length) return;
      mappingRules.splice(idx, 1);
      renderMappingRulesTable();
    }
  });

  $('mappingsTbody').addEventListener('click', async (e) => {
    const el = e.target;
    if (!(el instanceof HTMLElement)) return;
    const action = el.getAttribute('data-action');
    const id = el.getAttribute('data-id');
    if (!action || !id) return;

    if (action === 'delete-mapping') {
      if (!confirm(`确认删除映射 ID=${id} 吗？`)) return;
      try {
        await apiJson(`/api/mappings/${encodeURIComponent(id)}`, { method: 'DELETE' });
        await refreshMappings();
      } catch (err) {
        setStatus($('mappingStatus'), err.message || '删除失败', 'error');
      }
      return;
    }

    if (action === 'view-mapping') {
      try {
        const { data } = await apiJson(`/api/mappings/${encodeURIComponent(id)}`);
        const rulesHtml =
          (data.rules || [])
            .map(
              (r, idx) =>
                `<tr><td>${escapeHtml(idx + 1)}</td><td>${escapeHtml(r.templateHeader)}</td><td>${escapeHtml(
                  r.purchaseHeader
                )}</td></tr>`
            )
            .join('') || `<tr><td colspan="3">暂无规则</td></tr>`;

        openPreviewShell('映射详情', escapeHtml(data.name || ''));
        $('previewBody').innerHTML = `<div class="meta">合同模板标题：${escapeHtml(
          String((data.template_headers || []).join('、') || '')
        )}</div>
        <div class="meta">订购单标题：${escapeHtml(String((data.purchase_headers || []).join('、') || ''))}</div>
        <div class="previewBox">
          <table class="previewTable">
            <thead><tr><th>序号</th><th>合同模板标题</th><th>订购单标题</th></tr></thead>
            <tbody>${rulesHtml}</tbody>
          </table>
        </div>`;
      } catch (err) {
        setStatus($('mappingStatus'), err.message || '加载失败', 'error');
      }
    }
  });

  try {
    await Promise.all([refreshTemplates(), refreshPurchaseOrders(), refreshMappings()]);
    fillSelectOptions($('mappingTemplateSelect'), templatesCache, '请选择合同模板');
    fillSelectOptions($('mappingPurchaseSelect'), purchaseOrdersCache, '请选择订购单');
    renderMappingRulesTable();
  } catch (e) {
    setStatus($('templateStatus'), e.message || '加载失败', 'error');
    setStatus($('purchaseStatus'), e.message || '加载失败', 'error');
  }
}

init();
