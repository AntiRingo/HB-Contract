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
    html += `<th>${escapeHtml(String(c))}</th>`;
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
  tbody.innerHTML = '<tr><td colspan="5">正在加载...</td></tr>';
  const { data } = await apiJson('/api/templates');
  if (!data || data.length === 0) {
    tbody.innerHTML = '<tr><td colspan="5">暂无数据</td></tr>';
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
  } catch (e) {
    setStatus(status, e.message || '上传失败', 'error');
  } finally {
    btn.disabled = false;
  }
}

async function init() {
  $('uploadTemplateBtn').addEventListener('click', handleUploadTemplate);
  $('uploadPurchaseBtn').addEventListener('click', handleUploadPurchase);
  $('templateFile').addEventListener('change', () => {
    const input = $('templateName');
    const fileInput = $('templateFile');
    const file = fileInput.files && fileInput.files[0];
    if (!file) return;
    if (!String(input.value || '').trim()) input.value = stripFileExtension(file.name);
  });
  $('purchaseFile').addEventListener('change', () => {
    const input = $('purchaseName');
    const fileInput = $('purchaseFile');
    const file = fileInput.files && fileInput.files[0];
    if (!file) return;
    if (!String(input.value || '').trim()) input.value = stripFileExtension(file.name);
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
      } catch (err) {
        setStatus($('purchaseStatus'), err.message || '删除失败', 'error');
      }
    }
  });

  $('previewClose').addEventListener('click', closePreview);
  $('previewOverlay').addEventListener('click', (e) => {
    if (e.target === $('previewOverlay')) closePreview();
  });

  try {
    await Promise.all([refreshTemplates(), refreshPurchaseOrders()]);
  } catch (e) {
    setStatus($('templateStatus'), e.message || '加载失败', 'error');
    setStatus($('purchaseStatus'), e.message || '加载失败', 'error');
  }
}

init();
