const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const os = require('os');
const zlib = require('zlib');
const { execFile } = require('child_process');
const pool = require('../db');

const router = express.Router();

function quoteIdent(ident) {
  return `\`${String(ident).replace(/`/g, '``')}\``;
}

function pickColumn(fields, candidates) {
  const byLower = new Map(fields.map((f) => [String(f).toLowerCase(), f]));
  for (const c of candidates) {
    const hit = byLower.get(String(c).toLowerCase());
    if (hit) return hit;
  }
  return null;
}

function safeBasename(filename) {
  const base = path.basename(String(filename ?? 'file'));
  return base.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_').slice(0, 180) || 'file';
}

function safeBaseName(name) {
  const s = String(name ?? '').trim();
  const base = s.replace(/[<>:"/\\|?*\u0000-\u001F]/g, '_').replace(/\s+/g, ' ');
  return base.slice(0, 180).trim();
}

function escapeRegExp(str) {
  return String(str).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function normalizeWantedBaseName(inputName, uploadedExt, fallbackBaseName) {
  const clean = safeBaseName(inputName);
  if (!clean) return String(fallbackBaseName ?? '').trim();
  const ext = String(uploadedExt ?? '').replace(/^\./, '').trim();
  if (ext) {
    const reUploaded = new RegExp(`\\.${escapeRegExp(ext)}$`, 'i');
    if (reUploaded.test(clean)) return clean.replace(reUploaded, '');
  }
  const reKnown = /\.(xlsx|xls|docx|doc|pdf)$/i;
  if (reKnown.test(clean)) return clean.replace(reKnown, '');
  return clean;
}

function buildUniqueFilename(dir, baseName, ext) {
  const base = safeBaseName(baseName) || 'file';
  const cleanExt = String(ext ?? '').replace(/^\./, '').trim().toLowerCase() || 'bin';
  const tryPath = (n) =>
    path.join(dir, n === 0 ? `${base}.${cleanExt}` : `${base}_${n}.${cleanExt}`);

  for (let n = 0; n < 1000; n++) {
    const p = tryPath(n);
    if (!fs.existsSync(p)) return { filename: path.basename(p), absPath: p };
  }
  throw new Error('无法生成唯一文件名');
}

function scoreFilename(str) {
  const s = String(str ?? '');
  const cjk = (s.match(/[\u4E00-\u9FFF]/g) || []).length;
  const replacement = (s.match(/\uFFFD/g) || []).length;
  const mojibake = (s.match(/[ÃÂÐÑÕØÝÞ]/g) || []).length;
  return cjk * 10 - replacement * 20 - mojibake * 5 - s.length * 0.01;
}

function normalizeIncomingFilename(filename) {
  const base = path.basename(String(filename ?? 'file'));
  const latin1ToUtf8 = Buffer.from(base, 'latin1').toString('utf8');
  return scoreFilename(latin1ToUtf8) > scoreFilename(base) ? latin1ToUtf8 : base;
}

function decodeXmlEntities(str) {
  return String(str ?? '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

function normalizeHeadersFromDocTable(values) {
  if (!Array.isArray(values)) return [];
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

function extractTypeIdFromName(name) {
  const m = String(name ?? '').trim().match(/(\d{3})\s*$/);
  return m ? m[1] : null;
}

function normalizePurchaseTypeId(value) {
  if (value == null) return null;
  const s = String(value).trim();
  if (!s) return null;
  if (/^\d+$/.test(s)) return s.padStart(3, '0');
  return s;
}

function findZipEocdOffset(buf) {
  const sig = 0x06054b50;
  const maxBack = Math.min(buf.length, 0xffff + 22);
  for (let i = buf.length - 22; i >= buf.length - maxBack; i--) {
    if (i < 0) break;
    if (buf.readUInt32LE(i) === sig) return i;
  }
  return -1;
}

function readDocxFileFromZip(buf, entryName) {
  const eocd = findZipEocdOffset(buf);
  if (eocd < 0) throw new Error('docx 文件格式不正确（未找到 EOCD）');

  const cdSize = buf.readUInt32LE(eocd + 12);
  const cdOffset = buf.readUInt32LE(eocd + 16);
  const cdEnd = cdOffset + cdSize;
  if (cdOffset < 0 || cdEnd > buf.length) throw new Error('docx 文件格式不正确（CD 越界）');

  let p = cdOffset;
  const cdSig = 0x02014b50;
  while (p + 46 <= cdEnd) {
    if (buf.readUInt32LE(p) !== cdSig) break;
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commentLen = buf.readUInt16LE(p + 32);
    const localHeaderOffset = buf.readUInt32LE(p + 42);
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen);
    const next = p + 46 + nameLen + extraLen + commentLen;

    if (name === entryName) {
      const lhSig = 0x04034b50;
      if (localHeaderOffset + 30 > buf.length || buf.readUInt32LE(localHeaderOffset) !== lhSig) {
        throw new Error('docx 文件格式不正确（LH 无效）');
      }
      const lhNameLen = buf.readUInt16LE(localHeaderOffset + 26);
      const lhExtraLen = buf.readUInt16LE(localHeaderOffset + 28);
      const dataStart = localHeaderOffset + 30 + lhNameLen + lhExtraLen;
      const dataEnd = dataStart + compSize;
      if (dataEnd > buf.length) throw new Error('docx 文件格式不正确（数据越界）');
      const comp = buf.slice(dataStart, dataEnd);
      if (method === 0) return comp;
      if (method === 8) return zlib.inflateRawSync(comp);
      throw new Error(`docx 压缩算法不支持（method=${method}）`);
    }

    p = next;
  }

  throw new Error(`docx 中未找到 ${entryName}`);
}

function extractDocxFirstTableHeaders(buffer) {
  const xml = readDocxFileFromZip(buffer, 'word/document.xml').toString('utf8');
  const tbl = xml.match(/<w:tbl(?:\s|>)[\s\S]*?<\/w:tbl>/);
  if (!tbl) return [];
  const tr = tbl[0].match(/<w:tr(?:\s|>)[\s\S]*?<\/w:tr>/);
  if (!tr) return [];
  const rowXml = tr[0];
  const cells = rowXml.match(/<w:tc(?:\s|>)[\s\S]*?<\/w:tc>/g) || [];
  const headers = [];
  for (const cell of cells) {
    const parts = [];
    cell.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_m, t) => {
      parts.push(decodeXmlEntities(t));
      return _m;
    });
    cell.replace(/<w:instrText(?:\s[^>]*)?>([\s\S]*?)<\/w:instrText>/g, (_m, t) => {
      parts.push(decodeXmlEntities(t));
      return _m;
    });
    const text = parts.join('').replace(/\s+/g, ' ').trim();
    headers.push(text);
  }
  return normalizeHeadersFromDocTable(headers);
}

function extractDocxPreviewBlocks(buffer) {
  const xml = readDocxFileFromZip(buffer, 'word/document.xml').toString('utf8');
  const body = xml.match(/<w:body(?:\s|>)[\s\S]*?<\/w:body>/)?.[0] ?? xml;
  const blocks = [];

  const takeTextFromXml = (src) => {
    const parts = [];
    src.replace(/<w:tab\s*\/>/g, () => {
      parts.push('\t');
      return '';
    });
    src.replace(/<w:br\s*\/>/g, () => {
      parts.push('\n');
      return '';
    });
    src.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_m, t) => {
      parts.push(decodeXmlEntities(t));
      return _m;
    });
    const text = parts.join('').replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    return text;
  };

  const re = /<w:(p|tbl)(?:\s|>)[\s\S]*?<\/w:\1>/g;
  let m;
  while ((m = re.exec(body))) {
    const tag = m[1];
    const chunk = m[0];

    if (tag === 'p') {
      const text = takeTextFromXml(chunk).replace(/\s+/g, ' ').trim();
      if (text) blocks.push({ type: 'p', text });
    } else {
      const rows = [];
      const rowMatches = chunk.match(/<w:tr(?:\s|>)[\s\S]*?<\/w:tr>/g) || [];
      for (let i = 0; i < Math.min(rowMatches.length, 30); i++) {
        const rowXml = rowMatches[i];
        const cellMatches = rowXml.match(/<w:tc(?:\s|>)[\s\S]*?<\/w:tc>/g) || [];
        const row = [];
        for (let c = 0; c < Math.min(cellMatches.length, 20); c++) {
          const cellText = takeTextFromXml(cellMatches[c]).replace(/\s+/g, ' ').trim();
          row.push(cellText);
        }
        if (row.some((v) => String(v || '').trim())) rows.push(row);
      }
      if (rows.length) blocks.push({ type: 'table', rows });
    }

    if (blocks.length >= 120) break;
  }

  return blocks;
}

async function extractDocFirstTableHeaders(buffer) {
  if (process.platform !== 'win32') throw new Error('doc 格式解析仅支持在 Windows 环境运行');
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'doc_headers_'));
  const docPath = path.join(tmpDir, `input_${Date.now()}_${Math.random().toString(16).slice(2)}.doc`);
  fs.writeFileSync(docPath, buffer);

  const ps = `
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$path = '${docPath.replace(/'/g, "''")}'
$word = $null
$doc = $null
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $doc = $word.Documents.Open($path, $false, $true)
  if ($doc.Tables.Count -lt 1) { throw '文档中未找到表格' }
  $tbl = $doc.Tables.Item(1)
  if ($tbl.Rows.Count -lt 1) { throw '表格无行' }
  $row = $tbl.Rows.Item(1)
  $headers = @()
  foreach ($cell in $row.Cells) {
    $t = $cell.Range.Text
    $t = $t -replace \"[\\r\\a]\", ''
    $t = $t.Trim()
    if ($t) { $headers += $t }
  }
  $headers | ConvertTo-Json -Compress
} finally {
  if ($doc) { $doc.Close([ref]0) | Out-Null }
  if ($word) { $word.Quit() | Out-Null }
}
`;

  try {
    const headersJson = await new Promise((resolve, reject) => {
      execFile(
        'powershell.exe',
        ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', ps],
        { timeout: 20000, maxBuffer: 1024 * 1024 },
        (err, stdout, stderr) => {
          if (err) {
            reject(new Error((stderr || stdout || err.message || '').toString().trim() || 'doc 解析失败'));
            return;
          }
          resolve(String(stdout || '').trim());
        }
      );
    });
    const parsed = JSON.parse(headersJson || '[]');
    return normalizeHeadersFromDocTable(parsed);
  } finally {
    try {
      fs.unlinkSync(docPath);
    } catch {}
    try {
      fs.rmdirSync(tmpDir, { recursive: true });
    } catch {}
  }
}

function encodeRFC5987(str) {
  return encodeURIComponent(String(str))
    .replace(/['()]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`)
    .replace(/\*/g, '%2A');
}

function toSafeAsciiFilename(filename) {
  const s = String(filename ?? 'file');
  const ascii = s.replace(/[^\x20-\x7E]/g, '_');
  return ascii || 'file';
}

function buildContentDisposition(filename) {
  const ascii = toSafeAsciiFilename(filename);
  const utf8 = encodeRFC5987(filename);
  return `attachment; filename="${ascii}"; filename*=UTF-8''${utf8}`;
}

function guessMimeFromFilePath(filePath) {
  const ext = path.extname(String(filePath ?? '')).toLowerCase();
  if (ext === '.xlsx') return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (ext === '.xls') return 'application/vnd.ms-excel';
  if (ext === '.pdf') return 'application/pdf';
  if (ext === '.docx') return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (ext === '.doc') return 'application/msword';
  return 'application/octet-stream';
}

function extractTypeIdFromName(value) {
  const s = String(value ?? '').trim();
  const m = s.match(/(\d{3})$/);
  return m ? m[1] : null;
}

function resolveAbsFilePath(filePathValue) {
  const raw = String(filePathValue ?? '').trim();
  if (!raw) return null;
  if (/^https?:\/\//i.test(raw)) return null;
  if (path.isAbsolute(raw)) return fs.existsSync(raw) ? raw : null;

  const normalized = raw.replace(/^[/\\]+/, '');
  const candidates = [
    path.join(__dirname, '..', normalized),
    path.join(__dirname, '..', '..', normalized),
    path.join(__dirname, '..', '..', '合同生成', normalized)
  ];
  for (const p of candidates) {
    if (fs.existsSync(p)) return p;
  }
  return null;
}

function tryEnsureDbJunction() {
  const link = path.resolve(__dirname, '..', '数据库');
  const target = path.resolve(__dirname, '..', '..', '数据库');
  try {
    if (!fs.existsSync(link) && fs.existsSync(target)) {
      fs.symlinkSync(target, link, 'junction');
    }
  } catch {}
}

function pickWritableDir(candidates) {
  for (const dir of candidates) {
    try {
      fs.mkdirSync(dir, { recursive: true });
      const probe = path.join(
        dir,
        `.write_probe_${process.pid}_${Date.now()}_${Math.random().toString(16).slice(2)}`
      );
      const fd = fs.openSync(probe, 'wx');
      fs.closeSync(fd);
      fs.unlinkSync(probe);
      return dir;
    } catch {}
  }
  throw new Error('存储目录不可写');
}

function getPurchaseOrdersStorageRoot() {
  const preferred = path.resolve(__dirname, '..', '..', '数据库', '采购单列表');
  const viaLocal = path.resolve(__dirname, '..', '数据库', '采购单列表');
  tryEnsureDbJunction();
  return pickWritableDir([preferred, viaLocal]);
}

let purchaseConfigPromise = null;
async function resolvePurchaseOrdersConfig() {
  if (purchaseConfigPromise) return purchaseConfigPromise;
  purchaseConfigPromise = (async () => {
    const [tableRows] = await pool.query('SHOW TABLES');
    const tableNames = tableRows
      .map((r) => r[Object.keys(r)[0]])
      .filter(Boolean)
      .map((t) => String(t));

    const table =
      tableNames.find((t) => t.toLowerCase() === 'purchase_order') ||
      tableNames.find((t) => t === '采购单列表') ||
      tableNames.find((t) => t.includes('采购')) ||
      tableNames.find((t) => t.toLowerCase().includes('purchase')) ||
      null;

    if (!table) {
      throw new Error('未找到采购单表');
    }
                 
    const [descRows] = await pool.query(`DESCRIBE ${quoteIdent(table)}`);
    const fields = descRows.map((r) => r.Field);

    const columns = {
      id: pickColumn(fields, ['id', 'Id', 'ID']),
      name: pickColumn(fields, ['name', 'Name']),
      filePath: pickColumn(fields, ['file_path', 'File_path', 'filepath', 'path']),
      createdAt: pickColumn(fields, ['created_at', 'Created_at', 'create_time', 'createdTime']),
      typeId: pickColumn(fields, ['type_id', 'Type_id', 'type_ID', 'Type_ID'])
    };

    if (!columns.id || !columns.name || !columns.filePath) {
      throw new Error('采购单表缺少必要字段（Id/Name/File_path）');
    }

    return { table, columns };
  })();
  return purchaseConfigPromise;
}

async function ensurePurchaseCreatedAt(config) {
  const table = config.table;
  const [lowerRows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(table)} LIKE ?`, ['created_at']);
  const [upperRows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(table)} LIKE ?`, ['Created_at']);

  if (lowerRows.length > 0) {
    config.columns.createdAt = 'created_at';
    return config;
  }
  if (upperRows.length > 0) {
    config.columns.createdAt = 'Created_at';
    return config;
  }

  const col = 'created_at';
  await pool.query(
    `ALTER TABLE ${quoteIdent(table)} ADD COLUMN ${quoteIdent(col)} TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP`
  );
  config.columns.createdAt = col;
  return config;
}

async function ensurePurchaseTypeId(config) {
  const table = config.table;
  const [upperRows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(table)} LIKE ?`, ['type_ID']);
  const [lowerRows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(table)} LIKE ?`, ['type_id']);
  const [camelRows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(table)} LIKE ?`, ['Type_ID']);

  if (upperRows.length > 0) {
    config.columns.typeId = 'type_ID';
    return config;
  }
  if (lowerRows.length > 0) {
    config.columns.typeId = 'type_id';
    return config;
  }
  if (camelRows.length > 0) {
    config.columns.typeId = 'Type_ID';
    return config;
  }

  const col = 'type_ID';
  await pool.query(`ALTER TABLE ${quoteIdent(table)} ADD COLUMN ${quoteIdent(col)} VARCHAR(10) NULL`);
  config.columns.typeId = col;
  return config;
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024 }
});

router.get('/', async (req, res) => {
  try {
    const config = await ensurePurchaseTypeId(await ensurePurchaseCreatedAt(await resolvePurchaseOrdersConfig()));
    const { table, columns } = config;

    const select = [
      `${quoteIdent(columns.id)} AS id`,
      `${quoteIdent(columns.name)} AS name`,
      `${quoteIdent(columns.filePath)} AS file_path`,
      columns.createdAt ? `${quoteIdent(columns.createdAt)} AS created_at` : `NULL AS created_at`,
      columns.typeId ? `${quoteIdent(columns.typeId)} AS type_id` : `NULL AS type_id`
    ].join(', ');

    const [rows] = await pool.query(
      `SELECT ${select} FROM ${quoteIdent(table)} ORDER BY ${quoteIdent(columns.id)} ASC`
    );
    res.json({ success: true, data: rows ?? [] });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '服务器错误' });
  }
});

router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const file = req.file;
    if (!file) return res.status(400).json({ success: false, error: '未选择文件' });

    const config = await ensurePurchaseTypeId(await ensurePurchaseCreatedAt(await resolvePurchaseOrdersConfig()));
    const { table, columns } = config;

    const originalName = safeBasename(normalizeIncomingFilename(file.originalname));
    const originalBaseName = path.parse(originalName).name || originalName;
    const ext = path.extname(originalName).replace(/^\./, '').toLowerCase() || 'bin';
    const wantedBaseName = normalizeWantedBaseName(req.body?.name, ext, originalBaseName);
    const name = wantedBaseName;

    const purchaseDir = getPurchaseOrdersStorageRoot();
    const { filename: diskName, absPath } = buildUniqueFilename(purchaseDir, wantedBaseName, ext);
    fs.writeFileSync(absPath, file.buffer);

    const relativePath = path.join('数据库', '采购单列表', diskName).replace(/\\/g, '/');

    const typeId =
      extractTypeIdFromName(wantedBaseName) ||
      extractTypeIdFromName(originalBaseName) ||
      extractTypeIdFromName(path.parse(diskName).name);

    const [result] = await pool.query(
      `INSERT INTO ${quoteIdent(table)} (${quoteIdent(columns.name)}, ${quoteIdent(columns.filePath)}, ${quoteIdent(
        columns.createdAt
      )}, ${quoteIdent(columns.typeId)}) VALUES (?, ?, ?, ?)`,
      [name, relativePath, new Date(), typeId]
    );

    res.json({ success: true, data: { id: result.insertId } });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '上传失败' });
  }
});

router.get('/:id/download', async (req, res) => {
  const { id } = req.params;
  try {
    const config = await resolvePurchaseOrdersConfig();
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '采购单不存在' });

    const row = rows[0];
    const filePathValue = row[columns.filePath] ?? row.file_path ?? row.File_path;
    const abs = resolveAbsFilePath(filePathValue);
    if (!abs) return res.status(404).json({ success: false, error: '文件不存在' });

    const buffer = fs.readFileSync(abs);
    const filename = path.basename(abs) || `purchase_${id}.xlsx`;
    const mime = guessMimeFromFilePath(abs);

    res.setHeader('Content-Type', mime);
    res.setHeader('Content-Disposition', buildContentDisposition(filename));
    res.setHeader('Content-Length', buffer.length);
    res.status(200).send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '下载失败' });
  }
});

router.get('/:id/headers', async (req, res) => {
  const { id } = req.params;
  try {
    const config = await resolvePurchaseOrdersConfig();
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '采购单不存在' });

    const row = rows[0];
    const filePathValue = row[columns.filePath] ?? row.file_path ?? row.File_path;
    const abs = resolveAbsFilePath(filePathValue);
    if (!abs) return res.status(404).json({ success: false, error: '文件不存在' });

    const buffer = fs.readFileSync(abs);
    const ext = path.extname(abs).replace(/^\./, '').trim().toLowerCase();
    if (ext !== 'docx' && ext !== 'doc') {
      return res.status(415).json({ success: false, error: '仅支持 doc/docx 解析标题' });
    }

    const headers = ext === 'docx' ? extractDocxFirstTableHeaders(buffer) : await extractDocFirstTableHeaders(buffer);
    if (!headers.length) return res.status(400).json({ success: false, error: '未从文档表格第一行解析到标题' });

    res.json({ success: true, data: { headers, source: { type: ext, table: 1, row: 1 } } });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '解析标题失败' });
  }
});

router.get('/:id/preview', async (req, res) => {
  const { id } = req.params;
  try {
    const config = await resolvePurchaseOrdersConfig();
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '采购单不存在' });

    const row = rows[0];
    const filePathValue = row[columns.filePath] ?? row.file_path ?? row.File_path;
    const abs = resolveAbsFilePath(filePathValue);
    if (!abs) return res.status(404).json({ success: false, error: '文件不存在' });

    const ext = path.extname(abs).replace(/^\./, '').trim().toLowerCase();
    if (ext !== 'docx') return res.status(415).json({ success: false, error: '仅支持 docx 预览' });

    const buffer = fs.readFileSync(abs);
    const blocks = extractDocxPreviewBlocks(buffer);
    if (!blocks.length) return res.status(400).json({ success: false, error: '未解析到可预览内容' });

    res.json({ success: true, data: { type: 'docx', blocks } });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '预览解析失败' });
  }
});

router.delete('/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const config = await resolvePurchaseOrdersConfig();
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '采购单不存在' });

    const row = rows[0];
    const purchaseTypeId =
      normalizePurchaseTypeId(columns.typeId ? row[columns.typeId] : null) ||
      normalizePurchaseTypeId(extractTypeIdFromName(columns.name ? row[columns.name] : null));
    const filePathValue = row[columns.filePath] ?? row.file_path ?? row.File_path;
    let fileDeleted = false;
    let fileMissing = false;
    if (filePathValue) {
      const raw = String(filePathValue).trim().replace(/\\/g, '/');
      let abs = null;
      if (path.isAbsolute(raw)) {
        abs = raw;
      } else {
        const normalized = raw.replace(/^[/\\]+/, '');
        const localAbs = path.join(__dirname, '..', normalized);
        abs = fs.existsSync(localAbs) ? localAbs : resolveAbsFilePath(raw);
      }
      if (!abs) {
        fileMissing = true;
      } else {
        const resolved = path.resolve(abs);
        const allowedRoots = [
          path.resolve(__dirname, '..', '数据库', '采购单列表') + path.sep,
          path.resolve(__dirname, '..', '..', '数据库', '采购单列表') + path.sep,
          path.resolve(__dirname, '..', 'uploads') + path.sep,
          path.resolve(__dirname, '..', '..', '合同生成', 'templates') + path.sep
        ];
        const allowed = allowedRoots.some((r) => resolved.startsWith(r));
        if (!allowed) {
          return res.status(500).json({ success: false, error: '文件路径不允许删除' });
        }
        try {
          fs.unlinkSync(resolved);
          fileDeleted = true;
        } catch (e) {
          return res.status(500).json({ success: false, error: `删除文件失败：${e.code || 'ERROR'}` });
        }
      }
    }

    await pool.query(
      `CREATE TABLE IF NOT EXISTS ${quoteIdent('mapping_config')} (
        ${quoteIdent('ID')} INT NOT NULL AUTO_INCREMENT,
        ${quoteIdent('config_name')} VARCHAR(255) NULL,
        ${quoteIdent('Template_id')} INT NULL,
        ${quoteIdent('Purchase_id')} VARCHAR(10) NULL,
        ${quoteIdent('contract_headers')} JSON NULL,
        ${quoteIdent('purchase_headers')} JSON NULL,
        ${quoteIdent('created_at')} TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (${quoteIdent('ID')})
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`
    );

    await pool.query(
      `CREATE TABLE IF NOT EXISTS ${quoteIdent('mapping_rule')} (
        ${quoteIdent('ID')} INT NOT NULL AUTO_INCREMENT,
        ${quoteIdent('mapping_id')} INT NOT NULL,
        ${quoteIdent('contract_header')} VARCHAR(255) NOT NULL,
        ${quoteIdent('purchase_header')} VARCHAR(255) NOT NULL,
        PRIMARY KEY (${quoteIdent('ID')}),
        INDEX ${quoteIdent('idx_mapping_rule_mapping_id')} (${quoteIdent('mapping_id')})
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`
    );

    const conn = await pool.getConnection();
    let deletedMappingCount = 0;
    try {
      await conn.beginTransaction();

      if (purchaseTypeId) {
        const [mappingRows] = await conn.query(
          `SELECT ${quoteIdent('ID')} AS id FROM ${quoteIdent('mapping_config')} WHERE ${quoteIdent('Purchase_id')} = ?`,
          [purchaseTypeId]
        );
        const mappingIds = (mappingRows ?? []).map((r) => r.id).filter((v) => v != null);
        deletedMappingCount = mappingIds.length;

        if (mappingIds.length > 0) {
          await conn.query(
            `DELETE FROM ${quoteIdent('mapping_rule')} WHERE ${quoteIdent('mapping_id')} IN (${mappingIds
              .map(() => '?')
              .join(', ')})`,
            mappingIds
          );
        }
        await conn.query(`DELETE FROM ${quoteIdent('mapping_config')} WHERE ${quoteIdent('Purchase_id')} = ?`, [
          purchaseTypeId
        ]);
      }

      await conn.query(`DELETE FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`, [id]);
      await conn.commit();

      res.json({ success: true, data: { id: Number(id), fileDeleted, fileMissing, deletedMappingCount } });
    } catch (e) {
      try {
        await conn.rollback();
      } catch {}
      throw e;
    } finally {
      conn.release();
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '删除失败' });
  }
});

module.exports = router;
