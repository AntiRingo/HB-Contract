const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const zlib = require('zlib');
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

function normalizeToBuffer(value) {
  if (!value) return null;
  if (Buffer.isBuffer(value)) return value;
  if (value instanceof ArrayBuffer) return Buffer.from(value);
  if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
  if (typeof value === 'string') return Buffer.from(value, 'base64');
  return null;
}

function guessMimeFromExtOrType(fileType) {
  const t = String(fileType ?? '').trim().toLowerCase();
  if (t.includes('xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  if (t.includes('xls')) return 'application/vnd.ms-excel';
  if (t.includes('docx')) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (t.includes('doc')) return 'application/msword';
  if (t.includes('pdf')) return 'application/pdf';
  return 'application/octet-stream';
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

function getTemplatesStorageRoot() {
  const preferred = path.resolve(__dirname, '..', '..', '数据库', '合同模板列表');
  const viaLocal = path.resolve(__dirname, '..', '数据库', '合同模板列表');
  tryEnsureDbJunction();
  return pickWritableDir([preferred, viaLocal]);
}

let templatesConfigPromise = null;
async function resolveTemplatesConfig() {
  if (templatesConfigPromise) return templatesConfigPromise;
  templatesConfigPromise = (async () => {
    const [tableRows] = await pool.query('SHOW TABLES');
    const tableNames = tableRows
      .map((r) => r[Object.keys(r)[0]])
      .filter(Boolean)
      .map((t) => String(t));

    const table =
      tableNames.find((t) => t.toLowerCase() === 'contract_template') ||
      tableNames.find((t) => t.toLowerCase() === 'contract_templates') ||
      tableNames.find((t) => t === '合同模板列表') ||
      tableNames.find((t) => t.includes('合同模板')) ||
      tableNames.find((t) => t.includes('模板')) ||
      null;

    if (!table) {
      throw new Error('未找到公司合同模板表');
    }

    const [descRows] = await pool.query(`DESCRIBE ${quoteIdent(table)}`);
    const fields = descRows.map((r) => r.Field);

    const columns = {
      id: pickColumn(fields, ['id', 'Id', 'ID']),
      name: pickColumn(fields, ['name', 'Name']),
      filePath: pickColumn(fields, ['file_path', 'File_path', 'filepath', 'path']),
      fileType: pickColumn(fields, ['file_type', 'File_type', 'mime_type', 'content_type']),
      fileName: pickColumn(fields, ['file_name', 'File_name', 'original_name', 'Original_name']),
      fileBlob: pickColumn(fields, [
        'file_data',
        'file_blob',
        'file_content',
        'File_blob',
        'File_data',
        'File_content',
        'content',
        'data'
      ]),
      createdAt: pickColumn(fields, ['created_at', 'Created_at', 'create_time', 'createdTime'])
    };

    if (!columns.id || !columns.name) {
      throw new Error('公司合同模板表缺少必要字段（Id/Name）');
    }

    const byField = new Map(descRows.map((r) => [r.Field, r]));
    const idRow = byField.get(columns.id);
    const idAutoIncrement = String(idRow?.Extra ?? '').toLowerCase().includes('auto_increment');

    return { table, columns, idAutoIncrement };
  })();
  return templatesConfigPromise;
}

async function ensureBlobSupport(config) {
  if (config.columns.fileBlob && config.columns.fileType) return config;
  const table = config.table;

  const ensure = async (col, def) => {
    const [rows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(table)} LIKE ?`, [col]);
    if (rows.length === 0) {
      await pool.query(`ALTER TABLE ${quoteIdent(table)} ADD COLUMN ${quoteIdent(col)} ${def}`);
    }
  };

  if (!config.columns.fileBlob) {
    await ensure('File_blob', 'LONGBLOB NULL');
    config.columns.fileBlob = 'File_blob';
  }
  if (!config.columns.fileType) {
    await ensure('File_type', 'VARCHAR(50) NULL');
    config.columns.fileType = 'File_type';
  }
  if (!config.columns.fileName) {
    await ensure('File_name', 'VARCHAR(255) NULL');
    config.columns.fileName = 'File_name';
  }
  if (!config.columns.filePath) {
    await ensure('File_path', 'VARCHAR(255) NULL');
    config.columns.filePath = 'File_path';
  }

  return config;
}

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 30 * 1024 * 1024 }
});

router.get('/', async (req, res) => {
  try {
    const config = await resolveTemplatesConfig();
    const { table, columns } = config;

    const select = [
      `${quoteIdent(columns.id)} AS id`,
      `${quoteIdent(columns.name)} AS name`,
      columns.fileName ? `${quoteIdent(columns.fileName)} AS file_name` : `NULL AS file_name`,
      columns.filePath ? `${quoteIdent(columns.filePath)} AS file_path` : `NULL AS file_path`,
      columns.fileType ? `${quoteIdent(columns.fileType)} AS file_type` : `NULL AS file_type`,
      columns.createdAt ? `${quoteIdent(columns.createdAt)} AS created_at` : `NULL AS created_at`
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

    const config = await ensureBlobSupport(await resolveTemplatesConfig());
    const { table, columns } = config;

    const originalName = safeBasename(normalizeIncomingFilename(file.originalname));
    const ext = path.extname(originalName).replace(/^\./, '').toLowerCase() || 'bin';
    const originalBaseName = path.parse(originalName).name || originalName;
    const wantedBaseName = normalizeWantedBaseName(req.body?.name, ext, originalBaseName);
    const name = wantedBaseName;

    const templatesDir = getTemplatesStorageRoot();
    const { filename: diskName, absPath } = buildUniqueFilename(templatesDir, wantedBaseName, ext);
    const relativePath = path.join('数据库', '合同模板列表', diskName).replace(/\\/g, '/');
    fs.writeFileSync(absPath, file.buffer);

    const insertCols = [];
    const placeholders = [];
    const params = [];

    insertCols.push(columns.name);
    placeholders.push('?');
    params.push(name);

    if (columns.filePath) {
      insertCols.push(columns.filePath);
      placeholders.push('?');
      params.push(relativePath);
    }

    if (columns.createdAt) {
      insertCols.push(columns.createdAt);
      placeholders.push('?');
      params.push(new Date());
    }

    if (columns.fileBlob) {
      insertCols.push(columns.fileBlob);
      placeholders.push('?');
      params.push(file.buffer);
    }

    if (columns.fileType) {
      insertCols.push(columns.fileType);
      placeholders.push('?');
      params.push(ext);
    }

    if (columns.fileName) {
      insertCols.push(columns.fileName);
      placeholders.push('?');
      params.push(diskName);
    }

    const doInsert = async (executor) => {
      const sql = `INSERT INTO ${quoteIdent(table)} (${insertCols
        .map(quoteIdent)
        .join(', ')}) VALUES (${placeholders.join(', ')})`;
      return executor.query(sql, params);
    };

    try {
      if (config.idAutoIncrement) {
        const [result] = await doInsert(pool);
        res.json({ success: true, data: { id: result.insertId } });
        return;
      }

      const conn = await pool.getConnection();
      try {
        await conn.beginTransaction();
        const [rows] = await conn.query(
          `SELECT IFNULL(MAX(${quoteIdent(columns.id)}), 0) + 1 AS next_id FROM ${quoteIdent(table)} FOR UPDATE`
        );
        const nextId = rows?.[0]?.next_id;
        if (!nextId) throw new Error('无法生成新的 Id');

        insertCols.unshift(columns.id);
        placeholders.unshift('?');
        params.unshift(nextId);

        await doInsert(conn);
        await conn.commit();
        res.json({ success: true, data: { id: nextId } });
      } catch (e) {
        try {
          await conn.rollback();
        } catch {}
        throw e;
      } finally {
        conn.release();
      }
    } catch (e) {
      try {
        fs.unlinkSync(absPath);
      } catch {}
      throw e;
    }
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '上传失败', detail: err?.code || err?.message || String(err) });
  }
});

router.get('/:id/download', async (req, res) => {
  const { id } = req.params;
  try {
    const config = await ensureBlobSupport(await resolveTemplatesConfig());
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '模板不存在' });
    const row = rows[0];

    const fileType =
      (columns.fileType ? row[columns.fileType] : null) ??
      row.file_type ??
      row.File_type ??
      row.mime_type ??
      row.content_type;
    const filename =
      (columns.fileName ? row[columns.fileName] : null) ??
      row.file_name ??
      row.original_name ??
      row[columns.name] ??
      `template_${id}.bin`;

    const fileData =
      (columns.fileBlob ? row[columns.fileBlob] : null) ??
      row.file_data ??
      row.file_blob ??
      row.file_content ??
      row.File_blob ??
      row.File_data ??
      row.File_content ??
      row.content ??
      row.data;

    const buffer = normalizeToBuffer(fileData);
    if (buffer) {
      res.setHeader('Content-Type', guessMimeFromExtOrType(fileType));
      res.setHeader('Content-Disposition', buildContentDisposition(filename));
      res.setHeader('Content-Length', buffer.length);
      return res.status(200).send(buffer);
    }

    const filePathValue = (columns.filePath ? row[columns.filePath] : null) ?? row.file_path ?? row.File_path;
    if (!filePathValue) return res.status(500).json({ success: false, error: '模板文件数据缺失' });

    const abs = resolveAbsFilePath(filePathValue);
    if (!abs) return res.status(404).json({ success: false, error: '文件不存在' });

    const diskBuffer = fs.readFileSync(abs);
    const inferredType = fileType ?? path.extname(abs).replace(/^\./, '').toLowerCase() ?? 'bin';
    const inferredName = path.basename(abs) || filename;

    res.setHeader('Content-Type', guessMimeFromExtOrType(inferredType));
    res.setHeader('Content-Disposition', buildContentDisposition(inferredName));
    res.setHeader('Content-Length', diskBuffer.length);
    res.status(200).send(diskBuffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '下载失败' });
  }
});

router.delete('/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const config = await ensureBlobSupport(await resolveTemplatesConfig());
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '模板不存在' });
    const row = rows[0];

    const filePathValue = (columns.filePath ? row[columns.filePath] : null) ?? row.file_path ?? row.File_path;
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
          path.resolve(__dirname, '..', '数据库', '合同模板列表') + path.sep,
          path.resolve(__dirname, '..', '..', '数据库', '合同模板列表') + path.sep,
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

    await pool.query(`DELETE FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`, [id]);
    res.json({ success: true, data: { id: Number(id), fileDeleted, fileMissing } });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '删除失败' });
  }
});

function decodeXmlEntities(str) {
  return String(str ?? '')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
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
  while (p < cdEnd) {
    if (buf.readUInt32LE(p) !== 0x02014b50) break;
    const method = buf.readUInt16LE(p + 10);
    const compSize = buf.readUInt32LE(p + 20);
    const nameLen = buf.readUInt16LE(p + 28);
    const extraLen = buf.readUInt16LE(p + 30);
    const commLen = buf.readUInt16LE(p + 32);
    const localHeaderOffset = buf.readUInt32LE(p + 42);
    const name = buf.slice(p + 46, p + 46 + nameLen).toString('utf8');
    const next = p + 46 + nameLen + extraLen + commLen;

    if (name === entryName) {
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

function extractDocxPreviewBlocks(buffer) {
  const xml = readDocxFileFromZip(buffer, 'word/document.xml').toString('utf8');
  const bodyMatch = xml.match(/<w:body(?:\s|>)[\s\S]*?<\/w:body>/);
  const body = bodyMatch ? bodyMatch[0] : xml;
  const blocks = [];

  const takeTextFromXml = (src) => {
    const parts = [];
    // 提取所有文本节点的内容
    src.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_m, t) => {
      parts.push(decodeXmlEntities(t));
      return _m;
    });
    return parts.join('').trim();
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
      for (let i = 0; i < Math.min(rowMatches.length, 60); i++) {
        const rowXml = rowMatches[i];
        const cellMatches = rowXml.match(/<w:tc(?:\s|>)[\s\S]*?<\/w:tc>/g) || [];
        const row = [];
        for (let c = 0; c < Math.min(cellMatches.length, 30); c++) {
          const cellText = takeTextFromXml(cellMatches[c]).replace(/\s+/g, ' ').trim();
          row.push(cellText);
        }
        if (row.some((v) => String(v || '').trim())) rows.push(row);
      }
      if (rows.length) blocks.push({ type: 'table', rows });
    }

    if (blocks.length >= 200) break;
  }

  return blocks;
}

function extractDocxTableRows(buffer) {
  const xml = readDocxFileFromZip(buffer, 'word/document.xml').toString('utf8');
  const bodyMatch = xml.match(/<w:body(?:\s|>)[\s\S]*?<\/w:body>/);
  const body = bodyMatch ? bodyMatch[0] : xml;

  const takeTextFromXml = (src) => {
    const parts = [];
    // 提取所有文本节点的内容
    src.replace(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g, (_m, t) => {
      parts.push(decodeXmlEntities(t));
      return _m;
    });
    return parts.join('').trim();
  };

  const tblMatch = body.match(/<w:tbl(?:\s|>)[\s\S]*?<\/w:tbl>/);
  if (!tblMatch) return [];

  const rows = [];
  const rowMatches = tblMatch[0].match(/<w:tr(?:\s|>)[\s\S]*?<\/w:tr>/g) || [];
  for (let i = 0; i < Math.min(rowMatches.length, 100); i++) {
    const rowXml = rowMatches[i];
    const cellMatches = rowXml.match(/<w:tc(?:\s|>)[\s\S]*?<\/w:tc>/g) || [];
    const row = [];
    for (let c = 0; c < Math.min(cellMatches.length, 50); c++) {
      row.push(takeTextFromXml(cellMatches[c]).replace(/\s+/g, ' ').trim());
    }
    rows.push(row);
  }
  return rows;
}

function detectBestHeaderRow(rows) {
  if (!rows || rows.length === 0) return 1;

  const keywords = ['编号', '名称', '规格', '型号', '单位', '数量', '单价', '总价', '金额', '备注', '项目', '货物'];
  let bestRow = 1;
  let maxScore = -1;

  for (let i = 0; i < Math.min(rows.length, 10); i++) {
    const row = rows[i];
    const nonBlank = row.filter((c) => String(c || '').trim()).length;
    if (nonBlank === 0) continue;

    let score = nonBlank * 2;
    for (const cell of row) {
      const text = String(cell || '');
      if (keywords.some((k) => text.includes(k))) score += 5;
      if (text.length > 0 && text.length <= 10) score += 1;
    }

    if (score > maxScore) {
      maxScore = score;
      bestRow = i + 1;
    }
  }
  return bestRow;
}

router.get('/:id/headers', async (req, res) => {
  const { id } = req.params;
  const requestedRow = parseInt(req.query.row);

  try {
    const config = await resolveTemplatesConfig();
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '模板不存在' });
    const row = rows[0];

    const fileData =
      (columns.fileBlob ? row[columns.fileBlob] : null) ??
      row.file_data ??
      row.file_blob ??
      row.file_content ??
      row.File_blob ??
      row.File_data ??
      row.File_content ??
      row.content ??
      row.data;

    let buffer = normalizeToBuffer(fileData);
    if (!buffer) {
      const filePathValue = (columns.filePath ? row[columns.filePath] : null) ?? row.file_path ?? row.File_path;
      const abs = resolveAbsFilePath(filePathValue);
      if (!abs) return res.status(404).json({ success: false, error: '文件不存在' });
      buffer = fs.readFileSync(abs);
    }

    const tableRows = extractDocxTableRows(buffer);
    if (!tableRows.length) return res.status(400).json({ success: false, error: '未在文档中找到表格' });

    const bestRow = isNaN(requestedRow) ? detectBestHeaderRow(tableRows) : requestedRow;
    const rowIndex = Math.max(1, Math.min(bestRow, tableRows.length)) - 1;
    const headers = tableRows[rowIndex] || [];

    res.json({
      success: true,
      data: {
        headers,
        bestRow: rowIndex + 1,
        rowCount: tableRows.length
      }
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '解析标题失败' });
  }
});

router.get('/:id/preview', async (req, res) => {
  const { id } = req.params;
  try {
    const config = await resolveTemplatesConfig();
    const { table, columns } = config;

    const [rows] = await pool.query(
      `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '模板不存在' });
    const row = rows[0];

    const fileData =
      (columns.fileBlob ? row[columns.fileBlob] : null) ??
      row.file_data ??
      row.file_blob ??
      row.file_content ??
      row.File_blob ??
      row.File_data ??
      row.File_content ??
      row.content ??
      row.data;

    let buffer = normalizeToBuffer(fileData);
    if (!buffer) {
      const filePathValue = (columns.filePath ? row[columns.filePath] : null) ?? row.file_path ?? row.File_path;
      const abs = resolveAbsFilePath(filePathValue);
      if (!abs) return res.status(404).json({ success: false, error: '文件不存在' });
      buffer = fs.readFileSync(abs);
    }

    const blocks = extractDocxPreviewBlocks(buffer);
    if (!blocks.length) return res.status(400).json({ success: false, error: '未解析到可预览内容' });

    res.json({ success: true, data: { type: 'docx', blocks } });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '预览解析失败' });
  }
});

module.exports = router;
