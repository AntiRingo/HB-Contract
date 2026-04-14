const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
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

    return { table, columns };
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
      `SELECT ${select} FROM ${quoteIdent(table)} ORDER BY ${quoteIdent(columns.id)} DESC`
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
    const wantedInput = safeBaseName(String(req.body?.name ?? ''));
    const wantedBaseName = (wantedInput ? path.parse(wantedInput).name : '') || originalBaseName;
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

    const sql = `INSERT INTO ${quoteIdent(table)} (${insertCols
      .map(quoteIdent)
      .join(', ')}) VALUES (${placeholders.join(', ')})`;

    const [result] = await pool.query(sql, params);
    res.json({ success: true, data: { id: result.insertId } });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '上传失败' });
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

module.exports = router;
