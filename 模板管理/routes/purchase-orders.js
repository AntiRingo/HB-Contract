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

    await pool.query(`DELETE FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`, [id]);
    res.json({ success: true, data: { id: Number(id), fileDeleted, fileMissing } });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '删除失败' });
  }
});

module.exports = router;
