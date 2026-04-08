const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const pool = require('../db');

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
            tableNames.find((t) => t === '公司合同模板') ||
            tableNames.find((t) => t.includes('模板')) ||
            null;

        if (!table) {
            throw new Error('未找到模板表');
        }

        const [descRows] = await pool.query(`DESCRIBE ${quoteIdent(table)}`);
        const fields = descRows.map((r) => r.Field);

        const columns = {
            id: pickColumn(fields, ['id', 'Id', 'ID']),
            name: pickColumn(fields, ['name', 'Name']),
            description: pickColumn(fields, ['description', 'Description', 'desc']),
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
                'data',
            ]),
            createdAt: pickColumn(fields, ['created_at', 'Created_at', 'create_time', 'createdTime']),
        };

        if (!columns.id || !columns.name) {
            throw new Error('模板表缺少必要字段');
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

    return config;
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

function normalizeToBuffer(value) {
    if (!value) return null;
    if (Buffer.isBuffer(value)) return value;
    if (value instanceof ArrayBuffer) return Buffer.from(value);
    if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
    if (typeof value === 'string') return Buffer.from(value, 'base64');
    return null;
}

function guessMimeFromFileType(fileType) {
    const t = String(fileType ?? '').trim().toLowerCase();
    if (t.includes('xlsx')) return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if (t.includes('xls')) return 'application/vnd.ms-excel';
    if (t.includes('docx')) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    if (t.includes('doc')) return 'application/msword';
    if (t.includes('pdf')) return 'application/pdf';
    return 'application/octet-stream';
}

function appendExtIfMissing(name, fileType) {
    const base = String(name ?? 'template');
    const t = String(fileType ?? '').trim().toLowerCase();
    const m = base.match(/\.([A-Za-z0-9]+)$/);
    if (m) {
        const ext = String(m[1]).toLowerCase();
        if (ext === 'bin' && t && t !== 'bin') {
            if (t.includes('xlsx')) return base.replace(/\.[A-Za-z0-9]+$/, '.xlsx');
            if (t.includes('xls')) return base.replace(/\.[A-Za-z0-9]+$/, '.xls');
            if (t.includes('docx')) return base.replace(/\.[A-Za-z0-9]+$/, '.docx');
            if (t.includes('doc')) return base.replace(/\.[A-Za-z0-9]+$/, '.doc');
            if (t.includes('pdf')) return base.replace(/\.[A-Za-z0-9]+$/, '.pdf');
        }
        return base;
    }
    if (t.includes('xlsx')) return `${base}.xlsx`;
    if (t.includes('xls')) return `${base}.xls`;
    if (t.includes('docx')) return `${base}.docx`;
    if (t.includes('doc')) return `${base}.doc`;
    if (t.includes('pdf')) return `${base}.pdf`;
    return `${base}.bin`;
}

// ========== 1. 获取所有模板列表 ==========
router.get('/', async (req, res) => {
    try {
        const config = await resolveTemplatesConfig();
        const { table, columns } = config;

        const select = [
            `${quoteIdent(columns.id)} AS id`,
            `${quoteIdent(columns.name)} AS name`,
            columns.description ? `${quoteIdent(columns.description)} AS description` : `NULL AS description`,
            columns.fileType ? `${quoteIdent(columns.fileType)} AS file_type` : `NULL AS file_type`,
            columns.createdAt ? `${quoteIdent(columns.createdAt)} AS created_at` : `NULL AS created_at`,
        ].join(', ');

        const [rows] = await pool.query(`SELECT ${select} FROM ${quoteIdent(table)}`);
        res.json({ success: true, data: rows ?? [] });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, error: '服务器错误' });
    }
});

// ========== 2. 获取单个模板详情 ==========
router.get('/:id', async (req, res) => {
    const { id } = req.params;
    try {
        const config = await resolveTemplatesConfig();
        const { table, columns } = config;

        const select = [
            `${quoteIdent(columns.id)} AS id`,
            `${quoteIdent(columns.name)} AS name`,
            columns.description ? `${quoteIdent(columns.description)} AS description` : `NULL AS description`,
            columns.filePath ? `${quoteIdent(columns.filePath)} AS file_path` : `NULL AS file_path`,
            columns.fileType ? `${quoteIdent(columns.fileType)} AS file_type` : `NULL AS file_type`,
            columns.createdAt ? `${quoteIdent(columns.createdAt)} AS created_at` : `NULL AS created_at`,
        ].join(', ');

        const [rows] = await pool.query(
            `SELECT ${select} FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
            [id]
        );
        if (rows.length === 0) {
            return res.status(404).json({ success: false, error: '模板不存在' });
        }
        res.json({ success: true, data: rows[0] });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, error: '服务器错误' });
    }
});

// ========== 3. 下载模板文件 ==========
router.get('/:id/download', async (req, res) => {
    const { id } = req.params;
    try {
        const config = await resolveTemplatesConfig();
        const { table, columns } = config;

        const [rows] = await pool.query(
            `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
            [id]
        );
        if (rows.length === 0) {
            return res.status(404).json({ success: false, error: '模板不存在' });
        }
        const row = rows[0];

        const fileType =
            (columns.fileType ? row[columns.fileType] : null) ??
            row.file_type ??
            row.File_type ??
            row.mime_type ??
            row.content_type;
        const baseName =
            (columns.fileName ? row[columns.fileName] : null) ??
            row.file_name ??
            row.original_name ??
            row[columns.name] ??
            `template_${id}`;
        let effectiveType = fileType;
        if (!effectiveType) {
            const fromBase = path.extname(String(baseName)).replace(/^\./, '').toLowerCase();
            if (fromBase) effectiveType = fromBase;
        }
        const filename = appendExtIfMissing(baseName, effectiveType || 'xlsx');

        const fileData =
            (columns.fileBlob ? row[columns.fileBlob] : null) ??
            row.file_data ??
            row.file_blob ??
            row.file_content ??
            row.File_blob ??
            row.File_data ??
            row.File_content ??
            row.content ??
            row.blob_data ??
            row.file ??
            row.data;

        const buffer = normalizeToBuffer(fileData);
        if (buffer) {
            res.setHeader('Content-Type', guessMimeFromFileType(effectiveType));
            res.setHeader('Content-Disposition', buildContentDisposition(filename));
            res.setHeader('Content-Length', buffer.length);
            return res.status(200).send(buffer);
        }

        const filePath =
            (columns.filePath ? row[columns.filePath] : null) ?? row.file_path ?? row.File_path;
        if (!filePath) {
            return res.status(500).json({ success: false, error: '模板文件数据缺失' });
        }

        const absolutePath = path.join(__dirname, '../../', filePath);
        if (!fs.existsSync(absolutePath)) {
            return res.status(404).json({ success: false, error: '文件不存在' });
        }

        const ensured = await ensureBlobSupport(config);
        const extGuess = path.extname(String(absolutePath)).replace(/^\./, '').toLowerCase();
        const inferredType = effectiveType ?? extGuess ?? 'xlsx';
        const inferredFilename = appendExtIfMissing(baseName, inferredType);
        const diskBuffer = fs.readFileSync(absolutePath);

        await pool.query(
            `UPDATE ${quoteIdent(ensured.table)} SET ${quoteIdent(ensured.columns.fileBlob)} = ?, ${quoteIdent(
                ensured.columns.fileType
            )} = ?, ${quoteIdent(ensured.columns.fileName)} = ? WHERE ${quoteIdent(ensured.columns.id)} = ?`,
            [diskBuffer, inferredType, inferredFilename, id]
        );

        res.setHeader('Content-Type', guessMimeFromFileType(inferredType));
        res.setHeader('Content-Disposition', buildContentDisposition(inferredFilename));
        res.setHeader('Content-Length', diskBuffer.length);
        res.status(200).send(diskBuffer);
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, error: '下载失败' });
    }
});

module.exports = router;
