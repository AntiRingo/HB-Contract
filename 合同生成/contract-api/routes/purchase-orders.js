const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const pool = require('../db');

const projectRootDir = path.resolve(__dirname, '../..');
const purchaseStorageDir = path.resolve(projectRootDir, '../数据库/采购单列表');

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

function normalizeToBuffer(value) {
    if (!value) return null;
    if (Buffer.isBuffer(value)) return value;
    if (value instanceof ArrayBuffer) return Buffer.from(value);
    if (ArrayBuffer.isView(value)) return Buffer.from(value.buffer, value.byteOffset, value.byteLength);
    if (typeof value === 'string') return Buffer.from(value, 'base64');
    return null;
}

function guessMimeFromFilePath(filePath) {
    const ext = path.extname(String(filePath ?? '')).toLowerCase();
    if (ext === '.xlsx') return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    if (ext === '.xls') return 'application/vnd.ms-excel';
    return 'application/octet-stream';
}

function buildContentDisposition(filename) {
    const s = String(filename ?? 'file');
    const ascii = s.replace(/[^\x20-\x7E]/g, '_') || 'file';
    const utf8 = encodeURIComponent(s)
        .replace(/['()]/g, (c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`)
        .replace(/\*/g, '%2A');
    return `attachment; filename="${ascii}"; filename*=UTF-8''${utf8}`;
}

function resolveAbsoluteFilePath(filePathValue) {
    const raw = String(filePathValue ?? '').trim();
    if (!raw) return null;
    if (/^https?:\/\//i.test(raw)) return null;
    if (path.isAbsolute(raw)) return raw;
    const normalized = raw.replace(/^[/\\]+/, '');
    const baseName = path.basename(normalized);
    const candidates = [
        path.resolve(purchaseStorageDir, normalized),
        path.resolve(purchaseStorageDir, baseName),
        path.resolve(projectRootDir, normalized),
        path.resolve(projectRootDir, baseName),
    ];
    for (const p of candidates) {
        if (fs.existsSync(p)) return p;
    }
    return candidates[0] ?? null;
}

let purchaseConfigPromise = null;
async function resolvePurchaseConfig() {
    if (purchaseConfigPromise) return purchaseConfigPromise;
    purchaseConfigPromise = (async () => {
        const [rows] = await pool.query('SHOW TABLES');
        const tableNames = rows.map((r) => r[Object.keys(r)[0]]).map(String);
        const table =
            tableNames.find((t) => String(t).toLowerCase() === 'purchase_order') ||
            tableNames.find((t) => t === '采购单列表') ||
            tableNames.find((t) => String(t).includes('采购')) ||
            null;
        if (!table) throw new Error('未找到采购单表');

        const [desc] = await pool.query(`DESCRIBE ${quoteIdent(table)}`);
        const fields = desc.map((r) => r.Field);
        const columns = {
            id: pickColumn(fields, ['id', 'Id', 'ID']),
            name: pickColumn(fields, ['name', 'Name', 'title', 'Title']),
            filePath: pickColumn(fields, ['file_path', 'File_path', 'filepath', 'path']),
        };
        if (!columns.id || !columns.name || !columns.filePath) {
            throw new Error('采购单表缺少 Id / Name / File_path 字段');
        }
        return { table, columns };
    })();
    return purchaseConfigPromise;
}

router.get('/', async (req, res) => {
    try {
        const config = await resolvePurchaseConfig();
        const { table, columns } = config;
        const select = [
            `${quoteIdent(columns.id)} AS id`,
            `${quoteIdent(columns.name)} AS name`,
            `${quoteIdent(columns.filePath)} AS file_path`,
        ].join(', ');
        const [rows] = await pool.query(`SELECT ${select} FROM ${quoteIdent(table)} ORDER BY ${quoteIdent(columns.id)} DESC`);
        res.json({ success: true, data: rows ?? [] });
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, error: '服务器错误' });
    }
});

router.get('/:id/download', async (req, res) => {
    const { id } = req.params;
    try {
        const config = await resolvePurchaseConfig();
        const { table, columns } = config;
        const [rows] = await pool.query(
            `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
            [id]
        );
        if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '采购单不存在' });
        const row = rows[0];

        const filePathValue = row[columns.filePath] ?? row.file_path ?? row.File_path;
        const abs = resolveAbsoluteFilePath(filePathValue);
        if (!abs) return res.status(500).json({ success: false, error: '采购单文件地址无效' });
        if (!fs.existsSync(abs)) return res.status(404).json({ success: false, error: '文件不存在' });

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

module.exports = router;
