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

async function resolveContractsConfig() {
    const [rows] = await pool.query('SHOW TABLES');
    const tableNames = rows.map((r) => r[Object.keys(r)[0]]).map(String);
    const table =
        tableNames.find((t) => t === '合同列表') ||
        tableNames.find((t) => t.toLowerCase() === 'contracts') ||
        null;
    if (!table) throw new Error('未找到“合同列表”表');

    const [desc] = await pool.query(`DESCRIBE ${quoteIdent(table)}`);
    const fields = desc.map((r) => r.Field);
    const columns = {
        id: pickColumn(fields, ['id', 'Id', 'ID']),
        name: pickColumn(fields, ['name', 'Name', '标题', '合同名称']),
        filePath: pickColumn(fields, ['file_path', 'File_path', 'filepath', 'path', '文件路径']),
        createdAt: pickColumn(fields, ['created_at', 'Created_at', 'create_time', '创建时间']),
    };
    if (!columns.name || !columns.filePath) {
        throw new Error('“合同列表”表缺少 Name 或 File_path 字段');
    }
    return { table, columns };
}

function sanitizeFilename(input) {
    const s = String(input ?? '').trim();
    const cleaned = s
        .replace(/[\\/:*?"<>|]/g, '_')
        .replace(/\s+/g, '')
        .replace(/_+/g, '_')
        .replace(/^_+|_+$/g, '');
    return cleaned || '合同';
}

function ensureDirSync(dir) {
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

router.post(
    '/save',
    express.raw({ type: 'application/octet-stream', limit: '50mb' }),
    async (req, res) => {
        try {
            const nameParam = req.query.name || req.body?.name; // 兼容意外情况
            const baseName = sanitizeFilename(nameParam);
            const fileName = baseName.endsWith('.xlsx') ? baseName : `${baseName}.xlsx`;

            const buffer = Buffer.isBuffer(req.body)
                ? req.body
                : Buffer.from([]);
            if (!buffer || buffer.length === 0) {
                return res.status(400).json({ success: false, error: '文件数据为空' });
            }

            const targetDir = path.join(__dirname, '../../templates/合同列表');
            ensureDirSync(targetDir);
            const filePathAbs = path.join(targetDir, fileName);
            fs.writeFileSync(filePathAbs, buffer);

            const relPath = path.join('templates/合同列表', fileName).replace(/\\/g, '/');

            const config = await resolveContractsConfig();
            const { table, columns } = config;
            const fields = [columns.name, columns.filePath];
            const values = [baseName, relPath];
            let sql = `INSERT INTO ${quoteIdent(table)} (${fields.map(quoteIdent).join(', ')}`;
            let placeholders = '?, ?';
            if (columns.createdAt) {
                sql += `, ${quoteIdent(columns.createdAt)}`;
                placeholders += ', NOW()';
            }
            sql += `) VALUES (${placeholders})`;
            const [result] = await pool.query(sql, values);

            res.json({
                success: true,
                data: {
                    id: result?.insertId ?? null,
                    name: baseName,
                    file_path: relPath,
                },
            });
        } catch (e) {
            console.error(e);
            res.status(500).json({ success: false, error: '保存失败' });
        }
    }
);

module.exports = router;
