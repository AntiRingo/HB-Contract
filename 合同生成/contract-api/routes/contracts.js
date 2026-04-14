const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const pool = require('../db');

const projectRootDir = path.resolve(__dirname, '../..');
const contractStorageDir = path.resolve(projectRootDir, '../数据库/合同列表');

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
    if (!columns.id || !columns.name || !columns.filePath) {
        throw new Error('“合同列表”表缺少 Id / Name / File_path 字段');
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
        path.resolve(contractStorageDir, normalized),
        path.resolve(contractStorageDir, baseName),
        path.resolve(projectRootDir, normalized),
        path.resolve(projectRootDir, baseName),
    ];
    for (const p of candidates) {
        if (fs.existsSync(p)) return p;
    }
    return candidates[0] ?? null;
}

router.get('/', async (req, res) => {
    try {
        const config = await resolveContractsConfig();
        const { table, columns } = config;
        const select = [
            `${quoteIdent(columns.id)} AS id`,
            `${quoteIdent(columns.name)} AS name`,
            `${quoteIdent(columns.filePath)} AS file_path`,
            columns.createdAt ? `${quoteIdent(columns.createdAt)} AS created_at` : `NULL AS created_at`,
        ].join(', ');
        const [rows] = await pool.query(
            `SELECT ${select} FROM ${quoteIdent(table)} ORDER BY ${quoteIdent(columns.id)} DESC`
        );
        res.json({ success: true, data: rows ?? [] });
    } catch (e) {
        console.error(e);
        res.status(500).json({ success: false, error: '服务器错误' });
    }
});

router.get('/:id/download', async (req, res) => {
    const { id } = req.params;
    try {
        const config = await resolveContractsConfig();
        const { table, columns } = config;
        const [rows] = await pool.query(
            `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
            [id]
        );
        if (!rows || rows.length === 0) {
            return res.status(404).json({ success: false, error: '合同不存在' });
        }
        const row = rows[0];
        const filePathValue = row[columns.filePath] ?? row.file_path ?? row.File_path;
        const abs = resolveAbsoluteFilePath(filePathValue);
        if (!abs) return res.status(500).json({ success: false, error: '合同文件地址无效' });
        if (!fs.existsSync(abs)) return res.status(404).json({ success: false, error: '文件不存在' });

        const buffer = fs.readFileSync(abs);
        const name = row[columns.name] ?? row.name ?? `contract_${id}`;
        const filename = /\.xlsx$/i.test(String(name)) ? String(name) : `${String(name)}.xlsx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', buildContentDisposition(filename));
        res.setHeader('Content-Length', buffer.length);
        res.status(200).send(buffer);
    } catch (e) {
        console.error(e);
        res.status(500).json({ success: false, error: '下载失败' });
    }
});

router.delete('/:id', async (req, res) => {
    const { id } = req.params;
    try {
        const config = await resolveContractsConfig();
        const { table, columns } = config;
        const [rows] = await pool.query(
            `SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
            [id]
        );
        if (!rows || rows.length === 0) {
            return res.status(404).json({ success: false, error: '合同不存在' });
        }
        const row = rows[0];
        const filePathValue = row[columns.filePath] ?? row.file_path ?? row.File_path;
        const abs = resolveAbsoluteFilePath(filePathValue);
        if (abs && fs.existsSync(abs)) {
            try {
                fs.unlinkSync(abs);
            } catch {}
        }
        await pool.query(
            `DELETE FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`,
            [id]
        );
        res.json({ success: true });
    } catch (e) {
        console.error(e);
        res.status(500).json({ success: false, error: '删除失败' });
    }
});

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

            const targetDir = contractStorageDir;
            ensureDirSync(targetDir);
            const filePathAbs = path.join(targetDir, fileName);
            fs.writeFileSync(filePathAbs, buffer);

            const relPath = path.join('数据库/合同列表', fileName).replace(/\\/g, '/');

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
