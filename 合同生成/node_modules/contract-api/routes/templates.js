const express = require('express');
const router = express.Router();
const path = require('path');
const fs = require('fs');
const JSZip = require('jszip');
const pool = require('../db');

const projectRootDir = path.resolve(__dirname, '../..');
const templateStorageDir = path.resolve(projectRootDir, '../数据库/合同模板列表');

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
            tableNames.find((t) => t.toLowerCase() === 'contract_template') ||
            tableNames.find((t) => t.toLowerCase() === 'contract_templates') ||
            tableNames.find((t) => t === '合同模板列表') ||
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

function resolveAbsoluteFilePath(filePathValue) {
    const raw = String(filePathValue ?? '').trim();
    if (!raw) return null;
    if (/^https?:\/\//i.test(raw)) return null;
    if (path.isAbsolute(raw)) return raw;

    const normalized = raw.replace(/^[/\\]+/, '');
    const baseName = path.basename(normalized);
    const candidates = [
        path.resolve(templateStorageDir, normalized),
        path.resolve(templateStorageDir, baseName),
        path.resolve(projectRootDir, normalized),
        path.resolve(projectRootDir, baseName),
    ];
    for (const p of candidates) {
        if (fs.existsSync(p)) return p;
    }
    return candidates[0] ?? null;
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

        const absolutePath = resolveAbsoluteFilePath(filePath);
        if (!absolutePath) {
            return res.status(500).json({ success: false, error: '模板文件地址无效' });
        }
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

function normalizeText(value) {
    if (value === null || value === undefined) return '';
    return String(value).trim().replace(/\s+/g, ' ');
}

function escapeXml(text) {
    return String(text ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
}

function decodeXmlEntities(text) {
    return String(text ?? '')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&amp;/g, '&');
}

function extractTcText(tcXml) {
    const parts = [];
    const re = /<w:t(?:\s+[^>]*)?>([\s\S]*?)<\/w:t>/g;
    let m = null;
    while ((m = re.exec(tcXml))) {
        parts.push(decodeXmlEntities(m[1] ?? ''));
    }
    return normalizeText(parts.join(''));
}

function splitTrToTcs(trXml) {
    return trXml.match(/<w:tc[\s\S]*?<\/w:tc>/g) || [];
}

function updateTcText(tcXml, text) {
    const value = escapeXml(String(text ?? '').replace(/\r\n/g, '\n').replace(/\n/g, ' '));
    const tRe = /<w:t(\s+[^>]*)?>([\s\S]*?)<\/w:t>/g;
    const matches = [...tcXml.matchAll(tRe)];

    if (matches.length > 0) {
        // 使用一个对象来管理替换，避免 string.replace(string, string) 的 $ 符号解析问题
        let index = 0;
        return tcXml.replace(tRe, (match, attrs) => {
            const currentIdx = index++;
            const attrStr = attrs ?? '';
            const attrFixed = /xml:space=/.test(attrStr) ? attrStr : `${attrStr} xml:space="preserve"`;
            
            if (currentIdx === 0) {
                // 第一项填入实际值
                return `<w:t${attrFixed}>${value}</w:t>`;
            }
            // 后续项清空，防止内容重复
            return `<w:t${attrStr}></w:t>`;
        });
    }

    // 兜底逻辑：如果没有 w:t 但有 w:p，在第一个 w:p 内部插入内容
    if (/<w:p[\s>]/.test(tcXml)) {
        return tcXml.replace(/<w:p([\s>])/, (match, suffix) => {
            return `<w:p${suffix}><w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
        });
    }
    
    // 最终兜底：直接在单元格末尾追加内容
    return tcXml.replace(/<\/w:tc>/, () => `<w:p><w:r><w:t xml:space="preserve">${value}</w:t></w:r></w:p></w:tc>`);
}

function isBlankTr(trXml) {
    const tcs = splitTrToTcs(trXml);
    for (const tc of tcs) {
        if (extractTcText(tc)) return false;
    }
    return true;
}

function buildTableCandidate(tableXml, contractHeaders) {
    const rows = tableXml.match(/<w:tr[\s\S]*?<\/w:tr>/g) || [];
    if (rows.length === 0) return null;
    const want = contractHeaders.map((h) => normalizeText(h)).filter(Boolean);
    if (want.length === 0) return null;

    let best = null;
    for (let rIdx = 0; rIdx < rows.length; rIdx += 1) {
        const tr = rows[rIdx];
        const tcs = splitTrToTcs(tr);
        if (tcs.length === 0) continue;

        const colByHeader = new Map();
        for (let cIdx = 0; cIdx < tcs.length; cIdx += 1) {
            const txt = extractTcText(tcs[cIdx]);
            if (!txt) continue;
            colByHeader.set(txt, cIdx);
        }

        let hit = 0;
        for (const h of want) {
            if (colByHeader.has(h)) hit += 1;
        }
        if (hit === 0) continue;

        const cand = { headerRowIndex: rIdx, colByHeader, hit, rows };
        if (!best || cand.hit > best.hit) best = cand;
        if (hit === want.length) return cand;
    }
    return best;
}

function replaceTableAt(documentXml, start, end, newTableXml) {
    return documentXml.slice(0, start) + newTableXml + documentXml.slice(end);
}

async function loadTemplateById(id) {
    const config = await resolveTemplatesConfig();
    const { table, columns } = config;

    const [rows] = await pool.query(`SELECT * FROM ${quoteIdent(table)} WHERE ${quoteIdent(columns.id)} = ?`, [id]);
    if (!rows || rows.length === 0) return null;
    const row = rows[0];

    const fileType =
        (columns.fileType ? row[columns.fileType] : null) ??
        row.file_type ??
        row.File_type ??
        row.mime_type ??
        row.content_type ??
        null;
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

    let buffer = normalizeToBuffer(fileData);
    if (!buffer) {
        const filePath = (columns.filePath ? row[columns.filePath] : null) ?? row.file_path ?? row.File_path;
        const absolutePath = resolveAbsoluteFilePath(filePath);
        if (!absolutePath || !fs.existsSync(absolutePath)) return null;
        buffer = fs.readFileSync(absolutePath);
    }

    return {
        buffer,
        fileType: String(effectiveType ?? '').trim().toLowerCase(),
        baseName: String(baseName ?? '').trim(),
    };
}

router.post('/:id/merge-docx', async (req, res) => {
    const { id } = req.params;
    try {
        const outputBaseName = String(req.body?.output_base_name ?? '').trim();
        const contractHeaders = Array.isArray(req.body?.contract_headers) ? req.body.contract_headers : null;
        const rows = Array.isArray(req.body?.rows) ? req.body.rows : null;
        if (!contractHeaders || contractHeaders.length === 0) {
            return res.status(400).json({ success: false, error: '缺少 contract_headers' });
        }
        if (!rows || rows.length === 0) {
            return res.status(400).json({ success: false, error: '缺少 rows' });
        }

        const tpl = await loadTemplateById(id);
        if (!tpl || !tpl.buffer) return res.status(404).json({ success: false, error: '模板不存在或无法读取' });

        const ft = String(tpl.fileType ?? '').toLowerCase();
        if (ft.includes('doc') && !ft.includes('docx')) {
            return res.status(400).json({ success: false, error: 'DOC 格式暂不支持，请另存为 DOCX' });
        }
        if (!ft.includes('docx')) {
            return res.status(400).json({ success: false, error: '当前模板不是 DOCX 文件' });
        }

        const zip = await JSZip.loadAsync(tpl.buffer);
        const docEntry = zip.file('word/document.xml');
        if (!docEntry) return res.status(500).json({ success: false, error: 'DOCX 文件结构异常（缺少 document.xml）' });
        const documentXml = await docEntry.async('string');

        const tblRe = /<w:tbl[\s\S]*?<\/w:tbl>/g;
        let best = null;
        let bestPos = null;
        let m = null;
        while ((m = tblRe.exec(documentXml))) {
            const tableXml = m[0];
            const cand = buildTableCandidate(tableXml, contractHeaders);
            if (!cand) continue;
            const pos = { start: m.index, end: m.index + tableXml.length, tableXml };
            if (!best || cand.hit > best.hit) {
                best = cand;
                bestPos = pos;
            }
            const wantLen = contractHeaders.map((h) => normalizeText(h)).filter(Boolean).length;
            if (wantLen > 0 && cand.hit === wantLen) break;
        }

        if (!best || !bestPos) {
            return res.status(400).json({ success: false, error: '未在模板中找到匹配映射标题的表格' });
        }

        const tableXml = bestPos.tableXml;
        const trList = best.rows.slice();
        const headerRowIndex = best.headerRowIndex;
        if (headerRowIndex < 0 || headerRowIndex >= trList.length) {
            return res.status(500).json({ success: false, error: '模板表格标题行解析失败' });
        }

        const headerTcs = splitTrToTcs(trList[headerRowIndex]);
        const colIndexByHeader = new Map();
        for (let cIdx = 0; cIdx < headerTcs.length; cIdx += 1) {
            const txt = extractTcText(headerTcs[cIdx]);
            if (txt) colIndexByHeader.set(txt, cIdx);
        }

        const headerOrder = contractHeaders.map((h) => normalizeText(h));
        let matched = 0;
        for (const h of headerOrder) {
            if (h && colIndexByHeader.has(h)) matched += 1;
        }
        if (matched === 0) {
            return res.status(400).json({ success: false, error: '模板表格未找到任何可映射的列' });
        }

        const footerTextRe = /(合计|小计|总计|金额合计|总价|备注|说明|条款|结算|税|供方|需方|代表|日期|地址|电话|交货|付款)/;
        const isFooterTr = (trXml) => {
            const txt = extractTcText(trXml);
            return footerTextRe.test(txt);
        };

        const afterHeaderStart = headerRowIndex + 1;
        let insertAt = afterHeaderStart;
        let useOverwrite = false;

        // 判定插入位置：
        // 1. 如果标题行下一行就是空白，则从该行开始并覆盖
        // 2. 如果标题行下一行有内容且非汇总行，则向下扫描直到遇到空白行或汇总行
        if (afterHeaderStart < trList.length) {
            if (isBlankTr(trList[afterHeaderStart])) {
                insertAt = afterHeaderStart;
                useOverwrite = true;
            } else if (isFooterTr(trList[afterHeaderStart])) {
                insertAt = afterHeaderStart;
                useOverwrite = false;
            } else {
                // 有内容，寻找内容结束点
                let lastDataIdx = afterHeaderStart;
                for (let i = afterHeaderStart; i < trList.length; i++) {
                    if (isFooterTr(trList[i])) {
                        insertAt = i;
                        useOverwrite = false;
                        break;
                    }
                    if (isBlankTr(trList[i])) {
                        insertAt = i;
                        useOverwrite = true;
                        break;
                    }
                    lastDataIdx = i;
                    insertAt = i + 1;
                }
            }
        }

        // 选取模板行（用于克隆样式）：优先取标题行下一行，其次取标题行
        const baseRowTemplate = trList[headerRowIndex + 1] || trList[headerRowIndex];

        // 如果是覆盖模式且有多行空白，收缩多余空白行
        if (useOverwrite && insertAt < trList.length) {
            let blankCount = 0;
            for (let i = insertAt; i < trList.length; i++) {
                if (isBlankTr(trList[i])) blankCount++;
                else break;
            }
            if (blankCount > 1) {
                trList.splice(insertAt + 1, blankCount - 1);
            }
        }

        const templateRowXml = baseRowTemplate;
        const applyRowValues = (rowXml, values) => {
            const tcs = splitTrToTcs(rowXml);
            const nextTcs = tcs.slice();
            for (let i = 0; i < headerOrder.length; i += 1) {
                const h = headerOrder[i];
                if (!h) continue;
                const colIdx = colIndexByHeader.get(h);
                if (colIdx === undefined) continue;
                if (colIdx < 0 || colIdx >= nextTcs.length) continue;
                const v = Array.isArray(values) ? values[i] : '';
                nextTcs[colIdx] = updateTcText(nextTcs[colIdx], v);
            }
            return rowXml.replace(/<w:tc[\s\S]*?<\/w:tc>/g, () => nextTcs.shift() || '');
        };

        const inserts = rows.map((rowValues) => applyRowValues(templateRowXml, rowValues));

        if (useOverwrite && insertAt < trList.length) {
            trList[insertAt] = inserts[0];
            if (inserts.length > 1) {
                trList.splice(insertAt + 1, 0, ...inserts.slice(1));
            }
        } else {
            trList.splice(insertAt, 0, ...inserts);
        }

        const firstTr = best.rows[0];
        const lastTr = best.rows[best.rows.length - 1];
        const firstIdx = tableXml.indexOf(firstTr);
        const lastIdx = tableXml.lastIndexOf(lastTr);
        if (firstIdx < 0 || lastIdx < 0) {
            return res.status(500).json({ success: false, error: '模板表格结构解析失败' });
        }
        const prefix = tableXml.slice(0, firstIdx);
        const suffix = tableXml.slice(lastIdx + lastTr.length);
        const newTableXml = prefix + trList.join('') + suffix;

        const nextDocumentXml = replaceTableAt(documentXml, bestPos.start, bestPos.end, newTableXml);
        zip.file('word/document.xml', nextDocumentXml);
        const out = await zip.generateAsync({
            type: 'nodebuffer',
            compression: 'DEFLATE',
            compressionOptions: { level: 6 },
        });

        const safeBase = outputBaseName ? outputBaseName.replace(/[\\/:*?"<>|]/g, '_') : tpl.baseName.replace(/\.(docx|doc)$/i, '');
        const filename = `${safeBase || '合同'}.docx`;
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', buildContentDisposition(filename));
        res.setHeader('Content-Length', out.length);
        res.status(200).send(out);
    } catch (err) {
        console.error(err);
        res.status(500).json({ success: false, error: '生成失败' });
    }
});

module.exports = router;
