const express = require('express');
const router = express.Router();
const pool = require('../db');

function quoteIdent(ident) {
    return `\`${String(ident).replace(/`/g, '``')}\``;
}

router.get('/', async (req, res) => {
    const templateId = String(req.query.template_id ?? '').trim();
    const purchaseTypeId = String(req.query.purchase_type_id ?? '').trim();
    if (!templateId || !purchaseTypeId) {
        return res.status(400).json({ success: false, error: '缺少 template_id 或 purchase_type_id' });
    }

    try {
        const [configs] = await pool.query(
            `SELECT * FROM ${quoteIdent('mapping_config')} WHERE ${quoteIdent('Template_id')} = ? AND ${quoteIdent(
                'Purchase_id'
            )} = ? ORDER BY ${quoteIdent('ID')} DESC LIMIT 1`,
            [templateId, purchaseTypeId]
        );
        if (!configs || configs.length === 0) {
            return res.status(404).json({ success: false, error: '未找到映射配置' });
        }
        const config = configs[0];
        const configId = config.ID ?? config.Id ?? config.id;

        const [rules] = await pool.query(
            `SELECT * FROM ${quoteIdent('mapping_rule')} WHERE ${quoteIdent('mapping_id')} = ? OR ${quoteIdent(
                'config_ID'
            )} = ? ORDER BY ${quoteIdent('ID')} ASC`,
            [configId, configId]
        );

        const data = {
            config: {
                id: configId,
                config_name: config.config_name ?? null,
                contract_headers: config.contract_headers ?? null,
                purchase_headers: config.purchase_headers ?? null,
                Template_id: config.Template_id ?? null,
                Purchase_id: config.Purchase_id ?? null,
            },
            rules: (rules ?? []).map((r) => ({
                id: r.ID ?? r.Id ?? r.id ?? null,
                contract_header: r.contract_header ?? null,
                purchase_header: r.purchase_header ?? null,
            })),
        };

        res.json({ success: true, data });
    } catch (e) {
        console.error(e);
        res.status(500).json({ success: false, error: '服务器错误' });
    }
});

module.exports = router;
