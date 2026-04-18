const express = require('express');
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

let mappingsSchemaPromise = null;
async function ensureMappingsTables() {
  if (mappingsSchemaPromise) return mappingsSchemaPromise;
  mappingsSchemaPromise = (async () => {
    const configTable = 'mapping_config';
    const ruleTable = 'mapping_rule';

    await pool.query(
      `CREATE TABLE IF NOT EXISTS ${quoteIdent(configTable)} (
        ${quoteIdent('ID')} INT NOT NULL AUTO_INCREMENT,
        ${quoteIdent('config_name')} VARCHAR(255) NULL,
        ${quoteIdent('Template_id')} INT NULL,
        ${quoteIdent('Purchase_id')} INT NULL,
        ${quoteIdent('contract_headers')} JSON NULL,
        ${quoteIdent('purchase_headers')} JSON NULL,
        ${quoteIdent('created_at')} TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP,
        PRIMARY KEY (${quoteIdent('ID')})
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`
    );

    await pool.query(
      `CREATE TABLE IF NOT EXISTS ${quoteIdent(ruleTable)} (
        ${quoteIdent('ID')} INT NOT NULL AUTO_INCREMENT,
        ${quoteIdent('mapping_id')} INT NOT NULL,
        ${quoteIdent('contract_header')} VARCHAR(255) NOT NULL,
        ${quoteIdent('purchase_header')} VARCHAR(255) NOT NULL,
        PRIMARY KEY (${quoteIdent('ID')}),
        INDEX ${quoteIdent('idx_mapping_rule_mapping_id')} (${quoteIdent('mapping_id')})
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4`
    );

    const [descConfig] = await pool.query(`DESCRIBE ${quoteIdent(configTable)}`);
    const configFields = (descConfig ?? []).map((r) => r.Field);
    const configCols = {
      id: pickColumn(configFields, ['ID', 'Id', 'id']),
      name: pickColumn(configFields, ['config_name', 'Name', 'name']),
      templateId: pickColumn(configFields, ['Template_id', 'template_id']),
      purchaseId: pickColumn(configFields, ['Purchase_id', 'purchase_id']),
      contractHeaders: pickColumn(configFields, ['contract_headers', 'Template_headers_json', 'template_headers_json']),
      purchaseHeaders: pickColumn(configFields, ['purchase_headers', 'Purchase_headers_json', 'purchase_headers_json']),
      createdAt: pickColumn(configFields, ['created_at', 'Created_at', 'createdAt'])
    };

    const ensureConfigCol = async (col, def) => {
      const [rows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(configTable)} LIKE ?`, [col]);
      if (rows.length === 0) {
        await pool.query(`ALTER TABLE ${quoteIdent(configTable)} ADD COLUMN ${quoteIdent(col)} ${def}`);
      } else if (def) {
        const type = String(rows[0]?.Type ?? '').toLowerCase();
        const wantsVarchar = String(def).toLowerCase().includes('varchar');
        if (wantsVarchar && !type.includes('char') && !type.includes('text')) {
          await pool.query(`ALTER TABLE ${quoteIdent(configTable)} MODIFY COLUMN ${quoteIdent(col)} ${def}`);
        }
      }
    };

    if (!configCols.name) {
      await ensureConfigCol('config_name', 'VARCHAR(255) NULL');
      configCols.name = 'config_name';
    }
    if (!configCols.templateId) {
      await ensureConfigCol('Template_id', 'INT NULL');
      configCols.templateId = 'Template_id';
    }
    if (!configCols.purchaseId) {
      await ensureConfigCol('Purchase_id', 'VARCHAR(10) NULL');
      configCols.purchaseId = 'Purchase_id';
    } else {
      await ensureConfigCol(configCols.purchaseId, 'VARCHAR(10) NULL');
    }
    if (!configCols.contractHeaders) {
      await ensureConfigCol('contract_headers', 'LONGTEXT NULL');
      configCols.contractHeaders = 'contract_headers';
    }
    if (!configCols.purchaseHeaders) {
      await ensureConfigCol('purchase_headers', 'LONGTEXT NULL');
      configCols.purchaseHeaders = 'purchase_headers';
    }
    if (!configCols.createdAt) {
      await ensureConfigCol('created_at', 'TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP');
      configCols.createdAt = 'created_at';
    }

    const [descRule] = await pool.query(`DESCRIBE ${quoteIdent(ruleTable)}`);
    const ruleFields = (descRule ?? []).map((r) => r.Field);
    const ruleCols = {
      id: pickColumn(ruleFields, ['ID', 'Id', 'id']),
      mappingId: pickColumn(ruleFields, ['mapping_id', 'Mapping_id']),
      contractHeader: pickColumn(ruleFields, ['contract_header', 'Template_header', 'template_header']),
      purchaseHeader: pickColumn(ruleFields, ['purchase_header', 'Purchase_header', 'purchase_header']),
      createdAt: pickColumn(ruleFields, ['created_at', 'Created_at', 'createdAt'])
    };

    const ensureRuleCol = async (col, def) => {
      const [rows] = await pool.query(`SHOW COLUMNS FROM ${quoteIdent(ruleTable)} LIKE ?`, [col]);
      if (rows.length === 0) await pool.query(`ALTER TABLE ${quoteIdent(ruleTable)} ADD COLUMN ${quoteIdent(col)} ${def}`);
    };

    if (!ruleCols.mappingId) {
      await ensureRuleCol('mapping_id', 'INT NOT NULL');
      ruleCols.mappingId = 'mapping_id';
    }
    if (!ruleCols.contractHeader) {
      await ensureRuleCol('contract_header', 'VARCHAR(255) NOT NULL');
      ruleCols.contractHeader = 'contract_header';
    }
    if (!ruleCols.purchaseHeader) {
      await ensureRuleCol('purchase_header', 'VARCHAR(255) NOT NULL');
      ruleCols.purchaseHeader = 'purchase_header';
    }
    if (!ruleCols.createdAt) {
      await ensureRuleCol('created_at', 'TIMESTAMP NULL DEFAULT CURRENT_TIMESTAMP');
      ruleCols.createdAt = 'created_at';
    }

    return { configTable, ruleTable, configCols, ruleCols };
  })();
  return mappingsSchemaPromise;
}

function normalizeHeaders(value) {
  if (!Array.isArray(value)) return [];
  const out = [];
  const seen = new Set();
  for (const v of value) {
    const s = String(v ?? '').trim();
    if (!s) continue;
    if (seen.has(s)) continue;
    seen.add(s);
    out.push(s);
  }
  return out;
}

function normalizeRules(value) {
  if (!Array.isArray(value)) return [];
  const out = [];
  for (const r of value) {
    const templateHeader = String(r?.templateHeader ?? '').trim();
    const purchaseHeader = String(r?.purchaseHeader ?? '').trim();
    if (!templateHeader || !purchaseHeader) continue;
    out.push({ templateHeader, purchaseHeader });
  }
  return out;
}

router.get('/', async (req, res) => {
  try {
    const schema = await ensureMappingsTables();
    const [rows] = await pool.query(
      `SELECT ${quoteIdent(schema.configCols.id)} AS id,
              ${quoteIdent(schema.configCols.name)} AS name,
              ${quoteIdent(schema.configCols.templateId)} AS template_id,
              ${quoteIdent(schema.configCols.purchaseId)} AS purchase_id,
              ${quoteIdent(schema.configCols.createdAt)} AS created_at
       FROM ${quoteIdent(schema.configTable)}
       ORDER BY ${quoteIdent(schema.configCols.id)} ASC`
    );

    const ids = (rows ?? []).map((r) => r.id).filter(Boolean);
    if (ids.length === 0) return res.json({ success: true, data: [] });

    const [ruleRows] = await pool.query(
      `SELECT ${quoteIdent(schema.ruleCols.mappingId)} AS mapping_id,
              COUNT(1) AS rule_count
       FROM ${quoteIdent(schema.ruleTable)}
       WHERE ${quoteIdent(schema.ruleCols.mappingId)} IN (${ids.map(() => '?').join(', ')})
       GROUP BY ${quoteIdent(schema.ruleCols.mappingId)}`,
      ids
    );
    const byId = new Map((ruleRows ?? []).map((r) => [Number(r.mapping_id), Number(r.rule_count)]));

    res.json({
      success: true,
      data: (rows ?? []).map((r) => ({
        id: r.id,
        name: r.name,
        template_id: r.template_id,
        purchase_id: r.purchase_id,
        created_at: r.created_at,
        rule_count: byId.get(Number(r.id)) ?? 0
      }))
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '服务器错误' });
  }
});

router.get('/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const schema = await ensureMappingsTables();
    const [rows] = await pool.query(
      `SELECT ${quoteIdent(schema.configCols.id)} AS id,
              ${quoteIdent(schema.configCols.name)} AS name,
              ${quoteIdent(schema.configCols.templateId)} AS template_id,
              ${quoteIdent(schema.configCols.purchaseId)} AS purchase_id,
              ${quoteIdent(schema.configCols.contractHeaders)} AS contract_headers,
              ${quoteIdent(schema.configCols.purchaseHeaders)} AS purchase_headers,
              ${quoteIdent(schema.configCols.createdAt)} AS created_at
       FROM ${quoteIdent(schema.configTable)}
       WHERE ${quoteIdent(schema.configCols.id)} = ?`,
      [id]
    );
    if (!rows || rows.length === 0) return res.status(404).json({ success: false, error: '映射不存在' });
    const row = rows[0];

    const [ruleRows] = await pool.query(
      `SELECT ${quoteIdent(schema.ruleCols.id)} AS id,
              ${quoteIdent(schema.ruleCols.contractHeader)} AS contract_header,
              ${quoteIdent(schema.ruleCols.purchaseHeader)} AS purchase_header
       FROM ${quoteIdent(schema.ruleTable)}
       WHERE ${quoteIdent(schema.ruleCols.mappingId)} = ?
       ORDER BY ${quoteIdent(schema.ruleCols.id)} ASC`,
      [id]
    );

    const parseJson = (s) => {
      try {
        return JSON.parse(String(s ?? '[]'));
      } catch {
        return [];
      }
    };

    res.json({
      success: true,
      data: {
        id: row.id,
        name: row.name,
        template_id: row.template_id,
        purchase_id: row.purchase_id,
        template_headers: parseJson(row.contract_headers),
        purchase_headers: parseJson(row.purchase_headers),
        created_at: row.created_at,
        rules: (ruleRows ?? []).map((r) => ({
          id: r.id,
          templateHeader: r.contract_header,
          purchaseHeader: r.purchase_header
        }))
      }
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ success: false, error: '服务器错误' });
  }
});

router.post('/', async (req, res) => {
  try {
    const schema = await ensureMappingsTables();

    const name = String(req.body?.name ?? '').trim();
    if (!name) return res.status(400).json({ success: false, error: '映射名称不能为空' });

    const templateIdRaw = req.body?.templateId ?? req.body?.template_id;
    const purchaseTypeIdRaw = req.body?.purchaseTypeId ?? req.body?.purchase_type_id ?? req.body?.purchaseId ?? req.body?.purchase_id;
    const templateId = templateIdRaw != null && String(templateIdRaw).trim() !== '' ? Number(templateIdRaw) : null;
    const purchaseTypeId =
      purchaseTypeIdRaw != null && String(purchaseTypeIdRaw).trim() !== '' ? String(purchaseTypeIdRaw).trim() : null;
    if (templateId == null || Number.isNaN(templateId)) {
      return res.status(400).json({ success: false, error: '合同模板 template_id 不能为空' });
    }
    if (!purchaseTypeId) return res.status(400).json({ success: false, error: '订购单 type_ID 不能为空' });

    const templateHeaders = normalizeHeaders(req.body?.templateHeaders ?? req.body?.template_headers);
    const purchaseHeaders = normalizeHeaders(req.body?.purchaseHeaders ?? req.body?.purchase_headers);
    if (templateHeaders.length === 0) return res.status(400).json({ success: false, error: '合同模板标题行不能为空' });
    if (purchaseHeaders.length === 0) return res.status(400).json({ success: false, error: '采购单标题行不能为空' });

    const rules = normalizeRules(req.body?.rules);
    if (rules.length === 0) return res.status(400).json({ success: false, error: '请至少配置一条映射规则' });

    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();
      const [exists] = await conn.query(
        `SELECT ${quoteIdent(schema.configCols.id)} AS id
         FROM ${quoteIdent(schema.configTable)}
         WHERE ${quoteIdent(schema.configCols.purchaseId)} = ?
           AND ${quoteIdent(schema.configCols.templateId)} = ?
         LIMIT 1
         FOR UPDATE`,
        [purchaseTypeId, templateId]
      );
      if (exists && exists.length > 0) {
        await conn.rollback();
        res.status(409).json({
          success: false,
          error: '该 type_ID 已有映射关系，请删除后再建立映射关系',
          data: { existingId: exists[0].id }
        });
        return;
      }

      const [result] = await conn.query(
        `INSERT INTO ${quoteIdent(schema.configTable)} (${quoteIdent(schema.configCols.name)}, ${quoteIdent(
          schema.configCols.templateId
        )}, ${quoteIdent(schema.configCols.purchaseId)}, ${quoteIdent(schema.configCols.contractHeaders)}, ${quoteIdent(
          schema.configCols.purchaseHeaders
        )}) VALUES (?, ?, ?, ?, ?)`,
        [name, templateId, purchaseTypeId, JSON.stringify(templateHeaders), JSON.stringify(purchaseHeaders)]
      );
      const mappingId = result.insertId;

      const placeholders = rules.map(() => '(?, ?, ?, ?)').join(', ');
      const params = [];
      for (const r of rules) {
        params.push(mappingId, r.templateHeader, r.purchaseHeader, new Date());
      }
      await conn.query(
        `INSERT INTO ${quoteIdent(schema.ruleTable)} (${quoteIdent(schema.ruleCols.mappingId)}, ${quoteIdent(
          schema.ruleCols.contractHeader
        )}, ${quoteIdent(schema.ruleCols.purchaseHeader)}, ${quoteIdent(schema.ruleCols.createdAt)}) VALUES ${placeholders}`,
        params
      );

      await conn.commit();
      res.json({ success: true, data: { id: mappingId } });
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
    res.status(500).json({ success: false, error: '保存失败' });
  }
});

router.delete('/:id', async (req, res) => {
  const { id } = req.params;
  try {
    const schema = await ensureMappingsTables();
    const conn = await pool.getConnection();
    try {
      await conn.beginTransaction();
      await conn.query(
        `DELETE FROM ${quoteIdent(schema.ruleTable)} WHERE ${quoteIdent(schema.ruleCols.mappingId)} = ?`,
        [id]
      );
      const [result] = await conn.query(
        `DELETE FROM ${quoteIdent(schema.configTable)} WHERE ${quoteIdent(schema.configCols.id)} = ?`,
        [id]
      );
      await conn.commit();
      if (!result.affectedRows) return res.status(404).json({ success: false, error: '映射不存在' });
      res.json({ success: true, data: { id: Number(id) } });
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
