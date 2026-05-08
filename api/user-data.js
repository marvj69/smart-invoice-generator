const { requireUser } = require('../server/auth');
const { query } = require('../server/db');
const { methodAllowed, readJson, sendError, sendJson } = require('../server/http');

const DATA_BODY_LIMIT = 4 * 1024 * 1024;

function asObject(value) {
  return value && typeof value === 'object' && !Array.isArray(value) ? value : {};
}

function asArray(value) {
  return Array.isArray(value) ? value : [];
}

function normalizeStoredRow(row) {
  return {
    currentInvoice: asObject(row && row.current_invoice),
    templates: asArray(row && row.templates),
    defaultCompany: asObject(row && row.default_company),
    updatedAt: row && row.updated_at ? row.updated_at : null
  };
}

module.exports = async function userData(req, res) {
  if (!methodAllowed(req, res, ['GET', 'PUT'])) return;

  try {
    const user = await requireUser(req);

    if (req.method === 'GET') {
      const rows = await query`
        SELECT current_invoice, templates, default_company, updated_at
        FROM user_invoice_data
        WHERE user_id = ${user.id}
        LIMIT 1
      `;

      return sendJson(res, 200, {
        data: normalizeStoredRow(rows[0])
      });
    }

    const body = await readJson(req, { limit: DATA_BODY_LIMIT });
    const currentInvoice = JSON.stringify(asObject(body.currentInvoice));
    const templates = JSON.stringify(asArray(body.templates).slice(0, 250));
    const defaultCompany = JSON.stringify(asObject(body.defaultCompany));

    const rows = await query`
      INSERT INTO user_invoice_data (user_id, current_invoice, templates, default_company, updated_at)
      VALUES (${user.id}, ${currentInvoice}::jsonb, ${templates}::jsonb, ${defaultCompany}::jsonb, now())
      ON CONFLICT (user_id) DO UPDATE SET
        current_invoice = EXCLUDED.current_invoice,
        templates = EXCLUDED.templates,
        default_company = EXCLUDED.default_company,
        updated_at = now()
      RETURNING current_invoice, templates, default_company, updated_at
    `;

    sendJson(res, 200, {
      data: normalizeStoredRow(rows[0])
    });
  } catch (error) {
    sendError(res, error, 'Could not sync invoice data.');
  }
};
