const {
  createSession,
  normalizeEmail,
  validateEmail,
  verifyPassword
} = require('../../server/auth');
const { query } = require('../../server/db');
const { methodAllowed, readJson, sendError, sendJson } = require('../../server/http');

module.exports = async function login(req, res) {
  if (!methodAllowed(req, res, ['POST'])) return;

  try {
    const body = await readJson(req, { limit: 32 * 1024 });
    const email = normalizeEmail(body.email);
    const password = String(body.password || '');

    if (!validateEmail(email) || !password) {
      return sendJson(res, 401, { error: 'Email or password is incorrect.' });
    }

    const rows = await query`
      SELECT id, email, password_hash, created_at
      FROM app_users
      WHERE email = ${email}
      LIMIT 1
    `;

    const user = rows[0];
    if (!user || !(await verifyPassword(password, user.password_hash))) {
      return sendJson(res, 401, { error: 'Email or password is incorrect.' });
    }

    await createSession(req, res, user.id);
    sendJson(res, 200, {
      user: {
        id: user.id,
        email: user.email,
        created_at: user.created_at
      }
    });
  } catch (error) {
    sendError(res, error, 'Could not sign in.');
  }
};
