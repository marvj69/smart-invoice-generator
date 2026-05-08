const crypto = require('crypto');
const {
  createSession,
  hashPassword,
  normalizeEmail,
  validateEmail,
  validatePassword
} = require('../../server/auth');
const { query } = require('../../server/db');
const { methodAllowed, readJson, sendError, sendJson } = require('../../server/http');

module.exports = async function signup(req, res) {
  if (!methodAllowed(req, res, ['POST'])) return;

  try {
    const body = await readJson(req, { limit: 32 * 1024 });
    const email = normalizeEmail(body.email);
    const password = String(body.password || '');

    if (!validateEmail(email)) {
      return sendJson(res, 400, { error: 'Enter a valid email address.' });
    }

    if (!validatePassword(password)) {
      return sendJson(res, 400, { error: 'Password must be at least 8 characters.' });
    }

    const userId = crypto.randomUUID();
    const passwordHash = await hashPassword(password);
    const rows = await query`
      INSERT INTO app_users (id, email, password_hash)
      VALUES (${userId}, ${email}, ${passwordHash})
      RETURNING id, email, created_at
    `;

    await createSession(req, res, rows[0].id);
    sendJson(res, 201, { user: rows[0] });
  } catch (error) {
    if (error && error.code === '23505') {
      error.statusCode = 409;
      error.message = 'An account with that email already exists.';
    }
    sendError(res, error, 'Could not create account.');
  }
};
