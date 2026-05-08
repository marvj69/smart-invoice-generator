const bcrypt = require('bcryptjs');
const crypto = require('crypto');
const { query } = require('./db');

const SESSION_COOKIE_NAME = 'invoice_session';
const SESSION_MAX_AGE_SECONDS = 60 * 60 * 24 * 30;
const PASSWORD_MIN_LENGTH = 8;
const EMAIL_PATTERN = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

function normalizeEmail(value) {
  return String(value || '').trim().toLowerCase();
}

function validateEmail(email) {
  return EMAIL_PATTERN.test(email) && email.length <= 254;
}

function validatePassword(password) {
  return typeof password === 'string' && password.length >= PASSWORD_MIN_LENGTH && password.length <= 256;
}

function randomToken() {
  return crypto.randomBytes(32).toString('base64url');
}

function hashToken(token) {
  return crypto.createHash('sha256').update(token).digest('hex');
}

function parseCookies(req) {
  const header = req.headers.cookie || '';
  return header.split(';').reduce((cookies, part) => {
    const index = part.indexOf('=');
    if (index < 0) return cookies;
    const key = part.slice(0, index).trim();
    const value = part.slice(index + 1).trim();
    if (key) cookies[key] = decodeURIComponent(value);
    return cookies;
  }, {});
}

function shouldUseSecureCookie(req) {
  const host = String(req.headers.host || '').split(':')[0];
  const proto = String(req.headers['x-forwarded-proto'] || '').toLowerCase();
  if (host === 'localhost' || host === '127.0.0.1' || host === '::1') {
    return false;
  }
  return proto !== 'http';
}

function setSessionCookie(res, req, token) {
  const secure = shouldUseSecureCookie(req) ? '; Secure' : '';
  res.setHeader(
    'Set-Cookie',
    `${SESSION_COOKIE_NAME}=${encodeURIComponent(token)}; Path=/; HttpOnly; SameSite=Lax; Max-Age=${SESSION_MAX_AGE_SECONDS}${secure}`
  );
}

function clearSessionCookie(res, req) {
  const secure = shouldUseSecureCookie(req) ? '; Secure' : '';
  res.setHeader(
    'Set-Cookie',
    `${SESSION_COOKIE_NAME}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0${secure}`
  );
}

async function hashPassword(password) {
  return bcrypt.hash(password, 12);
}

async function verifyPassword(password, passwordHash) {
  return bcrypt.compare(password, passwordHash);
}

async function createSession(req, res, userId) {
  const token = randomToken();
  const tokenHash = hashToken(token);
  const sessionId = crypto.randomUUID();
  const expiresAt = new Date(Date.now() + SESSION_MAX_AGE_SECONDS * 1000).toISOString();

  await query`
    INSERT INTO app_sessions (id, user_id, token_hash, expires_at)
    VALUES (${sessionId}, ${userId}, ${tokenHash}, ${expiresAt})
  `;

  setSessionCookie(res, req, token);
}

async function destroySession(req, res) {
  const token = parseCookies(req)[SESSION_COOKIE_NAME];
  try {
    if (token) {
      await query`
        DELETE FROM app_sessions
        WHERE token_hash = ${hashToken(token)}
      `;
    }
  } finally {
    clearSessionCookie(res, req);
  }
}

async function getUserFromRequest(req) {
  const token = parseCookies(req)[SESSION_COOKIE_NAME];
  if (!token) {
    return null;
  }

  const rows = await query`
    SELECT app_users.id, app_users.email, app_users.created_at
    FROM app_sessions
    INNER JOIN app_users ON app_users.id = app_sessions.user_id
    WHERE app_sessions.token_hash = ${hashToken(token)}
      AND app_sessions.expires_at > now()
    LIMIT 1
  `;

  return rows[0] || null;
}

async function requireUser(req) {
  const user = await getUserFromRequest(req);
  if (!user) {
    const error = new Error('You need to sign in first.');
    error.statusCode = 401;
    throw error;
  }
  return user;
}

module.exports = {
  createSession,
  destroySession,
  getUserFromRequest,
  hashPassword,
  normalizeEmail,
  requireUser,
  validateEmail,
  validatePassword,
  verifyPassword
};
