const { destroySession } = require('../../server/auth');
const { methodAllowed, sendError, sendJson } = require('../../server/http');

module.exports = async function logout(req, res) {
  if (!methodAllowed(req, res, ['POST'])) return;

  try {
    await destroySession(req, res);
    sendJson(res, 200, { ok: true });
  } catch (error) {
    sendError(res, error, 'Could not sign out.');
  }
};
