const { getUserFromRequest } = require('../../server/auth');
const { methodAllowed, sendError, sendJson } = require('../../server/http');

module.exports = async function me(req, res) {
  if (!methodAllowed(req, res, ['GET'])) return;

  try {
    const user = await getUserFromRequest(req);
    sendJson(res, 200, { user });
  } catch (error) {
    sendError(res, error, 'Could not load account.');
  }
};
