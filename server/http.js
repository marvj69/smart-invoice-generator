function sendJson(res, statusCode, payload, headers = {}) {
  res.statusCode = statusCode;
  Object.entries(headers).forEach(([key, value]) => {
    res.setHeader(key, value);
  });
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.end(JSON.stringify(payload));
}

function methodAllowed(req, res, allowedMethods) {
  if (allowedMethods.includes(req.method)) {
    return true;
  }

  res.setHeader('Allow', allowedMethods.join(', '));
  sendJson(res, 405, { error: 'Method not allowed.' });
  return false;
}

function readJson(req, options = {}) {
  const limit = options.limit || 1024 * 1024;

  if (req.body !== undefined) {
    if (!req.body) {
      return Promise.resolve({});
    }
    if (Buffer.isBuffer(req.body)) {
      return parseJson(req.body.toString('utf8'));
    }
    if (typeof req.body === 'string') {
      return parseJson(req.body);
    }
    if (typeof req.body === 'object') {
      return Promise.resolve(req.body);
    }
  }

  return new Promise((resolve, reject) => {
    let raw = '';
    let size = 0;

    req.on('data', (chunk) => {
      size += chunk.length;
      if (size > limit) {
        const error = new Error('Request body is too large.');
        error.statusCode = 413;
        reject(error);
        req.destroy();
        return;
      }
      raw += chunk.toString('utf8');
    });

    req.on('end', () => {
      parseJson(raw).then(resolve).catch(reject);
    });

    req.on('error', reject);
  });
}

async function parseJson(raw) {
  const text = String(raw || '').trim();
  if (!text) {
    return {};
  }

  try {
    return JSON.parse(text);
  } catch (_error) {
    const error = new Error('Request body must be valid JSON.');
    error.statusCode = 400;
    throw error;
  }
}

function sendError(res, error, fallbackMessage = 'Something went wrong.') {
  const statusCode = Number.isInteger(error && error.statusCode) ? error.statusCode : 500;
  const canExposeMessage = statusCode < 500 || (error && error.code === 'DATABASE_NOT_CONFIGURED');
  const message = canExposeMessage ? error.message : fallbackMessage;

  if (statusCode >= 500 && !(error && error.code === 'DATABASE_NOT_CONFIGURED')) {
    console.error(error);
  }

  sendJson(res, statusCode, {
    error: message,
    code: error && error.code ? error.code : undefined
  });
}

module.exports = {
  methodAllowed,
  readJson,
  sendError,
  sendJson
};
