const http = require('http');
const fs = require('fs');
const path = require('path');

const PORT = process.env.PORT || 3000;
const ROOT_DIR = __dirname;
const GRAPH_ROOT = 'https://graph.microsoft.com/v1.0';
const WORKBOOK_NAME = process.env.MICROSOFT_GRAPH_WORKBOOK_NAME || '5p.xlsx';

class PublicError extends Error {
  constructor(message, status = 500) {
    super(message);
    this.status = status;
  }
}

function requireEnv(name) {
  const value = process.env[name];
  if (!value) {
    throw new PublicError(`Required environment variable is missing: ${name}`, 500);
  }
  return value;
}

function sendJson(res, status, payload) {
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Cache-Control': 'no-store'
  });
  res.end(JSON.stringify(payload));
}

async function getGraphToken() {
  const tenantId = requireEnv('MICROSOFT_TENANT_ID');
  const clientId = requireEnv('MICROSOFT_CLIENT_ID');
  const clientSecret = requireEnv('MICROSOFT_CLIENT_SECRET');

  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });

  const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok || !payload.access_token) {
    throw new PublicError('Microsoft Graph authentication failed.', 502);
  }

  return payload.access_token;
}

async function graphGet(token, url, notFoundMessage) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` }
  });

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    if (response.status === 404 && notFoundMessage) {
      throw new PublicError(notFoundMessage, 404);
    }

    const graphMessage = payload.error && payload.error.message ? ` ${payload.error.message}` : '';
    throw new PublicError(`Microsoft Graph request failed.${graphMessage}`, 502);
  }

  return payload;
}

async function findWorkbook(token) {
  const driveId = requireEnv('MICROSOFT_GRAPH_DRIVE_ID');
  const searchUrl = `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/root/search(q='${encodeURIComponent(WORKBOOK_NAME)}')?$select=id,name`;
  const payload = await graphGet(token, searchUrl, `${WORKBOOK_NAME} cannot be found.`);
  const workbook = Array.isArray(payload.value)
    ? payload.value.find((item) => item.name && item.name.toLowerCase() === WORKBOOK_NAME.toLowerCase())
    : null;

  if (!workbook) {
    throw new PublicError(`${WORKBOOK_NAME} cannot be found.`, 404);
  }

  return { driveId, itemId: workbook.id };
}

function parseCellValue(rangePayload, label) {
  const value = rangePayload && rangePayload.values && rangePayload.values[0] && rangePayload.values[0][0];
  const number = Number(value);

  if (value === null || value === undefined || value === '' || !Number.isFinite(number)) {
    throw new PublicError(`${label} cannot be read as a number.`, 502);
  }

  return number;
}

async function readCell(token, driveId, itemId, sheetName, address) {
  const url = `${GRAPH_ROOT}/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/workbook/worksheets/${encodeURIComponent(sheetName)}/range(address='${encodeURIComponent(address)}')`;
  const payload = await graphGet(token, url, `${sheetName} sheet cannot be found.`);
  return parseCellValue(payload, `${sheetName}!${address}`);
}

async function handleLeaderboard(req, res) {
  try {
    const token = await getGraphToken();
    const { driveId, itemId } = await findWorkbook(token);
    const [scholastic, specialists] = await Promise.all([
      readCell(token, driveId, itemId, 'Scholastic', 'F7'),
      readCell(token, driveId, itemId, 'Specialists', 'H7')
    ]);

    sendJson(res, 200, {
      section: '5P',
      scholastic,
      specialists,
      total: scholastic + specialists
    });
  } catch (error) {
    sendJson(res, error.status || 500, { error: error.message || 'Unable to load the 5P leaderboard score.' });
  }
}

function contentType(filePath) {
  if (filePath.endsWith('.html')) return 'text/html; charset=utf-8';
  if (filePath.endsWith('.css')) return 'text/css; charset=utf-8';
  if (filePath.endsWith('.js')) return 'text/javascript; charset=utf-8';
  if (filePath.endsWith('.json')) return 'application/json; charset=utf-8';
  return 'application/octet-stream';
}

function serveStatic(req, res) {
  const requestPath = decodeURIComponent(new URL(req.url, `http://${req.headers.host}`).pathname);
  const normalizedPath = path.normalize(requestPath).replace(/^([/\\])+/, '');
  const filePath = path.join(ROOT_DIR, normalizedPath || 'index.html');
  const resolvedPath = path.resolve(filePath);

  const relativePath = path.relative(ROOT_DIR, resolvedPath);
  if (relativePath.startsWith('..') || path.isAbsolute(relativePath)) {
    res.writeHead(403);
    res.end('Forbidden');
    return;
  }

  const finalPath = fs.existsSync(resolvedPath) && fs.statSync(resolvedPath).isDirectory()
    ? path.join(resolvedPath, 'index.html')
    : resolvedPath;

  fs.readFile(finalPath, (error, data) => {
    if (error) {
      fs.readFile(path.join(ROOT_DIR, 'index.html'), (indexError, indexData) => {
        if (indexError) {
          res.writeHead(404);
          res.end('Not found');
          return;
        }
        res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
        res.end(indexData);
      });
      return;
    }

    res.writeHead(200, { 'Content-Type': contentType(finalPath) });
    res.end(data);
  });
}

const server = http.createServer((req, res) => {
  const pathname = new URL(req.url, `http://${req.headers.host}`).pathname;

  if (req.method === 'GET' && pathname === '/api/leaderboard/5p') {
    handleLeaderboard(req, res);
    return;
  }

  if (req.method === 'GET' || req.method === 'HEAD') {
    serveStatic(req, res);
    return;
  }

  res.writeHead(405, { Allow: 'GET, HEAD' });
  res.end('Method not allowed');
});

server.listen(PORT, () => {
  console.log(`PathPride server listening on port ${PORT}`);
});
