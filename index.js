const path = require('path');
const fs = require('fs');
const http = require('http');
const { URL } = require('url');
const { DatabaseSync } = require('node:sqlite');

const rootDir = path.resolve(__dirname, '..');
const publicDir = path.join(rootDir, 'public');
const dataDir = path.join(rootDir, 'data');
const uploadsDir = path.join(rootDir, 'uploads');
const exportsDir = path.join(rootDir, 'exports');

for (const dir of [publicDir, dataDir, uploadsDir, exportsDir]) {
  fs.mkdirSync(dir, { recursive: true });
}

const dbPath = path.join(dataDir, 'projects.sqlite');
const db = new DatabaseSync(dbPath);

db.exec(`
  CREATE TABLE IF NOT EXISTS projects (
    id TEXT PRIMARY KEY,
    payload TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  );
`);

const getProjectStmt = db.prepare('SELECT id, payload, created_at, updated_at FROM projects WHERE id = ?');
const saveProjectStmt = db.prepare(`
  INSERT INTO projects (id, payload, created_at, updated_at)
  VALUES (?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
  ON CONFLICT(id) DO UPDATE SET
    payload = excluded.payload,
    updated_at = CURRENT_TIMESTAMP
`);

const port = Number(process.env.PORT || 3000);

function normalizeProjectId(rawId) {
  return String(rawId || 'default')
    .trim()
    .replace(/[^a-zA-Z0-9_-]/g, '-')
    .slice(0, 80) || 'default';
}

function loadProject(id) {
  const row = getProjectStmt.get(normalizeProjectId(id));
  if (!row) return null;
  return {
    id: row.id,
    data: JSON.parse(row.payload),
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function saveProject(id, data) {
  const projectId = normalizeProjectId(id);
  saveProjectStmt.run(projectId, JSON.stringify(data));
  return { ok: true, id: projectId };
}

function validateProjectData(data) {
  return data && typeof data === 'object' && !Array.isArray(data);
}

function createExpressServer() {
  const express = require('express');
  const app = express();

  app.use(express.json({ limit: '50mb' }));
  app.use(express.static(publicDir));
  app.use('/uploads', express.static(uploadsDir));
  app.use('/exports', express.static(exportsDir));

  app.get('/api/health', (req, res) => {
    res.json({ ok: true, mode: 'express', database: path.relative(rootDir, dbPath) });
  });

  app.get('/api/projects/:id', (req, res) => {
    const project = loadProject(req.params.id);
    if (!project) {
      res.status(404).json({ error: 'Project not found' });
      return;
    }
    res.json(project);
  });

  app.put('/api/projects/:id', (req, res) => {
    if (!validateProjectData(req.body?.data)) {
      res.status(400).json({ error: 'Request body must include an object field named data' });
      return;
    }
    res.json(saveProject(req.params.id, req.body.data));
  });

  app.post('/api/projects', (req, res) => {
    if (!validateProjectData(req.body?.data)) {
      res.status(400).json({ error: 'Request body must include an object field named data' });
      return;
    }
    res.status(201).json(saveProject(req.body.id || 'default', req.body.data));
  });

  app.get('*', (req, res) => {
    res.sendFile(path.join(publicDir, 'index.html'));
  });

  return app;
}

function sendJson(res, status, payload) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body)
  });
  res.end(body);
}

function contentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return {
    '.html': 'text/html; charset=utf-8',
    '.css': 'text/css; charset=utf-8',
    '.js': 'application/javascript; charset=utf-8',
    '.json': 'application/json; charset=utf-8',
    '.png': 'image/png',
    '.jpg': 'image/jpeg',
    '.jpeg': 'image/jpeg',
    '.svg': 'image/svg+xml'
  }[ext] || 'application/octet-stream';
}

function sendFile(res, baseDir, requestPath) {
  const safePath = path.normalize(decodeURIComponent(requestPath)).replace(/^(\.\.[/\\])+/, '');
  const filePath = path.join(baseDir, safePath === '/' ? 'index.html' : safePath);
  const resolved = path.resolve(filePath);
  if (!resolved.startsWith(path.resolve(baseDir))) {
    sendJson(res, 403, { error: 'Forbidden' });
    return;
  }
  const target = fs.existsSync(resolved) && fs.statSync(resolved).isFile()
    ? resolved
    : path.join(publicDir, 'index.html');
  res.writeHead(200, { 'Content-Type': contentType(target) });
  fs.createReadStream(target).pipe(res);
}

function readJsonBody(req) {
  return new Promise((resolve, reject) => {
    let body = '';
    req.on('data', chunk => {
      body += chunk;
      if (body.length > 50 * 1024 * 1024) {
        reject(new Error('Request body is too large'));
        req.destroy();
      }
    });
    req.on('end', () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (error) {
        reject(error);
      }
    });
    req.on('error', reject);
  });
}

function createFallbackServer() {
  return http.createServer(async (req, res) => {
    const url = new URL(req.url, `http://${req.headers.host || '127.0.0.1'}`);
    const projectMatch = url.pathname.match(/^\/api\/projects\/([^/]+)$/);

    try {
      if (url.pathname === '/api/health' && req.method === 'GET') {
        sendJson(res, 200, { ok: true, mode: 'native-fallback', database: path.relative(rootDir, dbPath) });
        return;
      }
      if (projectMatch && req.method === 'GET') {
        const project = loadProject(projectMatch[1]);
        sendJson(res, project ? 200 : 404, project || { error: 'Project not found' });
        return;
      }
      if (projectMatch && req.method === 'PUT') {
        const body = await readJsonBody(req);
        if (!validateProjectData(body.data)) {
          sendJson(res, 400, { error: 'Request body must include an object field named data' });
          return;
        }
        sendJson(res, 200, saveProject(projectMatch[1], body.data));
        return;
      }
      if (url.pathname === '/api/projects' && req.method === 'POST') {
        const body = await readJsonBody(req);
        if (!validateProjectData(body.data)) {
          sendJson(res, 400, { error: 'Request body must include an object field named data' });
          return;
        }
        sendJson(res, 201, saveProject(body.id || 'default', body.data));
        return;
      }
      if (url.pathname.startsWith('/uploads/')) {
        sendFile(res, uploadsDir, url.pathname.replace('/uploads/', ''));
        return;
      }
      if (url.pathname.startsWith('/exports/')) {
        sendFile(res, exportsDir, url.pathname.replace('/exports/', ''));
        return;
      }
      sendFile(res, publicDir, url.pathname);
    } catch (error) {
      sendJson(res, 500, { error: error.message || 'Server error' });
    }
  });
}

try {
  createExpressServer().listen(port, () => {
    console.log(`Server Express dang chay tai http://127.0.0.1:${port}`);
  });
} catch (error) {
  if (error.code !== 'MODULE_NOT_FOUND' || !String(error.message).includes('express')) {
    throw error;
  }
  console.warn('Chua cai package express. Dang chay fallback Node thuan; hay chay npm install de dung Express.');
  createFallbackServer().listen(port, () => {
    console.log(`Server fallback dang chay tai http://127.0.0.1:${port}`);
  });
}
