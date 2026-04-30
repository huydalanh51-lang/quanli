const path = require('path');
const fs = require('fs');
const http = require('http');
const https = require('https');
const { spawn } = require('child_process');
const { URL } = require('url');
const { DatabaseSync } = require('node:sqlite');

const rootDir = path.resolve(__dirname, '..');
const publicDir = path.join(rootDir, 'public');
const dataDir = path.join(rootDir, 'data');
const uploadsDir = path.join(rootDir, 'uploads');
const exportsDir = path.join(rootDir, 'exports');
const envPath = path.join(rootDir, '.env');

loadEnvFile(envPath);

process.on('uncaughtException', error => {
  console.error('UNCAUGHT_EXCEPTION', error && error.stack ? error.stack : error);
});

process.on('unhandledRejection', error => {
  console.error('UNHANDLED_REJECTION', error && error.stack ? error.stack : error);
});

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
const openaiModel = process.env.OPENAI_MODEL || 'gpt-4.1-mini';
const geminiModel = process.env.GEMINI_MODEL || 'gemini-2.5-flash';
const geminiFallbackModel = process.env.GEMINI_FALLBACK_MODEL || 'gemini-2.0-flash-lite';

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  const lines = fs.readFileSync(filePath, 'utf8').split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#') || !trimmed.includes('=')) continue;
    const index = trimmed.indexOf('=');
    const key = trimmed.slice(0, index).trim();
    let value = trimmed.slice(index + 1).trim();
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (key && process.env[key] === undefined) process.env[key] = value;
  }
}

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

function extractResponseText(payload) {
  if (typeof payload.output_text === 'string') return payload.output_text;
  const chunks = [];
  for (const item of payload.output || []) {
    for (const content of item.content || []) {
      if (typeof content.text === 'string') chunks.push(content.text);
    }
  }
  return chunks.join('\n').trim();
}

function buildAiInput(question, context) {
  return [
    {
      role: 'developer',
      content: [
        {
          type: 'input_text',
          text: [
            'Bạn là trợ lý AI cho phần mềm chu chuyển đất đai.',
            'Trả lời bằng tiếng Việt, ngắn gọn, đúng số liệu được cung cấp.',
            'Nếu dữ liệu chưa đủ, hãy nói rõ cần import hoặc nhập thêm dữ liệu nào.',
            'Ưu tiên kiểm tra: tổng hiện trạng, tổng quy hoạch, cộng tăng, cộng giảm, biến động, mã đất bất thường.',
            'Không bịa số liệu ngoài dữ liệu context.'
          ].join('\n')
        }
      ]
    },
    {
      role: 'user',
      content: [
        {
          type: 'input_text',
          text: `Câu hỏi: ${question}\n\nDữ liệu phần mềm:\n${JSON.stringify(context || {}, null, 2)}`
        }
      ]
    }
  ];
}

function postJson(url, headers, payload) {
  return new Promise((resolve, reject) => {
    const body = JSON.stringify(payload);
    const target = new URL(url);
    const req = https.request({
      method: 'POST',
      hostname: target.hostname,
      path: target.pathname + target.search,
      headers: {
        ...headers,
        'Content-Length': Buffer.byteLength(body)
      },
      timeout: 45000
    }, res => {
      let text = '';
      res.setEncoding('utf8');
      res.on('data', chunk => {
        text += chunk;
      });
      res.on('end', () => {
        let json = {};
        try {
          json = text ? JSON.parse(text) : {};
        } catch (error) {
          error.message = `OpenAI trả về dữ liệu không phải JSON: ${error.message}`;
          reject(error);
          return;
        }
        resolve({ status: res.statusCode || 0, ok: res.statusCode >= 200 && res.statusCode < 300, json });
      });
    });
    req.on('timeout', () => {
      req.destroy(new Error('Kết nối OpenAI quá thời gian chờ.'));
    });
    req.on('error', reject);
    req.write(body);
    req.end();
  });
}

async function askOpenAI(question, context) {
  const provider = process.env.GEMINI_API_KEY ? 'gemini' : 'openai';
  if (provider === 'openai' && !process.env.OPENAI_API_KEY) {
    const error = new Error('Chưa cấu hình OPENAI_API_KEY hoặc GEMINI_API_KEY trên server.');
    error.status = 503;
    throw error;
  }
  const request = {
    provider,
    model: provider === 'gemini' ? geminiModel : openaiModel,
    input: buildAiInput(question, context),
    max_output_tokens: 900
  };
  try {
    return await runOpenAIChild(request);
  } catch (error) {
    if (provider === 'gemini' && shouldRetryGemini(error) && geminiFallbackModel && geminiFallbackModel !== geminiModel) {
      return runOpenAIChild({ ...request, model: geminiFallbackModel });
    }
    throw error;
  }
}

function shouldRetryGemini(error) {
  const message = String(error && error.message || '').toLowerCase();
  return message.includes('high demand') ||
    message.includes('overloaded') ||
    message.includes('temporarily') ||
    message.includes('try again later') ||
    message.includes('503') ||
    message.includes('429');
}

function runOpenAIChild(payload) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [path.join(__dirname, 'openai-child.js')], {
      cwd: rootDir,
      env: process.env,
      stdio: ['pipe', 'pipe', 'pipe'],
      windowsHide: true
    });
    let stdout = '';
    let stderr = '';
    const timer = setTimeout(() => {
      child.kill();
      reject(new Error('AI quá thời gian phản hồi.'));
    }, 60000);
    child.stdout.setEncoding('utf8');
    child.stderr.setEncoding('utf8');
    child.stdout.on('data', chunk => {
      stdout += chunk;
    });
    child.stderr.on('data', chunk => {
      stderr += chunk;
    });
    child.on('error', error => {
      clearTimeout(timer);
      reject(error);
    });
    child.on('close', code => {
      clearTimeout(timer);
      if (code !== 0) {
        const error = new Error(stderr.trim() || `Tiến trình AI dừng với mã ${code}.`);
        error.status = 502;
        reject(error);
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        error.message = `AI trả về dữ liệu không hợp lệ: ${error.message}`;
        error.status = 502;
        reject(error);
      }
    });
    child.stdin.end(JSON.stringify(payload));
  });
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

  app.post('/api/ai', async (req, res) => {
    const question = String(req.body?.question || '').trim();
    if (!question) {
      res.status(400).json({ error: 'Thiếu câu hỏi AI.' });
      return;
    }
    if (!process.env.OPENAI_API_KEY) {
      res.status(503).json({ error: 'Chưa cấu hình OPENAI_API_KEY trên server. Hãy tạo file .env rồi chạy lại npm start.' });
      return;
    }
    try {
      res.json(await askOpenAI(question, req.body?.context || {}));
    } catch (error) {
      res.status(error.status || 500).json({ error: error.message || 'Không gọi được AI.' });
    }
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
      if (url.pathname === '/api/ai' && req.method === 'POST') {
        const body = await readJsonBody(req);
        const question = String(body.question || '').trim();
        if (!question) {
          sendJson(res, 400, { error: 'Thiếu câu hỏi AI.' });
          return;
        }
        if (!process.env.OPENAI_API_KEY) {
          sendJson(res, 503, { error: 'Chưa cấu hình OPENAI_API_KEY trên server. Hãy tạo file .env rồi chạy lại npm start.' });
          return;
        }
        sendJson(res, 200, await askOpenAI(question, body.context || {}));
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
  console.log('Chua cai package express. Dang chay fallback Node thuan; hay chay npm install de dung Express.');
  createFallbackServer().listen(port, () => {
    console.log(`Server fallback dang chay tai http://127.0.0.1:${port}`);
  });
}
