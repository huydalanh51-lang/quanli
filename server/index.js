const path = require('path');
const fs = require('fs');
const http = require('http');
const https = require('https');
const crypto = require('crypto');
const { spawn } = require('child_process');
const { URL } = require('url');
const { DatabaseSync } = require('node:sqlite');

const rootDir = path.resolve(__dirname, '..');
const envPath = path.join(rootDir, '.env');

loadEnvFile(envPath);

const storageRootDir = resolveStorageDir(
  process.env.APP_STORAGE_DIR || process.env.PERSISTENT_STORAGE_DIR || process.env.RENDER_DISK_PATH,
  rootDir
);
const publicDir = path.join(rootDir, 'public');
const dataDir = path.join(storageRootDir, 'data');
const uploadsDir = path.join(storageRootDir, 'uploads');
const exportsDir = path.join(storageRootDir, 'exports');
const protectedUploadsDir = path.join(storageRootDir, 'protected_uploads');
const libraryPdfDir = path.join(protectedUploadsDir, 'pdf');
const libraryCoverDir = path.join(protectedUploadsDir, 'covers');
const legacyDataDir = path.join(rootDir, 'data');
const legacyUploadsDir = path.join(rootDir, 'uploads');
const legacyExportsDir = path.join(rootDir, 'exports');
const legacyProtectedUploadsDir = path.join(rootDir, 'protected_uploads');

process.on('uncaughtException', error => {
  console.error('UNCAUGHT_EXCEPTION', error && error.stack ? error.stack : error);
});

process.on('unhandledRejection', error => {
  console.error('UNHANDLED_REJECTION', error && error.stack ? error.stack : error);
});

for (const dir of [publicDir, dataDir, uploadsDir, exportsDir, protectedUploadsDir, libraryPdfDir, libraryCoverDir]) {
  fs.mkdirSync(dir, { recursive: true });
}

if (!isSamePath(storageRootDir, rootDir)) {
  migrateLegacyStorage(legacyDataDir, dataDir);
  migrateLegacyStorage(legacyUploadsDir, uploadsDir);
  migrateLegacyStorage(legacyExportsDir, exportsDir);
  migrateLegacyStorage(legacyProtectedUploadsDir, protectedUploadsDir);
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

db.exec(`
  CREATE TABLE IF NOT EXISTS library_documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    title TEXT NOT NULL,
    author TEXT NOT NULL DEFAULT '',
    year INTEGER,
    description TEXT NOT NULL DEFAULT '',
    category TEXT NOT NULL DEFAULT '',
    pdf_path TEXT NOT NULL,
    pdf_name TEXT NOT NULL DEFAULT '',
    cover_path TEXT,
    cover_name TEXT NOT NULL DEFAULT '',
    visible INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  );
  CREATE INDEX IF NOT EXISTS idx_library_documents_visible ON library_documents(visible);
  CREATE INDEX IF NOT EXISTS idx_library_documents_category ON library_documents(category);
  CREATE INDEX IF NOT EXISTS idx_library_documents_year ON library_documents(year);
`);

const port = Number(process.env.PORT || 3000);
const openaiModel = process.env.OPENAI_MODEL || 'gpt-4.1-mini';
const geminiModel = process.env.GEMINI_MODEL || 'gemini-2.5-flash';
const geminiFallbackModel = process.env.GEMINI_FALLBACK_MODEL || 'gemini-2.0-flash-lite';
const libraryAdminSessions = new Map();
const libraryViewTokens = new Map();
const libraryTokenTtlMs = 10 * 60 * 1000;
const adminSessionTtlMs = 8 * 60 * 60 * 1000;
const supabaseUrl = String(process.env.SUPABASE_URL || '').trim().replace(/\s+/g, '').replace(/\/+$/, '');
const supabaseServiceRoleKey = String(process.env.SUPABASE_SERVICE_ROLE_KEY || '').trim();
const supabaseBucket = String(process.env.SUPABASE_BUCKET || 'library-documents').trim();
const supabaseLibraryIndexKey = '_metadata/library_documents.json';

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

function resolveStorageDir(rawPath, fallbackDir) {
  const value = String(rawPath || '').trim();
  if (!value) return fallbackDir;
  return path.isAbsolute(value) ? path.resolve(value) : path.resolve(fallbackDir, value);
}

function isSamePath(firstPath, secondPath) {
  return path.resolve(firstPath).toLowerCase() === path.resolve(secondPath).toLowerCase();
}

function hasUsefulFiles(dirPath) {
  if (!fs.existsSync(dirPath)) return false;
  return fs.readdirSync(dirPath).some(name => name !== '.gitkeep');
}

function copyDirectoryIfMissing(sourceDir, targetDir) {
  if (!fs.existsSync(sourceDir)) return;
  fs.mkdirSync(targetDir, { recursive: true });
  for (const entry of fs.readdirSync(sourceDir, { withFileTypes: true })) {
    if (entry.name === '.gitkeep') continue;
    const sourcePath = path.join(sourceDir, entry.name);
    const targetPath = path.join(targetDir, entry.name);
    if (fs.existsSync(targetPath)) continue;
    if (entry.isDirectory()) {
      copyDirectoryIfMissing(sourcePath, targetPath);
    } else if (entry.isFile()) {
      fs.copyFileSync(sourcePath, targetPath);
    }
  }
}

function migrateLegacyStorage(sourceDir, targetDir) {
  if (isSamePath(sourceDir, targetDir) || !hasUsefulFiles(sourceDir) || hasUsefulFiles(targetDir)) return;
  copyDirectoryIfMissing(sourceDir, targetDir);
  console.log(`Da copy du lieu cu sang thu muc luu tru ben vung: ${path.relative(rootDir, targetDir) || targetDir}`);
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

function safeText(value, max = 500) {
  return String(value ?? '').trim().slice(0, max);
}

function safeYear(value) {
  const year = Number(value);
  return Number.isFinite(year) && year >= 1800 && year <= 2300 ? Math.trunc(year) : null;
}

function randomToken(bytes = 24) {
  return crypto.randomBytes(bytes).toString('base64url');
}

function timingSafeStringEqual(a, b) {
  const left = Buffer.from(String(a || ''));
  const right = Buffer.from(String(b || ''));
  if (left.length !== right.length) return false;
  return crypto.timingSafeEqual(left, right);
}

function sanitizeFileName(name, fallback) {
  const ext = path.extname(String(name || '')).toLowerCase();
  const base = path.basename(String(name || fallback), ext)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60) || fallback;
  return `${base}${ext}`;
}

function parseDataUrl(dataUrl, allowedTypes) {
  const raw = String(dataUrl || '');
  const match = raw.match(/^data:([^;,]+);base64,(.+)$/);
  if (!match) throw new Error('File upload không hợp lệ.');
  const mime = match[1].toLowerCase();
  if (allowedTypes && !allowedTypes.includes(mime)) throw new Error(`Không hỗ trợ định dạng ${mime}.`);
  return { mime, buffer: Buffer.from(match[2], 'base64') };
}

function libraryUsesSupabase() {
  return Boolean(supabaseUrl && supabaseServiceRoleKey && supabaseBucket);
}

function publicLibraryDoc(row) {
  return {
    id: row.id,
    title: row.title,
    author: row.author,
    year: row.year,
    description: row.description,
    category: row.category,
    visible: Boolean(row.visible),
    coverUrl: row.cover_path ? `/api/library/documents/${row.id}/cover` : '',
    createdAt: row.created_at,
    updatedAt: row.updated_at
  };
}

function filterLibraryRows(rows, { includeHidden = false, q = '', category = '', year = '' } = {}) {
  const needle = String(q || '').toLowerCase();
  return rows
    .filter(row => includeHidden || Boolean(row.visible))
    .filter(row => !needle || [row.title, row.author, row.year, row.category].some(value => String(value || '').toLowerCase().includes(needle)))
    .filter(row => !category || row.category === category)
    .filter(row => !year || Number(row.year) === Number(year))
    .sort((a, b) => {
      const yearDiff = Number(b.year || 0) - Number(a.year || 0);
      if (yearDiff) return yearDiff;
      const dateDiff = String(b.updated_at || '').localeCompare(String(a.updated_at || ''));
      if (dateDiff) return dateDiff;
      return Number(b.id || 0) - Number(a.id || 0);
    });
}

function localLibraryQuery(options = {}) {
  const where = [];
  const params = [];
  if (!options.includeHidden) where.push('visible = 1');
  if (options.q) {
    where.push('(LOWER(title) LIKE ? OR LOWER(author) LIKE ? OR CAST(year AS TEXT) LIKE ? OR LOWER(category) LIKE ?)');
    const needle = `%${String(options.q).toLowerCase()}%`;
    params.push(needle, needle, `%${options.q}%`, needle);
  }
  if (options.category) {
    where.push('category = ?');
    params.push(options.category);
  }
  if (options.year) {
    where.push('year = ?');
    params.push(Number(options.year));
  }
  const sql = `
    SELECT * FROM library_documents
    ${where.length ? `WHERE ${where.join(' AND ')}` : ''}
    ORDER BY COALESCE(year, 0) DESC, updated_at DESC, id DESC
  `;
  return db.prepare(sql).all(...params).map(publicLibraryDoc);
}

async function libraryQuery(options = {}) {
  if (libraryUsesSupabase()) {
    return filterLibraryRows(await loadSupabaseLibraryRows(), options).map(publicLibraryDoc);
  }
  return localLibraryQuery(options);
}

async function libraryCategories() {
  if (!libraryUsesSupabase()) {
    return db.prepare(`
      SELECT category, COUNT(*) AS count
      FROM library_documents
      WHERE category <> ''
      GROUP BY category
      ORDER BY category COLLATE NOCASE
    `).all();
  }
  const counts = new Map();
  for (const row of await loadSupabaseLibraryRows()) {
    const category = String(row.category || '').trim();
    if (category) counts.set(category, (counts.get(category) || 0) + 1);
  }
  return [...counts.entries()]
    .sort((a, b) => a[0].localeCompare(b[0], 'vi'))
    .map(([category, count]) => ({ category, count }));
}

async function libraryYears() {
  if (!libraryUsesSupabase()) {
    return db.prepare(`
      SELECT DISTINCT year
      FROM library_documents
      WHERE year IS NOT NULL
      ORDER BY year DESC
    `).all().map(row => row.year);
  }
  return [...new Set((await loadSupabaseLibraryRows()).map(row => row.year).filter(year => year !== null && year !== undefined))]
    .sort((a, b) => Number(b) - Number(a));
}

async function getLibraryDocument(id, { includeHidden = false } = {}) {
  if (libraryUsesSupabase()) {
    const row = (await loadSupabaseLibraryRows()).find(item => Number(item.id) === Number(id));
    if (!row) return null;
    if (!includeHidden && !row.visible) return null;
    return row;
  }
  const row = db.prepare('SELECT * FROM library_documents WHERE id = ?').get(Number(id));
  if (!row) return null;
  if (!includeHidden && !row.visible) return null;
  return row;
}

function supabaseObjectUrl(key, { authenticated = false } = {}) {
  const encodedBucket = encodeURIComponent(supabaseBucket);
  const encodedKey = String(key).split('/').map(part => encodeURIComponent(part)).join('/');
  const route = authenticated ? 'object/authenticated' : 'object';
  return `${supabaseUrl}/storage/v1/${route}/${encodedBucket}/${encodedKey}`;
}

function supabaseBaseHeaders(extra = {}) {
  return {
    apikey: supabaseServiceRoleKey,
    Authorization: `Bearer ${supabaseServiceRoleKey}`,
    ...extra
  };
}

async function supabaseUploadObject(key, buffer, mime) {
  const response = await fetch(supabaseObjectUrl(key), {
    method: 'POST',
    headers: supabaseBaseHeaders({
      'Content-Type': mime || 'application/octet-stream',
      'Cache-Control': '3600',
      'x-upsert': 'true'
    }),
    body: buffer
  });
  if (!response.ok) {
    throw new Error(`Khong upload duoc len Supabase Storage: ${await response.text()}`);
  }
  return key;
}

async function supabaseDownloadObject(key) {
  if (!key) return null;
  const response = await fetch(supabaseObjectUrl(key, { authenticated: true }), {
    method: 'GET',
    headers: supabaseBaseHeaders()
  });
  if (response.status === 404) return null;
  if (!response.ok) {
    const errorText = await response.text();
    if (isSupabaseMissingObject(response, errorText)) return null;
    throw new Error(`Khong doc duoc file tu Supabase Storage: ${errorText}`);
  }
  return {
    buffer: Buffer.from(await response.arrayBuffer()),
    mime: response.headers.get('content-type') || contentType(key)
  };
}

function isSupabaseMissingObject(response, errorText) {
  if (response.status === 404) return true;
  try {
    const payload = JSON.parse(errorText);
    const statusCode = Number(payload.statusCode || payload.status || 0);
    const message = String(payload.message || payload.error || '').toLowerCase();
    return statusCode === 404 && message.includes('not found');
  } catch (error) {
    return false;
  }
}

async function supabaseDeleteObjects(keys) {
  const prefixes = keys.filter(Boolean);
  if (!prefixes.length) return;
  const response = await fetch(`${supabaseUrl}/storage/v1/object/${encodeURIComponent(supabaseBucket)}`, {
    method: 'DELETE',
    headers: supabaseBaseHeaders({ 'Content-Type': 'application/json' }),
    body: JSON.stringify({ prefixes })
  });
  if (!response.ok && response.status !== 404) {
    throw new Error(`Khong xoa duoc file Supabase Storage: ${await response.text()}`);
  }
}

async function loadSupabaseLibraryRows() {
  const object = await supabaseDownloadObject(supabaseLibraryIndexKey);
  if (!object) return [];
  try {
    const payload = JSON.parse(object.buffer.toString('utf8'));
    return Array.isArray(payload.documents) ? payload.documents : [];
  } catch (error) {
    throw new Error('File metadata thu vien tren Supabase bi loi JSON.');
  }
}

async function saveSupabaseLibraryRows(rows) {
  const body = Buffer.from(JSON.stringify({ documents: rows }, null, 2), 'utf8');
  await supabaseUploadObject(supabaseLibraryIndexKey, body, 'application/json; charset=utf-8');
}

function isoNow() {
  return new Date().toISOString();
}

function extensionFromMime(mime, fallback) {
  return {
    'application/pdf': '.pdf',
    'image/png': '.png',
    'image/jpeg': '.jpg',
    'image/webp': '.webp',
    'image/svg+xml': '.svg'
  }[mime] || fallback;
}

function saveBase64Upload(dataUrl, originalName, targetDir, allowedTypes, fallbackExt) {
  const parsed = parseDataUrl(dataUrl, allowedTypes);
  const original = sanitizeFileName(originalName || `upload${fallbackExt}`, `upload${fallbackExt}`);
  const ext = path.extname(original) || fallbackExt;
  const fileName = `${Date.now()}-${randomToken(8)}${ext.toLowerCase()}`;
  const filePath = path.join(targetDir, fileName);
  fs.writeFileSync(filePath, parsed.buffer);
  return { filePath, originalName: original, mime: parsed.mime };
}

function unlinkIfInside(filePath, baseDir) {
  if (!filePath) return;
  const resolved = path.resolve(filePath);
  if (resolved.startsWith(path.resolve(baseDir)) && fs.existsSync(resolved)) {
    fs.unlinkSync(resolved);
  }
}

async function upsertLibraryDocument(payload, existing = null) {
  const title = safeText(payload.title, 220);
  if (!title) throw new Error('Tên tài liệu là bắt buộc.');
  const author = safeText(payload.author, 220);
  const year = safeYear(payload.year);
  const description = safeText(payload.description, 1000);
  const category = safeText(payload.category, 120);
  const visible = payload.visible === false ? 0 : 1;
  if (libraryUsesSupabase()) {
    return upsertSupabaseLibraryDocument({ title, author, year, description, category, visible }, payload, existing);
  }
  let pdfPath = existing?.pdf_path || '';
  let pdfName = existing?.pdf_name || '';
  let coverPath = existing?.cover_path || '';
  let coverName = existing?.cover_name || '';

  if (payload.pdfDataUrl) {
    const upload = saveBase64Upload(payload.pdfDataUrl, payload.pdfName, libraryPdfDir, ['application/pdf'], '.pdf');
    if (existing?.pdf_path) unlinkIfInside(existing.pdf_path, libraryPdfDir);
    pdfPath = upload.filePath;
    pdfName = upload.originalName;
  }
  if (!pdfPath) throw new Error('File PDF là bắt buộc.');

  if (payload.coverDataUrl) {
    const upload = saveBase64Upload(payload.coverDataUrl, payload.coverName, libraryCoverDir, ['image/png', 'image/jpeg', 'image/webp', 'image/svg+xml'], '.png');
    if (existing?.cover_path) unlinkIfInside(existing.cover_path, libraryCoverDir);
    coverPath = upload.filePath;
    coverName = upload.originalName;
  }

  if (existing) {
    db.prepare(`
      UPDATE library_documents
      SET title = ?, author = ?, year = ?, description = ?, category = ?,
          pdf_path = ?, pdf_name = ?, cover_path = ?, cover_name = ?, visible = ?,
          updated_at = CURRENT_TIMESTAMP
      WHERE id = ?
    `).run(title, author, year, description, category, pdfPath, pdfName, coverPath || null, coverName, visible, existing.id);
    return publicLibraryDoc(await getLibraryDocument(existing.id, { includeHidden: true }));
  }

  const info = db.prepare(`
    INSERT INTO library_documents
      (title, author, year, description, category, pdf_path, pdf_name, cover_path, cover_name, visible)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `).run(title, author, year, description, category, pdfPath, pdfName, coverPath || null, coverName, visible);
  return publicLibraryDoc(await getLibraryDocument(info.lastInsertRowid, { includeHidden: true }));
}

async function upsertSupabaseLibraryDocument(normalized, payload, existing = null) {
  const rows = await loadSupabaseLibraryRows();
  const now = isoNow();
  const nextId = existing
    ? Number(existing.id)
    : rows.reduce((max, row) => Math.max(max, Number(row.id) || 0), 0) + 1;
  const row = {
    ...(existing || {}),
    id: nextId,
    title: normalized.title,
    author: normalized.author,
    year: normalized.year,
    description: normalized.description,
    category: normalized.category,
    visible: normalized.visible,
    pdf_path: existing?.pdf_path || '',
    pdf_name: existing?.pdf_name || '',
    cover_path: existing?.cover_path || '',
    cover_name: existing?.cover_name || '',
    created_at: existing?.created_at || now,
    updated_at: now
  };

  if (payload.pdfDataUrl) {
    const parsed = parseDataUrl(payload.pdfDataUrl, ['application/pdf']);
    const original = sanitizeFileName(payload.pdfName || 'document.pdf', 'document.pdf');
    const ext = path.extname(original) || extensionFromMime(parsed.mime, '.pdf');
    const key = `pdf/${Date.now()}-${randomToken(8)}${ext.toLowerCase()}`;
    await supabaseUploadObject(key, parsed.buffer, parsed.mime);
    if (existing?.pdf_path) await supabaseDeleteObjects([existing.pdf_path]);
    row.pdf_path = key;
    row.pdf_name = original;
  }
  if (!row.pdf_path) throw new Error('File PDF lĂ  báº¯t buá»™c.');

  if (payload.coverDataUrl) {
    const parsed = parseDataUrl(payload.coverDataUrl, ['image/png', 'image/jpeg', 'image/webp', 'image/svg+xml']);
    const original = sanitizeFileName(payload.coverName || 'cover.png', 'cover.png');
    const ext = path.extname(original) || extensionFromMime(parsed.mime, '.png');
    const key = `covers/${Date.now()}-${randomToken(8)}${ext.toLowerCase()}`;
    await supabaseUploadObject(key, parsed.buffer, parsed.mime);
    if (existing?.cover_path) await supabaseDeleteObjects([existing.cover_path]);
    row.cover_path = key;
    row.cover_name = original;
  }

  const nextRows = rows.filter(item => Number(item.id) !== Number(nextId));
  nextRows.push(row);
  await saveSupabaseLibraryRows(nextRows);
  return publicLibraryDoc(row);
}

async function deleteLibraryDocument(id) {
  const existing = await getLibraryDocument(id, { includeHidden: true });
  if (!existing) return false;
  if (libraryUsesSupabase()) {
    const rows = (await loadSupabaseLibraryRows()).filter(row => Number(row.id) !== Number(existing.id));
    await supabaseDeleteObjects([existing.pdf_path, existing.cover_path]);
    await saveSupabaseLibraryRows(rows);
    return true;
  }
  unlinkIfInside(existing.pdf_path, libraryPdfDir);
  unlinkIfInside(existing.cover_path, libraryCoverDir);
  db.prepare('DELETE FROM library_documents WHERE id = ?').run(existing.id);
  return true;
}

async function setLibraryVisibility(id, visible) {
  const existing = await getLibraryDocument(id, { includeHidden: true });
  if (!existing) return null;
  if (libraryUsesSupabase()) {
    const rows = await loadSupabaseLibraryRows();
    const index = rows.findIndex(row => Number(row.id) === Number(existing.id));
    if (index === -1) return null;
    rows[index] = { ...rows[index], visible: visible ? 1 : 0, updated_at: isoNow() };
    await saveSupabaseLibraryRows(rows);
    return publicLibraryDoc(rows[index]);
  }
  db.prepare('UPDATE library_documents SET visible = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?')
    .run(visible ? 1 : 0, existing.id);
  return publicLibraryDoc(await getLibraryDocument(existing.id, { includeHidden: true }));
}

async function getLibraryObject(document, type) {
  const filePath = type === 'cover' ? document.cover_path : document.pdf_path;
  if (!filePath) return null;
  if (libraryUsesSupabase()) return supabaseDownloadObject(filePath);
  if (!fs.existsSync(filePath)) return null;
  return { stream: fs.createReadStream(filePath), mime: contentType(filePath) };
}

function ensureAdminConfigured() {
  const user = process.env.LIBRARY_ADMIN_USER || process.env.ADMIN_USER;
  const password = process.env.LIBRARY_ADMIN_PASSWORD || process.env.ADMIN_PASSWORD;
  if (!user || !password) {
    const error = new Error('Chưa cấu hình tài khoản quản trị. Hãy đặt LIBRARY_ADMIN_USER và LIBRARY_ADMIN_PASSWORD trong biến môi trường.');
    error.status = 503;
    throw error;
  }
  return { user, password };
}

function createAdminSession() {
  const token = randomToken(32);
  libraryAdminSessions.set(token, Date.now() + adminSessionTtlMs);
  return token;
}

function isAdminToken(token) {
  const expires = libraryAdminSessions.get(String(token || ''));
  if (!expires) return false;
  if (Date.now() > expires) {
    libraryAdminSessions.delete(String(token || ''));
    return false;
  }
  return true;
}

function adminTokenFromAuthorization(headerValue) {
  const raw = String(headerValue || '');
  return raw.startsWith('Bearer ') ? raw.slice(7) : '';
}

function requireLibraryAdmin(req, res, next) {
  const token = adminTokenFromAuthorization(req.headers.authorization);
  if (!isAdminToken(token)) {
    res.status(401).json({ error: 'Bạn cần đăng nhập quản trị.' });
    return;
  }
  next();
}

function createViewToken(docId) {
  const token = randomToken(32);
  libraryViewTokens.set(token, { docId: Number(docId), expires: Date.now() + libraryTokenTtlMs });
  return token;
}

function validateViewToken(token, docId) {
  const record = libraryViewTokens.get(String(token || ''));
  if (!record || record.docId !== Number(docId) || Date.now() > record.expires) {
    if (record) libraryViewTokens.delete(String(token || ''));
    return false;
  }
  return true;
}

function samplePdfBuffer(title, subtitle) {
  const safeTitle = String(title).replace(/[()\\]/g, '');
  const safeSubtitle = String(subtitle).replace(/[()\\]/g, '');
  const stream = [
    'BT',
    '/F1 22 Tf',
    '72 740 Td',
    `(${safeTitle}) Tj`,
    '/F1 13 Tf',
    '0 -36 Td',
    `(${safeSubtitle}) Tj`,
    '0 -28 Td',
    '(Tai lieu mau trong Thu vien so PDF.) Tj',
    'ET'
  ].join('\n');
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
    `<< /Length ${Buffer.byteLength(stream)} >>\nstream\n${stream}\nendstream`
  ];
  let pdf = '%PDF-1.4\n';
  const offsets = [0];
  objects.forEach((obj, index) => {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${index + 1} 0 obj\n${obj}\nendobj\n`;
  });
  const xrefOffset = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach(offset => {
    pdf += `${String(offset).padStart(10, '0')} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF\n`;
  return Buffer.from(pdf, 'utf8');
}

function sampleCoverSvg(title, color) {
  const escaped = String(title).replace(/[<>&"]/g, ch => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' }[ch]));
  return Buffer.from(`<svg xmlns="http://www.w3.org/2000/svg" width="420" height="560" viewBox="0 0 420 560">
<rect width="420" height="560" rx="18" fill="${color}"/>
<rect x="28" y="28" width="364" height="504" rx="14" fill="rgba(255,255,255,0.86)"/>
<text x="50%" y="210" text-anchor="middle" font-family="Arial, sans-serif" font-size="28" font-weight="700" fill="#0f172a">${escaped}</text>
<text x="50%" y="260" text-anchor="middle" font-family="Arial, sans-serif" font-size="18" fill="#334155">Thu vien so PDF</text>
<text x="50%" y="472" text-anchor="middle" font-family="Arial, sans-serif" font-size="14" fill="#64748b">Chi doc truc tuyen</text>
</svg>`, 'utf8');
}

function seedLibrarySamples() {
  const count = db.prepare('SELECT COUNT(*) AS count FROM library_documents').get().count;
  if (count > 0) return;
  const samples = [
    {
      title: 'Tài liệu mẫu quy hoạch sử dụng đất',
      author: 'Nguyễn Quang Huy',
      year: 2025,
      category: 'Quy hoạch',
      description: 'Tài liệu mẫu phục vụ kiểm tra giao diện thư viện số và trình đọc trực tuyến.',
      color: '#d9f99d'
    },
    {
      title: 'Hướng dẫn lập biểu chu chuyển đất đai',
      author: 'Phần mềm đất đai',
      year: 2026,
      category: 'Hướng dẫn',
      description: 'Mẫu tài liệu hướng dẫn thao tác với biểu chu chuyển và dữ liệu GIS.',
      color: '#bfdbfe'
    },
    {
      title: 'Quy định quản lý tài liệu số',
      author: 'Bộ phận quản trị',
      year: 2024,
      category: 'Quy định',
      description: 'Tài liệu mẫu mô phỏng nhóm văn bản quy định trong thư viện PDF.',
      color: '#fde68a'
    }
  ];
  for (const item of samples) {
    const pdfPath = path.join(libraryPdfDir, `${randomToken(8)}.pdf`);
    const coverPath = path.join(libraryCoverDir, `${randomToken(8)}.svg`);
    fs.writeFileSync(pdfPath, samplePdfBuffer(item.title, item.description));
    fs.writeFileSync(coverPath, sampleCoverSvg(item.title, item.color));
    db.prepare(`
      INSERT INTO library_documents
        (title, author, year, description, category, pdf_path, pdf_name, cover_path, cover_name, visible)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
    `).run(item.title, item.author, item.year, item.description, item.category, pdfPath, `${item.title}.pdf`, coverPath, `${item.title}.svg`);
  }
}

if (!libraryUsesSupabase()) seedLibrarySamples();

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
    res.json({
      ok: true,
      mode: 'express',
      database: path.relative(rootDir, dbPath),
      storage: path.relative(rootDir, storageRootDir) || '.',
      persistentStorage: !isSamePath(storageRootDir, rootDir),
      libraryStorage: libraryUsesSupabase() ? 'supabase' : 'local'
    });
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

  app.post('/api/library/admin/login', (req, res) => {
    try {
      const { user, password } = ensureAdminConfigured();
      const username = String(req.body?.username || '');
      const inputPassword = String(req.body?.password || '');
      if (!timingSafeStringEqual(username, user) || !timingSafeStringEqual(inputPassword, password)) {
        res.status(401).json({ error: 'Sai tài khoản hoặc mật khẩu quản trị.' });
        return;
      }
      res.json({ token: createAdminSession(), username: user, expiresIn: Math.floor(adminSessionTtlMs / 1000) });
    } catch (error) {
      res.status(error.status || 500).json({ error: error.message || 'Không đăng nhập được.' });
    }
  });

  app.get('/api/library/documents', async (req, res) => {
    const includeHidden = req.query.includeHidden === '1' && isAdminToken(String(req.headers.authorization || '').replace(/^Bearer\s+/i, ''));
    try {
      res.json({
        documents: await libraryQuery({
          includeHidden,
          q: safeText(req.query.q, 160),
          category: safeText(req.query.category, 120),
          year: safeText(req.query.year, 12)
        }),
        categories: await libraryCategories(),
        years: await libraryYears()
      });
    } catch (error) {
      res.status(500).json({ error: error.message || 'Khong doc duoc thu vien tai lieu.' });
    }
  });

  app.get('/api/library/categories', async (req, res) => {
    try {
      res.json({ categories: await libraryCategories(), years: await libraryYears() });
    } catch (error) {
      res.status(500).json({ error: error.message || 'Khong doc duoc danh muc tai lieu.' });
    }
  });

  app.post('/api/library/documents', requireLibraryAdmin, async (req, res) => {
    try {
      res.status(201).json({ document: await upsertLibraryDocument(req.body || {}) });
    } catch (error) {
      res.status(400).json({ error: error.message || 'Không lưu được tài liệu.' });
    }
  });

  app.put('/api/library/documents/:id', requireLibraryAdmin, async (req, res) => {
    try {
      const existing = await getLibraryDocument(req.params.id, { includeHidden: true });
      if (!existing) {
        res.status(404).json({ error: 'Không tìm thấy tài liệu.' });
        return;
      }
      res.json({ document: await upsertLibraryDocument(req.body || {}, existing) });
    } catch (error) {
      res.status(400).json({ error: error.message || 'Không cập nhật được tài liệu.' });
    }
  });

  app.delete('/api/library/documents/:id', requireLibraryAdmin, async (req, res) => {
    const existing = await getLibraryDocument(req.params.id, { includeHidden: true });
    if (!existing) {
      res.status(404).json({ error: 'Không tìm thấy tài liệu.' });
      return;
    }
    await deleteLibraryDocument(existing.id);
    res.json({ ok: true });
  });

  app.patch('/api/library/documents/:id/visibility', requireLibraryAdmin, async (req, res) => {
    const existing = await getLibraryDocument(req.params.id, { includeHidden: true });
    if (!existing) {
      res.status(404).json({ error: 'Không tìm thấy tài liệu.' });
      return;
    }
    db.prepare('UPDATE library_documents SET visible = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?')
      .run(req.body?.visible === false ? 0 : 1, existing.id);
    res.json({ document: await setLibraryVisibility(existing.id, req.body?.visible !== false) });
  });

  app.post('/api/library/documents/:id/view-token', async (req, res) => {
    const document = await getLibraryDocument(req.params.id);
    if (!document) {
      res.status(404).json({ error: 'Không tìm thấy tài liệu hoặc tài liệu đang bị ẩn.' });
      return;
    }
    res.json({ token: createViewToken(document.id), expiresIn: Math.floor(libraryTokenTtlMs / 1000) });
  });

  app.get('/api/library/documents/:id/cover', async (req, res) => {
    const document = await getLibraryDocument(req.params.id, { includeHidden: true });
    const object = document ? await getLibraryObject(document, 'cover') : null;
    if (!object) {
      res.status(404).end();
      return;
    }
    res.setHeader('Cache-Control', 'private, max-age=300');
    res.setHeader('Content-Type', object.mime);
    if (object.buffer) res.end(object.buffer);
    else object.stream.pipe(res);
  });

  app.get('/api/library/documents/:id/pdf', async (req, res) => {
    // Không thể chống tải/copy tuyệt đối trên web: nếu trình duyệt xem được thì người dùng vẫn có thể chụp màn hình
    // hoặc dùng công cụ ngoài. Endpoint này chỉ không lộ đường dẫn thật, yêu cầu token ngắn hạn, tắt cache và dùng viewer canvas.
    const document = await getLibraryDocument(req.params.id);
    if (!document || !validateViewToken(req.query.token, document.id)) {
      res.status(403).json({ error: 'Token xem tài liệu không hợp lệ hoặc đã hết hạn.' });
      return;
    }
    const object = await getLibraryObject(document, 'pdf');
    if (!object) {
      res.status(404).json({ error: 'File PDF không tồn tại trên server.' });
      return;
    }
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', 'inline; filename="document.pdf"');
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, private');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('X-Robots-Tag', 'noindex, nofollow');
    if (object.buffer) res.end(object.buffer);
    else object.stream.pipe(res);
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
    '.webp': 'image/webp',
    '.pdf': 'application/pdf',
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
    const libraryDocMatch = url.pathname.match(/^\/api\/library\/documents\/(\d+)$/);
    const libraryViewTokenMatch = url.pathname.match(/^\/api\/library\/documents\/(\d+)\/view-token$/);
    const libraryCoverMatch = url.pathname.match(/^\/api\/library\/documents\/(\d+)\/cover$/);
    const libraryPdfMatch = url.pathname.match(/^\/api\/library\/documents\/(\d+)\/pdf$/);
    const libraryVisibilityMatch = url.pathname.match(/^\/api\/library\/documents\/(\d+)\/visibility$/);

    try {
      if (url.pathname === '/api/health' && req.method === 'GET') {
        sendJson(res, 200, {
          ok: true,
          mode: 'native-fallback',
          database: path.relative(rootDir, dbPath),
          storage: path.relative(rootDir, storageRootDir) || '.',
          persistentStorage: !isSamePath(storageRootDir, rootDir),
          libraryStorage: libraryUsesSupabase() ? 'supabase' : 'local'
        });
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
      if (url.pathname === '/api/library/admin/login' && req.method === 'POST') {
        const body = await readJsonBody(req);
        try {
          const { user, password } = ensureAdminConfigured();
          if (!timingSafeStringEqual(body.username, user) || !timingSafeStringEqual(body.password, password)) {
            sendJson(res, 401, { error: 'Sai tài khoản hoặc mật khẩu quản trị.' });
            return;
          }
          sendJson(res, 200, { token: createAdminSession(), username: user, expiresIn: Math.floor(adminSessionTtlMs / 1000) });
        } catch (error) {
          sendJson(res, error.status || 500, { error: error.message || 'Không đăng nhập được.' });
        }
        return;
      }
      if (url.pathname === '/api/library/documents' && req.method === 'GET') {
        const includeHidden = url.searchParams.get('includeHidden') === '1' && isAdminToken(adminTokenFromAuthorization(req.headers.authorization));
        sendJson(res, 200, {
          documents: await libraryQuery({
            includeHidden,
            q: safeText(url.searchParams.get('q'), 160),
            category: safeText(url.searchParams.get('category'), 120),
            year: safeText(url.searchParams.get('year'), 12)
          }),
          categories: await libraryCategories(),
          years: await libraryYears()
        });
        return;
      }
      if (url.pathname === '/api/library/categories' && req.method === 'GET') {
        sendJson(res, 200, { categories: await libraryCategories(), years: await libraryYears() });
        return;
      }
      if (url.pathname === '/api/library/documents' && req.method === 'POST') {
        if (!isAdminToken(adminTokenFromAuthorization(req.headers.authorization))) {
          sendJson(res, 401, { error: 'Bạn cần đăng nhập quản trị.' });
          return;
        }
        try {
          sendJson(res, 201, { document: await upsertLibraryDocument(await readJsonBody(req)) });
        } catch (error) {
          sendJson(res, 400, { error: error.message || 'Không lưu được tài liệu.' });
        }
        return;
      }
      if (libraryDocMatch && req.method === 'PUT') {
        if (!isAdminToken(adminTokenFromAuthorization(req.headers.authorization))) {
          sendJson(res, 401, { error: 'Bạn cần đăng nhập quản trị.' });
          return;
        }
        const existing = await getLibraryDocument(libraryDocMatch[1], { includeHidden: true });
        if (!existing) {
          sendJson(res, 404, { error: 'Không tìm thấy tài liệu.' });
          return;
        }
        try {
          sendJson(res, 200, { document: await upsertLibraryDocument(await readJsonBody(req), existing) });
        } catch (error) {
          sendJson(res, 400, { error: error.message || 'Không cập nhật được tài liệu.' });
        }
        return;
      }
      if (libraryDocMatch && req.method === 'DELETE') {
        if (!isAdminToken(adminTokenFromAuthorization(req.headers.authorization))) {
          sendJson(res, 401, { error: 'Bạn cần đăng nhập quản trị.' });
          return;
        }
        const existing = await getLibraryDocument(libraryDocMatch[1], { includeHidden: true });
        if (!existing) {
          sendJson(res, 404, { error: 'Không tìm thấy tài liệu.' });
          return;
        }
        await deleteLibraryDocument(existing.id);
        sendJson(res, 200, { ok: true });
        return;
      }
      if (libraryVisibilityMatch && req.method === 'PATCH') {
        if (!isAdminToken(adminTokenFromAuthorization(req.headers.authorization))) {
          sendJson(res, 401, { error: 'Bạn cần đăng nhập quản trị.' });
          return;
        }
        const existing = await getLibraryDocument(libraryVisibilityMatch[1], { includeHidden: true });
        if (!existing) {
          sendJson(res, 404, { error: 'Không tìm thấy tài liệu.' });
          return;
        }
        const body = await readJsonBody(req);
        db.prepare('UPDATE library_documents SET visible = ?, updated_at = CURRENT_TIMESTAMP WHERE id = ?')
          .run(body.visible === false ? 0 : 1, existing.id);
        sendJson(res, 200, { document: await setLibraryVisibility(existing.id, body.visible !== false) });
        return;
      }
      if (libraryViewTokenMatch && req.method === 'POST') {
        const document = await getLibraryDocument(libraryViewTokenMatch[1]);
        if (!document) {
          sendJson(res, 404, { error: 'Không tìm thấy tài liệu hoặc tài liệu đang bị ẩn.' });
          return;
        }
        sendJson(res, 200, { token: createViewToken(document.id), expiresIn: Math.floor(libraryTokenTtlMs / 1000) });
        return;
      }
      if (libraryCoverMatch && req.method === 'GET') {
        const document = await getLibraryDocument(libraryCoverMatch[1], { includeHidden: true });
        const object = document ? await getLibraryObject(document, 'cover') : null;
        if (!object) {
          res.writeHead(404);
          res.end();
          return;
        }
        res.writeHead(200, {
          'Content-Type': object.mime,
          'Cache-Control': 'private, max-age=300'
        });
        if (object.buffer) res.end(object.buffer);
        else object.stream.pipe(res);
        return;
      }
      if (libraryPdfMatch && req.method === 'GET') {
        const document = await getLibraryDocument(libraryPdfMatch[1]);
        if (!document || !validateViewToken(url.searchParams.get('token'), document.id)) {
          sendJson(res, 403, { error: 'Token xem tài liệu không hợp lệ hoặc đã hết hạn.' });
          return;
        }
        const object = await getLibraryObject(document, 'pdf');
        if (!object) {
          sendJson(res, 404, { error: 'File PDF không tồn tại trên server.' });
          return;
        }
        res.writeHead(200, {
          'Content-Type': 'application/pdf',
          'Content-Disposition': 'inline; filename="document.pdf"',
          'Cache-Control': 'no-store, no-cache, must-revalidate, private',
          'Pragma': 'no-cache',
          'X-Content-Type-Options': 'nosniff',
          'X-Robots-Tag': 'noindex, nofollow'
        });
        if (object.buffer) res.end(object.buffer);
        else object.stream.pipe(res);
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
