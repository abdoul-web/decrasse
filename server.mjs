// Serveur statique de production, sans dépendance (Render Web Service, Railway, VPS…).
// Écoute sur $PORT (fourni par l'hébergeur), sinon 5173 en local.
import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { extname, join, normalize, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = fileURLToPath(new URL('.', import.meta.url));
const PORT = Number(process.env.PORT) || 5173;
const HOST = process.env.HOST || '0.0.0.0';

const TYPES = {
  '.html': 'text/html; charset=utf-8', '.css': 'text/css; charset=utf-8', '.js': 'text/javascript; charset=utf-8',
  '.json': 'application/json; charset=utf-8', '.webmanifest': 'application/manifest+json; charset=utf-8',
  '.svg': 'image/svg+xml', '.png': 'image/png', '.ico': 'image/x-icon', '.txt': 'text/plain; charset=utf-8',
  '.xml': 'application/xml; charset=utf-8', '.woff2': 'font/woff2',
};
const SECURITY = {
  'Content-Security-Policy': "default-src 'self'; script-src 'self' https://cdn.sheetjs.com https://cdnjs.cloudflare.com; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src https://fonts.gstatic.com; img-src 'self' data: blob:; connect-src 'self'; object-src 'none'; base-uri 'self'; form-action 'none'; frame-ancestors 'none'",
  'X-Content-Type-Options': 'nosniff',
  'Referrer-Policy': 'strict-origin-when-cross-origin',
  'Permissions-Policy': 'camera=(), microphone=(), geolocation=()',
};
// Fichiers du dépôt qui ne doivent pas être servis
const PRIVATE = /^\/(\.git|tests|tools|node_modules|dist)(\/|$)|^\/(package(-lock)?\.json|yarn\.lock|render\.yaml|_headers|server\.mjs|README\.md|\.gitignore)$/;

async function send(res, status, file, extra = {}) {
  const body = await readFile(file);
  res.writeHead(status, { 'Content-Type': TYPES[extname(file)] || 'application/octet-stream', 'Content-Length': body.length, ...SECURITY, ...extra });
  res.end(body);
}

createServer(async (req, res) => {
  try {
    if (req.method !== 'GET' && req.method !== 'HEAD') { res.writeHead(405, { Allow: 'GET, HEAD' }); return res.end(); }
    let path = decodeURIComponent(new URL(req.url, 'http://x').pathname);
    if (path === '/healthz') { res.writeHead(200, { 'Content-Type': 'text/plain' }); return res.end('ok'); }
    if (PRIVATE.test(path)) return send(res, 404, join(ROOT, '404.html'));
    if (path.endsWith('/')) path += 'index.html';
    const file = normalize(join(ROOT, path));
    if (!file.startsWith(ROOT.endsWith(sep) ? ROOT : ROOT + sep)) { res.writeHead(400); return res.end(); }
    const s = await stat(file).catch(() => null);
    if (!s || !s.isFile()) return send(res, 404, join(ROOT, '404.html'));
    const cache = path.startsWith('/assets/') ? 'public, max-age=86400' : 'no-cache';
    return send(res, 200, file, { 'Cache-Control': cache });
  } catch (e) {
    console.error(e);
    res.writeHead(500); res.end('Erreur serveur');
  }
}).listen(PORT, HOST, () => console.log(`Décrasse en ligne sur http://${HOST}:${PORT}`));
