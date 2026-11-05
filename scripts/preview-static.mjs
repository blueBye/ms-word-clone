// Optional local preview of the exact files served by Cloudflare Pages.
// No external dependencies or server-side document processing.
import { createServer } from 'node:http';
import { readFile, realpath, stat } from 'node:fs/promises';
import { resolve, extname, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const output = process.env.PAGES_OUTPUT_DIR ? resolve(process.env.PAGES_OUTPUT_DIR) : fileURLToPath(new URL('../apps/web/out/', import.meta.url));
const port = Number(process.env.PORT ?? '3000');
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
  '.ico': 'image/x-icon',
  '.woff': 'font/woff',
  '.woff2': 'font/woff2',
  '.ttf': 'font/ttf',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
};

createServer(async (request, response) => {
  if (request.method !== 'GET' && request.method !== 'HEAD') {
    response.writeHead(405, { Allow: 'GET, HEAD' }).end();
    return;
  }
  let candidate;
  try {
    const path = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
    // Resolve within the exported directory; never serve arbitrary local files.
    candidate = resolve(output, `.${path}`, path.endsWith('/') ? 'index.html' : '');
    if (candidate !== output && !candidate.startsWith(output + sep)) throw Error('Invalid path');
    const info = await stat(candidate);
    if (info.isDirectory()) candidate = resolve(candidate, 'index.html');
    const [outputReal, fileReal] = await Promise.all([realpath(output), realpath(candidate)]);
    if (!fileReal.startsWith(outputReal + sep)) throw Error('Invalid path');
  } catch {
    response.writeHead(404).end('Not found');
    return;
  }
  try {
    const contents = await readFile(candidate);
    response.writeHead(200, {
      'Content-Type': types[extname(candidate)] ?? 'application/octet-stream',
      'X-Content-Type-Options': 'nosniff',
    });
    response.end(request.method === 'HEAD' ? undefined : contents);
  } catch {
    response.writeHead(404).end('Not found');
  }
}).listen(port, '127.0.0.1', () => {
  console.log(`[cloudflare] Serving apps/web/out at http://127.0.0.1:${port}/`);
});
