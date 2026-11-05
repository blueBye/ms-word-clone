import { existsSync, readdirSync, statSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const output = process.argv[2] ? resolve(process.argv[2]) : fileURLToPath(new URL('../apps/web/out/', import.meta.url));
const files = [
  'index.html',
  '404.html',
  'sample.docx',
  'icons/fluent/bold.svg',
  'icons/fluent/rtl.svg',
];

const missing = files.filter((name) => {
  const path = resolve(output, name);
  return !existsSync(path) || !statSync(path).isFile() || statSync(path).size === 0;
});

// A page can exist without the JavaScript necessary to run the DOCX editor.
// Verify that the client bundle and styles were exported, too.
const bundleDir = join(output, '_next', 'static');
function containsExtension(folder, ext) {
  if (!existsSync(folder)) return false;
  return readdirSync(folder, { withFileTypes: true }).some((entry) => {
    const path = join(folder, entry.name);
    return entry.isDirectory() ? containsExtension(path, ext) : entry.isFile() && entry.name.endsWith(ext);
  });
}
if (!containsExtension(bundleDir, '.js')) missing.push('_next/static/**/*.js');
if (!containsExtension(bundleDir, '.css')) missing.push('_next/static/**/*.css');

if (missing.length) {
  console.error('[cloudflare] Incomplete static export in apps/web/out:');
  for (const name of missing) console.error(`  - ${name}`);
  process.exitCode = 1;
} else {
  console.log(`[cloudflare] Verified ${files.length} public pages/assets and Next.js JS/CSS bundles in apps/web/out.`);
}
