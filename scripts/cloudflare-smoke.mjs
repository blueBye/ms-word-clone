// Dependency-free regression checks for the Cloudflare Pages output contract.
import assert from 'node:assert/strict';
import { execFileSync, spawnSync } from 'node:child_process';
import { mkdtempSync, mkdirSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const scripts = fileURLToPath(new URL('./', import.meta.url));
execFileSync(process.execPath, [join(scripts, 'check-cloudflare-config.mjs')], { stdio: 'inherit' });
const output = mkdtempSync(join(tmpdir(), 'docx-pages-test-'));
const verify = join(scripts, 'verify-cloudflare-export.mjs');
try {
  const missing = spawnSync(process.execPath, [verify, output], { encoding: 'utf8' });
  assert.notEqual(missing.status, 0, 'Empty static export must fail verification');
  assert.match(missing.stderr, /index\.html/, 'Missing root HTML should be reported');
  assert.match(missing.stderr, /sample\.docx/, 'Missing sample DOCX should be reported');
  assert.match(missing.stderr, /_next\/static/, 'Missing browser bundles should be reported');

  const required = [
    'index.html', '404.html', 'sample.docx',
    'icons/fluent/bold.svg', 'icons/fluent/rtl.svg',
    '_next/static/chunks/demo.js', '_next/static/css/demo.css',
  ];
  for (const asset of required) {
    const file = join(output, asset);
    mkdirSync(dirname(file), { recursive: true });
    writeFileSync(file, asset.endsWith('.html') ? '<!doctype html><title>test</title>' : 'test content');
  }
  const valid = spawnSync(process.execPath, [verify, output], { encoding: 'utf8' });
  assert.equal(valid.status, 0, `Complete export must be accepted: ${valid.stderr}`);

  rmSync(join(output, 'sample.docx'));
  const missingAsset = spawnSync(process.execPath, [verify, output], { encoding: 'utf8' });
  assert.notEqual(missingAsset.status, 0, 'Missing sample document must fail verification');
  console.log('[cloudflare] Static-export smoke tests passed (empty, complete, missing asset).');
} finally {
  rmSync(output, { recursive: true, force: true });
}
