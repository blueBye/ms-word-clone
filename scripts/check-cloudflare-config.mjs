import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import config from '../apps/web/next.config.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const file = (name) => resolve(root, name);
const readJson = (name) => JSON.parse(readFileSync(file(name), 'utf8'));
const rootPackage = readJson('package.json');
const webPackage = readJson('apps/web/package.json');
const build = rootPackage.scripts?.build ?? '';

assert.equal(config.output, 'export', 'Next.js must emit a static export');
assert.equal(config.trailingSlash, true, 'Static routes should use directory index files');
assert.equal(config.images?.unoptimized, true, 'Static exports cannot use the default Next.js image optimizer');
assert.ok(config.transpilePackages?.includes(rootPackage.name), 'Keep the engine workspace transpilation');
assert.ok(build.includes('build:engine'), 'Root build must compile the engine');
assert.ok(build.includes('build:web'), 'Root build must build the Next.js app');
assert.ok(build.indexOf('build:engine') < build.indexOf('build:web'), 'Engine must build before the web app');
assert.ok(build.includes('verify:pages'), 'Build must validate exported assets');
assert.match(webPackage.scripts?.build ?? '', /next build/, 'Web build must run Next.js');
assert.equal(webPackage.dependencies?.[rootPackage.name], 'workspace:*', 'Web must consume the workspace package');
assert.ok(existsSync(file('pnpm-lock.yaml')), 'Keep the workspace lockfile at repo root');
assert.ok(existsSync(file('apps/web/public/sample.docx')), 'Bundled DOCX must be available in the export');
assert.ok(existsSync(file('.node-version')), 'Pin a supported Node major version');
assert.ok(!existsSync(file('apps/web/app/api')), 'Unexpected API routes prevent a purely static deployment');
assert.ok(!existsSync(file('apps/web/middleware.ts')), 'Unexpected middleware requires a server runtime');
assert.ok(!existsSync(file('wrangler.toml')) && !existsSync(file('wrangler.jsonc')), 'This project must use Pages, not a Worker');
assert.match(readFileSync(file('.gitignore'), 'utf8'), /^out\/$/m, 'Ignore generated static output');

const guide = readFileSync(file('DEPLOY_CLOUDFLARE_PAGES.md'), 'utf8');
assert.match(guide, /apps\/web\/out/, 'Document the monorepo output path');
assert.match(guide, /pnpm run build/, 'Document the full root build command');

console.log('[cloudflare] Static-export configuration and workspace build order verified.');
