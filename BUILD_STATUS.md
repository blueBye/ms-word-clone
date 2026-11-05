# Cloudflare Pages deployment validation — 2026-11-05

## Changes made

- Converted the Next.js app to a Cloudflare Pages-compatible static export (`output: 'export'`, trailing slashes, unoptimized images).
- Preserved the existing pnpm workspace/engine setup. Root `build` now checks configuration, compiles the engine, builds the app, and validates the exported files.
- Added an offline static-export verification script, configuration smoke tests, local static preview, Pages deployment instructions, and a GitHub Actions CI workflow.
- Updated the root README, getting-started guide, publishing guide, `.gitignore`, and Node version pin.

## Checks actually performed in the preparation environment

| Check | Result | Notes |
| --- | --- | --- |
| `node scripts/cloudflare-smoke.mjs` | **PASS** | Verifies configuration and that an incomplete export is rejected and a complete fixture is accepted. |
| `node --check scripts/preview-static.mjs` | **PASS** | JavaScript syntax. |
| `node --check scripts/verify-cloudflare-export.mjs` | **PASS** | JavaScript syntax. |
| Local preview HTTP test (fixture site) | **PASS** | `200` on HTML and DOCX with correct MIME; `404` on missing path. |
| `npm run build` | **BLOCKED / NOT VERIFIED** | Missing installed dependencies; TypeScript reported `TS2307` for packages like `react` and missing JSX types. |
| `npm test` | **BLOCKED / NOT VERIFIED** | `vitest: not found` because dependencies are absent. |
| `npm ping` and `corepack pnpm --version` | **FAILED** | npm registry DNS lookup returned `EAI_AGAIN`, so dependency installation could not be completed here. |

**Do not treat this as a verified production build.** After installing dependencies, run `pnpm typecheck`, `pnpm test`, and `pnpm run build` locally or in GitHub Actions before publishing. The Cloudflare deployment also performs the build and static-export verification.

Cloudflare Pages expects root `/`, build `pnpm run build`, output `apps/web/out`, and **no deploy/start command**.
