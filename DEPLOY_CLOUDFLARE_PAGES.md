# Deploy the DOCX editor to Cloudflare Pages

The app is a **static Next.js export**. The DOCX parser, editor, and Word export run **in the visitor's browser**. Cloudflare Pages only hosts the generated HTML, JavaScript, CSS, icons, and example DOCX file. **No server runtime, database, Cloudflare Worker, or Wrangler deploy command is required.**

The repository is a **pnpm workspace/monorepo**: the reusable engine lives in `src/`, and the Next.js app lives in `apps/web/`. Unlike a single-app Next.js repository, the output directory is **`apps/web/out`**, not `out`.

Official Cloudflare guide: https://developers.cloudflare.com/pages/framework-guides/nextjs/deploy-a-static-nextjs-site/

## 1. Verify locally

Install Node.js 22 and pnpm 10 (Cloudflare Pages v3 currently offers pnpm 10), then run from the **repository root**:

```powershell
pnpm install --frozen-lockfile
pnpm typecheck
pnpm test
pnpm run build
pnpm start
```

The build command performs these steps in order:

1. Check Cloudflare configuration and the workspace build order.
2. Compile the TypeScript DOCX engine into `dist/`.
3. Run `next build` in `apps/web/` with `output: 'export'`.
4. Verify exported HTML, example DOCX, icons, and Next.js JS/CSS bundles in `apps/web/out/`.

`pnpm start` is an **optional local static-file preview** of `apps/web/out`, at `http://127.0.0.1:3000`. Cloudflare Pages does **not** run this command. For development with hot reload, use `pnpm dev` instead.

## 2. Push the source to GitHub

This repository should contain **all of** `src/`, `apps/web/`, `package.json`, `pnpm-lock.yaml`, `pnpm-workspace.yaml`, and `scripts/`. Do **not** set the repository root to `apps/web/` or push only that folder: the app depends on the root engine workspace.

If you already have Git configured, commit and push these changes:

```powershell
git add .
git commit -m "Configure Cloudflare Pages static deployment"
git push origin main
```

The `.gitignore` excludes generated `dist/`, `out/`, `.next/`, `node_modules/`, and local environment files. Cloudflare will create these generated outputs during its build.

## 3. Create a Cloudflare **Pages** project

1. Open the Cloudflare dashboard → **Workers & Pages** → **Create application**.
2. Choose **Pages** (not Workers), then **Import an existing Git repository**.
3. Select `blueBye/ms-word-clone` and branch `main`.
4. In build settings use the values below, then select **Save and Deploy**.

| Setting | Value |
| --- | --- |
| Framework preset | **Next.js (Static HTML Export)** |
| Production branch | `main` |
| Root directory | `/` (repository root; leave blank if already the default) |
| Build command | **`pnpm run build`** |
| Build output directory | **`apps/web/out`** |
| Node.js version | `22` (also specified by `.node-version`) |
| pnpm version (if an override is needed) | `10.11.1` (`PNPM_VERSION=10.11.1`) |
| Start command | **None** |
| Deploy command / Wrangler | **None** |

**Important:** Cloudflare's default Next.js static-export preset assumes the app is at the repository root and suggests `out`. Override both the **build command** and **output directory** with the monorepo values above. Running `next build` directly would **skip the engine build** and may fail to resolve `@local/docx-ir-engine`.

If the Cloudflare form asks you for a **deploy command** such as `npx wrangler deploy`, you have entered the **Workers** setup flow. Go back and choose **Pages → Import an existing Git repository**. If a framework preset alters your chosen command, set it back to `pnpm run build` before saving.

No application environment variables are required for the current browser-only editor. Cloudflare Pages v3 already supplies a suitable Node.js and pnpm version, but you can set these in **Settings → Environment variables** if its selected build image differs:

```text
NODE_VERSION=22
PNPM_VERSION=10.11.1
```

## 4. Verify the deployed site

The deployment will have an address similar to `https://<pages-project>.pages.dev/`. Check that:

- The editor opens instead of a 404 or empty page.
- The bundled sample document loads (`/sample.docx` returns a file, not a 404).
- **File → Open** loads a local DOCX, and editing/formatting works.
- **File → Download** produces a `.docx` that opens in Microsoft Word.
- Ribbon icons and styling load correctly, including on a hard refresh.

The browser performs parsing and export; the app does not upload documents to Cloudflare by default. However, users should still avoid using untrusted or sensitive files until they are comfortable with the project's experimental compatibility and security status.

## 5. Connect a domain (optional)

After the Pages deployment works, open **Pages project → Custom domains → Set up a custom domain** and follow Cloudflare's DNS instructions. You can leave a friend's domain in their own Cloudflare account and point it to the Pages deployment using the DNS guidance for that setup.

## Troubleshooting

| Symptom | Likely cause / fix |
| --- | --- |
| Cloudflare expects `wrangler deploy` | Wrong creation flow: choose **Pages**, not **Workers**. |
| Cannot find `@local/docx-ir-engine/core` | Wrong root/build command. Use root `/` and `pnpm run build`, which builds the engine first. |
| No `out` directory / output path not found | The export is under `apps/web/out`; use that as the **build output directory**. |
| `next start` complains about `output: 'export'` | This is expected. Use `pnpm start` (the local static-file preview), not `next start`. Cloudflare does not need a start command. |
| `sample.docx` or ribbon icons return 404 | Check `apps/web/public/` is committed and included in the static export. |
| Blank page or missing JS on Cloudflare | Inspect the browser console/network tab and verify `_next/static` files were deployed. The `verify:pages` build step detects a missing client bundle. |
| Build works locally but fails in Cloudflare | Check that Cloudflare uses Node 22, pnpm 10, repository root `/`, and the complete committed lockfile; inspect the first failed build-log step. |

## Further reading

- [Cloudflare Pages: Next.js static export](https://developers.cloudflare.com/pages/framework-guides/nextjs/deploy-a-static-nextjs-site/)
- [Cloudflare Pages: monorepos](https://developers.cloudflare.com/pages/configuration/monorepos/)
- [Cloudflare Pages: build image and versions](https://developers.cloudflare.com/pages/configuration/build-image/)
- [Next.js: static exports](https://nextjs.org/docs/app/guides/static-exports)
