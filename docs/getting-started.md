# Getting started

## Requirements

- Node.js **20.9 or newer**.
- **pnpm 9 or newer** (this repository has a pnpm lockfile v9).
- A current desktop browser for the best editor experience.
- No database or account is required for the standalone workspace.

This is a pnpm workspace with the private root package `@local/docx-ir-engine` and a Next.js app in `apps/web`.

## Install and run

```bash
pnpm install --frozen-lockfile
pnpm dev
```

Open **http://127.0.0.1:3000**. The `dev` script first compiles the engine and then starts Next.js. By default, the editor opens the bundled sample. Choose **File → Open** to use your own `.docx`, or create a blank document.

A typical first test:

1. Open the sample or a non-sensitive test DOCX.
2. Click a paragraph, type text, and change formatting from **Home**.
3. Select **Download DOCX** and open the saved file in Microsoft Word or another compatible viewer.
4. Verify content, pagination and any features important to your document.

**Do not overwrite your only copy of an important document.** The editor does not guarantee exact Word fidelity for every DOCX construct.

## Commands

| Command | Description |
| --- | --- |
| `pnpm dev` | Build engine + run Next.js dev server on `127.0.0.1:3000` |
| `pnpm build:engine` | Build the importable TypeScript package into `dist/` |
| `pnpm build:web` | Build the Next.js application (requires the engine first) |
| `pnpm build` | Build engine, then Next.js |
| `pnpm start` | Preview the static export locally after a successful build |
| `pnpm typecheck` | Type-check root engine and web app |
| `pnpm test` | Run the Vitest suite |
| `node scripts/create-sample.mjs` | Generate the bundled sample DOCX |
| `node scripts/check-samples.mjs path/to/file.docx` | Compare imported body text and create local HTML previews |
| `node scripts/benchmark-engine.mjs` | Run engine benchmark after `pnpm build:engine` |

`check-samples` writes local previews to `.cache/sample-previews/`, which is git-ignored. Its results are diagnostic rather than a proof of Word compatibility.

## Production run

```bash
pnpm install --frozen-lockfile
pnpm build
pnpm start
```

The web project uses **Next.js server mode**. It is **not configured as a static export**. Deploy to a host capable of running the Next.js server. Source documents do not require a backend storage service in the current app.

## Fonts and multilingual text

Document fonts need to exist in the browser environment. Use the editor's **Fonts** control to load a local font face (provide the family name expected by the document). Mixed-script rendering supports Latin and complex-script font channels, but browser font shaping and available glyphs affect the result. Missing font variants, embedded DOCX fonts, and Word-specific line breaks can change pagination.

Do not commit proprietary font files unless you have redistribution rights.

## Troubleshooting

**`pnpm: command not found`:** install pnpm or enable it via Node's Corepack, then retry. A package registry connection may be needed on first setup.

**Module resolution errors for `@local/docx-ir-engine`:** run `pnpm install` from the repository root and `pnpm build:engine` before the web command. The app imports built package outputs from `dist/`.

**Port 3000 occupied:** stop the other service, or change the app's configured dev/start port in `apps/web/package.json`.

**Missing fonts / different page breaks:** load the correct family, zoom to inspect the output, then verify in Word. Browser pagination is approximate.

**DOCX cannot be edited:** the session can lock protected/signed documents and some complex/unsupported paragraphs. Inspect the compatibility panel; do not remove protections to force editing.

**Build or test fails:** run the exact failing command alone and inspect the error. `pnpm build` compiles the app but does not run the test suite. Report the failing command, Node/pnpm versions, and an anonymized reproduction.

See [Using the editor](using-editor.md), [Feature status](feature-status.md), and [Known limitations](limitations.md).

## Cloudflare Pages hosting

The editor builds to `apps/web/out` using Next.js `output: 'export'`. Cloudflare must build from the repository root using `pnpm run build` (to compile the engine first) and deploy `apps/web/out`. No Node.js server or Wrangler deploy command is used. See [Cloudflare Pages deployment](../DEPLOY_CLOUDFLARE_PAGES.md).
