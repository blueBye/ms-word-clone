# Publishing the repository on GitHub

The project can be shared as an open-source repository, but **read this checklist before making it public**. Publishing code is not the same as establishing production security or Microsoft Word compatibility.

## Before you publish

- [ ] Confirm that all source files, tests and documentation may be made public.
- [ ] Replace the unnamed copyright holder in the root [MIT LICENSE](../LICENSE) with the correct author/organization, if appropriate. Do not change licensing terms without permission from the relevant rights holders.
- [ ] Review the separate license for the bundled [Microsoft Fluent System Icons](../apps/web/public/icons/fluent/README.md) and retain its copyright notices.
- [ ] Check the full project and Git history for tokens, private keys, `.env` files, sample documents with private content, proprietary fonts, customer information, and private URLs. The root `.gitignore` helps prevent **new** accidental additions, but **does not erase files previously committed**.
- [ ] Confirm the bundled `apps/web/public/sample.docx` contains only authorized demonstration material.
- [ ] Run `pnpm typecheck`, `pnpm test` and `pnpm build` on a machine with dependencies installed; record failures honestly.
- [ ] Decide which parts of the [feature status](feature-status.md) and [limitations](limitations.md) you want to highlight to first-time users.
- [ ] Enable GitHub's **Private vulnerability reporting** where available, and configure repository settings/issue policy for incoming reports.

## Create a repository

If this folder is **not already a Git repository**, create a new empty GitHub repository through the GitHub UI (avoid auto-generating a second README or license). Then run the following **from this project directory**:

```bash
git init
git add .
git status
git commit -m "Initial public repository with documentation"
git branch -M main
git remote add origin https://github.com/<USERNAME>/<REPOSITORY>.git
git push -u origin main
```

Replace `<USERNAME>` and `<REPOSITORY>` with your actual GitHub username/org and chosen repository name. If this project already has commits or a remote, **do not re-initialize or overwrite history**; review `git status`, `git remote -v`, and use your existing workflow.

## Suggested GitHub About text

**Description:** Browser-based DOCX editor with a portable TypeScript document engine, OOXML-preserving editing sessions, RTL typography, and a Next.js Word-style workspace.

**Topics:** `docx`, `ooxml`, `typescript`, `nextjs`, `react`, `document-editor`, `wordprocessingml`, `rtl`, `open-source`

Avoid claiming “Microsoft Word compatible” without qualification. The editor is a partial/experimental implementation, and it is not affiliated with Microsoft.

## After publishing

- Add real screenshots or a short demo GIF **only if** the images are from a working, up-to-date build and contain no confidential files.
- Pin documentation and report issues with reproducible synthetic samples.
- Consider GitHub Actions CI after dependency installation and the first verified clean test/build run.
- Tag releases only after running checks and documenting the feature/compatibility scope.

For GitHub-hosted deployment through Cloudflare Pages, follow [the monorepo deployment guide](../DEPLOY_CLOUDFLARE_PAGES.md) and verify the static export before tagging a release.
