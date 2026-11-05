# Contributing

Thanks for your interest in the DOCX IR Engine & Web Editor. This project is under active development and prioritizes **document integrity, well-scoped interoperability, and regression coverage** over broad but unsupported feature claims.

## Development setup

You need Node.js 20.9+ and pnpm 9+.

```bash
pnpm install --frozen-lockfile
pnpm dev
```

For a verification run:

```bash
pnpm typecheck
pnpm test
pnpm build
```

A build and a test run are independent; run both. If a command fails, include the failure in the pull-request description instead of reporting it as passing.

## Before submitting a change

1. Open an issue or describe the intended behavior clearly. Include the DOCX feature or browser behavior you want to address.
2. Add or update **automated regression tests** for the changed functionality. For DOCX writing, test edited export and reimport, preservation of unrelated source XML/parts and undo/redo where relevant.
3. Keep `src/core` independent of XML, React, Next.js and browser APIs. Format-specific work belongs in `src/docx`, browser adapters in `apps/web`, and presentation in `src/react`.
4. Update the [feature-status matrix](docs/feature-status.md), usage docs, and relevant [Home](docs/home-editor-status.md) or [Insert/Layout](docs/ribbon-status.md) status note when behavior changes.
5. Run and report `pnpm typecheck`, `pnpm test`, and `pnpm build`.

## Testing DOCX compatibility

For interoperability bugs, provide the smallest **anonymized** DOCX that reproduces the issue, the expected behavior, the current behavior, and the relevant Office/browser version. Avoid adding customer documents, copyrighted files without permission, private correspondence, proprietary fonts, or personal information as public fixtures.

Useful test locations:

- `tests/core.test.ts`, `tests/architecture.test.ts` — model commands and architecture.
- `tests/session.test.ts`, `tests/package.test.ts` — source-retaining session and package safety.
- `tests/home-editing.test.ts`, `tests/authoring-objects.test.ts`, `tests/fidelity.test.ts` — editing, objects and interoperability cases.
- `tests/renderer.test.ts`, `tests/measurement.test.ts` — rendering/measurement.

Synthetic fixtures live with tests and `tests/helpers/`. `scripts/check-samples.mjs` can create local comparison output in ignored `.cache/` directories.

## Coding and documentation

- Favor narrow changes and types over broad casts or assumptions about every paragraph being editable.
- Do not serialize HTML or a plain core snapshot as a replacement for a source-preserving DOCX session.
- Treat unknown XML and complex Word structures conservatively: restrict edits rather than silently dropping content.
- Document limitations and verify native Word rendering when a feature depends on it.
- Keep commits free from generated `dist/`, `.next/`, local samples, credentials and other ignored files.

## Security issues

Do **not** open a public issue containing an exploit, sensitive data or a malicious document. Follow [SECURITY.md](SECURITY.md).

## License

By contributing, you agree that your contributions are made available under the repository's [MIT License](LICENSE). Do not copy proprietary Microsoft Office interface assets or redistributable fonts into the project without an appropriate license.
