# DOCX engine and Next.js workspace

A browser-compatible TypeScript DOCX importer, portable document IR, and React preview renderer, with a Next.js host application. It now includes table-based documents, inherited formatting and RTL/LTR typography. DOCX editing and export are not implemented.

## Run the Next.js app

```powershell
pnpm install
pnpm dev
```

Open **http://127.0.0.1:3000** and click **Open DOCX**. An English/Persian/Arabic sample opens automatically so there is something to see immediately. **Custom fonts** lets you load a regular font face using the exact family name from your document. Documents are parsed locally in your browser.

The preview supports tables, merged cells, lists, inherited styles, images, page backgrounds and document margins. Floating drawings and tables use approximate browser layout; headers/footers appear once rather than repeating across automatically paginated pages. Compatibility notes explain known approximations for each document.

`pnpm dev` builds the engine and starts Next.js. Keep the terminal running while using the app; press Ctrl+C to stop it. Node.js 20.9 or newer is required.

| Command | Purpose |
| --- | --- |
| `pnpm dev` | Start Next.js at http://127.0.0.1:3000 |
| `pnpm build` | Build the engine and the Next.js app for production |
| `pnpm start` | Start the production app after building |
| `pnpm build:engine` | Build only the extractable engine into `dist/` |
| `pnpm typecheck` | Check the engine and app types |
| `pnpm test` | Run engine and sample-integration tests |

The application lives in `apps/web/`; the engine remains in `src/`. The app imports `@local/docx-ir-engine` and `@local/docx-ir-engine/react` through a workspace dependency, without reaching into private source files. Rebuild with `pnpm build:engine` after changing engine source while the app is running. Regenerate the sample using `node scripts/create-sample.mjs`.

## Review and roadmap

The [engine review](docs/engine-review.md) covers the parser, document IR, RTL/LTR typography and remaining compatibility risks. The [feature matrix](docs/compatibility.md) records current support, and the [development plan](docs/development-plan.md) defines staged engine work followed by the Next.js editor.

The build now uses the xmldom 0.9 callback API and DOM types. Synthetic/component and sample-integration checks cover the current subset; real Word interoperability, browser typography and export remain future validation work.

## Drop into your Next.js/React project

1. Extract this directory to your repository (for example `packages/docx-ir-engine`).
2. Build only the engine with `pnpm build:engine`.
3. Install from the package directory: `npm install ./packages/docx-ir-engine` (or use your preferred workspace tooling).

The ZIP includes TypeScript source but not `node_modules`. Build dependencies: `fflate`, `@xmldom/xmldom`; peer dependency: `react >=18`.

```tsx
'use client';
import { useState } from 'react';
import { importDocx, type DocumentIR } from '@local/docx-ir-engine';
import { DocumentRenderer } from '@local/docx-ir-engine/react';

export default function DocxPreview() {
  const [doc, setDoc] = useState<DocumentIR | null>(null);
  const [error, setError] = useState<string | null>(null);
  return <div>
    <input type="file" accept=".docx" onChange={async event => {
      const file = event.currentTarget.files?.[0];
      if (!file) return;
      try { setError(null); setDoc(await importDocx(await file.arrayBuffer())); }
      catch (e) { setError(e instanceof Error ? e.message : String(e)); }
    }} />
    {error && <p role="alert">{error}</p>}
    {doc && <DocumentRenderer document={doc} pageLayout />}
  </div>;
}
```

If a document uses `B Zar`, `B Nazanin`, or another proprietary font, **load the font in the host application** using `@font-face` or Next.js font loading. The engine does not embed or supply fonts.

## Supported subset

- OPC ZIP reading, root and source-relative part relationships, Transitional/Strict namespaces.
- Nested tables, grid widths, column/row merges, RTL column order, borders, shading and cell margins.
- Document defaults, paragraph/character/table style inheritance, style toggles, theme fonts/colors and direct overrides.
- Separate font slots and complex-script formatting, grapheme-safe mixed-script render segments, `w:rtl` and `w:cs`.
- Color, underline, highlight, strike, superscript/subscript, run position, letter spacing and paragraph mark typography.
- Paragraph direction/alignment, spacing, line height, indents, basic tabs, keep rules and typed page/column breaks.
- Bullet/numbered lists, multilevel counters and start overrides; safe hyperlinks.
- Embedded raster images, sizes/crops, basic rectangle/ellipse drawings, textboxes and positioned anchors.
- One-branch AlternateContent traversal, block/inline content controls and inserted text.
- Page size/margins/backgrounds, header/footer variants and their local relationships.

The current `DocumentIR` has `schemaVersion: 2`. `blocks` is a paragraph/table union and paragraph children can be text, drawings or typed breaks. Use `getParagraphs(document.blocks)` and `getTextRuns(document.blocks)` instead of assuming every block has text children. Resolved styles and direct overrides remain separate; this is still preview IR, not an OOXML preservation archive.

## Known boundaries — important

- Exact Word pagination, floating-object wrapping, all section transitions, repeating headers/footers and field recalculation remain incomplete. Page/column breaks are retained, but screen preview flows rather than recreating Word pages.
- Drawing support covers the sample documents' common image, rectangle, ellipse and textbox forms. Arbitrary geometry/effects, VML-only graphics, standalone SVG, charts and OLE are unsupported or approximate.
- Conditional table style regions, sophisticated tab alignment, all numbering formats and full Word font-selection behavior remain incomplete. Fonts must be available in the browser; embedded DOCX font decoding is not implemented.
- Comments/review editing, footnotes/endnotes, encrypted containers, UTF-16 XML decoding and DOCX export remain outside the supported subset. Unknown XML is not retained for round-trip saving.
- The archive reader preflights declared ZIP sizes, safe names, duplicate entries and entry count before inflation (defaults: 50 MiB and 2,000 entries), then checks extracted sizes. It still runs synchronously, trusts ZIP metadata for allocation, and lacks XML depth/node and CPU budgets; it is not yet a complete hostile-upload boundary.
- Unsupported constructs are intentionally skipped and cannot be round-tripped.

## Tests

```bash
npm install
npm run typecheck
npm test
npm run build
```

[Sample compatibility results](docs/sample-compatibility.md) record the four supplied documents' missing features and current checks. `node scripts/check-samples.mjs <file.docx> ...` compares source/imported body text and writes local HTML previews under ignored `.cache/sample-previews/`. Personal sample files are not copied into the app's public assets or committed as fixtures.

## Architecture

```text
src/model/        plain Word-independent IR
src/docx/         OPC package, XML and WordprocessingML import only
src/react/        presentation-only React component
apps/web/         Next.js upload/font-loading host
```

Open XML docs: [WordprocessingML](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-documents), [run fonts](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts), [font size](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fontsize). See ISO/IEC 29500 / ECMA-376 for authoritative specifications.
