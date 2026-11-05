# DOCX engine and Next.js workspace

A portable TypeScript document engine, an OOXML compatibility adapter, and a React rendering layer, with a Next.js host application. The core manages document state, styles, text, history and incremental dependencies using d2ts. The Word adapter retains original source parts for compatible editing/export. Editing is restricted to supported content; this is not full Word compatibility.

## Run the Next.js app

```powershell
pnpm install
pnpm dev
```

Open **http://127.0.0.1:3000** and use **File → Open**. An English/Persian/Arabic sample opens automatically so there is something to see immediately. **Fonts** lets you load a regular font face using the exact family name from your document. Documents are parsed locally in your browser.

The preview supports tables, merged cells, lists, inherited styles, image opacity, page backgrounds and document margins. The Next.js viewer measures content into separate sheets, repeats first/default/even headers and footers, and reflows after fonts load. Floating drawings and Word's exact pagination remain approximate; indivisible content taller than a page is retained with an oversize notice.

`pnpm dev` builds the engine and starts Next.js. Keep the terminal running while using the app; press Ctrl+C to stop it. Node.js 20.9 or newer is required.

| Command             | Purpose                                             |
| ------------------- | --------------------------------------------------- |
| `pnpm dev`          | Start Next.js at http://127.0.0.1:3000              |
| `pnpm build`        | Build the engine and the Next.js app for production |
| `pnpm start`        | Start the production app after building             |
| `pnpm build:engine` | Build only the extractable engine into `dist/`      |
| `pnpm typecheck`    | Check the engine and app types                      |
| `pnpm test`         | Run engine and sample-integration tests             |

The application lives in `apps/web/`; the engine remains in `src/`. The workspace uses public `/core`, `/docx`, and `/react` package entries without reaching into private source files. The core runs without React or XML. Rebuild with `pnpm build:engine` after changing engine source while the app is running. Regenerate the sample using `node scripts/create-sample.mjs`.

## Review and roadmap

The [current architecture review](docs/engine-architecture.md) documents the core/OOXML/UI boundaries, d2ts integration, performance evidence, preservation behavior and remaining work. Run `node scripts/benchmark-engine.mjs` after building to reproduce the core benchmark.

The [Word interface inventory](docs/word-ui-inventory.md) records the locally inspected File, Home, Insert and Layout commands, menus and dialogs. The [UI implementation plan](docs/word-ui-implementation-plan.md) specifies the ribbon, browser behavior, engine dependencies and phased editing/export gates. The [Home implementation status](docs/home-editor-status.md) and [Insert/Layout status](docs/ribbon-status.md) record working controls, verification and remaining gaps.

The [engine review](docs/engine-review.md) covers the parser, document IR, RTL/LTR typography and remaining compatibility risks. The [feature matrix](docs/compatibility.md) records current support, and the [development plan](docs/development-plan.md) defines staged engine work followed by the Next.js editor.

The build now uses the xmldom 0.9 callback API and DOM types. Synthetic/component and sample-integration checks cover the supported editing subset and package round trips. Word interoperability and browser typography still need a broader real-document corpus.

## Drop into your Next.js/React project

1. Extract this directory to your repository (for example `packages/docx-ir-engine`).
2. Build only the engine with `pnpm build:engine`.
3. Install from the package directory: `npm install ./packages/docx-ir-engine` (or use your preferred workspace tooling).

The ZIP includes TypeScript source but not `node_modules`. Runtime dependencies: `@electric-sql/d2ts`, `fflate`, `@xmldom/xmldom`. React is an optional peer required by the `/react` entry.

```tsx
"use client";
import { useState } from "react";
import { importDocx } from "@local/docx-ir-engine/docx";
import type { DocumentIR } from "@local/docx-ir-engine/core";
import { DocumentRenderer } from "@local/docx-ir-engine/react";

export default function DocxPreview() {
  const [doc, setDoc] = useState<DocumentIR | null>(null);
  const [error, setError] = useState<string | null>(null);
  return (
    <div>
      <input
        type="file"
        accept=".docx"
        onChange={async (event) => {
          const file = event.currentTarget.files?.[0];
          if (!file) return;
          try {
            setError(null);
            setDoc(await importDocx(await file.arrayBuffer()));
          } catch (e) {
            setError(e instanceof Error ? e.message : String(e));
          }
        }}
      />
      {error && <p role="alert">{error}</p>}
      {doc && <DocumentRenderer document={doc} pageLayout />}
    </div>
  );
}
```

If a document uses `B Zar`, `B Nazanin`, or another proprietary font, **load the font in the host application** using `@font-face` or Next.js font loading. The engine does not embed or supply fonts.

The Next.js app uses `PaginatedDocument` from the React entry point for measured page sheets. `DocumentRenderer` remains a single-sheet/flow primitive. `getRequiredFonts(document)` includes header/footer textboxes and list markers; the app probes local font availability and identifies missing fonts by location. Loading a font increments `fontRevision` so pages are measured again.

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

- Browser-measured page layout is not Word pagination. Pages split between paragraphs, table rows, and paragraphs within splittable table cells; vertical merges stay together. Cell padding and paragraph spacing count toward page capacity. Rows marked `cantSplit` or exact-height remain intact; paragraphs are not split into line fragments. Full section transitions and floating-object wrapping remain incomplete. Basic PAGE/NUMPAGES fields update per sheet; arbitrary fields do not.
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
