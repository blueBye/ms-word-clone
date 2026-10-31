# DOCX IR Engine — Milestone 1

A small, browser-compatible TypeScript importer and React renderer for a controlled subset of DOCX WordprocessingML, with explicit Persian/Arabic complex-script properties. **Not a Microsoft Word clone**; this milestone does not export DOCX.

## Drop into your Next.js/React project

1. Extract this directory to your repository (for example `packages/docx-ir-engine`).
2. Run `npm install` and `npm run build` inside this directory.
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
    {doc && <DocumentRenderer document={doc} />}
  </div>;
}
```

If a document uses `B Zar`, `B Nazanin`, or another proprietary font, **load the font in the host application** using `@font-face` or Next.js font loading. The engine does not embed or supply fonts.

## Supported subset

- OPC ZIP reading, root relationships, main officeDocument part discovery.
- Transitional and Strict WordprocessingML document namespaces.
- `w:document`, `w:body`, top-level `w:p`, `w:r`, `w:t`, `w:tab`, `w:br`, `w:cr`.
- Basic direct run properties: `w:rFonts` (`ascii`/`hAnsi`/`cs`), `w:b`, `w:i`, `w:bCs`, `w:iCs`, `w:sz`, `w:szCs`, `w:rtl`.
- Basic direct paragraph properties: `w:jc`, `w:bidi`.
- Standard XML namespace-aware DOM parsing; explicit `w:val="0"` off values and half-point font sizes.
- Empty paragraphs preserved, run and text order preserved, no merging of run fragments.
- Basic text traversals through `w:hyperlink`, insertion elements, and simple structured-document wrappers (metadata excluded).

## Known boundaries — important

- **No styles.xml resolution**: inherited/default paragraph and character styles, theme fonts, and document defaults are ignored. Many real DOCX files rely on these for styling; they will therefore render only approximately.
- **No tables**, pictures, images, lists, headers, footers, section/page layout, fields, comments, or export.
- The renderer selects complex-script typography for a whole run when it contains Persian/Arabic characters; **mixed-script text within one run is an approximation**, not per-glyph Open XML shaping. In complex-script-heavy runs, split into distinct segments later when editing is added.
- `w:ascii` vs `w:hAnsi` differences are collapsed into one base font family. `w:cs` stays separate.
- Paragraph `start`/`end` justification is resolved using explicit direct paragraph direction; inherited direction is not currently known.
- The archive reader decompresses ZIP entries in memory before enforcing the configured maximum uncompressed size. **Do not use it as-is for untrusted, adversarial DOCX files or arbitrarily large uploads**; add a streaming/ZIP-header budget precheck or sandbox the importer.
- Unsupported constructs are intentionally skipped and cannot be round-tripped.

## Tests

```bash
npm install
npm run typecheck
npm test
npm run build
```

## Architecture

```text
src/model/        plain Word-independent IR
src/docx/         OPC package, XML and WordprocessingML import only
src/react/        presentation-only React component
```

Open XML docs: [WordprocessingML](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-documents), [run fonts](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts), [font size](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fontsize). See ISO/IEC 29500 / ECMA-376 for authoritative specifications.
