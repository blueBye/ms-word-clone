# API and integration

The root package is named **`@local/docx-ir-engine`** and is currently **private** (`"private": true` in `package.json`), so it is not a published npm dependency. Build it within this repository or link/package it into another workspace yourself.

```bash
pnpm install --frozen-lockfile
pnpm build:engine
```

Public package entries:

| Import | Purpose | Notes |
| --- | --- | --- |
| `@local/docx-ir-engine/core` | Format-independent editing state, commands, styles, helpers | No OOXML, DOM or React dependency |
| `@local/docx-ir-engine/docx` | OOXML package, import and source-retaining editing sessions | Use when exporting imported Word files |
| `@local/docx-ir-engine/react` | DocumentRenderer, PaginatedDocument and view utilities | React peer dependency |
| `@local/docx-ir-engine` | Compatibility/root exports | Prefer targeted entries for layer isolation |

## Preview a DOCX

```tsx
"use client";

import { useState } from "react";
import { importDocx } from "@local/docx-ir-engine/docx";
import { DocumentRenderer } from "@local/docx-ir-engine/react";
import type { DocumentIR } from "@local/docx-ir-engine/core";

export function DocxPreview() {
  const [document, setDocument] = useState<DocumentIR | null>(null);
  const [error, setError] = useState("");

  return (
    <div>
      <input
        type="file"
        accept=".docx"
        onChange={async (event) => {
          const file = event.currentTarget.files?.[0];
          if (!file) return;
          try {
            setError("");
            setDocument(await importDocx(await file.arrayBuffer()));
          } catch (cause) {
            setError(cause instanceof Error ? cause.message : String(cause));
          }
        }}
      />
      {error && <p role="alert">{error}</p>}
      {document && <DocumentRenderer document={document} pageLayout />}
    </div>
  );
}
```

This API creates a **renderable document model**, not a source-preserving editing session. For paginated page measurement, the full app uses `PaginatedDocument`.

## Open, edit and export with preservation

```ts
import { openDocxSession } from "@local/docx-ir-engine/docx";
import { getParagraphs } from "@local/docx-ir-engine/core";

async function prependGreeting(file: File): Promise<Uint8Array> {
  const session = await openDocxSession(await file.arrayBuffer());
  const first = getParagraphs(session.document.blocks)[0];
  if (!first?.id) throw new Error("The document has no editable paragraph");
  if (!session.capability(first.id).text) {
    throw new Error("This paragraph is not editable");
  }

  await session.execute({
    type: "text",
    id: first.id,
    start: 0,
    end: 0,
    text: "Hello فارسی ",
  });

  return session.exportDocx();
}
```

`openDocxSession` retains the source package while the document is edited. Its `DocumentSession` contract exposes state, `execute` commands, selection, style operations, undo/redo, and capability checks. `DocxSession` also exposes:

- `exportDocx(): Promise<Uint8Array>` — download-ready OOXML ZIP bytes.
- `original(): Uint8Array` — the original document bytes.
- `compatibilityStatistics` — package parses, part writes and queued commands (diagnostics, not a performance SLA).

Never assume every imported paragraph can be edited. Check `session.capability(id)` and honor `session.lockedReason` when present. A digitally signed or enforced editing-protected document can be locked.

To download a resulting `Uint8Array` from a browser host:

```ts
const bytes = await prependGreeting(file);
const blob = new Blob([new Uint8Array(bytes)], {
  type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
});
const url = URL.createObjectURL(blob);
const link = document.createElement("a");
link.href = url;
link.download = "edited.docx";
link.click();
URL.revokeObjectURL(url);
```

The `new Uint8Array(bytes)` copy is intentional to make a browser-compatible Blob part.

## Create a DOCX from scratch

```ts
import { createBlankDocx, openDocxSession } from "@local/docx-ir-engine/docx";
import { getParagraphs } from "@local/docx-ir-engine/core";

const session = await openDocxSession(createBlankDocx());
const first = getParagraphs(session.document.blocks)[0];
if (!first?.id) throw new Error("Blank document is missing a paragraph");
await session.execute({
  type: "text",
  id: first.id,
  start: 0,
  end: 0,
  text: "A new document",
});
const output = await session.exportDocx();
```

## Work without Word/OOXML

```ts
import { createDocumentEngine } from "@local/docx-ir-engine/core";

const engine = createDocumentEngine();
await engine.execute({
  type: "text",
  id: "p1",
  start: 0,
  end: 0,
  text: "Hello world",
});
console.log(engine.statistics.words);

const snapshot = engine.snapshot();
// A JSON-serializable core document/state snapshot, not a DOCX file.
```

The portable engine's command offsets are UTF-16 offsets and are validated against grapheme boundaries. For command variants and document structures, consult `src/core/types.ts`, `src/model/document.ts`, and `tests/core.test.ts`.

## Boundaries

- React is optional for the core/docx package; required for `/react`.
- The current engine uses `@electric-sql/d2ts` for incremental dependencies, not persistence or collaboration.
- The parser has default archive budgets and does not accept every Word feature or arbitrary untrusted input without risk.
- A core JSON snapshot does **not** contain the opaque DOCX envelope needed for safe round-trips.
- Rendering depends on the host's available fonts; the engine does not distribute proprietary fonts.

Continue with [architecture](engine-architecture.md), [feature status](feature-status.md), and [limitations](limitations.md).
