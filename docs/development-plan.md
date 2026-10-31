# DOCX engine and Next.js development plan

UI planning update (2026-11-01): the [installed Word inventory](word-ui-inventory.md) and [UI implementation plan](word-ui-implementation-plan.md) now define File/Home/Insert/Layout. The next application batch is a Word-like shell around the approved preview, followed by an editing/preservation prototype. The engine gates below remain required before enabling edited DOCX export; this planning update does not complete them.

Date: 2026-10-31. This plan follows the [source review](engine-review.md) and [feature matrix](compatibility.md). The current change completes M0 only. M1–M6 describe future work, with gates that prevent unsupported content from being silently rewritten.

Implementation update: the local Next.js shell now exists in `apps/web/`, with DOCX upload, engine rendering, a bilingual sample and regular-face font loading. Run `pnpm dev` from the repository root. This advances the application shell portion of M5; the common-feature engine, editing and export gates below remain outstanding.

Sample-driven 0.2 update: import/preview now includes typed tables/drawings/breaks, style/default/theme resolution, separate font slots, numbering, header/footer parts, source page settings and diagnostics. The [sample report](sample-compatibility.md) records measured coverage. These complete parts of the M2–M4 import work; source preservation, writer, exact pagination and editing release gates remain open.

## Product contract

The long-term goal is a web document editor with strong DOCX interoperability, especially Persian/Arabic/Hebrew and English plus custom fonts. The first release should promise a named common-feature profile rather than universal Microsoft Word equivalence. Keep import, preview, editable support, exported support and opaque preservation independently measurable.

For any opened document, the application must be able to say which features it understands, which are approximate, which are locked but preserved, and which prevent safe saving. “Imported successfully” must never imply “no content was omitted.” The original file remains available until a verified replacement is saved.

Compatibility is based on [ECMA-376](https://ecma-international.org/publications-and-standards/standards/ecma-376/), supplemented by Word-authored examples and Microsoft's implementation notes. Tests should distinguish standard requirements, measured Word behavior and intentional product limitations.

## Structure and extraction boundary

Keep this repository as one small package for now. Moving everything into a monorepo before the app exists would create churn without improving the boundary. Preserve the current public imports:

```ts
import { importDocx, type DocumentIR } from '@local/docx-ir-engine';
import { DocumentRenderer } from '@local/docx-ir-engine/react';
```

Add directories only when their implementation starts:

```text
src/
  model/                 semantic types, validation, positions
  docx/
    package.ts           OPC parts and relationship resolution
    parser/              authored OOXML -> semantic nodes/source map
    writer/              semantic edits -> changed OOXML parts
    import-docx.ts       orchestration and compatibility wrapper
  resolve/               styles, theme, effective fonts/properties
  editing/               portable commands and history integration
  react/                 preview and browser editor adapters
tests/
  helpers/               small synthetic package builders
  fixtures/              approved source DOCX + expectations/manifest
docs/                    review, matrix and plan
```

Dependency direction: `model` has no imports from DOCX, React, Next.js or browser globals. DOCX parsing/writing uses the model. Resolution uses semantic style data. React consumes resolved data and calls commands. Next.js imports package exports; the package never imports the app. No private `src/` imports from the future host project.

The core may use xmldom internally; it must not require a browser `document`, Node filesystem, application router or global CSS. Keep React behind its existing subpath; make its peer optional when validating headless consumers, or split the renderer into a package when it actually has a separate lifecycle.

The app now exists in `apps/web` and consumes the root engine package through pnpm workspaces. Keep that working boundary for the UI work; there is no need to relocate the engine now. If it later moves to `packages/docx-ir-engine`, keep package exports/tests intact and add a consumer build/import smoke test at that move. No Turborepo, dependency injection container or plugin system is needed for this boundary.

## Model decisions before adding features

1. **Semantic IR and source session are separate.** The IR stays serializable and portable. A DOCX session retains original archive bytes, original part bytes, relationships, source mappings and diagnostics. Runtime font objects, object URLs, DOM nodes and ZIP buffers do not become paragraph properties.
2. **Authored and resolved properties are separate.** Preserve style IDs, explicit off/absent values, theme references, language and font slots. Compute effective properties without overwriting the authored form. This supports editing without destroying inheritance.
3. **Use discriminated block and inline unions.** Add paragraph/table blocks and text/tab/break/hyperlink/image/field inlines as needed. Keep explicit line/page/column break kinds. Do not represent every element as a generic property bag, and do not force OOXML into HTML.
4. **Use stable identities for edits.** Add deterministic import IDs and session-generated IDs for new nodes. Document model positions as UTF-16 offsets with grapheme-safe commands. Preserve logical Unicode text exactly, including ZWNJ, combining marks and directional controls.
5. **Retain logical layout intent.** Preserve start/end alignment and logical indents until direction is resolved. Record physical DOCX measurements explicitly and use named unit conversions at layout/export boundaries.
6. **Unknown content has an ownership rule.** Source mappings identify which parts/nodes are editable. Unsupported nodes retain source context and appear as locked placeholders where feasible. Editing a container with unknown children must either preserve them safely or produce a blocking diagnostic.
7. **Version the IR at the first actual schema expansion.** Provide migration/validation for persisted IR then. Do not store transient computed CSS as the saved document.

Introduce a richer session API alongside the existing `importDocx()` compatibility wrapper. The session can contain `document`, `diagnostics` and retained source state. Keep the wrapper returning `DocumentIR` until a deliberate breaking release. Diagnostic records should include a code, severity, feature, part path, node/source location and preservation/editability outcome. Never silently change the existing return type.

## M0 — Establish a runnable baseline (completed)

- [x] Fix xmldom callback and DOM typing throughout parser callers.
- [x] Fail on XML parse diagnostics; reject DTDs.
- [x] Preflight ZIP declared bytes, count, names and duplicates before inflation.
- [x] Validate configured limits and decoded main targets.
- [x] Restrict inline content-control traversal to `sdtContent`.
- [x] Expand synthetic checks and document remaining gaps.
- [x] Run build/typecheck/tests and import both built public exports.

Evidence: 81 passing checks in four test files. This milestone does not establish browser layout, font loading, export or Word compatibility.

## M1 — Safe package ingestion, diagnostics and preservation

**Depends on:** M0. **Purpose:** prevent invisible data loss before broadening editable coverage.

| Task | Main area | Acceptance criteria |
| --- | --- | --- |
| OPC-01 | `docx/package.ts` | Parse/validate content types and main relationship cardinality/type/mode; decode UTF-8 and UTF-16 XML correctly; diagnose unsupported containers. |
| OPC-02 | Package relationships | Resolve subordinate targets relative to their source part, including legal parent segments within package root; reject escaping, external and ambiguous targets. No external fetching during import. |
| XML-01 | XML boundary | Add per-part byte, nesting and node-count budgets; test malformed/duplicate attributes, unsupported encodings and excessive nesting. Validate ZIP integrity/CPU failure handling beyond declared-size preflight. |
| MCE-01 | Compatibility processing | Select exactly one supported AlternateContent choice/fallback using namespace capabilities; handle Ignorable/ProcessContent/MustUnderstand policies. Preserve originals. Never import both alternatives. |
| IR-01 | Model + import session | Add source mapping and structured diagnostics; represent/lock unsupported block/inline content instead of silently omitting it. Start schema versioning and node IDs. |
| PRES-01 | DOCX session/export | Unchanged save returns original bytes. Retain all original parts for edited exports; source mappings define safe write boundaries. Unknown XML in changed parts has dedicated tests. |

The distinction between selecting content for display and retaining original markup is important. Microsoft's [markup-compatibility documentation](https://learn.microsoft.com/en-us/office/open-xml/general/introduction-to-markup-compatibility) explains alternate-content processing; destructive preprocessing alone is insufficient for this preservation goal.

**Exit gate:** a mixed document with tables, drawings and comments reports its limitations; an unchanged export is byte-identical; unsupported content is never reported as editable. Fixtures cover both namespace families and relocated main parts. Export of unsafe edits is blocked rather than dropping source content.

## M2 — Styles, RTL/LTR typography, custom fonts and first edited export

**Depends on:** M1. **Purpose:** a complete import -> edit -> export slice for ordinary formatted paragraphs.

| Task | Main area | Acceptance criteria |
| --- | --- | --- |
| STYLE-01 | Related parts / resolver | Read defaults, paragraph/character style chains, theme references and explicit overrides; detect cycles/missing IDs; implement toggle semantics with fixtures. Keep table/numbering contributions extensible for M3 without pretending they exist now. |
| FONT-01 | IR / resolver | Retain ASCII, high ANSI, East Asian and complex-script slots, theme references, hints, language and forced-CS settings. Preserve all source settings on export. |
| BIDI-01 | Resolution / render segments | Separate paragraph direction, run RTL characteristics and CS formatting. Segment mixed runs by resolved font/format needs without changing source text or breaking combining sequences. Include Hebrew and neutral/whitespace runs. |
| FONT-02 | React/browser host | Add explicit font-family mapping and a font availability/loading state. Support uploaded licensed font variants, deterministic fallback and missing-glyph diagnostics; await font readiness before layout measurement. |
| TEXT-01 | IR / parser / writer | Add underline/color/highlight/strike/super/subscript and typed tabs/breaks; logical alignment, paragraph spacing/indents and empty-paragraph typography. No page-break flattening. |
| WRITE-01 | DOCX writer | Export paragraph text/style/font/direction changes while preserving untouched package parts and opaque source nodes. Create valid new DOCX files for the supported paragraph profile. |

Font family names alone do not provide font binaries. The host owns runtime loading, object URLs and cleanup; core receives stable family references and availability information. Keep regular/bold/italic/bold-italic variants explicit. Do not normalize Persian/Arabic characters or substitute digit forms just because the paragraph is RTL. Handle font names as literal names, not arbitrary CSS family lists.

Embedded font discovery/preservation starts here; activation follows verified decoding and embedding permissions. Browser font loading and DOCX font embedding are distinct operations. [Embedded font references](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.embedregularfont?view=openxml-3.0.1) and [font embedding guidance](https://support.microsoft.com/en-gb/office/fonts/benefits-of-embedding-custom-fonts) are starting points; use the normative font-part rules and binary fixtures for implementation.

Use browser shaping and bidi capabilities, with a tested OOXML-to-CSS mapping. Do not write an ad hoc Unicode bidi/shaping implementation. Font segmentation must not reorder text or become an editable node boundary by accident.

**Exit gate:** paragraph documents in English/Persian/Arabic/Hebrew render with recorded custom fonts and reopen in Word after edits without repair prompts. Test Latin/CS bold/italic independently, style-based RTL, `w:rtl` whitespace, `w:cs`, URLs, paired punctuation, Latin/Arabic/Persian digits, ZWNJ and combining marks. No unsupported-content loss on a supported text edit.

## M3 — Everyday document structures with matching writers

**Depends on:** M2. **Purpose:** the common DOCX feature profile becomes useful beyond plain text.

- **LIST-01:** numbering parts, abstract/instance IDs, levels, overrides, restart/start values, marker text/format and indents; RTL marker placement and digit behavior. Do not infer lists from visible text.
- **TABLE-01:** grids, rows/cells, nested paragraphs, widths, borders/shading, cell margins, horizontal/vertical merges, repeated headings and table RTL order. Handle unsupported floating/complex tables as preserved content.
- **LINK-01:** external hyperlinks with safe rendering policy, internal bookmarks/anchors and relationship IDs. Export retains targets; import never fetches them.
- **IMAGE-01:** inline raster images, dimensions, relationship-backed assets and alt text. Preserve source bytes; handle unsupported formats with a placeholder/diagnostic and no executable markup injection.
- **SDT-01:** visible block/inline content traversal with retained control metadata and explicit locked/editable policy.
- **STYLE-02:** supported table/numbering style contributions and conditional properties, with documented unsupported combinations.

Every task includes parser, IR, resolver/preview, writer and preservation tests in the same feature slice. A UI control does not become enabled just because the parser recognizes its markup.

**Exit gate:** a contract/report fixture with mixed RTL/LTR headings, multilevel lists, merged tables, links and inline images can be imported, edited, exported, reimported and reopened in Word. IDs and relationship references remain valid. Untouched source assets and unrelated parts retain their original uncompressed bytes.

## M4 — Sections, page layout and reference fidelity

**Depends on:** M3. **Purpose:** support normal document pages before presenting the product as a Word-like editor.

Implement body-final and paragraph-scoped section properties, page size/orientation/margins, page and section breaks, first/even/default headers/footers, supported page-number fields, paragraph keep/widow rules, and table row splitting. Add footnotes/endnotes with stable references. Keep cached field results and instructions separate; do not claim arbitrary field recalculation.

Build layout measurements only after required fonts have loaded. Decide whether the first UI provides a flowing editor plus print preview or true paginated editing through a measured prototype. A white page-shaped div is not evidence of Word pagination.

**Exit gate:** page size, orientation, section transitions, reference page count and text ordering match the approved corpus. Record known line-wrap/table/float differences as explicit limitations. Export must pass schema checks and Word reopen checks. Layout tolerances are established from measured fixtures using the same font files, not arbitrary screenshot similarity scores.

## M5 — Next.js document UI and editing

**Depends on:** the M1–M4 release gates for the selected profile. **Purpose:** a familiar Docs-style interface built on proven engine behavior.

First perform an editor-adapter spike with the exact target schema. Compare browser editing substrates on Persian/Arabic IME composition, bidirectional caret/selection, grapheme deletion, copy/paste, undo, marks, tables and source-preservation mapping. Select one based on working cases, then keep its data conversions localized. Establish one authoritative editing state and a documented mapping to the engine's serializable model; avoid two independently mutable documents.

Implement the app in this order:

1. `apps/web` package consumer, file open/import diagnostics, loaded-font status and a read-only preview.
2. Client editor boundary with commands, selection, undo/redo and local recovery. Parse/export in a worker with progress/cancel and stale-result protection.
3. Title/file actions; compact menus/toolbar; typography/style selectors; bold/italic/underline; independent paragraph direction and alignment controls; zoom and page canvas.
4. Lists, links, images and table commands backed by the supported writer. Show locked placeholders and compatibility details where required.
5. Save/export, unsaved-change status, offline recovery, keyboard navigation and accessibility. Add cloud persistence/collaboration only in a separate scoped phase.

Keep file/FontFace/selection/worker APIs inside the client side. Next.js documents the [client/server boundary](https://nextjs.org/docs/app/api-reference/directives/use-client); retained runtime sessions should not be passed through server-component props. Engine tests must continue running without Next.js.

UI locale, document paragraph direction, toolbar direction and page reading order are independent settings. User font uploads should expose availability/substitution clearly. Never label a fallback as the requested font when the requested file failed to load. Do not implement formatting by mutating preview DOM outside the editing adapter.

**Exit gate:** keyboard-only open/edit/save works; composition is not interrupted by rerenders; Persian ZWNJ input and mixed-script selection survive undo/export/reopen; unsupported content remains preserved. The extracted engine still builds and tests on its own.

## M6 — Broader compatibility, selected by corpus evidence

Prioritize gaps found in real files: font embedding/activation, floating drawings, more field updates/TOC, comments/revision editing, equations/charts and advanced tables/section layout. Preserve features before making them editable. Collaboration requires an operation/identity design review, not merely adding a transport to the preview.

“Fully compatible” remains a long-term verification program. Maintain a published matrix, producer/version coverage and known differences rather than declaring completion from one milestone.

## Test and release strategy

### Fixture corpus

Start with at least 40 small, independently understandable files: 12 typography/bidi/fonts, 8 styles/formatting, 8 structures, 6 sections/stories and 6 preservation/error cases. Add at least five multi-feature real documents from permitted sources. Include Word-authored Transitional files, Strict files and output from a second producer; record exact application versions instead of treating producers as interchangeable.

Each fixture manifest records provenance/license, producer/version, features, expected diagnostics, expected IR, required font files/hashes, expected export behavior and reference renders. Keep proprietary font binaries out of source control unless redistribution is authorized. Small redistributable fonts cover automated checks; locally licensed fonts cover additional manual comparisons.

### Validation layers

| Layer | What it establishes | Gate |
| --- | --- | --- |
| Unit/synthetic package tests | XML, relationships, property semantics and resource limits | Required on every engine change |
| Semantic corpus tests | Expected structure/text/style references/diagnostics | All supported profile cases pass |
| Round-trip tests | Supported edits survive; untouched/opaque content remains | No unreported loss; original-byte no-op export |
| External OOXML validation | Exported package/schema correctness | No validator errors for supported output |
| Word reopen check | Word accepts the exported file without repairing it | Required for each export feature and release corpus |
| Browser visual tests | Shaping/fonts/direction/selection/layout | Fixed browser/font references; reviewed differences |
| Browser editing tests | IME, graphemes, clipboard, history, accessibility | Chromium plus Firefox/WebKit where supported |
| Resource tests | Byte/count/depth budgets, cancellation, responsiveness | Deterministic limit failures; no main-thread hangs |
| Consumer tests | Built package works independently of app internals | Node import and Next production build |

Use a small .NET/Open XML SDK validation utility when writer work begins; it is test infrastructure, not a browser runtime dependency. Validation is necessary but does not prove Word layout or preservation. Export -> reimport comparisons alone are also insufficient: the parser and writer can share the same mistake.

Track import time, peak memory and interactive pauses for a fixed 1/20/100-page corpus on recorded hardware. Establish measured budgets before optimization; do not invent a throughput claim now. Record skipped manual/Word checks visibly in release evidence.

## Next engine implementation batch

For the next engine batch, start with OPC-01, OPC-02, IR-01 and PRES-01: content-type/encoding validation, related-part resolution, a diagnostic/session API and unchanged-file preservation. Then add a fixture containing style-based RTL and four font slots to drive M2. These remain prerequisites for safe edited exports. The next application batch is the preview ribbon shell described in the [UI plan](word-ui-implementation-plan.md); it can proceed before those editing/export gates are complete.
