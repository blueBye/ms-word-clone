# DOCX engine review

For the current document-core separation, d2ts integration and source-preserving adapter lifecycle, see the [2026-11-05 architecture review](engine-architecture.md). The assessment below is historical.

Historical baseline review: the subsequent 0.2 sample-driven pass implements several findings below. See [current support](compatibility.md) and [four-document results](sample-compatibility.md) for the updated state; unresolved preservation/export and visual-layout limits remain tracked.

Review date: 2026-10-31. Scope: every source module, package exports, TypeScript configuration, dependency lockfile, README and existing tests. This is a source review and a synthetic-fixture assessment, not certification against the entire OOXML schema or Microsoft Word's layout engine.

## Assessment

The separation between `model`, `docx` and `react` is a useful foundation. The current engine is a text importer with a React preview. It is not yet a document editor, a layout engine, or a DOCX writer. Keeping that distinction visible will prevent the application UI from promising behavior that the engine cannot save.

Full compatibility needs separate evidence for: understanding document structure, displaying it, editing it, exporting it, and preserving features that are not editable. Browser appearance alone proves none of the other four. The immediate objective should be a tested common-feature profile plus explicit preservation and diagnostics for everything outside it.

DOCX combines packaging, document markup, compatibility rules and Transitional features. Reviewing only `document.xml` misses important dependencies. The standards baseline is [ECMA-376 Parts 1–4](https://ecma-international.org/publications-and-standards/standards/ecma-376/); Office behavior also needs comparison with [Microsoft's implementation notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376).

## Fixed in this change

| Issue | Evidence before the change | Resolution |
| --- | --- | --- |
| Build and runtime failure | `build` and `typecheck` reported TS2353/TS2740; 3 of the original 4 tests failed because the parser constructor rejected the error-handler object. | Use xmldom's `onError` callback and its own `Document`/`Element` types throughout the parser. No casts to browser `Document`. |
| Recovered/prohibited XML accepted inconsistently | Error handling depended on an obsolete constructor API; DTDs were not prohibited. | Stop on parser diagnostics, wrap failures as `Invalid XML`, reject document types. |
| Decompression before limits | `unzipSync(input)` ran before checking total output. | First scan all central-directory metadata with a rejecting filter; enforce bytes, entry count, safe names and duplicates before extraction. Retain an output-size check. |
| Ambiguous archive names and encoded targets | Slash/backslash normalization could alias entries; URI checks happened before decoding. | Reject unsafe stored names, duplicates and encoded unsafe targets; allow a single package-root slash in the relationship target. |
| Loose content-control traversal | Traversal entered the entire inline `sdt`; malformed stray runs outside its content could become visible. Ordinary `sdtPr` metadata was already skipped. | Traverse only its `sdtContent`, including on malformed input. |
| Very narrow test baseline | Four tests; original fixture had an empty content-type manifest. | Reusable synthetic OPC fixtures and XML, package, import and React-output checks. |

The installed dependency is xmldom 0.9.12. Its declarations explicitly reject the old object handler; its DOM is not the browser's HTML DOM. This matches the upstream [0.9 migration discussion](https://github.com/xmldom/xmldom/discussions/435). The preflight uses the existing [fflate filter API](https://github.com/101arrowz/fflate#usage), without adding a ZIP dependency.

The preflight is an improvement, not a complete hostile-file validator. Declared ZIP sizes, decompression CPU, CRC integrity, XML object overhead and malformed-archive edge cases still need stronger controls before public uploads.

## Remaining findings, ranked

### R1 — P1: import silently discards ordinary document content

`src/docx/parser/document.ts` selects only direct body paragraphs. A table containing `TABLE TEXT` imports as `{ type: 'document', blocks: [] }`. A block-level content control containing a paragraph also imports with zero blocks. Inline wrappers have limited traversal, but block wrappers do not.

This is a compatibility gap in the declared subset, rather than an accidental claim in the README. It nevertheless blocks use as an editor: a successful import cannot mean that all visible text survived. Introduce diagnostics and source preservation before allowing edits/saves of such files. Add typed tables and explicit wrapper traversal afterward; do not recursively collect every descendant paragraph because that loses table structure and can include metadata or duplicate alternate content.

### R2 — P1: no style/default/theme resolution

`import-docx.ts` reads only the main document. `paragraph.ts` and `run.ts` read only direct properties. `docDefaults`, default styles, `pStyle`, `rStyle`, `basedOn`, themes and numbering/table contributions are absent. A document using an RTL paragraph style and a custom-font character style loses both settings even though its text is retained.

Keep authored style references separate from computed properties. Implement hierarchy and cycle detection; do not treat all property merges as object spread. Some run properties have style-level toggle semantics, while direct formatting supplies explicit values. Paragraph-mark run properties must also remain distinct from formatting for every existing run. See [paragraph styles](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-create-and-add-a-paragraph-style-to-a-word-processing-document) and [bold toggle semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bold?view=openxml-3.0.1).

### R3 — P1: font information is lost in the IR

`run.ts` collapses `ascii` and `hAnsi` into one family, preferring `hAnsi`. `eastAsia`, theme slots, font hints and language are dropped. `DocumentRenderer.tsx` selects one font for an entire run if any character matches the Arabic/Persian regex.

Runtime probe: `English فارسی`, with base `Latin` and complex-script `Persian`, receives `fontFamily: Persian` for the entire span. Hebrew receives the base font because the regex excludes it. These outcomes cannot be fixed by uploading a font alone. Preserve font slots first, then produce render segments while retaining the original run and logical text order. The [run-font specification](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts?view=openxml-3.0.1) permits different font slots within one run.

### R4 — P1: complex-script formatting is confused with text detection

`w:cs` is ignored. An explicitly RTL whitespace-only run uses the base font because it contains no Arabic letters. The renderer also falls back from missing complex-script bold/italic to base bold/italic; those independent formatting channels need their own resolved defaults.

`w:rtl` has effects on complex-script formatting and weak/neutral character treatment. It is more specific than CSS direction. Fixing this requires model/resolver work, not a larger regex. See [right-to-left run semantics](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.righttolefttext?view=openxml-3.0.1).

### R5 — P1: paragraph direction and browser bidi context can diverge

Paragraph `start`/`end` alignment is converted to physical alignment before inherited direction is available. Unspecified paragraph direction is left to the embedding page. A preview mounted in an RTL application can therefore inherit a direction unrelated to the document's resolved setting.

Every explicitly directed run becomes a CSS isolate. That is an approximation which can change punctuation/numeric behavior across adjacent runs; no browser/Word comparison currently establishes equivalence. Resolve paragraph direction independently of UI language, retain logical alignment, and test paired punctuation, numbers, URLs, mixed runs, ZWNJ and directional marks. Text must stay in logical order; do not reverse strings. The reference for browser reordering is [Unicode UAX #9](https://www.unicode.org/reports/tr9/), with OOXML properties interpreted separately.

### R6 — P1: the text-only IR destroys structural distinctions

`TextRun.text` stores both soft line breaks and page/column breaks as `\n`; all tabs become `\t`. Probe: `A<w:br w:type="page"/>B` becomes `A\nB`. Break type, clear behavior, tab stops, paragraph marks, fields, hyperlinks and image positions cannot be recovered from that string.

Introduce explicit inline nodes as each feature is implemented: text, tab, break, hyperlink, image and field. Preserve original unsupported nodes meanwhile. Keep UTF-16 model offsets documented, and make editing operations respect grapheme boundaries.

### R7 — P1: round-trip preservation is impossible with the current return type

`importDocx()` returns only paragraphs and text. The original package, relationships, style references, unknown XML, content controls, revisions, comments, fields, media and section properties become unreachable. Adding a serializer to this IR would reconstruct a much smaller document and discard the rest.

Use a separate DOCX import session/envelope for original part bytes, source mappings, diagnostics and dirty parts. Keep the portable IR independent of those package details. An unchanged export should return the original bytes; edited exports must preserve untouched part bytes and unsupported XML in changed parts. Opaque content should be locked until editing around it has a tested preservation rule. Digital signatures need explicit invalidation handling when content changes.

### R8 — P2: OPC support is discovery, not package conformance

`package.ts` checks manifest existence, not its contents or the main content type. Relationship IDs, duplicate officeDocument relationships, invalid modes and subordinate relationships are not validated. `readText()` assumes UTF-8, so valid UTF-16 XML parts are not decoded correctly. A general source-relative part resolver is needed for styles, numbering, fonts, images and headers; legal `../` targets must resolve within the package rather than being universally rejected.

The hardened current main-part resolver remains deliberately conservative. Do not reuse it unchanged as a subordinate relationship resolver. OPC naming/encoding rules, encrypted Office containers, macro-enabled formats and unsupported content types need explicit acceptance/rejection policies.

Primitive helpers are also a supported subset, not schema validation: `halfPoints()` accepts a narrow integer spelling, while invalid on/off values silently become unspecified and uppercase spellings are accepted permissively. Validate the intended lexical types against Strict/Transitional requirements and report invalid authored properties instead of quietly treating them as absent.

### R9 — P2: resource and responsiveness budgets remain incomplete

The importer is asynchronous at its API boundary but ZIP inflation, DOM parsing and traversal are synchronous. Large files can freeze the browser. Package bytes do not bound DOM object count, nesting depth or recursive wrapper stack depth. There is no timeout/cancel protocol or per-XML budget.

Move import/export work behind a host-owned worker boundary; add XML depth/node/part limits and deterministic diagnostics. A worker improves responsiveness, but does not substitute for byte and CPU limits. Do not introduce workers into the portable model.

### R10 — P2: preview typography has visible layout gaps

An empty paragraph has no line box content and can collapse. Paragraph spacing is a fixed `0.5em`; browser tabs do not implement Word tab stops. Line spacing, keep rules, widow/orphan behavior, page size/margins, section breaks, headers/footers and paragraph mark typography are absent. Font family names are inserted as CSS family strings without a literal-name mapping, which also needs tests for commas/quotes and generic names.

These require layout fixtures, loaded font variants and measured comparisons. Do not use screenshot similarity as proof that the underlying DOCX structure was preserved.

### R11 — P2: no editing identity or command layer

The current model has no stable IDs, selection/position contract, operations, history or validation. Array indices are adequate preview keys but unsuitable identities for mutable blocks and runs. A raw `contentEditable` surface would bypass the IR and jeopardize saved formatting.

Choose a browser editing substrate through a focused Persian/Arabic/English IME and selection spike. Keep its adapter outside the parser and model. The editable schema, importer and writer must share a documented mapping; do not normalize imported DOCX through HTML as the preservation model.

### R12 — P2: interoperability evidence is still missing

There are now 81 passing tests, but they use synthetic packages or component output. There are no real Word-authored fixtures, schema-validation reports, browser screenshots, actual custom-font files, editing/IME tests, export tests or Word reopen checks. No claim of percentage DOCX compatibility follows from the test count.

## Module disposition

| Module | Keep | Next responsibility |
| --- | --- | --- |
| `src/model/` | Plain serializable types; no framework imports | Versioned semantic nodes, style references, logical alignment, font slots and IDs |
| `src/docx/package.ts` | One OPC boundary | Content types, encoding-aware XML reads, relationship resolution, package preservation |
| `src/docx/parser/xml.ts` | Shared namespace-aware helpers | Bounded parsing and diagnostics with source context |
| `src/docx/parser/{document,paragraph,run}.ts` | Small parsing functions | Block/inline coverage; authored properties only |
| `src/docx/import-docx.ts` | Import orchestration | Related-part discovery, resolver context, session/diagnostic result |
| `src/react/` | Separate entry point, no parser UI coupling | Resolved render model, font mapping, browser preview |
| `src/index.ts` | Headless public API | Add explicit APIs when implemented; keep Next.js out |
| Package/config | ESM/declarations and existing exports | Consumer smoke tests; optional React peer or later renderer package |

## Verification performed

After the fixes: `npm run build`, `npm run typecheck`, and `npm test` pass (81 tests, four files). Both built public package entry points import successfully in Node. `pnpm` is unavailable in this execution shell; npm invoked the existing scripts using already-installed locked dependencies. No dependencies were added.

Additional built-code probes confirmed loss of table/block-control content, page-break flattening, list-identity loss, whole-run mixed-script font choice, and incorrect font selection for RTL whitespace and Hebrew. These are remaining findings, not completed fixes. Browser layout, Word application behavior and real font loading were not tested.

See [the feature matrix](compatibility.md) and [the development plan](development-plan.md) for implementation order and release gates.
