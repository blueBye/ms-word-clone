# DOCX compatibility matrix

Baseline: 2026-10-31, updated after the four-document compatibility pass. “Partial” means there are known lossy or approximate behaviors. “No” means absent, not merely untested. Editing/export and unknown-source preservation remain absent. See [sample results](sample-compatibility.md) for measured evidence and remaining limits.

| Feature | Import / IR now | Preview now | Target milestone |
| --- | --- | --- | --- |
| ZIP and main-part discovery | Partial, limits and safe-name checks | N/A | M1 conformance/diagnostics |
| Transitional / Strict namespaces | Supported tested subset | Supported tested subset | M1 mixed/compatibility validation |
| Content types and general relationships | Source-relative related parts; manifest validation partial | N/A | M1 |
| UTF-8 XML | Yes | N/A | M1 encoding coverage |
| UTF-16 XML | No | N/A | M1 |
| Text/run order, whitespace, ZWNJ | Yes in supported paragraphs | Partial, browser shaping unverified | M2 |
| Empty paragraphs | Yes, including paragraph marks | Basic line/mark typography | M2 |
| Direct bold/italic/size | Separate Latin/CS channels | Basic supported | M2 |
| Paragraph/character styles, defaults | BasedOn/defaults/toggles/direct overrides | Supported subset | M2 |
| Theme fonts/colors | Basic theme resolution | Supported subset | M2 |
| Four font slots and language | Retained | Grapheme segments; full Word algorithm incomplete | M2 |
| Custom font files and variants | No asset model | Host must load them | M2 |
| Embedded DOCX fonts | No | No | M2 discovery/preserve; M6 activation |
| RTL/LTR paragraph direction | Direct and inherited | Browser bidi mapping, Word equivalence unverified | M2 |
| Logical start/end alignment | Retained | Logical CSS | M2 |
| `w:rtl`, `w:cs`, mixed-script runs | Supported subset | Forced CS plus script segments | M2 |
| Arabic/Hebrew/Persian punctuation/digits | Text retained unchanged | Script-aware fonts; visual comparison pending | M2 |
| Color, underline, strike, highlight | Basic supported | Basic supported | M2 |
| Superscript/subscript | Supported | CSS positioning | M2 |
| Paragraph spacing, indentation, tabs | Properties retained | Spacing/indents; advanced tab alignment approximate | M2/M3 |
| Line / page / column break distinction | Typed page/column breaks | Flow/print hints, no exact pagination | M2/M4 |
| Hyperlinks and bookmarks | Safe external targets; bookmark definitions incomplete | Hyperlinks | M3 |
| Inline/block content controls | Visible text traversal | Controls not editable/preserved | M1 preservation; M3 text projection |
| Markup compatibility / AlternateContent | One supported choice/fallback | Supported subset, not full MCE | M1 |
| Bullets and multilevel numbering | Basic counters, overrides and restart rules | Markers and indentation | M3 |
| Tables, widths, merges, RTL order | Basic nested tables and merges | Flow layout; floating/conditional cases partial | M3 |
| Inline images, dimensions, alt text | Embedded rasters/crops | Basic supported | M3 |
| Sections, page size, margins | Final-section page settings | Flow preview with source margins | M4 |
| Headers/footers/page numbers | Story variants and cached text | Once per flow, no repeated pagination/recalculation | M4 |
| Footnotes/endnotes | No | No | M4 |
| Fields and TOC | Result text may survive, field structure lost | Partial cached text | M1 preserve; M4 display; M6 updates |
| Revisions/comments | Partial inserted text; metadata lost | No review UI | M1 preserve/diagnose; M6 editing |
| Floating drawings/charts/equations/OLE | Basic anchored raster/shape/textbox nodes; other types absent | Anchors approximate | M1 preserve/diagnose; M6 selective support |
| Unknown part / unknown XML preservation | No | No placeholder | M1 |
| Editing / history / IME | No | No | M5 |
| DOCX writer / Word reopen validation | No | N/A | M2 first slice, M3/M4 expand |
| Exact Word pagination | No | No | M4 measured profile, ongoing work |

## Proposed first release profile

Common editable documents: English, Persian, Arabic and Hebrew body text; mixed-script runs; headings and common character formatting; user-provided licensed fonts; explicit LTR/RTL paragraphs; basic multilevel lists; rectangular tables with supported merges; inline raster images; hyperlinks/bookmarks; ordinary page/section settings; headers/footers and cached page fields. The writer must support the same profile as import before its editing controls ship.

Preserve and expose as locked/unsupported initially: floating drawings, charts, equations, embedded objects, tracked revisions, advanced content controls, complex fields, advanced table/layout cases and unknown extension markup. A locked feature can still affect surrounding layout; preservation is not a claim of visual accuracy. Fail or warn with source locations if safe preservation is unavailable.

Macro-enabled/encrypted containers and legacy `.doc` are separate formats and need explicit product decisions. This profile is a proposed release scope, not a claim that all features have been implemented.

## Evidence required for each row

Track import, rendering, editing, export and preservation separately. Each promoted feature needs a minimal fixture, expected semantic IR, diagnostics, authoring/export round trip, external schema validation and a Word reopen check. Visual features additionally need fixed-font browser/Word references. Unsupported features need an explicit diagnostic and preservation test instead of silent success.

Coverage should be reported as passing feature cases and weighted real-document corpus results, with producers, browser versions and fonts recorded. Do not report a percentage of “the DOCX format” based on line coverage or synthetic test count.
