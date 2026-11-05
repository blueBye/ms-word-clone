# Feature status

**Scope:** code and documentation included in the repository archive dated 2026-11-05. This is an **implementation inventory, not a compatibility certification**. “Implemented” describes supported operations, not pixel-perfect Microsoft Word behavior. Features can import, display, edit and export with different levels of fidelity.

**Status key:**

- **Implemented (subset):** working code and dedicated tests for ordinary cases; complex variants may be limited.
- **Partial / approximate:** some behavior or fidelity is missing, or only common cases work.
- **Experimental:** available in the UI or writer, but native interoperability or browser behavior needs further verification.
- **Not supported:** no complete implementation is present.

| Area | Current status | Important details |
| --- | --- | --- |
| New/open/download DOCX | **Implemented (subset)** | `createBlankDocx`, `openDocxSession`, `exportDocx`; no-op export returns original bytes |
| Text editing, IME, undo/redo | **Implemented (subset)** | Common paragraph/range edits and browser selection; complex cross-structure editing remains restricted |
| Styles, fonts, character formatting | **Implemented (subset)** | Inheritance, direct overrides, separate complex-script font channels; exact Word font selection and some effects incomplete |
| Persian/Arabic/English RTL/LTR text | **Partial / approximate** | Direction and script-aware segments supported; shaping depends on browser, fonts and renderer |
| Find/replace and lists | **Implemented (subset)** | Current search options, numbering and common templates; advanced wildcard search and arbitrary list formats absent |
| Tables | **Implemented (subset)** | Insert/edit/merge/split, borders, shading and sizing for common tables; complex merged grids and Word layout differ |
| Page setup, sections and columns | **Partial / approximate** | Size, orientation, margins, section actions and columns; mixed continuous sections and layout fidelity limited |
| Headers, footers, page fields | **Implemented (subset)** | First/even/default stories and basic PAGE/NUMPAGES support; full Word field behavior absent |
| Images and drawings | **Partial / approximate** | Raster images, selected shapes/icons, text boxes, positioning and corrections; arbitrary DrawingML and wrapping are not fully compatible |
| Links, bookmarks and comments | **Implemented (subset)** | Basic actions and text-range comments; not collaborative reviews or full cross-reference field updates |
| Charts | **Partial / approximate** | Column, line, pie authoring with embedded workbook; imported complex charts, themes and other types need work |
| Equations | **Partial / approximate** | Basic linear/TeX editing with selected OMML constructs; no full TeX/Word equation model |
| Embedded XLSX/Office objects | **Experimental** | Package insertion/editor support exists; no browser Excel calculation and limited native Word interoperability evidence |
| 3D models | **Experimental** | GLB import, browser preview, native relationship/part; native Word rendering not established |
| Online video | **Experimental** | Browser click-to-load YouTube/Vimeo and DOCX metadata; native playback not verified |
| Digital signatures | **Not supported** | Visual signature lines only; protected/signed source documents can be read-only |
| Collaboration, accounts, cloud autosave | **Not supported** | Local single-user workflow; no backend persistence service |
| Exact Word pagination | **Not supported** | Browser-measured pages, not Word's line/page breaking engine |
| Arbitrary DOCX import/edit/export parity | **Not supported** | Unsupported content is restricted or approximate; source-preservation paths are bounded |

## Preservation and safe editing

- An untouched editing session returns the **exact input file bytes**.
- The Word adapter preserves unmodified package parts and certain unknown/unsupported XML while writing supported changes. Structural operations still reconcile through the Word adapter.
- Retention is **not** proof that every unsupported construct will survive every possible edit. Some imported paragraphs, fields, controls and revisions are deliberately not editable.
- The format-independent `DocumentEngine` JSON snapshot does **not** include opaque original DOCX parts; use `openDocxSession` to retain the source envelope.

## Verification evidence

- Tests include core command/state behavior, style propagation, package handling, round-trip preservation, authoring, RTL regressions, rendering and adapter integration.
- [Insert/Layout status](ribbon-status.md#validation) describes targeted Word reopen and native-rendering checks performed during development; it also states which browser-level checks remained pending.
- [Engine architecture](engine-architecture.md#validation-and-measured-limits) includes one-run performance measurements. Those figures are not benchmarks guaranteed across machines or documents.
- The [sample compatibility report](sample-compatibility.md) covers a limited historical sample set; it cannot establish general interoperability.

For operational risks and unsupported cases, see [Known limitations](limitations.md). For command-level detail, see [Home editing status](home-editor-status.md) and [Insert/Layout status](ribbon-status.md).
