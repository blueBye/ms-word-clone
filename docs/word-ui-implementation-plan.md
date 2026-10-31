# Word-like web UI implementation plan

Date: 2026-11-01. Status: **planned; no ribbon or editing implementation in this audit**. Based on the [installed Word inventory](word-ui-inventory.md), current source and [engine roadmap](development-plan.md).

## Outcome and scope

Build a familiar Word-like browser interface around the existing engine, with File backstage, Home, Insert and Layout, visible pages, useful side panes and a status bar. Retain the current sample rendering as a regression baseline. First deliver the interface around the working preview; then make editing features functional in complete import → edit → export slices.

The first common-document editor targets English/Persian/Arabic/Hebrew paragraphs, custom font variants, common marks/styles, lists, ordinary tables, links, inline raster pictures and normal section/header/footer settings. Preserve other content before offering edits around it. Charts, arbitrary shapes, equations, OLE, signatures, cloud collaboration and Office add-ins are separate milestones.

The inventory records what Word exposes. This plan decides what we implement and when. A visible button is not a commitment to immediate Word-equivalent functionality.

## Baseline and constraints

| Area | Current state | Consequence |
| --- | --- | --- |
| Portable engine | Root package `@local/docx-ir-engine`; public root and `/react` exports | Preserve package boundary; no Next imports in engine. No package move required for ribbon work. |
| Application | `apps/web/components/DocxWorkspace.tsx`: upload/sample, font loading/probes, diagnostics, paginated preview | Refactor incrementally when implementation begins; keep existing open/font lifecycle working. |
| Semantic model | IR v2, paragraphs/tables, text/drawings/breaks, direct/resolved properties, effective page setup/stories | Expand only for implemented feature slices. IDs/source positions/style definitions/sections are real prerequisites. |
| Layout | Hidden source measurement and repeated visible page renderings | Keep read-only preview working. Do not add `contentEditable` to every visible sheet or hidden source. |
| Fonts/bidi | Multiple font channels, mixed-script preview, regular uploaded font faces, local face probes | Editing needs variants, composition/selection and authoring rules beyond current rendering. |
| Save | No writer, source session, unknown-XML retention or edited-file validation | Save must not serialize preview HTML or rebuild a lossy DOCX. |
| Evidence | User-approved current samples; existing automated engine checks | This audit did not rerun builds/tests or establish universal layout parity. Add browser/editor tests when those behaviors are implemented. |

The original M1–M6 engine gates remain applicable. The P phases below order **UI delivery**, and do not mark the engine milestones complete.

## Proposed interface specification

### Window layout

| Region | Proposed contents and behavior |
| --- | --- |
| Title bar | App identity; filename; mode Preview/Editing; original/download/recovery state; Undo/Redo when supported; local command search later. No nonfunctional Share/AutoSave claim. |
| Tabs | File opens backstage; Home/Insert/Layout select ribbon. Start in Home. Selected tab remains stable on document changes; contextual groups appear with eligible selection. |
| Ribbon | Named groups in Word order; larger insertion/clipboard buttons; compact formatting rows; group separators/captions; dialog launcher at group's corner; split-button primary action distinct from dropdown. |
| Notices | Compact missing-font/compatibility summary below ribbon; detail pane on demand. No raw IR on the normal authoring path. |
| Workspace | Neutral background; centered source-sized pages with visible gaps/shadows; scroll container; optional ruler; panes reduce available space without changing page dimensions. |
| Navigation pane | Find results and page navigation first; headings/bookmarks later. Jump preserves semantic target after reflow. |
| Right pane | One active pane: Fonts, Compatibility, Styles, Selection or object properties. Each has name/close and remembers its own view state. |
| Status bar | Page x of y, content word count/language as available; font/layout status; zoom slider/value and Fit Width/Page. Mark provisional page count while fonts/layout are pending. |
| Debug tools | IR inspection and diagnostics source detail behind a developer/details entry, separate from main document workflow. |

Initial desktop dimensions are design targets, not Word measurements: title row about 40 px, tabs 36 px, ribbon 100–112 px, status 28 px, canvas gaps about 24 px. Use app CSS tokens so these can be tuned. Document content retains its own font sizes, line heights, margins and units.

At narrow widths, groups progressively collapse into labeled group menus in the same order; keep the ribbon horizontally scrollable as a fallback. On small screens, selected-tab actions use a compact toolbar/overflow and dialogs become sheets. Never shrink document content to make ribbon controls fit. Fit Width changes zoom, not the page's physical width.

Application locale and document direction are independent. English UI may display an RTL paragraph; Persian UI may edit an LTR paragraph. Use isolated direction for filenames, mixed labels and numeric fields. Page progression, section direction, paragraph direction and table order remain separate settings.

### Interaction contract

| Control type | Required behavior |
| --- | --- |
| Toggle | On/off/mixed; disabled reason; live state derived from selection. With a collapsed caret, set pending typing marks rather than altering unrelated text. |
| Split button | Primary applies current choice; arrow opens menu. Separate keyboard/accessible names and hit targets. |
| Menu/gallery | Arrow-key navigation, typeahead where useful, Escape/back, focus return; selected/mixed item; scrollable bounds; no hover-only access. |
| Font selector | Search document/available/uploaded families; requested and substituted faces distinguished; preserve theme reference; advanced Latin/CS controls. Font file loading is a separate action. |
| Size/unit input | Editable value, units, min/max validation and step; empty/mixed state; commit on Enter/blur, Escape restores value; convert once to model units. |
| Color | Shared theme/standard/recent/custom palette; Automatic versus No Color; keyboard labels; theme token and tint retained. Highlight is a distinct property. |
| Dialog | Draft values; Apply/OK produces one transaction; Cancel produces none; errors stay local; remember scope explicitly. A style dialog may reuse font/paragraph form sections. |
| Task pane | Nonmodal, keyboard reachable; does not lose document selection; resize/close; virtualize long lists only after measurement shows need. |
| Unavailable feature | Disabled with a concise reason in tooltip/help, available to keyboard/screen-reader discovery. Unsupported advanced/service features can be omitted from the primary ribbon but recorded in compatibility/help. Never open an empty fake tool. |

Selection survives toolbar/menu/dialog focus. Closing a tool restores the caret to the logical document position, including a header/footer/textbox story. Formatting previews, if later added, are transient and cannot accidentally commit on hover.

Use native button/input/dialog semantics and established accessible patterns. Test keyboard tab/arrow behavior, focus visibility, 200% browser zoom, high contrast and reduced motion. Shortcuts route through the same commands as buttons. Handle browser conflicts deliberately; no blanket prevention of Ctrl+S/F/P before the corresponding feature works.

## Structure and ownership

Add only directories whose features are being implemented. Proposed end-state organization:

```text
apps/web/
  components/
    workspace/       title, page workspace, status, notices and pane host
    ribbon/          tabs/groups/button/split button/menu/gallery
    dialogs/         file, font, paragraph, table, page and link forms
    panes/           fonts, compatibility, find, styles, selection
    editor/          browser adapter and selection/focus integration
  lib/
    commands.ts      small typed command table and UI metadata
    document-session.ts  app lifecycle, dirty/recovery/worker coordination
    fonts.ts         FontFace assets/availability and cleanup
    files.ts         picker/download and optional file handle adapter
src/
  model/             serializable nodes, positions, styles/sections as needed
  editing/           pure operations, transactions and position mapping
  docx/              package/source retention, parser, changed-part writer
  react/             preview; separate browser editor adapter if reusable
```

The app imports public package exports only. Core model/editing code uses no DOM, Next, storage or filesystem APIs. Browser adapter owns caret/selection/composition; host owns FontFace, file handles, object URLs, recovery and worker lifecycle. OOXML source archive stays in a runtime session beside serializable IR. Add public APIs alongside `importDocx()` rather than breaking existing consumers.

Use one authoritative semantic document. The editing adapter maps to it and transactions update it; preview is derived. Keep stable IDs and source maps in the engine session. If an editor library owns the runtime tree, establish one canonical mapping and explicit serialization boundary in the prototype; do not maintain two independently editable trees.

Start with one typed command table: ID, label, shortcut, control/menu metadata, state/value selector, capability predicate, disabled reason and execute handler. Handlers call portable operations or host services. Avoid a plugin registry, event bus, generic property schema or dependency-injection framework. Shared command families use parameters rather than one handler per gallery preset.

## Capability and saving rules

Track `import`, `preview`, `edit`, `export` and `preserve` independently for the feature and the active document. Enable an edit only if its affected source region is owned safely. A document-wide green state must not be inferred from a toolbar button.

Opening keeps original bytes. P0 can offer **Download original** with exact bytes; label this distinctly from saving edits. First edited export needs package retention, safe changed-node boundaries, relationship/content-type updates, opaque-node policy and writer validation. If a supported edit intersects unknown XML, block that edit/export with a useful explanation until preservation is implemented.

Local draft recovery is not a DOCX save. Once editing exists, show “Changes in this browser” versus “DOCX downloaded.” A download trigger is not proof the OS wrote a file successfully. Browser filesystem access is optional and must have a download fallback; it cannot be assumed on every browser.

## Phases and exit gates

| Phase | Deliverable | Dependencies | Exit gate | Relative scope |
| --- | --- | --- | --- | --- |
| **P0 — ribbon around preview** | Familiar shell; four tabs/backstage; Open/sample/original download; Fonts/Compatibility panes; zoom/fit/status; page navigation/basic Find | Current importer/renderer; original-byte retention in host | Existing sample visuals/page geometry preserved; working keyboard navigation and responsive ribbon; unavailable editing controls honestly disabled | Medium |
| **P1a — editor and preservation prototype** | Source session/IDs/positions; editing-substrate comparison; selection/IME/history; one plain-text/formatting edit with safe export | Engine M1 retention and first writer slice | Persian/English composition, grapheme deletion, selection and one edit/export/reopen pass; unknown content retained or explicit block | Large, highest uncertainty |
| **P1b — text and Home basics** | Text editing; clipboard supported subset; Latin/CS font variants/marks; paragraph direction/alignment/spacing; basic named style application; undo/redo; recovery | P1a; authored property/style model and writer | Bilingual formatted paragraph fixtures reopen in Word with no repair; inheritance/explicit off preserved | Large |
| **P2 — common structures** | Lists; table insert/edit/context tools; inline images; hyperlinks/bookmarks; symbols; supported borders/shading | P1; numbering/style definitions, relationship/media writer | Report/resume/form round trips with structure, list identity and source assets preserved | Large |
| **P3 — pages and stories** | Section model; size/margins/orientation/columns/breaks; header/footer editing; page fields; paragraph/table pagination rules; browser print | P2; engine M4 layout/model work | Mixed-section corpus, per-story fonts and page transitions checked against Word; no header overlap/empty-page regressions | Large |
| **P4 — advanced authoring** | Rich style management; advanced search/replace/sort; tab stops/leaders; supported floats/textboxes/arrange; selected fields/cross-references/notes/comments | Earlier feature ownership and writer coverage | Each promoted feature passes explicit import/edit/export/preservation and visual gates | Large, split per feature |
| **P5 — optional specialist/services** | Arbitrary drawing/effects, SmartArt/charts/OMML, OLE previews, signatures, cloud/share/collaboration, proofing/add-ins | Product decision plus dedicated models/services | Separate requirements and validation for each; unsupported originals remain preserved | Separate projects |

These are relative effort/risk categories, not calendar estimates. P0 can ship before editing. P1a must answer the editor architecture question before a large formatting UI is made interactive. No service work is required to deliver a local document editor.

## Command and engine work breakdown

IDs below are proposed **application IDs**, not Word automation IDs. `*` denotes a parameterized family. Every document mutation shares transaction/history/selection mapping, direct-versus-inherited formatting rules and the export/preservation gate.

### File and shell

| IDs | Phase | Behavior and dependencies | Acceptance |
| --- | --- | --- | --- |
| `file.open`, `file.sample` | P0 | Existing local import; retain original filename/bytes; cancel/stale-result guards; resource limit/error UI | Cancel keeps current document; two rapid opens cannot display stale results; malformed input reports failure. |
| `file.original.download` | P0 | Exact opened bytes, original MIME/name | Download bytes hash equals opened file; never rebuild from IR. |
| `file.new` | P1 | Construct supported semantic document with explicit page/default styles; separate from sample | Blank document editable and valid DOCX export; replacing dirty document invokes recovery/discard flow. |
| `file.save`, `file.saveAs` | P1 onward | Writer; download filename choice; optional filesystem handle; capability report | Changed text/formatting persists on Word reopen; unrelated/unknown content retained; unsafe edits blocked. |
| `file.info`, `pane.compatibility` | P0, expand P1 | Current diagnostics plus later structured source/editability/preservation metadata | Clearly distinguishes rendered, approximate, locked, omitted and export-blocking content. |
| `file.recover`, `file.close` | P1 | Versioned local session drafts/asset retention; dirty state; explicit close/discard | Crash/reload recovery restores a consistent revision; clearing/replacing uses deliberate user action. |
| `file.print`, `file.export.pdf` | P3 | Print stylesheet/font readiness; browser print dialog first | Physical source size and margins tested; ribbon/panes hidden; no duplicate measurement content; no claim of fixed-layout PDF parity. |
| `file.preferences`, `app.about` | P0 small, expand later | UI language, units, theme/reduced-motion preferences; version/compatibility | App defaults do not rewrite opened document; reset is explicit. |
| `view.zoom`, `view.fit*`, `navigate.page`, `pane.fonts` | P0 | View transform/scroll anchoring, measured page count; existing font management | Zoom changes apparent scale only; font load repaginates once ready; focus target remains visible. |
| `file.share`, `file.account`, `file.history.cloud` | P5 | Storage/auth/access/versioning service, separate from local lifecycle | Implement only with concrete service contract; do not show fictional presence/history. |

### Home

| IDs | Phase | Behavior, semantic target / OOXML area | Acceptance |
| --- | --- | --- | --- |
| `edit.insertText`, `edit.delete`, `edit.break`, `edit.undo/redo` | P1a/b | Logical positions, pending marks, split/join paragraphs, mapping, composition grouping; `w:p/r/t/br/tab` | ZWNJ/combining emoji/graphemes preserved; IME commit is not replayed twice; redo restores text and formatting/caret. |
| `clipboard.cut/copy/paste`, `clipboard.pasteText` | P1 | Plain text first; sanitized HTML subset later; cross-paragraph/list/table boundaries; browser events | Supported rich content converts predictably; text-only doesn't invent list identities; clipboard failure is recoverable. |
| `format.paint`, `format.clear`, `text.case` | P2/P4 | Snapshot defined formatting layer; clear direct overrides; locale-aware case transformation | Does not flatten named styles or transform Persian characters; one transaction per action. |
| `font.family/size`, `font.bold/italic` | P1 | ASCII/hAnsi/eastAsia/CS and theme references; size versus CS size, bold/italic versus CS variants; `w:rPr/rFonts` | Latin/CS selection combinations and inherited/explicit-off values survive export. Missing regular/bold/italic variants remain visible to user. |
| `font.underline/strike/script/color/highlight` | P1/P2 | Run marks; automatic/theme RGB/tint distinction; vertical position; `w:rPr` | Mixed selection state correct; split-run formatting preserves adjacent source and paragraph mark typography. Unsupported underline variants stay disabled. |
| `font.dialog`, `font.effects` | P2/P4/P5 | Reuse basic marks; hidden/caps/spacing/position first; kerning/OpenType/effects only with added IR/render/writer coverage | Preview uses correct script channel; Cancel leaves no undo entry; unsupported settings not silently dropped. |
| `paragraph.direction/align` | P1 | Paragraph bidi/alignment/logical indents; `w:pPr/bidi/jc/ind`; run RTL remains independent | RTL text, English phrase, URL, parentheses, digits and table markers retain ordering. No manual text reversal. Low/Medium/High justification stays gated pending shaping support. |
| `paragraph.indent/spacing`, `paragraph.dialog` | P1/P3 | Logical before/after, first-line/hanging, before/after spacing and auto/exact/at-least; keep/widow/page rules later | Single spacing uses font metrics; empty paragraphs/cell padding count; toolbar and dialog produce same properties. |
| `list.bullet/number/level/restart/continue/define` | P2; full definitions P4 | Editable abstract/instance/level definitions and overrides; `numbering` part plus paragraph references | Separate lists don't merge accidentally; Persian markers inside RTL cells; restart/export/reimport stable. Never save precomputed marker strings as list data. |
| `paragraph.tabs`, `paragraph.borders/shading` | P2/P4 | Stops/alignment/leaders, paragraph edges and table context; `w:tabs/pBdr/shd` | Decimal/right tabs and leaders measured; apply scope explicit; table borders use table command target. |
| `style.apply`, `style.create/modify/manage` | P1 basic; P4 management | Style identity/type/basedOn/next/defaults; separate authored/direct/effective layers; `styles` part | Editing shared style updates all dependents; unrelated direct overrides remain; inheritance cycles rejected. |
| `view.marks`, `select.all/objects/similar` | P1/P4 | View-only mark overlay; semantic text/cell/object selection; not source text changes | Marks absent from export/search counts; selection works across supported stories/structures. |
| `find.text`, `find.next`, `navigate.*` | P0 basic; P3/P4 advanced | Logical index with source positions; displayed results map to page fragments; selection editor integration later | Search is not duplicated by hidden/visible preview copies; mixed-script query hits correct source range. |
| `replace.one/all`, `find.advanced`, `paragraph.sort` | P2 basic replacement; P4 advanced | Scope, normalization-to-source mapping, format predicates, structural tokens; sort retains blocks/metadata | Replace All one undo action; preview/count before large changes; cross-run matches preserve unaffected marks; no fake Word-wildcard equivalence. |

### Insert

| IDs | Phase | Behavior, semantic target / OOXML area | Acceptance |
| --- | --- | --- | --- |
| `insert.pageBreak/blankPage/columnBreak` | P2/P3 | Typed breaks and paragraph split, section-aware layout; define blank-page structural behavior | Reopen retains break type; deleting break has predictable paragraph/section effects; no spurious trailing sheet. |
| `insert.table`, `table.*` | P2 | Grid/dialog with rows/columns/fixed-auto width; insert/delete row/column, merge/split, widths/height, padding/borders/shading/alignment; `w:tbl/tblPr/tr/tc` | Table remains rectangular and valid; nested content retained; merged-cell selection/logical order works RTL/LTR; no hidden cell-padding change. |
| `insert.image`, `image.size/crop/alt/opacity` | P2 | Browser picker; retained raster asset relationship/content type; DrawingML extent/transform | Original bytes retained, replace cleans relationships safely, aspect ratio/opacity/crop/export match; header/footer media uses owning story relationships. |
| `insert.link`, `insert.bookmark` | P2 | Inline link ranges/external relationship or internal anchor; stable named bookmark ranges | Edit/remove preserves display text; stale reference handled; no external fetch on import; safe click behavior. |
| `insert.symbol` | P2 | Actual Unicode sequence and selected font; special character menu | ZWJ/ZWNJ/nonbreaking characters preserved verbatim; font glyph squares are distinct from form checkbox controls and emoji presentation. |
| `insert.header/footer/pageNumber` | P3 | Section/story identity, variants/link-to-previous, supported field instruction/result | Editing header on one section/variant doesn't mutate all variants; correct fonts/page totals and no overlap. |
| `insert.dateTime` | P2 static; P4 field | Literal formatted text first; field instruction/cached value later; language/calendar explicitly chosen | Fixed date stays fixed; automatic field visibly identified; UI Persian does not silently switch calendar. |
| `insert.field/property/crossReference` | P4 | Field instruction/result, bookmarks/caption targets; document properties/relations | Broken references diagnosed; supported recalculation bounded; opaque field instructions retained. |
| `insert.textBox`, `insert.shape` | P4 limited; P5 arbitrary | Editable sub-story and supported geometry, anchors/wrap/transform; DrawingML/VML preservation | Unsupported source shapes locked/preserved; resizing/dragging commits once; text direction/font preserved. |
| `insert.cover/quickTable/autoText` | P4 | App-owned semantic templates/building blocks with IDs/style/relationship remapping | No duplicate bookmark/drawing/relationship IDs on repeated insertion; inherited defaults deliberate. |
| `insert.comment` | P4 | Anchored range, comments part/relationships/author metadata; separate comments pane | Anchor survives edits; export/reopen retains comment text/association; user chooses authorship settings. |
| `insert.chart/smartArt/equation/object/3d/video/signature` | P5 | Dedicated models/parts and preservation policy; external services where applicable | Preserve originals first; no execution of OLE/macro payloads; no arbitrary web embed mistaken for OOXML compatibility. |

Add contextual **Table Design/Table Layout**, **Picture Format**, and **Header & Footer** subsets alongside their first editing feature, despite their being outside the four audited tabs. In particular, cell margins, repeat heading rows, row split rules and vertical alignment cannot all be controlled from Insert Table.

### Layout

| IDs | Phase | Behavior, semantic target / OOXML area | Acceptance |
| --- | --- | --- | --- |
| `page.size/orientation/margins`, `page.setup` | P3 | Per-section `pgSz/pgMar`; gutter/mirror/scope; custom units | A4/Letter/custom sizes persist; landscape swaps dimensions deliberately; selected-section scope and RTL inside/outside correct. |
| `section.insert/type/direction` | P3 | Paragraph/final section boundaries, next/continuous/odd/even, section bidi, linked stories | Mixed page sizes/orientations and odd/even starts export/reopen; intentional parity blank pages distinguished from pagination bugs. |
| `page.columns` | P3 | Count, widths/gaps/equality/separator, RTL order; `sectPr/cols` | Unequal widths and RTL column reading sequence preserved; column break honors current section. |
| `page.lineNumbers/hyphenation/verticalAlign/borders` | P4 | Section/paragraph/document settings; line layout/shaping dependencies | Only supported line/word algorithms enabled; preserve imported settings even when approximate; compatibility switches not guessed. |
| `object.position/wrap/size/rotate` | P4 | Reference frames, offsets/alignment, text distances, inline versus anchor, overlap/lock, transform | Moving/resizing at zoom maps to physical units; reflow respects text wrap and story bounds; rotation not CSS-only export loss. |
| `object.forward/backward/align/distribute/group` | P4/P5 | Stable object identities, stacking, multi-object selection/group transforms | Correct eligibility; inline objects not falsely treated as floating; group export/preservation validated before enabling. |
| `pane.selection`, `view.guides/grid` | P4 | Object naming/selection/visibility/order; app view state versus document hidden state | Visibility changes have explicit persistence semantics; selection pane and canvas select same node. |

## Font and RTL acceptance programme

Run these against fixed known font files, plus the user's locally installed fonts without committing proprietary binaries. Imported run-font slots and theme references are distinct; Microsoft's [RunFonts reference](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.runfonts?view=openxml-3.0.1) is a starting point for the writer/resolver cases. The exact mapping to UI controls below is a proposed design.

1. Separate Latin/CS family, size, bold and italic; theme Body/Heading identity; regular/bold/italic/bold-italic asset availability; header/footer/textbox and list-marker fonts.
2. RTL paragraph with English acronym, URL, numbers, punctuation, brackets, ZWNJ/ZWJ and diacritics; LTR paragraph with a Persian phrase; Arabic and Hebrew cases; logical text unchanged through copy/export.
3. IME composition, Shift+arrow ranges, caret at bidi boundaries, backspace/delete over graphemes, paragraph split/join, list and cell transitions; expected logical/visual behavior recorded per browser.
4. RTL tables with internal numbered lists, hanging indents, merged cells, cell margins, empty paragraphs, exact/at-least/automatic line spacing, row splits and repeated headers.
5. Font load/failure/replacement while pane open; pagination waits for readiness; missing-face status remains honest; zoom doesn't change wrapping or page count at a stable layout.
6. Section/page changes, default/first/even headers and footers, page counters, terminal paragraphs/breaks; no incidental empty page or content behind header.

Do not normalize Persian letters/digits globally to simulate Word display. Search comparison normalization, list-number formatting, numeral-display preferences and actual source characters are separate concerns. Browser shaping/bidi should be used and verified; do not introduce a home-grown Unicode shaping engine.

## Editor prototype decision

Before choosing a large editor dependency, implement a disposable bilingual/table/font test document in the candidate substrate(s). Compare native editing adapter versus an established editor framework only where it materially reduces complexity. Record supported schema mapping, IME results, caret/selection, clipboard/history, source preservation and package size/maintenance cost.

Choose the smallest approach that passes those cases. The prototype must demonstrate one authoritative state and a safe writer round trip. Decide whether the first editor uses a flowing editing surface with print preview or true paginated editing. Paginated preview is already useful; paginated editing requires stable position-to-fragment mapping, caret overlays/DOM ranges and reflow behavior. Do not force that architectural decision with toolbar markup.

## Verification and release evidence

| Layer | Required checks |
| --- | --- |
| P0 shell | Keyboard/focus/ARIA, responsive groups and panes, file cancel/errors, zoom/fit, font readiness, no regression on approved sample pages. Engine headless build remains independent of Next. |
| Operations | Transaction invariants, stable IDs/position mapping, mixed marks, grapheme/composition/history, structural validity after table/list edits. |
| Preservation | No-op original export exact bytes; changed-part tests preserve unknown children/attributes and untouched part contents; ownership failures explicit. |
| OOXML writer | Relationship/content-type integrity; external Open XML validation; separate Word reopen check with no repair prompt. Reimport alone is insufficient. |
| Visual/editor | Recorded Word build, browser/version, source font hashes and page geometry; reference screenshots/page counts; reviewed tolerances; mixed-script caret and clipboard checks. |
| Runtime | Stale worker result/cancel; recoverable errors; font/object URL cleanup; local recovery failure/quota; larger-document responsiveness measured before budgets are promised. |

Each feature ticket owns its remaining Word-dialog inspection, minimal fixture, authored/resolved semantics, import/preview/edit/export/preserve states, tests and user-facing limitation. Promote support only when the whole selected slice passes. Private samples remain local unless redistribution is explicitly authorized.

## First implementation batch

1. Read `apps/web/AGENTS.md` and the installed Next.js guidance before app changes. Establish current build and sample screenshots/page counts.
2. Extract file/font lifecycle helpers and small workspace regions from `DocxWorkspace`; preserve behavior/public engine imports. Add original-byte retention for correctly labeled download.
3. Build the shared button/split/menu/dialog/pane primitives and small command table, then title/tab/ribbon/status shell. Home groups: Clipboard, Font, Paragraph, Styles, Editing. Insert/Layout groups follow the inventory.
4. Connect Open, Sample, Download Original, Fonts, Compatibility, zoom/fit and page navigation. Add basic logical-text Find with a preview target mapping. Show Preview mode; editing controls remain disabled with reasons.
5. Implement File backstage with local actions and document information. Keep service/specialist entries out of the primary flow. Keep raw IR under developer details.
6. Verify existing documents at original geometry/fonts, keyboard-only usage and narrow layouts. Deliver this shell for visual review, then start P1a's editing/preservation prototype.

This batch changes the interface while keeping the approved preview useful. It does not enable formatted editing or edited DOCX downloads before the engine can save them safely.
