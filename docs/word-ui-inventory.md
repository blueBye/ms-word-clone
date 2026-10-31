# Word interface inventory

Inspected locally on 2026-10-31–2026-11-01. Reference installation: **Microsoft Office LTSC Professional Plus 2024**, Word **Version 2610, Build 20527.20088 Click-to-Run**, Current Channel, as reported by File → Account. English interface with Persian authoring enabled. This describes this installation, not every Word edition or Microsoft 365 service entitlement.

Read the [implementation plan](word-ui-implementation-plan.md) for priorities, behavior, engine work and release gates. The [existing compatibility matrix](compatibility.md) remains the authority for current engine support.

## Evidence and scope

The four requested top-level surfaces were inspected in the running Windows application. Menus, galleries, task panes and important dialog tabs were opened without applying formatting, saving, printing, signing or sharing the source document. Account identities, recent-document lists, document contents and filesystem paths are omitted here. Private accessibility observations remain in ignored local cache files; they are not application assets or redistributable fixtures.

Evidence labels in the tables:

| Label | Meaning |
| --- | --- |
| **D** | Dialog or pane opened; controls listed below were read from accessibility output or screenshots. |
| **M** | Menu/gallery opened; its choices or categories were inspected. |
| **E** | Entry observed in the ribbon/navigation; destination or deeper settings were not inspected. |
| **Dynamic** | Contents depend on installed fonts, clipboard, document, printer, online catalogue, account or selection. |

These are inspection labels, **not implementation status**. A D label does not mean every nested picker was expanded. Unvisited destinations are listed explicitly at the end. Large dynamic catalogues are represented by their category and data contract rather than claiming to enumerate every possible value.

The ribbon changes with window width and selection. Dialog launchers, split-button arrows, gallery expansion and contextual tabs are separate interaction targets. A disabled command may become available after selecting text, a floating picture or multiple objects.

## Shared window and document chrome

| Surface | Observed controls | Web implication |
| --- | --- | --- |
| Title/Quick Access | AutoSave switch; Save; Undo with history; Redo/Repeat; toolbar customization; filename and storage/saved state | Distinguish local recovery, downloaded DOCX and cloud saving. No generic “Saved” claim based only on preview state. |
| Tab row | File, Home, Insert, Layout among other Word tabs; contextual Table Design/Table Layout and Picture Format | Build requested tabs first. Contextual table/picture tools are a subsequent necessary editing surface. |
| Global actions | Search, Share, account/profile, ribbon display controls | Command search can be local; sharing and account services are separate scope. |
| Workspace | Page sheets, inter-page space, margins, scrollbars; Selection and Styles panes | Keep page geometry and zoom separate from application typography and spacing. |
| Status bar | Current/total page, word count, language, accessibility status; view shortcuts; zoom percentage and slider | Page count comes from layout; words come from semantic content. Counts exclude hidden measurement copies. |

## File: backstage

### Navigation and pages

| Page/action | Evidence | Observed choices and fields | Planned web disposition |
| --- | --- | --- | --- |
| Back | E | Return to document | Close backstage and restore focus/scroll/selection. |
| Home | D, Dynamic | New/blank document, templates, Recent, Favorites, more templates/documents | Local start page; never infer permission to access arbitrary recent files. |
| New | D, Dynamic | Blank document; online template search; suggested Business, Cards, Flyers, Letters, Education, Resumes and Cover Letters, Holiday; template cards and pinning | Start with blank supported document and app-owned templates. No copied Office template assets. |
| Open | D, Dynamic | Recent documents/folders, favorites; OneDrive; This PC; Add a Place; Browse; Recover Unsaved Documents | Browser file picker/drop target and later local draft recovery. Cloud providers deferred. |
| Share | E | Navigation action and Info-page Share action | Destination not opened; future explicit cloud capability. |
| Info | D | Location, Upload/Share, copy path/local path, open location; Protect Document; Check for Issues; Manage Document; Properties | Compatibility, source metadata, fonts, preservation/editability and storage state. Browser cannot expose Word's arbitrary OS paths. |
| Save | E | Immediate save action | Original-byte download first; edited DOCX only after writer gate. Do not invoke desktop Save during inspection. |
| Save As | D, Dynamic | Recent/favorite folders, OneDrive, This PC, Add a Place, Browse; Copy Link unavailable for local document | Download a copy; optional filesystem handle support with fallback. OS Save As dialog not inspected. |
| Save as Foxit PDF | E | Third-party add-in command | Exclude from core product. |
| History | E | Disabled for inspected local document | Local history/recovery and cloud version history are different features. |
| Print | D, Dynamic | Print, copies; printer and properties; page scope/range; one/two-sided printing; collation; orientation; size; margins; pages per sheet; Page Setup; preview previous/next/current page, zoom/fit | Browser print dialogue owns printers/copies/duplex. App owns printable content and page geometry. No printing was triggered. |
| Export | D | Create PDF/XPS Document; Change File Type | DOCX and browser Print/PDF first; fixed PDF export requires its own fidelity gate. |
| Change File Type | M | DOCX, legacy DOC, ODT, DOTX, TXT, RTF, MHT/MHTML; another file type; Save As | Do not offer unsupported formats. Plain text, if added, explicitly discards formatting. |
| Close | E | Close document action | Close session; unsaved-change/recovery workflow once editing exists. |
| Account | D | User information; sign-out/switch account; privacy entry; background/theme; connected services; activation/product info; updates; add-ins; About Word/version | Product preferences/About locally; no imitation Microsoft login, activation or update controls. |
| Options | D | Categories and inspected controls below | Small app preferences surface; separate document settings from app defaults. |

### Info submenus and metadata

| Entry | Evidence | Details |
| --- | --- | --- |
| Protect Document | E | Security/protection entry recorded; protection submenu/settings not explored or changed. Preserve/diagnose relevant source restrictions before enabling editing. |
| Check for Issues | M | Inspect Document; Check Accessibility; Check Compatibility. Inspections/removal tools not executed. |
| Manage Document | M | Recover Unsaved Documents in this local-file context. |
| Properties | D | Size, pages, words, editing time; title, tags, comments; created/modified/printed dates; author/last modified by; related documents; Show All Properties. Values intentionally omitted. Properties menu and expanded property editor not opened. |

### Word Options

All categories were identified. The following pages were read; nested management dialogs, privacy/security settings and desktop configuration changes were not performed. Long pages list inspected controls, not a universal list for every language/version.

| Category | Evidence | Inspected settings |
| --- | --- | --- |
| General | D | Multi-display appearance/compatibility; mini toolbar on selection; live preview; update content while dragging; auto-collapse ribbon/search; ScreenTip style; author name/initials; Office background/theme; keep document page color; Privacy Settings entry; startup screen. |
| Display | D | White space between pages; highlights; document tooltips; show tabs, spaces, paragraph marks, hidden text, optional hyphens, anchors/all marks; print drawings/backgrounds/properties/hidden text; update fields/linked data before printing. |
| Proofing | D | AutoCorrect Options entry; ignore uppercase/numeric words/addresses; repeated-word flagging; accented French uppercase; main/custom dictionaries; French/Spanish modes; spelling/grammar as typed; confused words; grammar with spelling; readability statistics; writing-style/settings; recheck. Bottom exceptions and nested AutoCorrect/dictionary dialogs not expanded. |
| Save | D | Cloud AutoSave default; default format; AutoRecover interval/last recovered version/location; opening/saving backstage behavior; extra places; default computer/local/templates locations; server-draft/cache information; embed fonts; subset characters; omit common system fonts. Cache settings below the visible area not expanded. |
| Language | D | Independent display and authoring/proofing language lists; add/remove/prefer/reorder; installed/available proofing state; cloud language preference. No language pack installed. |
| Accessibility | D | Status-bar/inline accessibility status; sound/scheme and animation feedback; ScreenTips/shortcuts; startup screen; AI alt-text option; expand headings; reading-order arrow navigation entry. No AI or privacy setting changed. |
| Advanced: editing | D | Typing replaces selection; whole-word selection; drag/drop; Ctrl+click links; drawing canvas; smart paragraph selection/cursoring; overtype/Insert key; prompt/update style; Normal style for lists; track formatting; style update numbering behavior; click-and-type/default paragraph style; autocomplete. |
| Advanced: bidi | D/M | Logical/Visual cursor movement; Block/Continuous visual selection; automatic keyboard language; show directional controls, diacritics/color; document view RTL/LTR; numeral choices **Arabic, Hindi, Context, System**; month names; Font Substitution entry. These Word labels are not app locale names or permission to rewrite stored Unicode digits. |
| Advanced: clipboard | D | Paste defaults within/between documents, conflicting styles, other programs; default picture layout; preserve bullets/numbers in text-only paste; include control characters; Insert-key paste; Paste Options button; smart cut/paste settings entry. |
| Advanced: images/content | D | Pen interaction/shape gesture; image editing data/compression/resolution; chart property tracking; backgrounds, window wrapping, drawings/textboxes, controls, bookmarks, text/crop boundaries; field codes/shading; draft font/name/size; heading expansion. |
| Advanced: display/typography | D | Recent-document/folder counts; measurement units; style-area width; HTML pixels; tooltips; scrollbars/left scrollbar; vertical/right ruler; layout-oriented character positioning; live dragging/subpixel fonts; table insertion buttons; advanced typography, smart justification and hyphenation. |
| Advanced: print/save/general | D | Draft/background/reverse/XML/field-code printing; changed-field update; duplex front/back; A4/Letter scaling; tray; PostScript/forms; Normal-template prompt, backups, remote copy/background saving; form/linguistic data; conversion prompt, links on open, add-in errors; encrypted-file/capture-app entries; mailing address; File Locations/Web Options entries. No setting changed or external link followed. |
| Advanced: layout compatibility | D | Underline spacing/trailing spaces; table grid line height; hyphenation across pages/columns; SBCS/DBCS balancing; backslash/yen conversion; exact line-height centering; Shift+Return expansion; HTML paragraph auto spacing; extra line spacing at page edges; line-breaking rules; Word 2013 hyphenation rules; compatibility-options scope. Preserve source compatibility intent; no global web “fix” switch. |
| Customize Ribbon | D | Command category/list; tab/group tree and visibility; add/remove/reorder; new tab/group, rename; reset; import/export; keyboard shortcut customization entry. No customization applied. |
| Quick Access Toolbar | D | Command category/list; document/global scope; add/remove/reorder; AutoSave/Save/Undo/Redo; toolbar visibility/above-below position; labels; reset/import/export. |
| Add-ins | D | Active/inactive/document-related/disabled lists; name/location/type/publisher/compatibility/description; Manage selector/Go. No add-in installed, enabled or executed. |
| Trust Center | E | Category label only. Security/privacy configuration is outside the web UI replication scope. |

## Home

### Clipboard

| Command | Evidence | Options/behavior to represent |
| --- | --- | --- |
| Paste split button | M, Dynamic | Primary last/default paste; menu offered Keep Text Only, Paste Special, Set Default Paste for current clipboard. Other modes depend on source data. |
| Cut, Copy | E | Text, cell or object selection; browser gesture/clipboard limitations. |
| Format Painter | E | Copy formatting; destination selection; cancel mode. Persistent/double-click mode needs later interaction design. |
| Clipboard launcher | E | Office Clipboard pane entry; pane contents not inspected. Browser app should not promise global Office clipboard history. |

Keep Source Formatting, Merge Formatting and Use Destination Styles were also seen as defaults in Advanced Options. A future paste menu must advertise only conversions the app actually supports. Paste Special's data-format dialog was not successfully inspected.

### Font

| Command/menu | Evidence | Choices |
| --- | --- | --- |
| Font family/size | E, Dynamic | Editable family and size selectors; grow/shrink. Installed fonts and source theme labels are data, not a fixed font list. |
| Change Case | M | Sentence case; lowercase; UPPERCASE; Capitalize Each Word; tOGGLE cASE. |
| Clear All Formatting | E | Remove direct formatting while respecting intended style behavior; define supported semantics before enabling. |
| Bold, Italic | E | Toggle/mixed state; Latin and complex-script channels can differ. |
| Underline split button | M | Single, double, thick, dotted, dashed, dot-dash, dot-dot-dash, wave; None; More Underlines; Underline Color. |
| Strike, Subscript, Superscript | E | Formatting toggles; mutually exclusive vertical positions. Double strike is in Font dialog. |
| Text Effects/Typography | M | Presets; Outline, Shadow, Reflection, Glow; Number Styles; Ligatures; Stylistic Sets. Some font-dependent choices disabled. |
| Highlight split button | M | Yellow, Bright Green, Turquoise, Pink, Blue, Red, Dark Blue, Teal, Green, Violet, Dark Red, Dark Yellow, Gray 50%, Gray 25%, Black; No Color; Stop Highlighting; high-contrast-only filter. |
| Font Color split button | M | Automatic; theme base colors/tints/shades; standard colors; More Colors; Gradient; high-contrast-only filter. |
| More Colors | D | Standard swatch spectrum/grays and Custom tab; spectrum/luminance control; color-model selector, RGB inputs and Hex; Current/New preview; OK/Cancel. Other model values not expanded. |
| Font launcher | D | Font and Advanced tabs below; Preview, Set as Default, Text Effects, OK/Cancel. |

**Font tab:** separate Latin text and Complex scripts family/style/size. All text color, underline style/color; strike/double strike, super/subscript, small/all caps, hidden.

**Advanced tab:** scale percentage; spacing/amount; position/amount; kerning threshold; OpenType ligatures, number spacing/forms, stylistic sets and contextual alternates. These require font capability checks; a generic CSS setting is not proof of Word-equivalent output.

The shared standard-color row contains Dark Red, Red, Orange, Yellow, Light Green, Green, Light Blue, Blue, Dark Blue and Purple. Theme choices must retain their theme reference where appropriate, instead of eagerly replacing everything with RGB.

### Paragraph

| Command/menu | Evidence | Choices |
| --- | --- | --- |
| Bullets | M | None; solid/hollow circles, square, diamonds, arrow, checkmark, picture/document bullets; Change List Level; Define New Bullet. Document-specific glyphs/fonts are dynamic. |
| Define New Bullet | D | Symbol, Picture, Font; alignment; preview; OK/Cancel. Nested pickers not opened here. |
| Numbering | M | None; decimal/punctuation variants; upper/lower Roman and Latin; document formats; Change List Level; Define New Number Format. |
| Multilevel List | M | Current List; List Library; decimal/letter/Roman/bullet patterns and heading-linked outlines; gallery filter; Change List Level; Define New Multilevel List; Define New List Style. |
| Decrease/Increase Indent | E | Distinguish paragraph indent from changing an actual list level. |
| LTR/RTL paragraph | E | Two explicit direction commands, independent of alignment and interface locale. |
| Sort | D | Paragraphs/type, ascending/descending; up to two Then-by criteria; header/no-header; Options; OK/Cancel. Disabled criteria depend on selection; Options subdialog not opened. |
| Show All | E | Formatting-mark visibility, never literal insertion of marks. |
| Alignment | M/E | Left, Center, Right; Justify menu: **Justify, Justify Low, Justify Medium, Justify High**. |
| Line/Paragraph Spacing | M | 1.0, 1.15, 1.5, 2.0, 2.5, 3.0; Line Spacing Options; Add Space Before/After Paragraph in inspected context. Add/remove wording depends on current spacing. Detailed numeric rules in Paragraph dialog. |
| Shading | M | Theme/tint/shade and standard palette; No Color; More Colors; high-contrast-only filter. |
| Borders | M | Bottom/top/left/right; none/all/outside/inside; inside horizontal/vertical; diagonal down/up; Horizontal Line; Draw Table; View Gridlines; Borders and Shading. Last three depend on selection/feature scope. |
| Paragraph launcher | D | Indents and Spacing; Line and Page Breaks; Tabs; Set as Default; preview; OK/Cancel. |

**Indents and Spacing:** alignment, outline level, collapsed by default; RTL/LTR direction; Before/After text indents; special first-line/hanging/amount; mirror indents; before/after spacing; line-spacing rule and amount; don't add spacing between same-style paragraphs.

**Line and Page Breaks:** widow/orphan control; keep with next; keep lines together; page break before; suppress line numbers; don't hyphenate; textbox tight wrapping (disabled in inspected context).

**Tabs dialog:** positions/default interval; Right, Center, Left, Decimal, Bar alignment; None/dot/dash/underline leaders; Set, Clear, Clear All.

List-marker font, marker alignment, hanging indent, paragraph direction, numbering locale/format and list-instance restart are distinct data. The web's RTL number placement must not be implemented by swapping arbitrary CSS margins.

### Styles

| Surface | Evidence | Choices |
| --- | --- | --- |
| Quick Styles gallery | M/E, Dynamic | Normal, No Spacing, Heading 1/2, Title, Subtitle, Subtle Emphasis, Emphasis, Intense Emphasis, Strong, Quote, Intense Quote, Subtle Reference, Intense Reference, Book Title, List Paragraph; expansion/scrolling. Document styles can differ. |
| Styles pane | D | Style list, Clear All, Show Preview, Disable Linked Styles; New Style, Style Inspector/management icon entries, Options. |
| Create New Style from Formatting | D | Name; style type; based-on style; following-paragraph style; family/size/bold/italic/underline/color/script selector; paragraph formatting; preview/description; add to gallery; auto-update; current-document/template scope. Cancelled. |
| Style Format menu | M | Font, Paragraph, Tabs, Border, Language, Frame, Numbering, Shortcut key, Text Effects. |
| Manage Styles | D | Edit, Recommend, Restrict, Set Defaults tab labels; sort/recommended-only; style list/preview; Modify/Delete/New; document/template scope; Import/Export. Restrict tab not opened. |
| Recommend tab | D | Select All/Built-in; priority up/down/last/value; Show/Hide Until Used/Hide. |
| Set Defaults tab | D | Separate Latin/CS family/size, color; paragraph left/right/special indent/alignment; before/after and line spacing. |

Style Inspector, Styles Options, Modify Style and template organizer were not expanded. A style is not just a set of resolved toolbar values: identity and inheritance must survive edits/export.

### Editing and Add-ins

| Command | Evidence | Choices |
| --- | --- | --- |
| Find | M | Find; Advanced Find; Go To. |
| Find/Replace | D | Find, Replace, Go To tabs; query/replacement; Find Next, Replace, Replace All; More/Less; search direction/scope selector; Format/Special; No Formatting. |
| Advanced match options | D | Case; whole words; wildcards; sounds-like/all word forms (English); kashida; diacritics; alef hamza; control characters; prefix/suffix; ignore punctuation/whitespace. |
| Format search menu | M | Font, Paragraph, Tabs, Language, Frame, Style, Highlight. |
| Special search menu | M | Paragraph Mark, Tab, Any Character/Digit/Letter, Caret, section/paragraph characters, Column Break, Em/En dash, Endnote/Footnote Mark, Field, Graphic, Manual Line/Page Break, Nonbreaking Hyphen/Space, Optional Hyphen, Section Break, White Space, RTL/LTR Mark, ZWJ/ZWNJ. |
| Go To | D | Page, Section, Line, Bookmark, Comment, Footnote, Endnote, Field, Table, Graphic, Equation, Object; number/name/relative offset; Previous/Next/Close. |
| Select | M | Select All; Select Objects; Select All Text With Similar Formatting (No Data in inspected context); Selection Pane. |
| Add-ins | E, Dynamic | Add-ins entry; service/store destination not opened. |

Advanced Find must operate over logical document text, with a mapping back to positions. Optional matching normalization must not mutate the underlying Persian/Arabic text. Word wildcards are not equivalent to JavaScript regular expressions.

## Insert

| Group / command | Evidence | Menu, gallery or dialog inventory |
| --- | --- | --- |
| Pages / Cover Page | M | Built-in gallery; More Cover Pages from Office.com; Remove Current; Save Selection (disabled). Examples: Austin, Banded, Facet, Filigree, Grid, Integral, Ion Dark/Light, Motion, Retrospect, Semaphore, Sideline, Slice Dark/Light, ViewMaster, Whisp. |
| Pages / Blank Page, Page Break | E | Separate commands. A blank page has document structure; an explicit page break has position/type. |
| Tables / Table | M | Hover grid (7×6 visible at inspected width); Insert Table; Draw Table; Convert Text to Table (disabled); Excel Spreadsheet; Quick Tables. Grid size is responsive, not a file-format limit. |
| Insert Table | D | Columns/rows; fixed column width/Auto; AutoFit contents/window; remember dimensions; OK/Cancel. |
| Pictures | M | This Device; Stock Images; Online Pictures. |
| Shapes | M | Recently Used; Lines; Rectangles; Basic Shapes; Block Arrows; Equation Shapes; Flowchart; Stars and Banners; Callouts; New Drawing Canvas. Presets are shape geometry, not separate core commands. |
| Icons | E, Dynamic | Catalogue entry; destination not opened. |
| 3D Models | M | This Device; Stock 3D Models. |
| SmartArt | D, Dynamic | All, List, Process, Cycle, Hierarchy, Relationship, Matrix, Pyramid, Picture, Timeline, Meet The Team, Text Card, Office.com; layout gallery, preview/description, OK/Cancel. |
| Chart | D | Recent/Templates; Column, Line, Pie, Bar, Area, XY Scatter, Map, Stock, Surface, Radar, Treemap, Sunburst, Histogram, Box & Whisker, Waterfall, Funnel, Combo. Column variants shown; other subtype galleries not expanded. Cancelled before creating chart/Excel data. |
| Screenshot | M, Dynamic | Available Windows gallery; Screen Clipping. No capture inserted. |
| Media / Online Videos | E | Destination not opened. |
| Links / Link | D | Display text; ScreenTip; Existing File/Web, Place in This Document, Create New Document, Email; address; bookmark; target frame; browse/folder/recent controls; OK/Cancel. Only initial destination page inspected. |
| Bookmark | D | Name; Add/Delete/Go To; sort Name/Location; Hidden bookmarks; Cancel. |
| Cross-reference | D/M | Reference type, result form, target list; Insert as hyperlink; Include above/below; separator; Insert/Cancel. Visible types included numbered items, headings, bookmarks, footnotes/endnotes; remaining caption labels not fully scrolled. |
| Comments / Comment | E | Inserts a comment; not invoked on source document. Comment anchors, authorship and comments-part writer required. |
| Header | M, Dynamic | Built-in gallery (including Blank/Three Columns and themed odd/even variants); More from Office.com; Edit Header; Remove Header; Save Selection. |
| Footer | M, Dynamic | Corresponding gallery; Edit Footer; Remove Footer; Save Selection. |
| Page Number | M | Top/Bottom/Page Margins/Current Position; Format Page Numbers; Remove Page Numbers. Placement subgalleries not expanded. |
| Format Page Numbers | D | Number format; include chapter number; heading style/separator; continue previous section or start at; OK/Cancel. |
| Text Box | M, Dynamic | Quote/sidebar/template gallery; More from Office.com; Draw Text Box; Save Selection. Imported textboxes need their own editable story. |
| Quick Parts | M | AutoText; Document Property; Field; Building Blocks Organizer; Save Selection. Nested dialogs not expanded. |
| WordArt | M | Fifteen visible fill/outline/shadow/reflection-style presets in inspected gallery. Preset appearance is not support for arbitrary text transforms. |
| Drop Cap | E | Disabled in inspected table context; destination not inspected. |
| Signature Line | M | Microsoft Office Signature Line; Add Signature Services. No signing/setup action invoked. |
| Date & Time | D | Available formats; language (Persian Iran in context); calendar type (Gregorian in context); Update automatically; Set as Default; OK/Cancel. Other calendar choices not expanded; do not infer Jalali support from UI language. |
| Object menu | M | Object; Text from File. |
| Object dialog | D, Dynamic | Create New: installed object-type list, Display as icon, result. Create from File: filename/Browse, Link to file, Display as icon. Never instantiate or execute embedded content. |
| Equation | M | Built-ins for area of circle, binomial/expansion, Fourier, Pythagorean, quadratic, Taylor, trigonometric identities; More Office.com; New Equation; Ink Equation; Save Selection. |
| Symbol | M | Recent-symbol gallery; More Symbols. |
| Symbol dialog | D, Dynamic | Font, subset, glyph grid, recent glyphs, Unicode name/code, encoding source; AutoCorrect/Shortcut Key; Alt+X hint; Insert/Cancel. |
| Special Characters tab | D | Em/En dash; nonbreaking/optional hyphen; Em/En/quarter-em/nonbreaking spaces; copyright/registered/trademark; section/paragraph/ellipsis; opening/closing quotes; no-width optional/non-break; LTR/RTL marks, embeddings and overrides; Pop Directional Formatting. Shortcuts displayed. |
| eSignature fields | E | Separate service entry. No signing service opened. |

Gallery names are useful for grouping and interaction patterns. Recreating every Office template, chart style, SmartArt layout, stock image or equation example is not part of the common-document editing milestone. The engine must first preserve unsupported source content.

## Layout

### Page Setup

| Command | Evidence | Choices and fields |
| --- | --- | --- |
| Margins | M | Normal (all 1 inch); Narrow (all 0.5); Moderate (top/bottom 1, left/right 0.75); Wide (top/bottom 1, left/right 2); Mirrored (top/bottom 1, inside 1.25/outside 1); Office 2003 Default (top/bottom 1, left/right 1.25); Custom Margins. |
| Orientation | M | Portrait, Landscape. |
| Size | M, Dynamic | Letter, Legal, A4, Executive, 8.5×13 custom, A5, A6, JIS B5, 16K variants, Japanese postcard variants, envelope variants; More Paper Sizes. Printer-dependent list. |
| Columns | M | One, Two, Three, Left, Right; More Columns. |
| More Columns | D | Presets; count; RTL column order; line between; per-column widths/spaces; equal width; Apply to; Start new column; preview. Some controls disabled for one column. |
| Breaks | M | Page, Column, Text Wrapping; section Next Page, Continuous, Even Page, Odd Page. |
| Line Numbers | M | None, Continuous, Restart Each Page, Restart Each Section, Suppress for Current Paragraph; Line Numbering Options. Options destination not opened. |
| Hyphenation | M | None, Automatic, Manual, Hyphenation Options. |
| Hyphenation Options | D | Automatically hyphenate; words in CAPS; zone (disabled in context); consecutive-hyphen limit; Manual; OK/Cancel. |
| Page Setup launcher / Custom Margins | D | Margins, Paper, Layout tabs; preview, Apply to, Set as Default, OK/Cancel. |

**Page Setup → Margins:** top/bottom/left/right/gutter, gutter position; portrait/landscape; multiple-page layout; scope. Multiple-page alternatives beyond current Normal were not expanded.

**Page Setup → Paper:** size and custom width/height; first/other page paper source; printer trays; Print Options. The inspected assignment is **Letter, 8.5×11 inches**. Other samples use different geometry; never turn Letter or A4 into a global assumption.

**Page Setup → Layout:** section start, suppress endnotes, **section direction**; different first/odd/even stories; header/footer distance; vertical alignment; scope; Line Numbers and Borders entries. A section can be LTR while its paragraphs are RTL.

### Paragraph and Arrange

| Command | Evidence | Choices and fields |
| --- | --- | --- |
| Paragraph indent | E | Before/After text spinners; semantics follow paragraph direction, regardless of UIA's physical Left/Right names. |
| Paragraph spacing | E | Before/After spinners in points; same underlying properties as Home's dialog. |
| Paragraph launcher | D | Shared Paragraph dialog, not a second model implementation. |
| Position | M | In Line with Text; nine wrapped placement presets; More Layout Options. |
| Wrap Text | M | In Line, Square, Tight, Through, Top and Bottom, Behind, In Front; Edit Wrap Points; Move with Text; Fix Position on Page; More Layout Options; Set as Default Layout. Several anchor controls disabled for inline picture. |
| Bring Forward / Send Backward | E | Split controls; disabled for selected inline picture. Floating-object submenus not inspected. |
| Selection Pane | D | Named object list, Show All/Hide All, ordering buttons, pane close/options. List changed with page/context; selection activated Picture Format. |
| Align | M | Left/Center/Right/Top/Middle/Bottom; distribute horizontal/vertical; align to Page/Margin/Selected Objects; alignment guides; gridlines; Grid Settings. Alignment/distribution disabled without eligible objects. |
| Group | E | Disabled for current selection; group/ungroup/regroup semantics need multiple compatible floating objects. Submenu not inspected. |
| Rotate | M | Right/Left 90°; Flip Vertical/Horizontal; More Rotation Options. |

**More Layout Options → Position:** horizontal alignment/book layout/absolute/relative position and reference frame; vertical alignment/absolute/relative position and reference frame; Move with text, Lock anchor, Allow overlap, Layout in table cell. Disabled for inspected inline picture; labels still visible.

**Text Wrapping tab:** seven wrap styles; Both sides/Left only/Right only/Largest only; top/bottom/left/right distance from text.

**Size tab:** absolute/relative height/width and reference; rotation; height/width scale; Lock aspect ratio; Relative to original picture size; original dimensions; Reset. All inspected and cancelled without resizing.

## Inspection limits and follow-up destinations

The top-level command structure is covered; **this is not an exhaustive traversal of every recursive dialog or dynamic catalogue**. Before implementing the following advanced destinations, inspect them in a disposable document with the relevant selection and capture their exact control values:

| Surface | Remaining inspection |
| --- | --- |
| File | Share destination; protection menu; expanded/all/advanced properties; final Save As format/options dialogue; print setting dropdown domains; PDF publishing options; About Word details beyond Account version; nested Options dialogs; Trust Center labels/settings beyond category. Save/Close are intentional actions, not menus to execute on the source. |
| Home | Office Clipboard pane; Paste Special formats; font/size dropdown domains; deeper effects/gradient/underline-color pickers; Define New Number Format/Multilevel List/List Style; restart/continue/set numbering value context commands; Sort Options; Borders and Shading; Style Inspector/Options/Modify/Restrict/organizer; Add-ins destination. |
| Insert | Quick Tables; Convert Text to Table; icon/stock/online catalogues/video; remaining hyperlink destination pages; complete cross-reference types/result variants; page-number placement galleries; Quick Parts property/field/organizer dialogs; Drop Cap outside a table; signature setup; expanded calendars; equation authoring tools/Ink; Symbol AutoCorrect/Shortcut dialogs. |
| Layout | Multiple pages/book fold/mirror variants; Apply-to domains; line-number detail dialogue; Grid Settings; floating-object z-order menus; Group for multiple floating objects; Manual Hyphenation. |

This follow-up is part of each feature's implementation ticket, not a reason to delay the first usable ribbon shell. Dynamic catalogues and external services are separate scope; they cannot be exhaustively enumerated from one local file.

## Current engine implications

The code exposes a preview IR, not a Word command model. `DocumentIR` has no stable node IDs, selection/history, editable style/numbering definitions, source-preservation session or writer. `page` describes one effective page setup, not all section transitions. Header/footer variants and basic page fields exist; they do not constitute a complete story/field editor.

The current renderer already displays much of the common formatting, tables, drawings and bidi content above. **Import/preview support does not imply edit/export support.** In particular, editing a paragraph's resolved properties would flatten style inheritance; editing a rendered list marker would lose list identity; editing duplicated paginated DOM would create competing document copies.

OOXML keeps main text, styles, comments and other stories in related package parts. That informs the planned separation between editor session and serializable IR. See Microsoft's [document structure reference](https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document). Section settings occur at section boundaries as well as at the end of the body; this requires a richer section model before multi-section controls are enabled. See [SectionProperties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectionproperties?view=openxml-3.0.1).
