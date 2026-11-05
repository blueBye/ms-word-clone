# Insert and Layout implementation status

Updated 2026-11-05. This records implemented behavior and remaining differences from Microsoft Word. The earlier inventory is a reference, not a claim of feature parity.

## Run and use

Run `pnpm dev` from the repository root, then open http://localhost:3000. If the development server is already running, refresh the page. `pnpm build` compiles the application; it does not launch it.

Open a DOCX, choose the sample, or create a blank document. Click the page to type. Select text across paragraphs and styles using the mouse. Enter carries the current typing format. The ruler changes first-line, hanging, start and end indents, including RTL paragraphs and lists. Double-click a page's top or bottom margin to edit its header or footer. Download DOCX saves an edited copy.

## Insert

| Controls                                  | Implemented behavior                                                                                             | Current boundary                                                                                                            |
| ----------------------------------------- | ---------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| Cover, Blank Page, Page Break             | Three cover designs, remove cover, paragraph/page breaks                                                         | Does not reproduce Word's full template catalogue                                                                           |
| Table                                     | Grid and dimension insertion, header row, text conversion, divider drawing, embedded worksheet                   | Draw Table creates rectangular row/column grids; arbitrary cell merging and Word's AutoFit variants are separate work       |
| Pictures, Screenshot                      | Local PNG/JPEG upload and browser screen capture                                                                 | Screen capture requires the user's browser chooser; no stock-image search service                                           |
| Shapes, Icons                             | Local preset gallery, native editable shape objects                                                              | Limited preset collection; no Office cloud catalogue or freehand drawing canvas                                             |
| SmartArt                                  | List, process and cycle diagrams with editable text, exported as grouped shapes                                  | No native SmartArt data/layout model or complete Office diagram catalogue                                                   |
| Chart                                     | Column, line and pie; edit category/series data; native chart XML and embedded XLSX                              | Other chart families, Office themes, chart formatting and complex imported charts need further work                         |
| 3D Models                                 | Self-contained GLB import, browser preview and native model relationship/part                                    | Experimental native Word rendering; no online model service, animation authoring or saved interactive camera poses          |
| Online Video                              | YouTube/Vimeo URL, click-to-load browser player, poster and DOCX video metadata/hyperlink                        | Native Word playback remains unverified; no local video embedding                                                           |
| Link, Bookmark, Cross-reference           | Editable hyperlinks, bookmark ranges, navigation/removal; references to bookmarked text                          | Cross-reference currently inserts a hyperlink with cached text, not a dynamic REF/PAGEREF field                             |
| Comment                                   | Text-range comments, author/date, edit/delete, comments pane                                                     | No threaded replies, resolved state or collaboration service                                                                |
| Header, Footer, Page Number               | Create/edit/remove stories, PAGE fields, first/even/default variants, double-click entry                         | Linked stories follow existing DOCX relationships; no dedicated Link to Previous control yet                                |
| Text Box, WordArt                         | Native text boxes and outlined display text, editable inner paragraphs                                           | Limited designs; no curved WordArt transformations                                                                          |
| Quick Parts                               | Title, author, date, PAGE and NUMPAGES fields; core document properties                                          | No complete building-block/AutoText catalogue                                                                               |
| Drop Cap                                  | Add, update and remove; grapheme-aware first character                                                           | Browser float layout approximates Word's frame metrics                                                                      |
| Signature Line                            | Editable visual signature block                                                                                  | No cryptographic signing, certificates or signing service                                                                   |
| Date and Time, Symbol                     | Date formats/locales/calendars, common symbols, Unicode character entry                                          | No exhaustive font glyph browser; inserted dates are static text                                                            |
| Object, Excel Spreadsheet, Text from File | Native embedded XLSX/Office packages, OLE containers for other files; worksheet grid editor; text-file insertion | Text from File accepts TXT/CSV/Markdown as plain paragraphs. No DOCX fragment merge or in-browser Excel formula calculation |
| Equation                                  | Editable basic linear/TeX input; native OMML fractions, roots, powers, subscripts and common symbols             | No complete TeX implementation, matrix editor or ink recognition                                                            |

Chart, equation and worksheet content can be reopened by double-clicking the object or choosing **Edit content** in the Selection pane. Text inside text boxes remains directly editable.

## Layout

| Controls                   | Implemented behavior                                                                                                                               | Current boundary                                                                                                                                                                                      |
| -------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Margins, Orientation, Size | Presets, custom dimensions/margins, mirrored margins and gutter; selected-section scope                                                            | Browser layout is not certified to match Word's printer/font metrics                                                                                                                                  |
| Layout Options             | Header/footer distances, first/even-page settings, vertical alignment, section direction and page-number restart                                   | Mirroring, hyphenation and even/odd-header settings follow OOXML's document-wide settings                                                                                                             |
| Columns                    | Equal-width columns, gap and separator; selected-section scope                                                                                     | Unequal widths and final-column balancing remain incomplete                                                                                                                                           |
| Breaks                     | Page/column/text-wrapping breaks and next/continuous/even/odd section breaks; physical parity pages and logical page-number restarts               | Continuous sections with changed geometry currently start a separate rendered page; identical-layout continuous sections flow together                                                                |
| Line Numbers               | Continuous/page/section modes, count interval/distance and paragraph suppression                                                                   | Browser counts use measured text lines, excluding tables; continuous counting across separate sections is not implemented                                                                             |
| Hyphenation                | None/automatic browser hyphenation and optional-hyphen insertion                                                                                   | Browser dictionaries differ from Word; no Word-style word-by-word manual hyphenation wizard                                                                                                           |
| Paragraph                  | Direction-aware indents, first/hanging indents, before/after spacing and paragraph dialog                                                          | Existing engine/renderer font and layout limits apply                                                                                                                                                 |
| Arrange                    | Selection pane, geometry, wrapping properties, position/alignment, distribution, z-order, rotation/flips, group/ungroup pictures/shapes/text boxes | Native DOCX anchors are saved, but browser floating-object text exclusion/positioning is approximate. Grouping charts, equations and embedded media is rejected to preserve their native object types |

## Architecture

- `src/core`: portable document data, styles, ranges, sections, validation, commands, history and d2ts-backed dependencies. It imports no React, DOM, Next.js or OOXML code.
- `src/docx`: source-preserving OOXML adapter, native drawing/chart/equation writers, annotations, properties, sections and OLE packaging. New parts and binary assets participate in undo/redo and export.
- `src/react`: document/object rendering, browser measurement and pagination. Three.js is loaded only for 3D previews.
- `apps/web/components/editor`: ribbon controls, authoring dialogs, ruler, clipboard, selection and editing adapters. The workspace coordinates commands with the DOCX session.

CFB handles the compound containers needed by Windows OLE packages. The existing OPC/ZIP layer remains responsible for DOCX/XLSX packaging. Worksheet formulas are saved for Excel to evaluate; they are not executed by the web editor.

## Validation

All 239 tests, the engine/UI TypeScript checks and the production build pass. The automated checks cover core independence, style inheritance, editing regressions, source preservation, native object/annotation round trips, section scope, cover removal, break behavior, OLE filenames, numbering and history. Run `pnpm test`, `pnpm typecheck` and `pnpm build`.

A generated document was opened in the installed Word without a repair prompt. Native text-box, chart, equation and table rendering were inspected. After correcting the OLE frame, an exported worksheet activated Excel inside Word. This is targeted native verification, not schema-wide or full Word compatibility certification.

The browser access host was unavailable during this pass. Interactive browser verification of the new controls remains outstanding; compilation and engine tests do not substitute for it. Video playback, 3D native rendering and the floating-layout/continuous-section limits above require further interoperability work.

Format references: [WordprocessingML document structure](https://learn.microsoft.com/en-us/office/open-xml/word/structure-of-a-wordprocessingml-document), [section properties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectionproperties), [Microsoft 3D model schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/fd676695-83fa-4e92-8b65-4503ecbff9cc), [Open XML SDK model relationship](https://github.com/dotnet/Open-XML-SDK), [CFB library API](https://github.com/SheetJS/js-cfb), [embedded-object format implementation reference](https://github.com/dolanmiu/docx/tree/master/src/embedded-files).
