# Compatibility work driven by the four supplied samples

The files are read from their original locations. Personal document text, images and full DOCX packages are not added to tracked tests or public app assets. Synthetic regression fixtures reproduce the relevant markup with neutral text.

## Main causes and implemented changes

| Sample group | Cause of incorrect output | Added support |
| --- | --- | --- |
| Both résumés | Nearly all text is inside a layout table; the previous importer only kept direct body paragraphs. | Table blocks, columns, merged header cells, nested paragraph traversal and recursive renderer. |
| Persian reports | RTL table order and default paragraph bidi were ignored; merged cells, borders and table styles were absent. | Independent table/paragraph direction, document defaults, style inheritance, grid spans, row spans, widths, shading, borders and margins. |
| All four | List metadata, inherited typography, colors and spacing were lost. | Numbering definitions/instances, counters/restarts, paragraph/character/table style chains, theme references and direct overrides. |
| All four | Font channels were collapsed and a single Arabic character switched an entire mixed run's font. | Separate ASCII/high-ANSI/East Asian/CS slots, forced CS/RTL handling and grapheme-safe render segments without changing logical text. |
| Résumés | Embedded icons, pictures, dark background and drawing decoration were skipped. | Raster image relationships, dimensions/crops, page background, basic shapes, fill references, drawing anchors and safe hyperlinks. |
| Persian reports | Header logos and textboxes lived in separate parts; cached footer text was ignored. | Source-relative header/footer relationships, story variants, raster logos, textbox content, single-branch AlternateContent and cached visible field text. |
| PhD report | The row immediately below the thesis title contains 47 paragraphs. The paginator moved the whole row to the next page even though it permits splitting. | Browser-measured cell block boundaries let the row start below the title and continue on later pages, preserving source paragraphs, numbering and cell formatting. |
| Quarterly report | The apparent checkboxes are ordinary text: `□` in Arial, `⬜` in Segoe UI Emoji, and `☐` in Segoe UI Symbol. There are no checkbox form controls or `w:sym` elements. The emoji-font boxes appear colored in the browser. | Standalone box glyphs in emoji fonts use a monochrome symbol-font fallback and CSS text presentation; characters, source font metadata, size and spacing remain in the IR. Normal text-font symbols and explicit emoji presentation remain unchanged. |

Table direction and paragraph direction are separate OOXML properties; see [BiDiVisual](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bidivisual?view=openxml-3.0.1). Merged columns follow [GridSpan](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.gridspan?view=openxml-3.0.1). RTL run formatting also uses the complex-script channel, including neutral runs; see [RightToLeftText](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.righttolefttext?view=openxml-3.0.1). The implementation and regressions distinguish these cases.

The quarterly report's box symbols are decorative text rather than interactive controls. The preview does not invent checked states or edit the source. The monochrome fallback applies only to standalone box symbols in an emoji font; explicit U+FE0F emoji presentation and unrelated emoji are preserved. See [Unicode presentation sequences](https://unicode.org/emoji/charts/emoji-variants.html) and [CSS font-variant-emoji](https://www.w3.org/TR/css-fonts-4/#font-variant-emoji-prop).

The affected PhD report row has a minimum height of 155.45 pt and no `cantSplit` restriction. The minimum height does not require keeping all its text on one page. OOXML permits a row to span pages when no applicable restriction is set; see [CantSplit](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.cantsplit?view=openxml-3.0.1). Layout fragments are copies; the imported document and original file remain unchanged.

Follow-up page-count and overlap corrections: this report's source paper is A4 (595.3 × 841.9 pt), with 56.7 pt top/bottom margins. Paragraphs retain inherited paragraph-mark typography even without a direct `w:rPr`, and use it for the line box as well as empty paragraphs. The mark's `rtl`/`cs` run flags select its font channel independently of paragraph bidi; forcing the CS font from paragraph direction inflated neutral/Latin line boxes and blank lines. Styled runs retain their own formatting. See [RightToLeftText](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.righttolefttext?view=openxml-3.0.1) and the CSS [line-box strut rules](https://www.w3.org/TR/CSS2/visudet.html#propdef-line-height). Header/footer reservations include floating drawings and textbox contents, measured relative to each paper; background drawings do not reserve body space. The report's header logo ends at approximately 100.55 pt from the paper top. Anchor-only paragraphs retain their normal line, and measured column widths stay fixed across table fragments to prevent rewrapping after pagination.

The screenshot of page one's end also exposed a table-continuation issue: the report has three real blank paragraphs after the final activity bullet, but the paginator could move them to the next sheet and close the row border immediately below the bullet. Blank paragraphs remain in the IR and page fragments; a split row's first fragment now occupies the remaining content height so its border continues to the page's content bottom. Font-channel tests distinguish empty RTL paragraphs with and without mark-level `rtl`/`cs`, and the continuation regression retains all three blank paragraphs without inserting spacer content. Geometry, rendering and source-retention checks pass; the final three-page match has not been visually verified in a live browser.

## Structural checks

The exercise sample `تمرین دوم.docx` uses Pofak in two shaded tables. Its inherited cell margins are already represented: 5.4 pt left/right and zero top/bottom. The apparent missing vertical padding came from rendering OOXML single line spacing (`line="240"`, `lineRule="auto"`) as CSS `line-height: 1`, which is one font-size rather than the natural font line height. Single spacing now uses `normal`, allowing Pofak's ascent/descent to provide the intended breathing room without inventing cell padding. Exact point spacing remains explicit; automatic multiples other than 1 still use CSS ratios and are approximate. See [OOXML single-line units](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.spacingbetweenlines.line?view=openxml-3.0.1) and [CSS line-height](https://www.w3.org/TR/CSS22/visudet.html#propdef-line-height).

Pofak is installed per user on the supplied machine, rather than physically in `C:\Windows\Fonts`. The app's font probe also tries regular full-face and PostScript names because this font's family is `Pofak`, full face is `Pofak Regular`, and PostScript name is `Pofak-Regular`; probing the family name alone can report an installed face as missing. No font file is copied into the repository or public assets. Windows upload help points to `%LOCALAPPDATA%\Microsoft\Windows\Fonts`, and explains that Word's “(Body)” annotation is not part of the family name. See [CSS local font names](https://www.w3.org/TR/css-fonts-4/#local-font-fallback).

The report also ends with an undecorated empty body paragraph after its final table. If that paragraph cannot fit, the preview keeps it in the source model without generating a separate header-only sheet. It still renders when it fits. Explicit page breaks, list markers, drawings, fields, visible shading/borders and interior blank paragraphs remain unaffected. The regression recreates three full table pages followed by the empty end paragraph and verifies that the layout stays at three sheets.

| File | Paragraphs retained | Tables | List paragraphs | Body text retained |
| --- | ---: | ---: | ---: | --- |
| PhD_Report_1403_V1(1).docx | 120 | 12 | 26 | All |
| فرم گزارش پیشرفت سه ماهه.docx | 138 | 9 | 27 | All |
| resume.docx | 51 | 1 | 22 | All |
| resume-light.docx | 51 | 1 | 22 | All |

The check compares hashes of source-visible and imported body text after removing whitespace; character counts also match exactly for these files. It excludes metadata/deleted text and selects one markup-compatibility branch. This establishes text retention, not visual equivalence or export fidelity. All 23 tables and 97 list paragraphs are represented. Each report's selected header includes its logo; each résumé includes six raster pictures.

The IR is now version 2 with typed table and inline nodes. Public traversal helpers keep the Next.js app independent of the expanded internal tree. Page rendering uses source dimensions, margins and background rather than a host-defined paper template.

## Remaining limitations

Follow-up corrections: image-wide alpha (including the résumé's `alphaModFix amt="20000"`) is retained and rendered; color-specific alpha is not mistaken for whole-image transparency. Legacy RTL indents now use the paragraph's leading edge, markers inherit paragraph-mark fonts, and Persian RTL list labels use Persian digits without rewriting body text. Run spans no longer create separate bidi embeddings. The viewer now measures separate page sheets and repeats header/footer variants; basic page-number fields update on each sheet. Missing fonts such as the distinct header face B Compset are identified by location and can be loaded through the font control.

- The app uses browser-measured page sheets with first/even/default stories and basic PAGE/NUMPAGES results. Splittable table cells continue between paragraphs, with padding and paragraph spacing included in the measurements. Exact Word pagination, line-level paragraph splitting and full multi-section layout remain incomplete. Indivisible content, vertical merge groups, exact-height rows and rows marked `cantSplit` stay together and may be marked as oversize.
- Positioned drawings retain offsets, extents and stacking intent, but Word's wrapping/coordinate system is only approximated. The floating table in the quarterly report is placed in document order. The viewer displays compatibility notes for these cases.
- Arbitrary DrawingML geometry/effects and VML-only content are incomplete. The résumé SVG drawing uses its supplied raster fallback; the engine does not inject SVG into the page.
- B Nazanin, Pofak, Bahnschrift and other requested fonts must be installed or loaded using the app's font control. The supplied packages reference fonts but do not provide an embedded-font loading solution.
- Conditional table styles, advanced tabs and some numbering/font-selection rules remain incomplete. Editing/export and unknown-XML preservation remain separate work.

Validation includes synthetic regression tests, typecheck, production build, imports of the four actual files, and React static-render checks. Browser automation is unavailable in this environment, so these changes have not been visually certified against Word. No original document was modified.

Run `pnpm dev`, open the app, and re-upload a sample to see the updated renderer. Existing servers use built engine files; restart `pnpm dev` after engine changes to ensure a fresh build.
