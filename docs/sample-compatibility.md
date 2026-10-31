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

Table direction and paragraph direction are separate OOXML properties; see [BiDiVisual](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bidivisual?view=openxml-3.0.1). Merged columns follow [GridSpan](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.gridspan?view=openxml-3.0.1). RTL run formatting also uses the complex-script channel, including neutral runs; see [RightToLeftText](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.righttolefttext?view=openxml-3.0.1). The implementation and regressions distinguish these cases.

## Structural checks

| File | Paragraphs retained | Tables | List paragraphs | Body text retained |
| --- | ---: | ---: | ---: | --- |
| PhD_Report_1403_V1(1).docx | 120 | 12 | 26 | All |
| فرم گزارش پیشرفت سه ماهه.docx | 138 | 9 | 27 | All |
| resume.docx | 51 | 1 | 22 | All |
| resume-light.docx | 51 | 1 | 22 | All |

The check compares hashes of source-visible and imported body text after removing whitespace; character counts also match exactly for these files. It excludes metadata/deleted text and selects one markup-compatibility branch. This establishes text retention, not visual equivalence or export fidelity. All 23 tables and 97 list paragraphs are represented. Each report's selected header includes its logo; each résumé includes six raster pictures.

The IR is now version 2 with typed table and inline nodes. Public traversal helpers keep the Next.js app independent of the expanded internal tree. Page rendering uses source dimensions, margins and background rather than a host-defined paper template.

## Remaining limitations

- This is a flowing preview. Exact Word pagination, repeated first/even/default stories, updated page numbers and full multi-section layout are not implemented.
- Positioned drawings retain offsets, extents and stacking intent, but Word's wrapping/coordinate system is only approximated. The floating table in the quarterly report is placed in document order. The viewer displays compatibility notes for these cases.
- Arbitrary DrawingML geometry/effects and VML-only content are incomplete. The résumé SVG drawing uses its supplied raster fallback; the engine does not inject SVG into the page.
- B Nazanin, Pofak, Bahnschrift and other requested fonts must be installed or loaded using the app's font control. The supplied packages reference fonts but do not provide an embedded-font loading solution.
- Conditional table styles, advanced tabs and some numbering/font-selection rules remain incomplete. Editing/export and unknown-XML preservation remain separate work.

Validation includes synthetic regression tests, typecheck, production build, imports of the four actual files, and React static-render checks. Browser automation is unavailable in this environment, so these changes have not been visually certified against Word. No original document was modified.

Run `pnpm dev`, open the app, and re-upload a sample to see the updated renderer. Existing servers use built engine files; restart `pnpm dev` after engine changes to ensure a fresh build.
