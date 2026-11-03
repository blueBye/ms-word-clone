# Home editing status

The Next.js Home ribbon uses the portable `DocxSession` commands. Documents remain local. The editing session retains original package parts and unknown XML rather than rebuilding the file from the preview IR. Unsupported structured paragraphs are read-only for text/run edits.

## Current controls

| Area            | Implemented behavior                                                                                                                                                                                                                                                                                                                                                                       |
| --------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Editing surface | Direct typing, IME composition, paragraph split/join, undo/redo, plain-text paste. Enter continues lists; Enter on an empty item ends the list. Tab/Shift+Tab at an item start changes levels.                                                                                                                                                                                             |
| Font            | Latin/complex-script family and size, bold/italic, separate underline and solid strikethrough, sub/superscript, case, highlight/color, spacing, caps and hidden text. Split buttons retain the last chosen underline, highlight, color and list preset.                                                                                                                                    |
| Paragraph       | Explicit RTL/LTR, alignment, spacing, indents, shading and basic paragraph borders.                                                                                                                                                                                                                                                                                                        |
| Lists           | Word-like Bullet/Numbering/List libraries, custom Unicode bullet, number format, common letter/Roman/decimal formats, three multilevel templates, level changes, restart/start value and continuation. A collapsed caret's library choice restyles its current list; a text selection scopes the change to selected paragraphs. Level changes retain the original definition and instance. |
| Effects         | Preset gallery and Outline/Shadow/Reflection/Glow submenus; numeral forms/spacing, ligatures and stylistic sets. Office 2010 `w14` markup is written with qualified attributes and `mc:Ignorable`.                                                                                                                                                                                         |
| Styles          | Previews use resolved document styles. Apply, Update to Match Selection, Modify, Select All Instances, Rename and Remove from Gallery. Definition edits affect all instances; distinct inline direct formatting survives. Whole-paragraph uniform direct font overrides yield when applying a style.                                                                                       |
| Editing tools   | Literal Find/Replace, case sensitivity, atomic Replace All, sort consecutive paragraphs, Select All and style-based selection.                                                                                                                                                                                                                                                             |

Ribbon arrows open anchored dropdowns with Escape/outside-click dismissal and keyboard navigation. Detailed commands open modal dialogs. The Word desktop Home ribbon was inspected locally, including the style context menu and numbering/effects galleries. The icons use familiar letter/pilcrow/symbol drawings and vendored Microsoft Fluent assets; they are not extracted Office assets.

## Remaining limits

- This is a working subset, not full Word parity. Rich clipboard formats, image bullets, arbitrary per-level list definition editing, linked character styles and new style creation remain unfinished.
- Effect rendering is a CSS approximation. Inner/perspective shadows, gradient fills, dashed text outlines and 3D effects are not authorable. Imported unsupported XML is retained. Reflection may alter wrapping because the browser renders it on an inline block. OpenType results depend on font support.
- The styles UI edits common paragraph/run settings in this document. It does not update Word templates, implement automatic style updates, or customize a Quick Access Toolbar.
- Paragraph dialog spacing remains less complete than Word's exact/at-least line-height controls. Advanced search/wildcards, object selection and full cross-paragraph native editing remain further work.
- Imported drawings, fields, controls, revisions and complex table/layout cases retain their existing editing restrictions. Pagination remains browser-measured.

## Validation

`tests/home-editing.test.ts` checks bullet changes, shared numbering, level identity, hierarchical markers, restart/continue, empty-list Enter behavior, style definition updates and direct overrides, export/reimport, effect namespaces/units, effect clearing/undo, and independent underline/strike rendering. The wider test suite covers package/XML safety and renderer regressions. These checks do not replace a broad Word reopen and visual comparison corpus.

OOXML references: [Microsoft's run-property extensions](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/cdfcbd94-746b-42c4-8a0a-85efd7d24985) and [Office 2010 Word schema](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/9704b59f-bc49-4618-ac66-41beb82a0d7f).
