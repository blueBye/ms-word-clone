# Using the web editor

This is a Word-inspired interface for **local DOCX editing**. It is not Microsoft Word and does not implement cloud collaboration or autosave to a user account.

## Open, create, and download

1. Start the app using `pnpm dev` and visit **http://127.0.0.1:3000**.
2. The bundled sample opens automatically. Use **File → Open** to select another `.docx`, or choose the blank-document action to start fresh.
3. Click inside the page to place the caret. Type, select text and use ribbon tools.
4. Use **Download DOCX** to save the edited file. Opening the result in Word is recommended for important documents.

The original source package is held by the editor session; no-op export returns the original bytes. Edit/export support depends on the capabilities of the selected content. Unsupported structured content may remain visible but read-only.

## Home

- Type and edit plain text, including IME composition and mixed RTL/LTR text.
- Undo/redo, copy/paste plain text, find and replace.
- Format fonts, emphasis, highlights, text effects, paragraph alignment, direction, spacing and indentation.
- Apply and modify styles and change lists/numbering. Some advanced style and numbering behaviors are not implemented.

The page ruler edits start/end, first-line and hanging indents. The application models RTL direction separately from alignment. For detailed boundaries, see [Home editing status](home-editor-status.md).

## Insert

- Tables and basic structural table editing; pictures and local screen captures.
- Shapes, selected vector icons, text boxes, WordArt-style content, links, bookmarks and comments.
- Basic equations, chart types, and embedded objects/spreadsheets.
- Header/footer content, page numbering, symbols, date/time, signatures and selected document properties.

These features vary in fidelity. For example, charts are limited to selected families; spreadsheet formulas are not calculated by the browser; signature lines do not perform cryptographic signing; and native Word rendering of 3D models/video remains experimental. See [Insert/Layout status](ribbon-status.md).

## Layout and contextual tools

Use **Layout** for page margins, size, orientation, columns and other implemented page settings. Click a table for contextual table tools; click a picture or other eligible object for its format controls. Page measurement, columns and floating objects are approximations of Word, not a reproduction of its layout engine.

Double-click near a page's top or bottom margin to access header/footer editing. Existing section structure can affect where edits apply.

## Fonts, privacy and safe editing

- Load local font faces from the **Fonts** pane if your document uses unavailable families. Uploaded font content is used by the browser rendering session; proprietary fonts are not bundled.
- Document processing is local to the web app in the current architecture; the editor does not expose an upload or synchronization API. External URLs and embedded online video can still trigger network access when activated.
- Use **non-sensitive test files** when evaluating behavior; keep backup originals. The DOCX parser has size and entry-count limits but is not a fully hardened untrusted-file processing boundary.
- Some files (such as enforced editing-protected and digitally signed documents) are read-only to avoid invalid edits or signature damage.

Read [Feature status](feature-status.md) and [Known limitations](limitations.md) before using advanced Word features.
