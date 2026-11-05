# Known limitations and compatibility policy

This project is an **experimental DOCX editor**, not a full Microsoft Word implementation. Import success, preview appearance, editable selection, and safe edited DOCX export are **different guarantees**.

## Word interoperability

**Pagination and typography**

- The renderer measures pages in the browser. It can split across paragraph/table boundaries, but it does not replicate Word's line-breaking and layout algorithms.
- Complex continuous sections, unequal-width columns, float wrapping, nested drawing effects and very tall indivisible content may render differently or overflow.
- Missing fonts, complex-script shaping, language-dependent glyph selection, field updates and paragraph mark metrics can change the appearance of Persian/Arabic/English documents.
- Exact Word page parity is not an advertised feature.

**DOCX import, editing and export**

- OOXML is a large specification. The parser/rendering layer covers only a tested subset; unknown or unsupported structures may be skipped or approximated in the preview.
- The editing session retains original package parts. A no-op export returns the original bytes. Supported edits aim to preserve unrelated source content, but complex structure changes can still interact with unsupported XML.
- Some fields, content controls, revisions, protected paragraphs, and complex drawing structures are locked or restricted to prevent accidental loss.
- A standalone core model snapshot is **not** an OOXML-preserving archive; do not use it as a lossless DOCX serializer.
- The full suite of Word-version reopen tests and document corpus needed for production compatibility is not available. Verify outputs using the target Word/LibreOffice version before relying on them.

**Advanced objects**

- Chart authoring covers selected chart families; spreadsheet formulas are authored but not evaluated in the browser.
- Equations support a limited TeX/linear subset translated to OMML. Unsupported TeX commands or complex imported equation structures need separate handling.
- OLE packaging, video metadata, and 3D model parts are experimental; native handling in Word depends on platform/version and has not been broadly verified.
- Visual signature lines are **not cryptographic signatures**. Digitally signed documents are read-only to avoid invalidating their signatures.
- This is not a cloud collaboration, tracked-review or Office add-in platform.

## Security and privacy

- The current Next.js workspace processes opened document bytes in the browser and does not provide a document upload/storage API, authentication, or cloud document synchronization.
- `openDocxPackage` checks ZIP signatures, declared and extracted sizes, entry counts, duplicate names and unsafe paths. Default limits are **50 MiB** and **2,000 entries**. These controls **do not** make the parser a fully hardened boundary for hostile DOCX files; XML complexity, CPU and memory/resource exhaustion require further work.
- Online video embeds and external links can contact third parties when activated. Browser screen capture requires user consent through the browser's own selection prompt.
- Do not publicly upload sensitive documents, real employee/customer records, licensed font binaries or proprietary Office attachments as debugging fixtures.
- Deploying this code publicly requires your own hosting, operational security review, HTTP security policy and abuse/resource limits. Do not infer security readiness from successful unit tests alone.

## Platform and maintenance

- The web app is a **Next.js server deployment**. The configuration does not enable `output: "export"` and cannot be hosted as a static-only site without additional changes.
- Local editing does not imply automatic recovery of unsaved changes. Download an updated DOCX to keep your work.
- Features documented in historical plans are not necessarily current. Refer to [Feature status](feature-status.md) and the implementation code.

## Suggested compatibility checklist

For a new DOCX feature or bug report, test: (1) import, (2) preview, (3) select/edit controls, (4) no-op export identity where applicable, (5) edited export and reimport, (6) unrelated package-part retention, and (7) native Word reopen/rendering when available. Use an anonymized minimal document that exercises the issue.

See [Security](../SECURITY.md), [Feature status](feature-status.md), and [Contributing](../CONTRIBUTING.md).
