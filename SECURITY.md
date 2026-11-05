# Security policy

## Status

**Experimental / not security audited.** This project processes DOCX packages and XML in the browser. It is not yet certified or hardened for hostile documents or high-assurance confidential workflows. There is no stated long-term support or guaranteed security response window.

## Reporting a vulnerability

If the GitHub repository enables **Private vulnerability reporting**, use the repository's **Security → Report a vulnerability** flow. Otherwise, contact the maintainer **privately** through an existing trusted channel before disclosing technical details. Do not include exploit payloads, private documents, credentials, or personal information in public GitHub issues.

Please describe the affected source area, reproduction steps using non-sensitive sample data, impact, and any suggested mitigation. The maintainers have not announced a disclosure SLA.

## Security boundaries

- The DOCX archive reader enforces limits on input/expanded sizes and entry counts (default 50 MiB and 2,000 entries), checks duplicate names and rejects unsafe paths. This is a useful preflight but **not a complete decompression, XML-complexity, CPU, or memory-exhaustion defense**.
- Untrusted hyperlinks and media relationships require caution. Browser video embeds and user-activated external links may contact third-party domains.
- Enforced editing protection and digital signatures can make a session read-only. The application does not break protection or create legally binding signatures.
- The standalone web app does not implement authentication or a cloud document store. Hosting the app publicly does **not** add access control, private storage, or enterprise security features.
- Do not add production secrets, private fixtures, personal documents or licensed font binaries to the repository.

See [Known limitations](docs/limitations.md) for interoperability and operational considerations.
