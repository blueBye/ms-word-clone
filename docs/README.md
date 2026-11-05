# Documentation

This directory contains the **current usage and API guides** as well as **historical research and implementation notes**. The product is under active development; prefer the current status guides over earlier plans when they disagree.

## Start here

| Document | Audience | Contents |
| --- | --- | --- |
| [Getting started](getting-started.md) | Users, contributors | Prerequisites, installation, commands, troubleshooting |
| [Using the editor](using-editor.md) | Users | Opening, editing, downloading, fonts and object editing |
| [Feature status](feature-status.md) | Users, integrators | What is implemented, partial, experimental, or not supported |
| [Known limitations](limitations.md) | Everyone | Fidelity, import/export, security and deployment caveats |
| [API and integration](api.md) | Developers | Package exports, session, engine, renderer examples |
| [Engine architecture](engine-architecture.md) | Developers | Data model, dependencies, OOXML adapter, preservation, measured evidence |

For public release, read the [GitHub publishing checklist](publishing.md).

## Implementation detail

- [Home editing status](home-editor-status.md) — editing, formatting, lists, effects, search and style limitations.
- [Insert and Layout status](ribbon-status.md) — ribbons, tables, drawings, equations, charts, objects and validation.
- [Sample compatibility report](sample-compatibility.md) — snapshot from a four-document compatibility exercise, **not** a current universal conformance report.
- [DOCX compatibility baseline](compatibility.md) — **historical import/preview matrix**. Its editing/export rows predate the implemented editor; use [feature status](feature-status.md) for current product claims.

## Historical planning and research

These documents capture earlier thinking, research and requirements. Their “not implemented” statements may no longer describe the current code.

- [Development plan](development-plan.md) — initial milestone gates and compatibility strategy.
- [Engine review](engine-review.md) — original parser/renderer assessment.
- [Word UI inventory](word-ui-inventory.md) — observed Word 2024 controls, **not** an implemented-feature list.
- [Word UI implementation plan](word-ui-implementation-plan.md) — original proposed application and ribbon phases.

## Development conventions

- Prefer the **source code and tests** as the final authority when documentation is ambiguous.
- Be specific about the distinction between **import**, **render**, **edit**, **export**, and **native Word verification**.
- Do not claim exact Word pagination, complete OOXML preservation in arbitrary scenarios, or full Word feature parity.
- When functionality changes, update [feature status](feature-status.md), the relevant ribbon status page, and public usage docs together.

Return to the [repository README](../README.md).
