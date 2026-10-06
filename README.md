# Letter Wizard Template Bookmark Manager

A small Windows/C# development tool for creating, reviewing, maintaining, and testing Word bookmarks used by the Excel/VBA Letter Wizard templates.

The tool is intentionally deterministic: **AI is not used for bookmark placement.** Bookmark ranges are selected by the user in Microsoft Word.

## Project status

Architecture/design phase. No application implementation should begin until the document-preservation approach in the design has been accepted.

## Authoritative project documents

- [Design](docs/design.md) — architecture, UI, document-safety rules, and core workflows.
- [Implementation Plan](docs/implementation-plan.md) — phased build plan.
- [Development Progress](docs/progress.md) — current state and handoff notes. **Read this first when resuming development.**
- [Agent Instructions](AGENTS.md) — rules for any AI/dev agent working on the repository.

## Core architectural decision

The application will be a standalone C# WinForms application that automates the locally installed Microsoft Word application using Word Interop.

Word remains responsible for rendering, editing, and saving Word documents. The application does not attempt to import/rebuild/reformat .docx contents.

Original templates are not opened for editing during normal work. A session operates on temporary working copies and only replaces the source templates through an explicit save/finalize action.
