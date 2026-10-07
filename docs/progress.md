# Development Progress

**Last updated:** 2026-10-07

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 6 implemented; awaiting user manual testing.

Phases 1-5 are complete and user-tested. Phase 6 adds the dedicated sequential `INFO_n` workflow using the same live Word-selection model as normal bookmarks.

Test Mode has not started.

---

## Completed

### Phase 1 — Application shell + Word control

User-tested and complete.

Implemented:

- C# WinForms application;
- dedicated Microsoft Word instance through `Microsoft.Office.Interop.Word`;
- safe temporary `Working_E.docx` / `Working_F.docx`;
- source templates are not opened for editing;
- Side by Side using native Windows 50/50 positioning;
- English Focus / French Focus;
- clean managed Word/session shutdown.

Top / Bottom was removed after manual testing because it was not useful with the available vertical screen space.

### Phase 2 — Template Type/Subtype configuration

User-tested and complete.

Implemented:

- Template Type and subtype selectors;
- create/edit Template Type;
- create/edit subtype;
- configured normal bookmark list per Template Type;
- EN/FR source paths per subtype;
- JSON persistence under `%LOCALAPPDATA%\LWTemplateHelper\TemplateTypes`;
- Word bookmark-name validation.

### Phase 3 — Bookmark discovery/status

User-tested and complete.

Implemented:

- enumerate visible bookmarks in both working documents;
- configured bookmark EN/FR status;
- **Other Bookmarks**;
- Refresh Status;
- clear mismatch highlighting.

Phase 3 implementation commit:

`3ca870fc18830d5e4f6fb77b1abfbc50829c28e1`

### Phase 4 — Locate + Word selection handling

User-tested and complete.

Implemented:

- Locate English;
- Locate French;
- Locate Both;
- clear missing-side messages.

Selection workflow was simplified after user testing.

The original explicit Capture buttons/range storage were removed. Word's current selection in each document window is now the source of truth. Add Bookmark, Replace Range, and Add INFO read those two selections when the action is clicked.

Relevant simplification commits:

- `3c55e51034a66056a46104db1988e61a35c17fe2`
- `a2068b23719ecb529332cc0bbd914f395144deff`
- `1453880b211790192ec8ee82b02c5333d33bda99`

### Phase 5 — Bookmark creation/editing

User-tested and complete.

Implemented:

- **Add Bookmark** using current EN/FR Word selections;
- **Replace Range** using current EN/FR Word selections;
- **Rename** while preserving bookmark ranges;
- **Delete Bookmark** removes markers only, never bookmarked text;
- automatic status refresh after app-driven edits;
- duplicate/invalid operation protection;
- paired EN/FR operations attempt rollback if the second side fails.

Important behavior:

- Add refuses to overwrite an existing bookmark;
- Replace Range can repair an EN/FR mismatch;
- Rename changes the working documents only and does not silently change the shared Template Type configuration.

Phase 5 commits:

- `b937bc344384ccdecd0156c1a417eea5138cbb48`
- `4035f3c52908940b245dde61ef9959f9cdabc496`
- `f992e634a2885e8d8f1af0ad394f882f2bedef4d`

### Phase 6 — INFO bookmarks

Implemented on 2026-10-07.

Features:

- recognizes `INFO_n` bookmarks case-insensitively where `n >= 1`;
- INFO bookmarks are displayed in a dedicated status grid rather than under Other Bookmarks;
- INFO rows show EN/FR existence;
- mismatched INFO rows are highlighted red;
- matched INFO rows are highlighted green;
- INFO summary shows matched count, mismatch count, and the next INFO name;
- **Add INFO**:
  - reads the current English and French Word selections;
  - scans both documents;
  - finds the highest INFO number present in either document;
  - creates `INFO_(highest + 1)` in both documents;
  - never fills old gaps automatically;
  - refreshes status immediately;
- normal Locate English/French/Both works on selected INFO rows;
- Replace Range works on INFO rows;
- Delete Bookmark works on INFO rows and preserves text;
- Rename is intentionally disabled for INFO rows so sequential INFO naming is not accidentally broken.

Example numbering:

`INFO_1`, `INFO_2`, `INFO_4` => next is `INFO_5`.

Important files:

- `src/LWTemplateHelper/WordSession.cs`
  - `AddNextInfoBookmark()`
  - `TryGetInfoNumber()`
- `src/LWTemplateHelper/MainForm.cs`
  - INFO status grid
  - Add INFO
  - INFO mismatch display
  - INFO exclusion from Other Bookmarks

Phase 6 commits:

- sequential INFO creation: `4aad24631568c7f087bf81174e194939f371ff79`
- INFO UI/status: `f2bf667eee4bee5d5f5cee5bc2b92736d85a8b1d`
- INFO status hardening: `2c671a291a9e315c8c62e779568bffc16ccdcd72`

---

## User-tested functionality

### Phases 1-4

Confirmed working by the user, including:

- Word lifecycle/layout;
- Template Type/Subtype configuration and persistence;
- bookmark discovery/status;
- locate operations;
- Word selection handling.

### Phase 5

Confirmed working by the user after the Capture step was removed:

- Add Bookmark;
- Replace Range;
- Rename;
- Delete Bookmark;
- live EN/FR Word selections work without explicit capture;
- automatic status refresh remains usable.

### Phase 6

Not yet manually tested.

---

## Known Issues / Risks

### Phase 6 has not been compiled/run by the implementation agent

Changes were made directly through GitHub. The user workstation remains the authoritative build/runtime test environment.

### INFO numbering is global across the EN/FR pair

The next INFO number is based on the highest valid `INFO_n` found in either working document.

This deliberately preserves numbering across mismatches and gaps.

### INFO rename is disabled

INFO rows can be located, moved with Replace Range, or deleted, but not renamed through the normal Rename command.

This keeps INFO numbering predictable.

### Live Word selections are the source of truth

For Add Bookmark, Replace Range, and Add INFO:

1. select the English text in the English Word window;
2. select the French text in the French Word window;
3. return to the manager;
4. click the operation.

If either selection is collapsed, the operation is rejected.

### Working-copy edits are disposable until Finalize

All editing phases currently modify only the session working copies.

Closing the Word session uses Do Not Save. Source replacement/backups remain Phase 8.

---

## Manual Test Checklist — Phase 6

1. Sync latest `main`, build, and run.
2. Open a subtype that already contains one or more `INFO_n` bookmarks if available.
3. Confirm INFO bookmarks appear under the dedicated **INFO bookmarks** section and no longer appear under **Other Bookmarks**.
4. Confirm INFO bookmarks present in both EN/FR are green.
5. If an INFO exists on only one side, confirm it is red and the EN/FR status is clear.
6. Confirm the displayed **Next** value is highest INFO number in either document + 1.
7. If practical, create a gap manually (for example INFO_1 and INFO_4) and confirm Next is INFO_5, not INFO_2.
8. Select new text in both Word documents and click **Add INFO**.
9. Confirm the expected INFO name is created in both documents and status refreshes immediately.
10. Select different EN/FR text and click **Add INFO** again; confirm the number increments.
11. Select an INFO row and test Locate English / French / Both.
12. Select replacement EN/FR text and use **Replace Range** on the INFO row; confirm the INFO name stays the same and its ranges move.
13. Confirm **Rename** is disabled when an INFO row is selected.
14. Use **Delete Bookmark** on an INFO row and confirm only the bookmark markers disappear, not the text.
15. Confirm the next INFO number still follows highest existing number + 1 after deletions/gaps.
16. Close the session and confirm the original source templates remain unchanged.

---

## Exact Next Step

**User manually tests Phase 6.**

If Phase 6 passes, implement **Phase 7 — Test Mode**:

- save the current working document state through Word;
- create disposable `Test_E.docx` / `Test_F.docx`;
- open the test pair;
- clearly indicate Test Mode;
- support Set Value and Delete-content simulation for a selected bookmark;
- Reset Test;
- Close Test;
- guarantee test operations never change source or active working copies.

Before implementing exact Set Value/Delete behavior, inspect the existing Letter Wizard VBA semantics if available.

---

## Handoff Notes

Keep the implementation direct and lightweight.

The architecture remains:

- WinForms control panel;
- real Microsoft Word windows;
- Word Interop;
- safe temporary working copies;
- JSON configuration;
- deterministic user-controlled bookmark placement.

Do not introduce custom Word rendering, fuzzy matching, or automatic bookmark placement.
