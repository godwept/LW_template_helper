# Development Progress

**Last updated:** 2026-10-07

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 7 implemented; awaiting user manual testing.

Phases 1-6 are complete and user-tested. Phase 7 now provides disposable Test Mode copies for simulating the Letter Wizard's Set Value and Delete-content behavior without modifying working or source templates.

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
- Locate Both no longer repositions Word windows: `6c3b02414dbeada0de892ddfffeaf292a8f6bfa4`

### Phase 7 — Test Mode

Implemented on 2026-10-07.

Test Mode lifecycle:

- **Start Test Mode**
  - saves the current working copies through Word;
  - byte-copies them to `Test_E.docx` / `Test_F.docx` in the session directory;
  - opens the test pair in the managed Word instance;
  - leaves source templates untouched;
  - leaves the working-copy documents separate from all test mutations.
- **Reset Test**
  - closes/discards the current test documents;
  - saves the current working copies;
  - recreates fresh test copies from the working documents.
- **Close Test**
  - closes the disposable test documents with Do Not Save;
  - deletes the test files where possible;
  - returns the manager to normal bookmark-editing mode.

Test operations for the selected bookmark:

- **Set Value**
  - applies the entered test value to the bookmark in each test document where it exists;
  - mirrors the existing Letter Wizard helper by assigning the bookmark's `Range.Text`;
  - the bookmark marker is consumed by the Word range replacement and is not recreated;
  - disabled for INFO rows because Letter Wizard INFO bookmarks are deletion-only.
- **Delete Content**
  - calls `Range.Delete` on the selected bookmark in each test document where it exists;
  - deletes the bookmarked content/range, not merely the marker;
  - available for normal, INFO, and Other Bookmark rows.

Existing VBA behavior checked before implementation:

- `setBMvalue` replaces `Bookmarks(name).Range.Text` without recreating the bookmark;
- `deleteBookmark` deletes `Bookmarks(name).Range`;
- INFO cleanup likewise uses `bm.Range.Delete`.

UI behavior:

- prominent red **TEST MODE ACTIVE** state;
- normal bookmark mutation and working-copy locate/layout controls are disabled while Test Mode is active;
- bookmark rows remain selectable so Set Value/Delete Content can target a bookmark;
- the normal status grids continue to describe the working copies, not mutated test copies;
- test operations report whether EN, FR, or both test documents contained the bookmark.

Important files:

- `src/LWTemplateHelper/WordSession.cs`
  - test document lifecycle;
  - test Set Value;
  - test Delete Content;
- `src/LWTemplateHelper/MainForm.cs`
  - Test Mode controls;
  - state locking;
  - selected-bookmark test operations.

Phase 7 commits:

- test document lifecycle/operations: `b9df6c8ab3038cf1df4c6a728d26dd1625b221a8`
- Test Mode UI: `2708256155f8cd9ddf4d579ac99840aa76b86121`

---

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

Confirmed working by the user:

- dedicated INFO section;
- INFO numbering uses highest existing number + 1;
- Add INFO works from live EN/FR Word selections;
- INFO mismatch/status display works;
- Locate/Replace/Delete behavior works on INFO rows;
- INFO bookmarks stay out of Other Bookmarks;
- Locate Both no longer repositions Word windows.

### Phase 7

Not yet manually tested.

---

## Known Issues / Risks

### Phase 7 has not been compiled/run by the implementation agent

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

## Manual Test Checklist — Phase 7

1. Sync latest `main`, build, and run.
2. Open a subtype and make at least one temporary bookmark edit in the working copies.
3. Click **Start Test Mode**.
4. Confirm:
   - `Test_E.docx` and `Test_F.docx` open in Word;
   - the manager clearly says **TEST MODE ACTIVE**;
   - normal Add/Replace/Rename/Delete/Locate/layout controls are disabled;
   - the original working documents remain open and unchanged.
5. Select a normal bookmark row and leave the test value as `TEST VALUE`.
6. Click **Set Value**.
7. Inspect both test documents and confirm the bookmarked content was replaced by `TEST VALUE`.
8. Confirm the working documents still contain their original bookmark/content.
9. Confirm Set Value cannot simply be repeated on that same test bookmark without Reset, because the real Letter Wizard-style Range.Text replacement consumes the bookmark.
10. Click **Reset Test**.
11. Confirm the test documents return to the current working-copy state.
12. Select a bookmark and click **Delete Content**.
13. Confirm the bookmarked text/range disappears in the test documents, while the working documents remain untouched.
14. Select an INFO row:
   - confirm **Set Value** is disabled;
   - confirm **Delete Content** works.
15. If a bookmark exists on only one side, test it and confirm the manager reports EN-only or FR-only rather than modifying the other side.
16. Click **Close Test** and confirm:
   - the test documents close;
   - normal bookmark-editing controls become available again;
   - working documents remain intact.
17. Start Test Mode again to confirm a fresh pair can be created after closing.
18. Close the entire Word session and confirm source templates remain unchanged.

---

## Exact Next Step

**User manually tests Phase 7 Test Mode.**

If Phase 7 passes, implement **Phase 8 — Validation + final save**:

- validation summary;
- configured EN/FR mismatch warnings;
- INFO mismatch warnings;
- Other Bookmark warnings;
- explicit Save/Finalize;
- save working documents through Word;
- back up source templates;
- replace sources from working copies;
- clear success/failure reporting.

Do not begin the Phase 9 usability pass until finalization is safe and user-tested.

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
