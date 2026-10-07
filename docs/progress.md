# Development Progress

**Last updated:** 2026-10-07

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 9 usability redesign implemented; awaiting user visual/manual testing.

Phases 1-8 are complete and user-tested. Phase 9 now applies the requested layout and visual cleanup without changing bookmark/document behavior.

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
  - automatically positions Test_E.docx and Test_F.docx side by side in a 50/50 split;
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
- initial test-window Side by Side layout: `01e15b0fdca8848aaab246e7f28675e4353030af`
- test windows overlay their respective working-copy windows: `ef7adfe32c4c8eada406906993daf7d8bb8839a2`

---

### Phase 8 — Validation + final save

Implemented on 2026-10-07.

Validation:

- **Validate** reads the current working-copy bookmark state directly from Word;
- reports configured bookmarks present in both documents;
- reports configured bookmarks unused in both documents;
- warns on configured EN/FR mismatches;
- warns on INFO EN/FR mismatches;
- lists/warns about Other Bookmarks;
- validates that both configured source paths still exist;
- blocks finalization if the English/French source paths resolve to the same file;
- blocks validation/finalization while Test Mode is active;
- warnings inform the user but do not automatically modify documents or prevent finalization.

Finalize:

- explicit **Save / Finalize** action;
- shows the current validation summary before overwrite confirmation;
- default confirmation action is **No**;
- saves both working documents through Word before copying;
- rejects read-only source templates;
- creates timestamped backups beside each source template;
- backup naming convention:
  - `<filename>.backup-YYYYMMDD-HHMMSS.docx`
  - numeric suffix added if a backup name already exists;
- overwrites the configured English/French source templates from `Working_E.docx` / `Working_F.docx`;
- if either source replacement throws, attempts to restore **both** originals from the backups;
- leaves backups in place after successful finalization;
- reports the exact backup paths after success;
- working documents remain open after finalize, allowing review or another explicit finalize if further edits are made.

Important files:

- `src/LWTemplateHelper/WordSession.cs`
  - `SaveWorkingCopies()`
  - `FinalizeToSources()`
  - backup naming and rollback
- `src/LWTemplateHelper/MainForm.cs`
  - validation summary
  - Validate action
  - Save / Finalize confirmation and result UI

Phase 8 commits:

- safe Word save / backup / source replacement: `f6360299f5a37dc2b9a1ff5b6555708d475fc337`
- validation + finalize UI: `a51157b7a2f2dbf5fdf8336886b51900eae0c0f2`
- rollback hardening: `7ad5a180b118a9167b4344b8f81ef0cea90cb86c`

---

### Phase 9 — Usability / visual pass

Implemented on 2026-10-07.

Requested workflow/layout changes:

- moved **Open Working Copies**, **Close Word Session**, **Validate**, and **Save / Finalize** into the Template Configuration section;
- moved **Side by Side**, **English Focus**, and **French Focus** into the same top configuration area;
- removed the old separate session/layout button rows from the bottom of the form;
- reorganized the bookmark workspace so all three grids are visible side by side:
  - Configured bookmarks
  - INFO bookmarks
  - Other Bookmarks
- widened the default form to better suit the three-column workspace while keeping the window resizable.

Visual language:

- added shared dark navy/slate theme inspired by the user's existing Pay Centre Automation tool;
- teal outline/accent buttons;
- teal-filled primary actions;
- red destructive/session-discard actions;
- dark data grids with blue headers;
- dark green/red/gray/amber bookmark status rows instead of the previous light pastel colors;
- matching theme applied to Template Type, Template Subtype, and Rename Bookmark dialogs;
- best-effort Windows dark native title bar through DWM, with graceful fallback on unsupported Windows builds.

Important files:

- `src/LWTemplateHelper/AppTheme.cs`
  - shared palette;
  - button/grid/control styling;
  - best-effort dark title bar.
- `src/LWTemplateHelper/MainForm.cs`
  - top action relocation;
  - three-column bookmark workspace;
  - dark status colors.
- `src/LWTemplateHelper/ConfigurationDialogs.cs`
- `src/LWTemplateHelper/BookmarkNameDialog.cs`

Phase 9 commits:

- shared dark theme: `ccf05ccc6686fe8bf58392e25767e750bc9d24ab`
- main workspace redesign: `3b82ddaaaf5623195873b4d563aa7d8e83319236`
- theme hardening: `a982ac5a491cf7cbc19a347478e03e873a6c6cca`
- configuration dialogs themed: `85e2c92859ffec1bea7ae159074b00addcb6bf5d`
- bookmark dialog themed: `959e41be3dba2fadbfd08207f12a6acd6c509ee5`
- readable disabled-button rendering: `b83b26ede8ebe7cc9f247125e842c7cb158e48ae`
- main buttons switched to themed control: `26abbe9e35ea82a848a58a347c00edbfa0ad8ebf`

Visual test feedback:

- user liked the overall new layout/theme direction;
- disabled buttons were too dark to read because WinForms used its system disabled-text rendering;
- main-window buttons now use a small custom themed button control that paints disabled states with muted light text and a subdued border/background while preserving real `Enabled = false` behavior;
- Side by Side / English Focus / French Focus were moved onto the same top action row as Open / Close / Validate / Save, aligned to the far right.

Additional Phase 9 layout commits:

- top action row alignment: `2fe0c38524f6b13c30564fa8075d09fef3933869`
- bookmark/test/status alignment refinements: `d13466f34b627fed61f7dbc755703634d3c31b88`

Latest visual refinements:

- Refresh Status is now a compact refresh icon at the top-right of the Bookmark workspace;
- bookmark editing actions start the action row at the left, with Locate English/French/Both aligned at the far right;
- Test value label/textbox/Set Value/Delete Content are vertically middle-aligned;
- bottom session status is now a plain muted label rather than a titled GroupBox;
- fixed excess vertical gaps introduced by implicit TableLayoutPanel row sizing in Template Configuration and Bookmark workspace.

Latest layout fix commits:

- explicit row sizing: `c61f4d97abf0bd7a3deae6fef7021e3c71193bd9`
- bookmark action-row sizing cleanup: `9b2fc1e86eec806c386cd9f57f2b6e1a261199b4`
- simplified section layout: `186db07759e195a70281307a5d52a35dd2964e8c`

The first row-sizing attempt did not fix the visible gaps on the user's maximized layout. The follow-up intentionally removed the problematic nested TableLayoutPanel behavior:
- Template Configuration now occupies a fixed-height root row instead of sizing itself recursively;
- Bookmark workspace now uses simple Dock=Top status/action panels and a Dock=Fill grid area;
- this is intentionally simpler and should avoid WinForms preferred-size/layout feedback loops.

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

Confirmed working by the user:

- Test Mode starts correctly;
- Set Value works;
- Delete Content works;
- Reset Test works;
- Close Test works;
- working/source documents remain protected;
- Test_E.docx opens directly over Working_E.docx;
- Test_F.docx opens directly over Working_F.docx;
- test-window placement follows the working windows rather than the manager UI.

---

### Phase 8

Confirmed working by the user:

- Validate works;
- Save / Finalize works;
- finalized source templates are updated correctly;
- backup-and-replace workflow behaves as intended.

One observed Word edge case in Test Mode: deleting one INFO bookmark's Range removed its text but Word retained the bookmark as a collapsed bookmark. The user accepts this behavior because it mirrors Word/Letter Wizard Range.Delete semantics closely enough for this utility.

## Known Issues / Risks

### Phase 8 was manually tested by the user

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

### Finalize is the only source-overwrite path

Normal editing and Test Mode continue to operate only on temporary session documents.

Only the explicit **Save / Finalize** action copies working files back over the configured source templates.

After a successful finalize, additional edits made to the still-open working copies are not written to the sources unless the user explicitly finalizes again.

---

## Manual Test Checklist — Phase 8

Use disposable/copy source templates for the first finalize test.

1. Sync latest `main`, build, and run.
2. Configure a subtype whose EN/FR source paths point to expendable copies of real templates.
3. Open working copies and make a visible bookmark change.
4. Click **Validate**.
5. Confirm the summary correctly reports:
   - configured both/unused/mismatch counts;
   - INFO mismatch count;
   - Other Bookmark count/names;
   - the exact EN/FR source paths.
6. Create or preserve an intentional EN/FR mismatch and confirm Validate warns but does not alter the documents.
7. Click **Save / Finalize**.
8. Confirm the validation summary is shown again and the overwrite prompt clearly states that source templates will be replaced.
9. Choose **No** once and confirm no source files/backups are changed.
10. Run **Save / Finalize** again and choose **Yes**.
11. Confirm both source templates now contain the working-copy bookmark changes.
12. Confirm timestamped backup files exist beside both original source templates.
13. Open the backups and confirm they contain the pre-finalize originals.
14. Confirm the working Word documents remain open after finalize.
15. Make one more working-copy edit but do **not** finalize; confirm the already-finalized source template does not receive that later change.
16. If practical, make one source template read-only and confirm Finalize refuses before overwriting it.
17. Confirm Validate / Save / Finalize are disabled while Test Mode is active.
18. Close the session and reopen the subtype; confirm the newly finalized source state is what gets copied into the new working session.

---

## Exact Next Step

**User syncs and visually/manual-tests the Phase 9 UI redesign.**

Check:

- top Template Configuration action flow;
- three side-by-side bookmark grids at normal and maximized sizes;
- dark theme readability;
- disabled/enabled button states, especially readability of disabled labels;
- primary/destructive action emphasis;
- themed configuration/rename dialogs;
- no regression to the existing Word/bookmark/Test Mode/finalize workflows.

If the layout feels good, V1 is functionally complete. Only fix concrete issues found during use.

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
