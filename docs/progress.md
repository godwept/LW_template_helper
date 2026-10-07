# Development Progress

**Last updated:** 2026-10-07

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 5 core operations user-tested; simplified live-selection workflow awaiting quick retest.

Phases 1-4 are complete and user-tested. Phase 5 now adds the core bookmark-authoring operations against the temporary working copies.

INFO bookmarks and Test Mode have not started.

---

## Completed

### Phase 1 — Application shell + Word control

User-tested and complete.

Implemented:

- C# WinForms application;
- dedicated Microsoft Word instance through `Microsoft.Office.Interop.Word`;
- safe temp working copies;
- originals are never opened by the managed Word session;
- Side by Side using native Windows 50/50 positioning;
- English Focus;
- French Focus;
- clean managed Word/session shutdown.

Top / Bottom was removed after manual testing because it was not useful with the available vertical screen space.

### Phase 2 — Template Type/Subtype configuration

User-tested and complete.

Implemented:

- Template Type and subtype selectors;
- create/edit Template Type;
- create/edit subtype;
- configured bookmark-name list per Template Type;
- EN/FR source paths per subtype;
- JSON persistence under `%LOCALAPPDATA%\LWTemplateHelper\TemplateTypes`;
- basic Word bookmark-name validation;
- subtype configuration feeds the existing working-copy workflow.

### Phase 3 — Bookmark discovery/status

User-tested and complete.

Implemented:

- enumerate visible bookmarks from `Working_E.docx`;
- enumerate visible bookmarks from `Working_F.docx`;
- compare them case-insensitively against configured bookmarks;
- show EN/FR status;
- show **Other Bookmarks**;
- Refresh Status;
- clear visual mismatch indication.

Phase 3 implementation commit:

`3ca870fc18830d5e4f6fb77b1abfbc50829c28e1`

### Phase 4 — Locate + Word selection handling

Implemented on 2026-10-07 and later simplified during Phase 5.

Features:

- click a configured or Other Bookmark row to select that bookmark;
- **Locate English** selects/highlights that bookmark range in the English working copy;
- **Locate French** selects/highlights it in the French working copy;
- **Locate Both** locates the same bookmark in both documents and restores Side by Side;
- clear message if the selected bookmark does not exist in the requested language.

Selection workflow refinement:

The original Phase 4 implementation added explicit **Capture English** / **Capture French** buttons and stored duplicated Word Range objects. After user testing, this was judged unnecessary.

The current workflow relies on Word preserving the current selection independently in each document window. **Add Bookmark** and **Replace Range** read both live Word selections at the moment the action is clicked. This removes the extra capture step and reduces UI/state complexity.

Original Phase 4 implementation commit:

`ab3954d30a3214c4763db98b8daaf7f2805a7488`

Live-selection simplification commits:

- `3c55e51034a66056a46104db1988e61a35c17fe2`
- `a2068b23719ecb529332cc0bbd914f395144deff`
- `1453880b211790192ec8ee82b02c5333d33bda99`

### Phase 5 — Bookmark creation/editing

Implemented on 2026-10-07.

Features:

- **Add Bookmark**
  - available for configured Template Type bookmark rows;
  - uses the current English and French Word selections;
  - creates the same bookmark around both current Word selections;
  - refuses to overwrite if the bookmark already exists in either working copy.
- **Replace Range**
  - requires valid captured English and French selections;
  - recreates the selected bookmark around both current Word selections;
  - can repair an EN/FR mismatch by creating the missing side while moving the existing side;
  - does not change the selected document text.
- **Rename**
  - preserves each existing bookmark range;
  - renames the bookmark in whichever working copies currently contain it;
  - validates the new Word bookmark name;
  - refuses a target name already present in either working copy;
  - intentionally does **not** change the shared Template Type configuration.
- **Delete Bookmark**
  - removes bookmark markers only;
  - never deletes bookmarked text;
  - asks for confirmation first;
  - removes the selected bookmark from whichever working copies contain it.
- bookmark status automatically refreshes after every successful app-driven edit;
- obvious duplicate/invalid operations are rejected with clear messages;
- paired operations attempt to restore prior bookmark markers if the second document operation fails.

Important files:

- `src/LWTemplateHelper/WordSession.cs`
  - Add Bookmark
  - Replace Range
  - Rename
  - Delete bookmark markers
  - paired EN/FR operation rollback
- `src/LWTemplateHelper/MainForm.cs`
  - Phase 5 edit toolbar/actions
  - confirmations/status refresh
- `src/LWTemplateHelper/BookmarkNameDialog.cs`
  - bookmark rename UI and validation

Phase 5 commits:

- Word bookmark operations: `b937bc344384ccdecd0156c1a417eea5138cbb48`
- editing UI: `4035f3c52908940b245dde61ef9959f9cdabc496`
- rename dialog: `f992e634a2885e8d8f1af0ad394f882f2bedef4d`

---

---

## User-tested functionality

### Phase 1

Confirmed working by the user:

- application builds/runs;
- EN/FR templates open as temporary working copies;
- originals remain unchanged;
- Side by Side 50/50 layout works;
- English/French Focus work;
- managed Word session closes cleanly.

### Phase 2

Confirmed working by the user:

- Template Types/bookmark lists can be created and edited;
- subtypes and EN/FR paths can be created and edited;
- configuration persists after restart;
- selected subtype opens through the Word working-copy workflow.

### Phase 3

Confirmed working by the user:

- configured bookmark EN/FR status matches the open working documents;
- mismatch highlighting works;
- Other Bookmarks display works;
- Refresh Status updates after manual bookmark changes;
- existing Word/session behavior remains intact.

### Phase 4

Confirmed working by the user:

- Locate English works;
- Locate French works;
- Locate Both works;
- explicit English/French selection capture works;
- captured selection previews remain usable;
- Refresh Status does not wipe valid captures;
- Phase 1-3 behavior remains intact.

---

## Known Issues / Risks

### Phase 5 has not been compiled/run by the implementation agent

Changes were made directly through GitHub. The user workstation remains the authoritative build/runtime test environment.

### Live Word selections are now the source of truth

The explicit Capture step was removed after user testing.

For Add Bookmark and Replace Range, the user selects the desired text in both Word windows, returns to the manager, and clicks the action. The app reads each Word window's current selection at that moment.

If either selection is collapsed, the operation is rejected with a clear message.

### Locate Both requires the bookmark in both languages

If a bookmark is missing from one document, Locate Both reports that missing side. The individual Locate English/French commands remain available.

### Rename does not update Template Type configuration

Renaming a bookmark changes the working Word documents only.

This is intentional because a Template Type is shared by multiple subtypes; silently renaming the Template Type bookmark definition while editing one subtype could make other subtypes inconsistent.

If the shared configured name itself should change, edit the Template Type explicitly.

### Working-copy edits are still disposable

Phase 5 edits only `Working_E.docx` and `Working_F.docx`.

Closing the managed session still discards unsaved working-copy edits. Final source replacement remains Phase 8.

### INFO bookmarks and Test Mode remain future phases

Phase 5 does not add the sequential INFO workflow or any Letter Wizard simulation.

---

## Manual Test Checklist — Phase 5

Use temporary working copies opened through the app.

1. Sync latest `main`, build, and run.
2. Open a subtype with at least one configured bookmark missing from both documents.
3. Select the desired text in the English working copy and the desired text in the French working copy.
4. Select the missing configured bookmark row and click **Add Bookmark**.
5. Confirm status immediately changes to **Exists / Exists** and Locate Both selects the new ranges.
6. Try **Add Bookmark** again for the same bookmark and confirm it refuses to overwrite it.
7. Select two different EN/FR ranges directly in Word, select an existing bookmark, and click **Replace Range**.
8. Locate the bookmark afterward and confirm it moved to the new ranges without changing text.
9. If practical, test Replace Range on an EN/FR mismatch and confirm the missing side is repaired.
10. Select an existing bookmark and click **Rename**.
11. Confirm its text/range is unchanged and the new name appears in status.
12. Confirm a configured rename leaves the Template Type configuration unchanged (old configured name becomes missing; new name appears under Other Bookmarks).
13. Try renaming to an invalid name or an already-existing bookmark and confirm it is rejected.
14. Select a bookmark and click **Delete Bookmark**.
15. Confirm the confirmation dialog clearly states that text will be preserved.
16. Confirm only the bookmark markers disappear and all document text remains.
17. Confirm status refreshes automatically after Add, Replace, Rename, and Delete.
18. Close the Word session and confirm the original source templates remain unchanged.

---

## Exact Next Step

**User quickly retests the simplified Phase 5 live-selection workflow.**

If Phase 5 passes, implement **Phase 6 — INFO bookmarks**:

- detect `INFO_n` in both working documents;
- calculate next INFO number as highest found in either document + 1;
- add the same INFO bookmark around captured EN/FR ranges;
- show INFO EN/FR discrepancies;
- reuse normal locate/delete/change-range behavior where practical.

Do not implement Test Mode yet.

---

## Handoff Notes

Keep the implementation direct and lightweight.

The architecture remains unchanged:

- WinForms control panel;
- real Microsoft Word windows;
- Word Interop;
- safe temporary working copies;
- JSON configuration;
- deterministic user-controlled bookmark placement.

Do not introduce custom Word rendering, fuzzy matching, or automatic bookmark placement.
