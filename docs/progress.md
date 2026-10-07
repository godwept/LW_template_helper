# Development Progress

**Last updated:** 2026-10-07

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 3 complete and user-tested; ready for Phase 4.

Phases 1 and 2 are complete and user-tested. Phase 3 now reads bookmark names from the open English/French working documents and compares them with the selected Template Type configuration.

Bookmark discovery/status is read-only. Bookmark creation/editing has not started.

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

Implemented on 2026-10-07.

Features:

- enumerate visible bookmarks from `Working_E.docx`;
- enumerate visible bookmarks from `Working_F.docx`;
- compare them case-insensitively against the configured Template Type bookmark list;
- show EN and FR status for each configured bookmark;
- show unconfigured bookmarks separately under **Other Bookmarks**;
- automatically refresh status after opening the working copies;
- manual **Refresh Status** button;
- visual status:
  - green-tinted row = configured bookmark exists in both;
  - gray row = configured bookmark exists in neither;
  - red-tinted row = EN/FR mismatch;
  - yellow-tinted Other Bookmark = unconfigured bookmark exists in both;
  - red-tinted Other Bookmark = unconfigured EN/FR mismatch;
- summary counts for both/mismatch/unused/other.

Word hidden/system bookmarks are intentionally excluded from discovery so generated internal bookmarks such as TOC/navigation bookmarks do not flood **Other Bookmarks**.

Important Phase 3 files:

- `src/LWTemplateHelper/WordSession.cs`
  - `GetBookmarkNames()`
  - read-only Word bookmark enumeration with COM cleanup
- `src/LWTemplateHelper/MainForm.cs`
  - bookmark status grids
  - Other Bookmarks grid
  - Refresh Status
  - mismatch styling and summary

Phase 3 implementation commit:

`3ca870fc18830d5e4f6fb77b1abfbc50829c28e1`

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
- Refresh Status updates after manual bookmark changes in a working copy;
- existing Phase 1 Word/session behavior remains intact.

---

## Known Issues / Risks

### Phase 3 has not been compiled/run by the implementation agent

Changes were made directly through GitHub. The user workstation is the authoritative build/runtime test environment.

### Visible bookmarks only

Phase 3 intentionally reads Word's normal visible bookmark collection. Hidden/system bookmarks are not shown.

If a Letter Wizard template later relies on a hidden bookmark, this decision should be revisited, but visible bookmarks are the expected template-authoring use case.

### Status is snapshot-based

Bookmark status refreshes when the working copies open and when **Refresh Status** is clicked.

Phase 3 does not subscribe to Word bookmark-change events. Automatic refresh after future app-driven bookmark edits can be added in the relevant editing phases.

### No bookmark mutation yet

Phase 3 must remain read-only. Locate, selection capture, add, rename, replace-range, delete, INFO bookmarks, and Test Mode remain future phases.

---

## Manual Test Checklist — Phase 3

1. Sync latest `main`, build, and run.
2. Select a Template Type/subtype whose Word files contain known bookmarks.
3. Click **Open Working Copies**.
4. Confirm bookmark status populates automatically.
5. Verify a bookmark present in both documents shows **Exists / Exists**.
6. Verify a configured bookmark missing from both shows **Missing / Missing**.
7. Use a pair where one configured bookmark exists in only EN or only FR and confirm the mismatch row is clearly highlighted.
8. Confirm bookmarks present in Word but absent from the Template Type configuration appear under **Other Bookmarks**.
9. Make a bookmark change manually in a working Word document, click **Refresh Status**, and confirm the display updates.
10. Confirm Phase 1 layout/focus/close behavior still works.
11. Confirm no source document is modified by the status feature.

---

## Exact Next Step

**Implement Phase 4 — Locate + selection capture**:

- Locate English;
- Locate French;
- Locate Both if practical;
- capture and retain meaningful EN/FR selections;
- show selection previews;
- detect clearly invalid/stale selections where practical.

Do not implement bookmark creation/editing, INFO bookmarks, or Test Mode yet.

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

Do not introduce custom Word rendering or automatic/fuzzy bookmark placement.
