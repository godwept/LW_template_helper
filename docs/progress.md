# Development Progress

**Last updated:** 2026-10-07

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 4 complete and user-tested; ready for Phase 5.

Phases 1-3 are complete and user-tested. Phase 4 now supports locating existing bookmarks and explicitly capturing independent English/French Word selections for later bookmark operations.

Bookmark creation/editing has not started.

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

### Phase 4 — Locate + selection capture

Implemented on 2026-10-07.

Features:

- click a configured or Other Bookmark row to select that bookmark;
- **Locate English** selects/highlights that bookmark range in the English working copy;
- **Locate French** selects/highlights it in the French working copy;
- **Locate Both** locates the same bookmark in both documents and restores Side by Side;
- clear message if the selected bookmark does not exist in the requested language;
- explicit **Capture English** and **Capture French** workflow;
- captured English/French selections are stored independently as duplicated Word `Range` objects;
- short previews are shown in the manager;
- recapturing replaces the previously stored range for that language;
- collapsed/no-text-cursor selections are rejected;
- captured ranges are re-checked when bookmark status refreshes;
- clearly invalid/collapsed captured ranges are discarded;
- captured COM ranges are explicitly released when replaced or when the Word session closes.

Implementation decision:

Use explicit capture buttons instead of Word selection-change events. This is the least complicated/reliable approach allowed by the design and avoids fragile event handling while switching between two Word windows.

Important files:

- `src/LWTemplateHelper/WordSession.cs`
  - Locate English/French/Both
  - captured Word Range storage
  - capture/read validation
  - selection preview generation
- `src/LWTemplateHelper/MainForm.cs`
  - bookmark-row selection
  - Locate buttons
  - Capture buttons
  - EN/FR selection previews

Phase 4 implementation commit:

`ab3954d30a3214c4763db98b8daaf7f2805a7488`

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

### Phase 4 has not been compiled/run by the implementation agent

Changes were made directly through GitHub. The user workstation remains the authoritative build/runtime test environment.

### Explicit capture is intentional

The app does not automatically follow every Word selection change.

Workflow is:

1. select text in the English or French Word working copy;
2. return to the manager;
3. click the matching Capture button.

This keeps range capture deterministic and avoids event/lifecycle complexity.

### Stored Word ranges can move with edits

Microsoft Word Range objects can adjust their positions as nearby text is edited. Phase 4 re-checks that captured ranges are still non-collapsed and accessible, but it cannot prove that a heavily edited range still represents the user's original semantic intent.

The preview is therefore important. The user can simply recapture whenever needed.

### Locate Both requires the bookmark in both languages

If a bookmark is missing from one document, Locate Both reports that missing side. The individual Locate English/French commands remain available.

### No bookmark mutation yet

Phase 4 does not add, rename, replace, or delete bookmarks. Those belong to Phase 5.

INFO bookmarks and Test Mode also remain future phases.

---

## Manual Test Checklist — Phase 4

1. Sync latest `main`, build, and run.
2. Open a configured subtype with known bookmarks.
3. Click a configured bookmark row that exists in both documents.
4. Click **Locate English** and confirm Word selects/highlights the correct English range.
5. Click **Locate French** and confirm the correct French range.
6. Click **Locate Both** and confirm both Word windows retain the bookmark selections in Side by Side view.
7. Select a bookmark that exists in only one language and confirm the missing-side Locate command gives a clear message.
8. In the English working copy, manually select a meaningful text range.
9. Click **Capture English** and confirm the preview shows the selected text.
10. Repeat with a different range in French and **Capture French**.
11. Switch between Word windows and use layout/focus commands; confirm both captured previews remain.
12. Recapture English with different text and confirm only the English preview changes.
13. Click **Refresh Status** and confirm valid captures remain visible.
14. Click in Word without selecting text, then try Capture; confirm it rejects the collapsed selection.
15. Close the Word session and confirm both capture previews reset to **Not captured**.
16. Confirm no source template is modified by locate/capture operations.

---

## Exact Next Step

**Implement Phase 5 — Bookmark creation/editing**:

- add configured bookmark around captured EN/FR ranges;
- delete bookmark markers;
- replace bookmark ranges;
- rename bookmarks;
- auto-refresh status after app-driven operations;
- guard against duplicates/invalid names.

Do not implement INFO bookmarks or Test Mode yet.

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
