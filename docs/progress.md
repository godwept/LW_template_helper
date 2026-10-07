# Development Progress

**Last updated:** 2026-10-06

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 2 implemented; awaiting user manual testing.

Phase 1 is complete and user-tested. The application now has the basic Word lifecycle plus JSON-backed Template Type / Template Subtype configuration.

No bookmark discovery or bookmark editing has been implemented yet.

---

## Completed

### Phase 1 — Application shell + Word control

User-tested and complete.

Implemented:

- C# WinForms application;
- dedicated Microsoft Word instance through `Microsoft.Office.Interop.Word`;
- safe temp working copies:
  - `Working_E.docx`
  - `Working_F.docx`;
- originals are never opened by the managed Word session;
- Side by Side layout;
- English Focus;
- French Focus;
- clean Word/session shutdown;
- native Windows positioning for an even 50/50 Side by Side layout.

The original Top / Bottom layout was removed after user testing because it was not useful with the available vertical screen space.

The revised native Windows Side by Side layout was manually tested by the user and confirmed working.

Important Phase 1 commits include:

- initial implementation: `740e0f31d5bb005de1c5d4a737020e9e4bda4599`
- working Office interop dependency fix: `b2d476322bb07ad5e20a607e561e914124ac0eea`
- native Windows Side by Side positioning: `08a2ec2cfbbdfbfb43c35e1054c1dac65f8327d7`
- remove Top / Bottom: `4b3e45d23704b5652f9c219029f706a72b8c9c80`

### Phase 2 — Template Type/Subtype configuration

Implemented on 2026-10-06.

Features:

- Template Type selector;
- Template Subtype selector;
- create/edit Template Type;
- create/edit Template Subtype;
- bookmark-name list per Template Type;
- English/French source paths per subtype;
- simple JSON persistence;
- Word bookmark-name validation;
- duplicate Template Type / subtype / bookmark checks;
- selected subtype feeds directly into the existing Phase 1 **Open Working Copies** workflow;
- configuration controls are disabled while a managed Word session is open.

Configuration storage:

`%LOCALAPPDATA%\LWTemplateHelper\TemplateTypes`

Each Template Type is stored as its own JSON file.

Bookmark validation currently requires:

- 1-40 characters;
- first character must be a letter;
- remaining characters may be letters, numbers, or underscores;
- duplicate names are rejected case-insensitively.

Important files/classes:

- `src/LWTemplateHelper/TemplateModels.cs`
  - `TemplateTypeConfig`
  - `TemplateSubtypeConfig`
  - bookmark/path validation
- `src/LWTemplateHelper/TemplateConfigStore.cs`
  - JSON load/save
  - LocalAppData storage
- `src/LWTemplateHelper/ConfigurationDialogs.cs`
  - Template Type editor
  - Template Subtype editor
- `src/LWTemplateHelper/MainForm.cs`
  - selectors
  - configuration workflow
  - bookmark list display
  - Phase 1 integration

Phase 2 implementation commit:

`ae83deafd26634ee21d5c4a0ed689331c75f8f66`

Follow-up configuration save handling:

`21c22bc4d30443a6232b07b5693d1e0dfb55d445`

---

## User-tested functionality

### Phase 1

Confirmed working by the user:

- project builds and runs;
- EN/FR templates can be opened as temporary working copies;
- originals remain unchanged;
- Side by Side uses an even 50/50 split;
- English Focus works;
- French Focus works;
- Close Word Session works;
- closing the manager cleans up the managed Word session.

### Phase 2

Not yet manually tested.

---

## Known Issues / Risks

### Phase 2 has not been compiled/run by the implementation agent

Changes were made directly through GitHub. The user workstation remains the authoritative build/runtime test environment.

### Configuration location

Configuration is intentionally stored in `%LOCALAPPDATA%` rather than under the repository/build output so normal rebuilds do not erase saved Template Types and subtypes.

### No delete actions yet

Phase 2 requires create/edit behavior, not delete behavior. There is currently no Delete Template Type or Delete Subtype command.

If deletion becomes useful later, it can be added without changing the storage model.

### Bookmark status is not implemented yet

The bookmark list shown in Phase 2 is only the configured bookmark-name list for the selected Template Type.

It does not yet inspect the currently opened Word documents. EN/FR existence/status belongs to Phase 3.

### Source path changes

If a configured source file is moved or deleted after the subtype is saved, opening the working copies will fail until the subtype is edited with a valid path.

---

## Manual Test Checklist — Phase 2

1. Sync the latest `main` and build/run the application.
2. Create a Template Type named `Leave Without Pay`.
3. Add several bookmark names, for example:
   - `EMPLOYEE_NAME`
   - `START_DATE`
   - `END_DATE`
   - `RETURN_DATE`
4. Confirm the bookmark names appear in the main window.
5. Try an invalid bookmark such as `BAD BOOKMARK` and confirm it is rejected.
6. Create at least two subtypes under Leave Without Pay.
7. Assign real English and French `.docx` files to each subtype.
8. Switch between subtypes and confirm the displayed EN/FR paths change correctly.
9. Edit the Template Type and add/remove a configured bookmark name.
10. Edit a subtype and confirm name/path changes persist.
11. Select a subtype and click **Open Working Copies** to confirm the Phase 1 Word workflow still works from saved configuration.
12. Close and restart the application.
13. Confirm the Template Type, bookmark list, subtypes, and source paths are still present.

---

## Exact Next Step

**User manually tests Phase 2.**

If Phase 2 passes, update this file with the result and implement **Phase 3 — Bookmark discovery/status**:

- enumerate bookmarks from both working Word documents;
- compare them with the selected Template Type bookmark list;
- show EN/FR existence status;
- show Other Bookmarks;
- add Refresh Status.

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
