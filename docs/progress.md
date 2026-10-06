# Development Progress

**Last updated:** 2026-10-06

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 1 lifecycle passed manual testing; one Side-by-Side layout refinement awaiting retest.

The C# WinForms application shell now proves the basic Word/document lifecycle. No bookmark-management functionality has been implemented yet.

### Phase 1 implementation

- Select an English `.docx` source template.
- Select a French `.docx` source template.
- Create a unique session directory under the user's temp directory.
- Byte-copy the source files to `Working_E.docx` and `Working_F.docx`.
- Start a dedicated Microsoft Word instance through `Microsoft.Office.Interop.Word`.
- Open only the working copies in Word.
- Arrange the two managed Word windows using:
  - Side by Side
  - English Focus
  - French Focus
- Close managed documents and Word cleanly.
- Close documents with `wdDoNotSaveChanges` during Phase 1 so source files cannot be changed through the managed close path.

### Important files/classes

- `LWTemplateHelper.sln` — Visual Studio solution.
- `src/LWTemplateHelper/LWTemplateHelper.csproj` — .NET 8 Windows WinForms project and Word Interop dependency.
- `src/LWTemplateHelper/Program.cs` — application entry point.
- `src/LWTemplateHelper/MainForm.cs` — Phase 1 manager UI and file/layout/session controls.
- `src/LWTemplateHelper/WordSession.cs` — working-copy creation, Word COM lifecycle, and Word-window layout control.

### Commit

Phase 1 implementation commit:

`740e0f31d5bb005de1c5d4a737020e9e4bda4599`

A documentation-only follow-up commit updates this file after the implementation commit.

---

### Layout refinement commits

- Native Windows Side by Side positioning: `08a2ec2cfbbdfbfb43c35e1054c1dac65f8327d7`
- Remove Top / Bottom from UI: `4b3e45d23704b5652f9c219029f706a72b8c9c80`
- Design updated for revised layouts: `c21b81bb6924ceae4343d450a5131f136d6b603a`
- Implementation plan updated: `9127f668c26147fef3e645b5551c9340e5e4a0b1`

## Architecture

The architecture remains unchanged from `docs/design.md`.

- Standalone C# WinForms application.
- Microsoft Word remains the renderer/editor.
- Word automation through `Microsoft.Office.Interop.Word`.
- Original templates are protected by session working copies.
- Bookmark placement will remain deterministic and user-controlled.
- No custom .docx renderer/editor.

No technical limitation requiring an architecture change was encountered during Phase 1 implementation.

---

## User-tested functionality

### 2026-10-06 — Phase 1 lifecycle pass

The user manually confirmed:

- the project builds and launches;
- English and French templates can be selected;
- both temporary working copies open successfully in Word;
- English Focus and French Focus work;
- the managed Word session closes cleanly;
- closing the manager while the Word session is open cleans up correctly;
- source templates remain unchanged;
- no orphaned managed Word process was observed.

The original Top / Bottom layout also worked technically, but the user found it impractical because of limited vertical screen real estate. It has therefore been removed by user request.

The original Side by Side layout worked, but did not divide the usable monitor area evenly. The first implementation mixed Windows pixel coordinates with Word window sizing. Side by Side now uses native Windows window positioning against each Word window handle so the windows should occupy exact left/right halves of the monitor work area.

**Retest pending:** verify the revised Side by Side layout after syncing the latest `main`.

---

## Phase 1 Manual Test Findings

### 2026-10-06 — Office interop dependency issues

The first manual run built successfully but failed at **Open Working Copies** with:

`Could not load file or assembly 'office, Version=15.0.0.0 ...'`

The Word interop package requires the Office Core PIA (`office.dll`) at runtime.

Two attempted fixes were rejected during manual testing:

- direct `COMReference` — does not build under the .NET 8 SDK/MSBuild path because `ResolveComReference` is unsupported there;
- `Microsoft.Office.Core` NuGet reference — invalid package ID; restore fails because no such package exists on NuGet.

Current fix:

- keep `Microsoft.Office.Interop.Word` `15.0.4797.1004`;
- add `MSOfficeCore.Interop` `15.0.2`, which packages the required Office Core PIA and explicitly supports modern .NET targets including .NET 8.

Current fix commit:

`b2d476322bb07ad5e20a607e561e914124ac0eea`

**Result:** user confirmed the project now builds and both working copies open successfully in Word.

---

## Known Issues / Risks

### Build/runtime not exercised by the implementation agent

The repository changes were made directly through GitHub. The application was not compiled or run on a Windows machine with Microsoft Word installed.

The first user test therefore also serves as the first real compile/runtime verification.

### Word window positioning

The initial Word-property-based Side by Side positioning did not produce a true 50/50 split during manual testing.

The implementation now uses native Windows `MoveWindow` against each Word window handle, using the monitor working area in pixels. This avoids mixing Windows pixel dimensions with Word's window coordinate system.

The revised Side by Side behavior still needs one user retest, particularly if multiple monitors or unusual DPI scaling are involved.

### Focus layout

English Focus / French Focus maximizes the selected Word document window. Returning to Side by Side restores both windows to normal state and repositions them.

### COM lifecycle

The implementation explicitly closes documents, quits the dedicated Word application, releases the primary COM references, and performs final GC cleanup.

Manual testing should confirm that closing the manager does not routinely leave a managed WINWORD.EXE process behind.

### Manual Word closure

The close path tolerates the user having manually closed a managed document or Word window where Word returns a COM exception. Other unusual manual interference with the dedicated Word instance may still expose lifecycle edge cases.

### Working-copy persistence

Phase 1 does not provide Save/Finalize. Closing the managed session intentionally discards unsaved Word edits to the working copies. Final source replacement belongs to Phase 8.

---

## Manual Test Checklist

The original Phase 1 lifecycle checklist has passed.

Current retest only:

1. Sync the latest `main`.
2. Build and launch the app.
3. Select English and French templates and open the working copies.
4. Confirm they automatically open in an even 50/50 Side by Side layout across the usable monitor area.
5. Try English Focus, French Focus, then Side by Side again and confirm the 50/50 layout is restored.

---

## Exact Next Step

**User retests the revised native Windows Side by Side layout.**

If the revised 50/50 layout behaves correctly, Phase 1 is complete and the next development task is **Phase 2 — Template Type/Subtype configuration** exactly as defined in `docs/implementation-plan.md`.

Do not begin bookmark CRUD, INFO bookmarks, or Test Mode yet.

---

## Handoff Notes

Keep development quick and direct. Do not introduce strict TDD or broad automated testing.

The current implementation intentionally proves only the Word lifecycle and layout workflow. Fix Phase 1 issues found during manual testing before building configuration or bookmark features on top of it.
