# Development Progress

**Last updated:** 2026-10-06

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Phase 1 implemented; awaiting user manual testing.

The C# WinForms application shell now proves the basic Word/document lifecycle. No bookmark-management functionality has been implemented yet.

### Phase 1 implementation

- Select an English `.docx` source template.
- Select a French `.docx` source template.
- Create a unique session directory under the user's temp directory.
- Byte-copy the source files to `Working_E.docx` and `Working_F.docx`.
- Start a dedicated Microsoft Word instance through `Microsoft.Office.Interop.Word`.
- Open only the working copies in Word.
- Arrange the two managed Word windows using:
  - Top / Bottom
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

None yet for Phase 1. The user should sync the repository and manually test the milestone below.

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

**Retest required:** sync `main`, restore/build, then select both templates and click **Open Working Copies**.

---

## Known Issues / Risks

### Build/runtime not exercised by the implementation agent

The repository changes were made directly through GitHub. The application was not compiled or run on a Windows machine with Microsoft Word installed.

The first user test therefore also serves as the first real compile/runtime verification.

### Word window positioning

Layout uses the working area of the monitor containing the manager window and Word's `Window.Left/Top/Width/Height` properties.

Top/Bottom and Side-by-Side need real testing on the user's workstation, especially with multiple monitors or non-default DPI scaling.

### Focus layout

English Focus / French Focus maximizes the selected Word document window. Returning to Top/Bottom or Side-by-Side restores both windows to normal state and repositions them.

### COM lifecycle

The implementation explicitly closes documents, quits the dedicated Word application, releases the primary COM references, and performs final GC cleanup.

Manual testing should confirm that closing the manager does not routinely leave a managed WINWORD.EXE process behind.

### Manual Word closure

The close path tolerates the user having manually closed a managed document or Word window where Word returns a COM exception. Other unusual manual interference with the dedicated Word instance may still expose lifecycle edge cases.

### Working-copy persistence

Phase 1 does not provide Save/Finalize. Closing the managed session intentionally discards unsaved Word edits to the working copies. Final source replacement belongs to Phase 8.

---

## Manual Test Checklist

1. Open `LWTemplateHelper.sln` in Visual Studio/VS Code on Windows.
2. Restore/build the project.
3. Launch the app.
4. Choose a representative complex English `.docx` template.
5. Choose its French `.docx` template.
6. Click **Open Working Copies**.
7. Confirm Word opens two documents named `Working_E.docx` and `Working_F.docx`, not the source files.
8. Confirm the status shows a unique temp session directory.
9. Visually inspect both documents for intact formatting/structure.
10. Try **Top / Bottom**, **Side by Side**, **English Focus**, and **French Focus**.
11. Make an obvious edit in a working copy.
12. Close the Word session from the manager.
13. Reopen the original source file separately and confirm the edit is absent.
14. Confirm the managed Word process closes and no orphaned WINWORD.EXE remains.
15. Repeat once by closing the manager while the Word session is still open.

Report any compile error, Word startup/open error, awkward layout behavior, or orphaned Word process before Phase 2 begins.

---

## Exact Next Step

**User manually tests Phase 1 on the actual Windows/Word workstation.**

If Phase 1 passes, update this file with the test result and begin **Phase 2 — Template Type/Subtype configuration** exactly as defined in `docs/implementation-plan.md`.

Do not begin bookmark CRUD, INFO bookmarks, or Test Mode yet.

---

## Handoff Notes

Keep development quick and direct. Do not introduce strict TDD or broad automated testing.

The current implementation intentionally proves only the Word lifecycle and layout workflow. Fix Phase 1 issues found during manual testing before building configuration or bookmark features on top of it.
