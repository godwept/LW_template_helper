# Subtype Bookmark Selection Implementation Plan

**Date:** 2026-10-09  
**Design doc:** [Approved subtype bookmark selection design](../specs/2026-10-09-subtype-bookmarks-design.md)  
**Status:** Implementation complete; 15 regression checks and Release build pass; user manual acceptance pending

## Overview

Add a normal-bookmark checklist to subtype configuration and persist each subtype's selections. New subtypes start with none selected; legacy missing/null selections mean all master bookmarks until edited. The configured grid and validation use the subtype selection, while Other Bookmarks continues to use the entire Template Type master list. Known unselected bookmarks are ignored. INFO discovery and Word operations retain their existing behavior.

## Execution conventions

- Run commands from `D:\source\repos\LW_template_helper` on Windows with the .NET 8 SDK/runtime or a compatible newer SDK.
- Every task below is a 2–5 minute change. Follow the order, add the stated check before implementation, confirm it fails for the intended reason, and rerun it after the change. If a task takes longer, split its listed checks into smaller increments.
- Reuse the package-free WinForms runner at `tests/BookmarkGridRegression/Program.cs`. It compiles the real form/dialog/model sources with in-memory Word/configuration services. Keep those services isolated from user data.
- Add one small package-free configuration runner for the real JSON store. Do not introduce a test framework, COM mocks, or a new application layer.
- `UI checks` means `dotnet run --project tests/BookmarkGridRegression/BookmarkGridRegression.csproj`.
- `Configuration checks` means `dotnet run --project tests/BookmarkConfigurationRegression/BookmarkConfigurationRegression.csproj`.
- The working tree already contains the earlier scroll fix, its regression runner, and approved design notes. Preserve them and review the starting `git status --short` before editing.

## Tasks

### Task 1: Represent legacy and explicit subtype selections

**Time:** 5 minutes.  
**Files:** `src/LWTemplateHelper/TemplateModels.cs`, `tests/BookmarkConfigurationRegression/BookmarkConfigurationRegression.csproj` (new), `tests/BookmarkConfigurationRegression/Program.cs` (new).

**Test first:** Create a console regression project targeting `net8.0`, with nullable references and implicit usings enabled. Link the real `../../src/LWTemplateHelper/TemplateModels.cs` and `../../src/LWTemplateHelper/TemplateConfigStore.cs`. Use named `Action` checks, assertions, PASS/FAIL output, and a nonzero exit code on failure, following the existing runner. Add checks for master list `[A, B, C]`: missing/null selection resolves to `[A, B, C]`; explicit `[]` resolves to `[]`; `[c, a, STALE]` resolves to `[A, C]`; adding `D` does not expand that explicit selection; an empty master resolves to empty.

**Implementation:** Add nullable `List<string>? Bookmarks` to `TemplateSubtypeConfig`, leaving its default null for old JSON compatibility. Add `TemplateTypeConfig.GetBookmarksForSubtype(TemplateSubtypeConfig subtype)` returning a new `List<string>`. For a null selection, copy the master list. Otherwise, use a case-insensitive set of the subtype names and filter the master list in its original order. Do not initialize a deserialized missing list to empty. New-subtype creation will explicitly assign the dialog's empty/selected list in Task 6.

**Verify:** Run Configuration checks. Confirm both empty and legacy cases pass and the master/subtype source lists remain unchanged by resolution.

---

### Task 2: Verify JSON persistence in an isolated directory

**Time:** 5 minutes.  
**Files:** `src/LWTemplateHelper/TemplateConfigStore.cs`, `tests/BookmarkConfigurationRegression/Program.cs`.

**Test first:** Add a store round-trip check using a unique fixture directory under the regression executable's `AppContext.BaseDirectory` (ignored build output). Write legacy JSON with one subtype lacking `Bookmarks`, another with `Bookmarks: null`, one with `Bookmarks: []`, and one with `Bookmarks: ["A"]`. Load, save, and reload using the real store. Assert the distinction survives, selections stay independent, and resolving legacy selections still yields the master list. Inspect the saved JSON to ensure the explicit empty list is `[]`. Never use the user's `%LOCALAPPDATA%` directory in the runner.

**Implementation:** Give `TemplateConfigStore` an optional constructor argument `string? configDirectory = null`. Move the current default path assignment into that constructor and use the supplied directory only when present. Preserve the existing default path, directory creation, JSON options, file naming, and save/error behavior. The new public model property is serialized automatically; do not add a migration that materializes legacy selections.

**Verify:** Run Configuration checks. Confirm fixture paths reside under the test project's `bin` directory and that no application configuration files are read or changed. Leave fixtures in ignored build output or delete only the exact files created by the check.

---

### Task 3: Remove stale saved selections when saving configuration

**Time:** 3 minutes.  
**Files:** `src/LWTemplateHelper/TemplateConfigStore.cs`, `tests/BookmarkConfigurationRegression/Program.cs`.

**Test first:** Start with master `[A, B]`, an explicit subtype selection `[a, B]`, an empty subtype, and a legacy null subtype. Replace the master with `[A, C]`, save/reload, and expect explicit `[A]`, empty `[]`, and legacy null. Verify `C` remains unchecked for the explicit subtype and the legacy subtype still resolves to all current master names. Include a removed-and-replaced name to cover rename-as-removal-plus-addition.

**Implementation:** Before serializing in `TemplateConfigStore.Save`, normalize only subtypes with non-null `Bookmarks` through `config.GetBookmarksForSubtype(subtype)`. This removes stale names, preserves master casing/order, and keeps null legacy selections null. Do not change the master list or Word documents.

**Verify:** Run Configuration checks; inspect the fixture JSON for removed names and preserved null/empty state.

---

### Task 4: Add and initialize the subtype checklist

**Time:** 5 minutes.  
**Files:** `src/LWTemplateHelper/ConfigurationDialogs.cs`, `src/LWTemplateHelper/MainForm.cs`, `tests/BookmarkGridRegression/Program.cs`.

**Test first:** Extend the runner's reflection helpers to accept an arbitrary owner object and use its runtime type, so they work for both the form and dialogs. Add constructor checks using master `[A, B, C]`: new dialog has no checked items; legacy existing subtype has all checked; explicit `[c, a, STALE]` checks `[A, C]` in master order; explicit empty and empty-master cases are valid. Assert the existing name/path fields are restored. Do not open Word.

**Implementation:** Change the dialog signature to `TemplateSubtypeDialog(TemplateTypeConfig type, TemplateSubtypeConfig? existing = null)`. Add a `CheckedListBox` field `_bookmarksList` with `Dock = Fill`, `CheckOnClick = true`, `IntegralHeight = false`, and the existing `AppTheme.InputBack`/`AppTheme.Text` colors. Populate it from `type.Bookmarks`, checking only effective existing selections; use an empty selection when `existing` is null. Add a label `Bookmarks for this subtype`, increase the dialog to approximately `720 x 480` with a `620 x 400` minimum, and give the checklist the flexible row in an explicitly sized layout. Keep name/path rows and Save/Cancel visible. Update the two MainForm constructor call sites immediately to pass `type`, so the app still compiles; selection saving is handled next.

**Verify:** Run UI checks and `dotnet build LWTemplateHelper.sln --configuration Release`. Inspect the dialog at normal/minimum sizes during final manual verification; avoid adding bulk-selection controls or unrelated theme changes.

---

### Task 5: Return checked names without mutating the existing subtype

**Time:** 3 minutes.  
**Files:** `src/LWTemplateHelper/ConfigurationDialogs.cs`, `tests/BookmarkGridRegression/Program.cs`.

**Test first:** Create isolated `.docx` fixture files under the UI runner's build output; they only need to exist for the dialog's current path validation and are never opened by Word. Fill valid name/path fields, check `A` and `C`, invoke `SaveAndClose`, and assert the dialog returns `[A, C]` with OK. Repeat with no checks and expect an explicit empty list. Change checks on an existing subtype and cancel/dispose; assert the source subtype's name, paths, and selection remain unchanged. Also assert invoking Save does not directly modify the source subtype; MainForm applies the result after OK.

**Implementation:** Add a non-null result property `List<string> BookmarkNames { get; private set; } = []` on `TemplateSubtypeDialog`. In `SaveAndClose`, after existing name/path validation succeeds, assign `_bookmarksList.CheckedItems.Cast<string>().ToList()` before setting OK. Do not add a minimum-selection rule or mutate the object supplied to the constructor.

**Verify:** Run UI checks. Confirm empty selections save successfully and cancellation leaves the supplied configuration untouched.

---

### Task 6: Save the selection when creating a subtype

**Time:** 5 minutes.  
**Files:** `src/LWTemplateHelper/MainForm.cs`, `tests/BookmarkGridRegression/TestServices.cs`, `tests/BookmarkGridRegression/Program.cs`.

**Test first:** Let the in-memory `TemplateConfigStore` fake remember the last saved type, count saves, and return the saved type from `LoadAll`; keep its state per instance. Add a small modal-dialog driver: a WinForms timer checks for the open `TemplateSubtypeDialog`, fills its name/fixture paths, changes checks, and invokes Save or Cancel. Use a stopwatch deadline, capture callback exceptions for reporting after the modal call, close the owned dialog on failure, and stop/dispose the timer in `finally`; wait on the dialog condition, never a fixed sleep. Invoke `MainForm.CreateSubtype` with a selected type. Assert saving no checks creates an explicit `[]`, saving a subset creates exactly that list, and cancelling adds no subtype. These tests must exercise the MainForm call path, not only construct a model object.

**Implementation:** In the `TemplateSubtypeConfig` initializer in `CreateSubtype`, assign `Bookmarks = dialog.BookmarkNames.ToList()` alongside name/paths. Keep existing duplicate checks, configuration persistence, and selection/reload behavior. The dialog already receives the parent type from Task 4.

**Verify:** Run UI checks and Configuration checks. The newly created subtype remains empty after Save/reload; it does not inherit the master list.

---

### Task 7: Save edits independently and materialize legacy selections

**Time:** 3 minutes.  
**Files:** `src/LWTemplateHelper/MainForm.cs`, `tests/BookmarkGridRegression/Program.cs`.

**Test first:** Reuse the modal driver to edit a selected legacy subtype, uncheck all but `B`, save, and expect explicit `[B]`. Reopen and save with none checked to expect `[]`. Verify another subtype's selection remains unchanged. Cancel another edit and verify the selected subtype is unchanged and no extra configuration save occurs.

**Implementation:** After existing validation succeeds in `EditSubtype`, copy `dialog.BookmarkNames` into `subtype.Bookmarks` together with the existing name/path updates. Keep duplicate-name guards, SaveConfiguration error reporting, and reload behavior. Do not infer selections from Word or change the paired bookmark editing operations.

**Verify:** Run UI checks and Configuration checks. Legacy selection changes only when the subtype edit is successfully accepted; editing another subtype does not copy selections between them.

---

### Task 8: Show the selected subtype's rows before Word opens

**Time:** 5 minutes.  
**Files:** `src/LWTemplateHelper/MainForm.cs`, `tests/BookmarkGridRegression/Program.cs`.

**Test first:** Add a type with master `[A, B, C]` and three subtypes: explicit `[C, A]`, explicit `[]`, and legacy null. Switch the subtype combo with no Word session and expect `[A, C]`, empty, and `[A, B, C]`. Verify selected-bookmark state is cleared when switching, and selecting a type with no subtypes leaves no configured rows. Update the existing `CreateForm` scroll fixture to include a legacy subtype, so it continues representing a configured session rather than a type with no subtype.

**Implementation:** Use `GetBookmarksForSubtype` in `PopulateConfiguredBookmarkRows` when both a type and subtype are selected; otherwise clear the rows. Call that population and `ClearSelectedBookmark` from `SubtypeChanged`, including its null-selection path. Remove the premature master-list population in `TemplateTypeChanged` and let the final subtype selection drive it. Keep the existing null-type prompt and show a concise select/create-subtype prompt when appropriate. Ensure closing a session repopulates the selected subtype's filtered rows. Keep session-based configuration locks intact.

**Verify:** Run UI checks. Confirm no stale selection/actions survive a subtype switch and the existing scroll scenarios still pass.

---

### Task 9: Filter refreshed rows and counts while preserving Other/INFO discovery

**Time:** 3 minutes.  
**Files:** `src/LWTemplateHelper/MainForm.cs`, `tests/BookmarkGridRegression/Program.cs`.

**Test first:** Set master `[A, B, C]`, subtype `[A, C]`, English names `[A, B, INFO_1, UNKNOWN]`, and French names `[A, INFO_1]`. After refresh, expect configured rows `[A, C]`, counts `1 both, 0 mismatch, 1 unused`, INFO `[INFO_1]`, and Other `[UNKNOWN]`. Assert the unselected English-only `B` appears nowhere. Repeat with empty subtype selection and expect zero configured counts while INFO/Other still appear.

**Implementation:** In `RefreshBookmarkStatus`, obtain a selected subtype and resolve its effective list for the configured-row loop. Keep a separate case-insensitive set of the full `type.Bookmarks` for Other Bookmarks exclusion. Do not substitute the subtype list into that exclusion. Keep INFO discovery, styles, summaries, and scroll capture/restore in their existing locations. With no subtype selected, return rather than displaying the master list as a configured session.

**Verify:** Run UI checks, including all three original scroll checks. Confirm `B` is ignored, `UNKNOWN` is still Other, and INFO remains independent.

---

### Task 10: Filter validation without adding new warning rules

**Time:** 4 minutes.  
**Files:** `src/LWTemplateHelper/MainForm.cs`, `tests/BookmarkGridRegression/Program.cs`.

**Test first:** With valid distinct fixture paths, use master `[A, B, C, salutation]`, selection `[A, C, salutation]`, English `[A, B, salutation, INFO_1, UNKNOWN]`, and French `[A, INFO_1]`. Invoke `BuildValidationSummary` and assert `1 in both, 1 unused, 1 mismatch`; the mismatch names `salutation (EN only)`, unselected `B` is absent from warnings/Other, and `UNKNOWN` is reported as Other. Repeat with `[]` and expect zero configured counts/warnings while unknown/INFO reporting remains. Add an INFO-only mismatch case and confirm its existing warning is still generated.

**Implementation:** Change only the configured-bookmark loop in `BuildValidationSummary` to use `type.GetBookmarksForSubtype(subtype)`. Retain the full-master set for unknown-bookmark detection, existing INFO logic, unused classification, source-path checks, Test Mode protection, and finalization warnings/confirmation. Do not report known unselected bookmarks or add language exceptions/required-bookmark rules.

**Verify:** Run UI checks. Confirm `salutation` keeps the accepted French-missing warning and empty selections introduce no new warning category.

---

### Task 11: Verify scrolling with a filtered, shortened, or empty configured list

**Time:** 3 minutes.  
**Files:** `tests/BookmarkGridRegression/Program.cs`; `src/LWTemplateHelper/MainForm.cs` only if a regression is found.

**Test first:** Add checks using a long explicit subtype selection rather than only a legacy all-bookmarks fixture. Scroll all grids, add a configured bookmark, and verify each first displayed row remains unchanged. Add INFO and verify its new last row is visible while the other grids stay put. While scrolled, shorten the effective configured selection to a few rows, refresh, and confirm the stored first row clamps to the new final row. Then set `[]`, refresh, and confirm an empty configured grid without an exception. Manipulating fixture selections while a fake session is open is only a test of the refresh boundary; real configuration editing remains locked.

**Implementation:** Reuse the existing scroll preservation code. Make a minimal correction only if a failing check demonstrates a new interaction introduced by filtering; do not redesign selection or scrolling behavior.

**Verify:** Run UI checks and Configuration checks. Every original and added scenario passes with no fixed-delay waits.

---

### Task 12: Build and record the implementation handoff

**Time:** 3 minutes.  
**Files:** `docs/progress.md`, `docs/design.md`, `docs/implementation-plan.md`, `docs/specs/2026-10-09-subtype-bookmarks-design.md`, `docs/plans/2026-10-09-subtype-bookmarks.md`.

**Test first:** Run both regression executables and the Release build before marking any task complete. Inspect the final diff against this plan and the approved design. Confirm no implementation edits to `src/LWTemplateHelper/WordSession.cs`, no user configuration writes from checks, and no unrelated changes beyond the starting working-tree state.

**Implementation:** Record completed work, changed classes/files, the null-versus-empty compatibility decision, actual verification commands/results, known issues, and the exact next manual test in `docs/progress.md`. Change the approved-extension status in `docs/design.md` to implemented only when accurate. Update the Phase 10 status in `docs/implementation-plan.md` and this plan; tick design success criteria only when verified. Separate local automated results from user manual results.

**Verify:** Run:

```powershell
dotnet run --project tests/BookmarkConfigurationRegression/BookmarkConfigurationRegression.csproj
dotnet run --project tests/BookmarkGridRegression/BookmarkGridRegression.csproj
dotnet build LWTemplateHelper.sln --configuration Release
git diff --check
git status --short
```

Record unresolved manual testing rather than claiming it passed.

## Manual acceptance checks

Use disposable source-template copies for Word testing. These checks are for the user workstation after implementation; do not treat pending user results as automated verification.

1. Open an existing legacy subtype's editor: all master bookmarks are checked. Cancel and confirm its current full list is preserved.
2. Create a subtype: all boxes start unchecked. Save with none selected and confirm the configured grid is empty.
3. Edit that subtype, select a subset, save, switch away/back, and restart the app. Confirm only those names are shown in master order.
4. Edit a different subtype and confirm its choices remain independent. Check the dialog at its minimum size for accessible checklist and Save/Cancel controls.
5. Add a master bookmark: it is unchecked for subtypes with explicit saved selections. Remove a selected master name: it disappears from their selections. A legacy subtype continues using the current master list until its own selection is saved.
6. Open Word working copies containing a known unselected normal bookmark, a truly unknown bookmark, and INFO bookmarks. Confirm the known unselected name is ignored, the unknown name appears under Other Bookmarks, and INFO remains separate.
7. Select English-only `salutation`; confirm its existing EN/FR discrepancy warning remains. Confirm selected bookmarks absent on both sides remain unused rather than becoming required.
8. Confirm the three-grid scroll behavior when adding configured bookmarks and INFO. Changing checklist selections does not modify Word document content/bookmarks; configuration editing remains locked while the session is open.

## Definition of Done

- [x] All implementation tasks completed in order, with failing checks added before their fixes.
- [x] Both regression runners pass and the Release build succeeds.
- [x] New subtypes start empty; legacy missing/null selections and explicit empty selections stay distinct across persistence.
- [x] Grid/status/validation filtering matches the approved design, including the unchanged meaning of Other Bookmarks and independent INFO behavior.
- [x] Known unselected bookmarks produce no rows, warnings, or extra reports.
- [x] No unrelated files modified; earlier scroll-fix changes preserved.
- [x] Documentation reflects implemented behavior and accurately identifies any pending manual acceptance checks.
- [ ] User manual acceptance results recorded when available.

## Execution results — 2026-10-09

Tasks 1–12 are implemented. The configuration runner passes 3 checks; the WinForms runner passes 12, including the original 3 scroll regressions, Create/Edit modal paths, Cancel, filtering, empty selections, and dialog bounds at normal/minimum sizes. The Release build reports zero warnings/errors, and `git diff --check` passes. The shortening-list check validates a legal visible viewport because WinForms can normalize the first displayed row to zero when all remaining rows fit; no additional scrolling code change was needed. Real Word integration code is unchanged. User manual acceptance is pending.
