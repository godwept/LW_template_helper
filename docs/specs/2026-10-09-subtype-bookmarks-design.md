# Subtype Bookmark Selection Design

**Date:** 2026-10-09
**Status:** Approved and implemented; local checks pass; user manual testing pending

## Goal

Let each Template Subtype select the normal bookmarks that apply to its English/French template pair. Keep the Template Type as the master list, and show only the selected bookmarks in the main grid. The user is responsible for choosing the correct bookmarks.

## Success Criteria

- [x] The subtype editor offers a checklist drawn from the Template Type's bookmark list.
- [x] New subtypes start with no bookmarks selected; an empty selection is valid.
- [x] Existing configurations retain all Template Type bookmarks until the subtype selection is edited and saved.
- [x] Saved selections survive restarting the app and remain independent between subtypes.
- [x] The main grid and configured-bookmark validation use only the selected subtype's bookmarks.
- [x] Unselected normal bookmarks are ignored, even when present in Word.
- [x] Other Bookmarks contains only document bookmarks absent from the Template Type's master list, with INFO handled separately.
- [x] INFO behavior and Word document operations remain unchanged.

Verified locally with real JSON-store reloads, actual WinForms form/dialog code and in-memory Word data, and a successful Release build. Manual acceptance with the user's Word templates remains pending.

## Scope

**In scope:**

- A bookmark checklist in New/Edit Template Subtype.
- Per-subtype selection persistence and compatibility with existing JSON configurations.
- Filtering the configured grid, counts, and validation by the selected subtype.
- Keeping selections consistent when the master bookmark list changes.

**Out of scope:**

- Reporting, warnings, or an extra grid for unselected normal bookmarks.
- Per-language applicability settings or exceptions for English-only bookmarks.
- New required-bookmark rules or stricter finalization checks.
- Inferring selections from document contents or changing Word documents when configuration changes.

## Design

### Configuration and state

The Template Type continues to own its existing ordered `Bookmarks` list. Each Template Subtype gains a nullable bookmark-name list representing its selection.

- A missing/null subtype list means a legacy subtype uses all current master bookmarks until its selection is explicitly saved.
- New subtypes have an explicit empty list.
- An explicit list, including an empty list, is the saved selection.
- Match names case-insensitively and display selected names in master-list order.

After a subtype has a saved selection, adding a master bookmark does not automatically select it. Removing a master bookmark removes that name from subtype selections when the Template Type configuration is saved. Changing a master name is treated as removal plus addition; the new name must be selected explicitly for subtypes with saved selections.

### Subtype editor and workspace

The subtype editor receives the parent Template Type's bookmark list and shows one checkbox per normal bookmark alongside the existing name and EN/FR path fields. New subtypes start unchecked; existing explicit selections are restored, and legacy subtypes initially show all checked. Saving records the checked names. Cancelling leaves the configuration unchanged.

Selecting a subtype updates the configured grid to its effective selection, including before working copies are opened. A subtype with no selections has an empty configured grid and zero configured counts. Keep the established configuration-editing restrictions while a Word session is open.

Changing the selection updates configuration only. It does not add, remove, rename, or move bookmarks in Word.

### Discovery and validation

Use the effective subtype selection for configured rows, status counts, and configured-bookmark validation. Retain existing behavior: bookmarks absent on both sides are unused, and EN/FR discrepancies generate the existing warning. For example, selected `salutation` can exist only in English and retain its French-missing warning; the user accepts that warning.

Use the entire Template Type master list, rather than the subtype selection, when deciding whether a document bookmark belongs in Other Bookmarks. Consequently, a known but unselected normal bookmark appears in neither the configured grid nor Other Bookmarks and generates no validation report or warning.

Continue discovering `INFO_n` bookmarks from Word independently of the checklist. Preserve their separate grid, numbering, mismatch warnings, and operations.

### Error handling and edge cases

- Empty selections and empty master lists are valid.
- Preserve existing subtype-name and template-path validation and configuration-save error handling.
- Ignore stale selection names that are no longer in the master list when computing the effective selection; clean them when saving the relevant configuration.
- Preserve the bookmark-grid scrolling fix when refreshing filtered rows, including when the resulting list is shorter or empty.

## Testing Strategy

Use focused checks for legacy missing/null lists versus explicit empty lists, JSON round trips, new-subtype defaults, restoring checkbox selections, cancellation, independent subtype selections, master-list additions/removals, and master ordering. Verify grid/status/validation filtering, that known unselected bookmarks are ignored, that truly unknown bookmarks remain Other Bookmarks, and that English-only selected bookmarks retain existing warnings. Reuse the WinForms scroll regression checks and build the application. Manually confirm subtype editing/switching and bookmark operations with Word.

## Open Questions

None. The user approved the design with the corrected default of no bookmarks selected for new subtypes.
