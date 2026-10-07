# Letter Wizard Template Bookmark Manager — Design

Status: **Accepted architecture; implementation in progress (Phases 1-7 implemented)**

## 1. Purpose

The Letter Wizard Template Bookmark Manager is a dedicated Windows utility for creating, inspecting, maintaining, validating, and testing bookmarks in English and French Microsoft Word templates used by the existing Excel/VBA Letter Wizard.

The Letter Wizard generates final letters by setting bookmark values or deleting bookmarked content. Building and maintaining those bookmarks manually in Word is slow and error-prone.

The manager should make bookmark work faster while keeping placement completely deterministic and under developer control.

AI will not be used for bookmark placement.

---

## 2. Primary workflow

The intended workflow is:

1. Select a Template Type.
2. Select or create a Template Subtype.
3. Load its English and French templates.
4. Create safe working copies.
5. Open both working copies in Microsoft Word.
6. Select corresponding English and French text.
7. Click a bookmark name in the manager.
8. Repeat for required bookmarks.
9. Add sequential INFO bookmarks as required.
10. Review English/French bookmark status and discrepancies.
11. Locate, change, rename, or delete bookmarks when required.
12. Test bookmark behavior using separate temporary test copies.
13. Validate the pair of templates.
14. Explicitly save/finalize the templates.

---

## 3. Critical document-preservation requirement

The application is a **bookmark editor, not a Word document reformatter**.

Templates may contain important Word structures unrelated to bookmarks, including:

- styles;
- outline/navigation headings;
- table-of-contents structures and fields;
- headers and footers;
- sections;
- page and section breaks;
- tables;
- accessibility-related structure;
- fields;
- metadata;
- other Word-specific formatting and document structures.

The manager must modify only the bookmark structures/ranges needed for the requested operation.

### Architectural rule

**Microsoft Word is the document rendering and editing engine.**

The application must not build its own .docx editor or round-trip templates through HTML, RTF, a custom renderer, or another document model for routine bookmark editing.

Word should open and save the documents through its native object model.

This does not promise byte-for-byte package preservation because Microsoft Word itself can update internal package metadata during a normal save. The preservation target is equivalent to manually opening the document in Word, editing bookmarks, and saving it.

---

## 4. Application architecture

### 4.1 Desktop application

Use a standalone **C# WinForms** application.

Reasons:

- fast to build;
- appropriate for an internal Windows developer utility;
- simple desktop UI;
- straightforward COM/Word Interop integration;
- no need for heavier WPF/WinUI architecture.

### 4.2 Word integration

Use `Microsoft.Office.Interop.Word` to control a locally installed Microsoft Word instance.

The manager owns/coordinates:

- launching or attaching to its Word instance;
- opening the English/French working documents;
- reading Word selections/ranges;
- enumerating bookmarks;
- adding bookmarks around user-selected ranges;
- locating/selecting bookmark ranges;
- deleting bookmarks and bookmarked content when requested;
- arranging Word windows;
- saving working/test copies through Word;
- closing managed documents cleanly.

Do not attempt unsupported embedding of Word into a custom WinForms document surface.

### 4.3 Word windows and layout

English and French templates remain real Word windows.

The manager provides layout commands instead of imposing one layout:

- **Side by Side**
- **English Focus**
- **French Focus**

Side by Side should use native Windows window positioning so the two Word windows occupy the left and right halves of the monitor work area accurately.

The selected layout should eventually be remembered as a user preference.

The bookmark manager should remain relatively narrow so most screen area is available to Word.

---

## 5. Safe editing model

Original source templates should not normally be opened for editing.

When a subtype is loaded:

1. Create a unique session directory under the user's temporary/local application area.
2. Byte-copy the English source template to a working English document.
3. Byte-copy the French source template to a working French document.
4. Open only the working copies in Word.

Example:

```
Source
  LWOP_Sick_E.docx
  LWOP_Sick_F.docx

Session
  Working_E.docx
  Working_F.docx
```

### Benefits

- crashes do not directly corrupt source templates;
- Cancel can discard the session;
- test operations cannot affect source files;
- working documents can be saved frequently;
- final template replacement is explicit.

### Final save

A final save/finalize action should:

1. ask Word to save working documents;
2. validate that the intended source paths still exist/are writable;
3. create backups of the current source templates;
4. copy/replace the sources using the working files;
5. report success/failure clearly.

Backups should be simple and easy to locate.

---

## 6. Template model

### 6.1 Template Type

A Template Type defines the complete set of normal bookmark names available to a family of templates.

Example:

`Leave Without Pay`

Possible bookmark definitions:

- EMPLOYEE_NAME
- START_DATE
- END_DATE
- RETURN_DATE
- EMPLOYEE_SHARE
- EMPLOYER_SHARE

### 6.2 Template Subtype

A subtype identifies a specific English/French template pair.

Example:

- Sick Leave
- Personal Leave
- Care of Family
- Other

A subtype may use only a subset of the Template Type's bookmarks.

### 6.3 Storage

Use simple JSON configuration rather than a database.

Suggested structure:

```
TemplateTypes/
  LeaveWithoutPay.json
  Termination.json
  Hire.json
```

A subtype can store references to its English/French template paths once selected.

Configuration details can evolve as actual usage clarifies requirements.

---

## 7. Selection model

The user selects bookmark ranges directly in Word.

English and French are separate Word windows, and Word preserves the current selection in each window independently.

For operations that require new ranges, such as Add Bookmark and Replace Range:

1. select the desired English text in the English working copy;
2. select the desired French text in the French working copy;
3. return to the manager and choose the bookmark action;
4. the manager reads both Word selections at that moment;
5. reject the operation if either selection is collapsed;
6. create/recreate the bookmark around those ranges;
7. refresh bookmark status.

No explicit Capture step is required and the manager does not maintain duplicate selection state.

No fuzzy matching or automatic translation alignment is performed.

---

## 8. Bookmark list and status

The primary manager view contains the Template Type bookmark list.

Each row should show:

- bookmark name;
- English status;
- French status;
- relevant actions.

Example:

| Bookmark | EN | FR |
|---|---|---|
| EMPLOYEE_NAME | Exists | Exists |
| START_DATE | Exists | Exists |
| RETURN_DATE | Exists | Missing |
| EMPLOYEE_SHARE | Missing | Missing |

States should be visually obvious:

- both present;
- neither present;
- English only;
- French only.

### Other bookmarks

The application must also enumerate bookmarks that exist in either document but are not defined by the selected Template Type.

Display these separately as **Other Bookmarks**.

This prevents existing or legacy bookmarks from being silently hidden.

---

## 9. Bookmark operations

Required operations:

### Add

Create the selected configured bookmark around the current English and French Word selections.

Add must not silently overwrite an existing bookmark. If the name already exists in either working copy, the user should use Replace Range instead.

### Locate

Navigate Word to a bookmark and select/highlight its range.

Support:

- Locate English
- Locate French
- Locate Both

Locate commands must not reposition Word windows. Window placement changes are reserved for explicit layout commands such as **Side by Side**.

### Change range

User selects replacement English/French ranges directly in Word and chooses Replace Range.

The existing bookmark markers are replaced with bookmarks around the new ranges without changing the selected document text.

Replace Range may also repair an EN/FR mismatch by recreating the selected bookmark on both current Word selections.

### Rename

Where appropriate, preserve the bookmarked ranges while recreating them under a new valid bookmark name.

Rename behavior must guard against duplicate bookmark names.

Renaming a bookmark in the working documents does not automatically rename the shared Template Type configuration. Template Type definitions are shared across subtypes and must be edited explicitly when the configured name itself should change.

### Delete bookmark

Support removal of bookmark markers only. The bookmarked text must remain untouched.

The UI should distinguish this editing operation from the Letter Wizard **test delete** behavior that removes bookmarked content.

---

## 10. INFO bookmarks

INFO bookmarks are sequential informational regions removed by Letter Wizard during final letter generation.

Names use:

`INFO_1`, `INFO_2`, `INFO_3`, etc.

The manager provides a prominent **Add INFO** action.

Workflow:

1. select matching English text in the English working copy;
2. select matching French text in the French working copy;
3. click **Add INFO**;
4. scan both documents for existing `INFO_n` bookmarks;
5. determine the next number;
6. create the same INFO bookmark around the two current Word selections;
7. refresh INFO status.

INFO bookmarks are displayed in their own status area rather than under Other Bookmarks. Normal locate, replace-range, and delete-marker behavior is reused. Rename is disabled for INFO rows so sequential INFO naming remains predictable.

### Numbering rule

Use:

**highest INFO number found in either document + 1**

Do not fill old gaps automatically.

Example:

`INFO_1`, `INFO_2`, `INFO_4` => next is `INFO_5`.

INFO English/French discrepancies should be prominently reported.

---

## 11. Testing architecture

Bookmark testing must never operate on source templates or the active working copies.

When Test Mode begins:

1. save the current working document state through Word;
2. create a second set of temporary copies:
   - `Test_E.docx`
   - `Test_F.docx`
3. capture the current native screen bounds of the English and French working-copy windows;
4. open/display the test copies through Word;
5. position Test_E.docx directly over Working_E.docx and Test_F.docx directly over Working_F.docx using those captured bounds;
6. bring the test windows forward;
7. perform all simulated Letter Wizard operations on those test documents.

Starting Test Mode must not reposition the existing working-copy windows. Test-window placement follows the working windows rather than the bookmark-manager window, so multi-monitor/manual layouts are preserved.

### Required V1 test operations

For an individual bookmark:

**Set Value**
- replace bookmarked content with a supplied test value by assigning the bookmark range's Text;
- match the current Letter Wizard VBA behavior: do not recreate the bookmark after the replacement;
- INFO rows do not expose Set Value because INFO content is deletion-only in the Letter Wizard.

**Delete**
- simulate the Letter Wizard's deletion by deleting the bookmarked range/content;
- match the current VBA `Range.Delete` behavior.

### Reset test

Reset should discard test copies and recreate them from the current working copies.

Working documents are saved through Word before the new test snapshot is created.

### Close test

Close the test documents and return to bookmark-editing mode without touching working/source files.

### Future extension

A later version can support named test scenarios that apply combinations of Set Value/Delete operations to simulate a complete generated letter.

This is not required for the first functional release.

---

## 12. Validation

Before final save, provide a validation summary.

Initial checks:

- required manager configuration is valid;
- English source exists;
- French source exists;
- configured bookmarks show their EN/FR presence;
- mismatches are highlighted;
- INFO mismatches are highlighted;
- unexpected/Other bookmarks are listed;
- duplicate/conflicting situations are reported where Word permits detection.

Validation should inform the user rather than automatically changing document content.

---

## 13. UI concept

The manager is a control panel beside the actual Word windows.

Primary areas:

### Header

- Template Type
- Template Subtype
- Load/Create subtype
- current source paths
- layout controls
- Test Mode
- Save/Finalize

### Selection status

- English current selection
- French current selection

### Bookmark list

- name
- English state
- French state
- actions

### INFO area

- current INFO bookmarks/status
- Add INFO

### Validation/status area

- discrepancies
- warnings
- session status

The first version should favor clarity and speed over visual polish.

---

## 14. Scope deliberately excluded from V1

Do not initially build:

- custom Word rendering;
- embedded Word editing surfaces;
- AI bookmark placement;
- automatic English/French text matching;
- database storage;
- cloud synchronization;
- elaborate test automation;
- full Letter Wizard generation engine;
- complex template version control;
- Office web integration.

The goal is a small reliable developer utility.

---

## 15. Open implementation questions

These should be resolved through small manual prototypes during implementation:

1. Best reliable mechanism to capture and retain EN/FR Word ranges while the user switches between Word windows.
2. Exact window positioning behavior across multiple monitors/DPI settings.
3. How Word handles selection/range references after intervening edits and when a stored range should be invalidated.
4. ~~Exact Letter Wizard semantics for Set Value/Delete so Test Mode mirrors the VBA behavior.~~ Resolved in Phase 7: Set Value assigns bookmark `Range.Text` without recreating the bookmark; Delete removes the bookmarked range/content with `Range.Delete`.
5. Backup filename/location convention for final saves.

None of these changes the overall architecture.
