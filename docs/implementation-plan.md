# Letter Wizard Template Bookmark Manager — Implementation Plan

This plan prioritizes a quick, stable internal tool over framework complexity or broad automated testing.

See `docs/design.md` for the authoritative architecture.

## Development approach

- C# WinForms.
- Microsoft Word Interop.
- Manual user testing in short iterations.
- Small commits.
- Minimal automated test investment.
- Build the smallest usable vertical slice first.
- Update `docs/progress.md` after each meaningful development session.

---

## Phase 0 — Architecture/design

**Goal:** establish the safe Word-document approach before implementation.

Status: **Complete**

Decisions:

- standalone WinForms manager;
- real Microsoft Word windows, not a custom .docx renderer;
- Word Interop for bookmark/range operations;
- temporary working copies protect originals;
- separate temporary copies for Test Mode;
- selectable Word layouts: Side by Side, English Focus, French Focus;
- JSON configuration;
- deterministic user-selected bookmark ranges.

Exit criterion:

Architecture accepted by user.

---

## Phase 1 — Application shell + Word control

**Status:** Complete / user-tested

**Goal:** prove that the app can reliably own the basic Word workflow.

Build:

- WinForms solution/project.
- Main manager window.
- Word Interop service/wrapper.
- Start a dedicated Word Application instance.
- Select English and French .docx files.
- Create unique session working directory.
- Copy sources to `Working_E.docx` / `Working_F.docx`.
- Open both working documents in Word.
- Close Word/documents safely when session exits.
- Basic layout commands:
  - Side by Side
  - English Focus
  - French Focus

Manual verification:

- complex existing templates visually remain intact;
- original files remain untouched;
- Word can save/reopen working copies normally;
- Side by Side and focus layouts are usable.

Do not proceed to complex bookmark UI until this basic Word lifecycle works reliably.

---

## Phase 2 — Template Type/Subtype configuration

**Status:** Complete / user-tested

**Goal:** establish the template organization model.

Build:

- Template Type selector.
- Subtype selector.
- Create/edit Template Type.
- Create/edit subtype.
- Bookmark-name list per Template Type.
- English/French source paths per subtype.
- JSON persistence.
- basic validation of bookmark names.

Keep configuration UI simple.

Manual verification:

- create Leave Without Pay type;
- add several subtypes;
- restart app;
- confirm configuration persists.

---

## Phase 3 — Bookmark discovery/status

**Status:** Complete / user-tested

**Goal:** make the application immediately useful for auditing templates.

Build:

- enumerate bookmarks in English document;
- enumerate bookmarks in French document;
- compare against Template Type bookmark list;
- display EN/FR status for each configured bookmark;
- display `Other Bookmarks`;
- refresh status command;
- clear visual mismatch indication.

Manual verification against real Letter Wizard templates.

This phase should not alter bookmark ranges yet.

---

## Phase 4 — Locate + selection handling

**Status:** Complete / user-tested

**Goal:** connect Word selections to the manager reliably.

Build:

- Locate English.
- Locate French.
- Locate Both if practical.
- use Word's independent per-window selections as the source for later editing actions;
- reject collapsed selections when an edit action requires a range;
- avoid extra manager-side capture state unless later testing proves it necessary.

The original explicit Capture workflow was removed after user testing because Word already preserves each document window's selection.

---

## Phase 5 — Bookmark creation/editing

**Status:** Complete / user-tested

**Goal:** deliver the core bookmark-authoring workflow.

Build:

- click bookmark name to add around the current EN/FR Word selections;
- delete bookmark markers;
- replace bookmark range;
- rename bookmark;
- refresh status automatically after operations;
- protect against obvious duplicates/invalid names;
- report failures without leaving the UI in an unknown state.

Manual verification should use a disposable copy of a real template and inspect it directly in Word.

---

## Phase 6 — INFO bookmarks

**Status:** Complete / user-tested

**Goal:** make repetitive informational bookmarks fast.

Build:

- detect `INFO_n` bookmarks in both documents;
- display INFO status;
- calculate next number as highest found in either document + 1;
- Add INFO using the current EN/FR Word selections;
- flag EN/FR INFO discrepancies;
- locate/delete/change INFO ranges using normal bookmark management where practical.

---

## Phase 7 — Test Mode

**Status:** Complete / user-tested

**Goal:** safely simulate Letter Wizard manipulation.

Build:

- save working documents;
- create `Test_E.docx` / `Test_F.docx`;
- open test pair;
- open Test_E.docx directly over Working_E.docx and Test_F.docx directly over Working_F.docx using the working windows' current native bounds;
- clear visual indication that app is in Test Mode;
- Set Value for selected bookmark;
- Delete bookmarked content;
- Reset Test;
- Close Test;
- ensure source and working documents are never changed by test operations.

Existing VBA semantics were reviewed before implementation: Set Value assigns bookmark Range.Text without recreating the bookmark; Delete uses Range.Delete on the bookmarked content/range.

---

## Phase 8 — Validation + final save

**Status:** Complete / user-tested

**Goal:** safely turn working copies into finalized templates.

Build:

- validation summary;
- EN/FR mismatch warnings;
- INFO mismatch warnings;
- Other Bookmark listing/warning;
- explicit Save/Finalize action;
- save through Word first;
- backup current source files;
- replace source files with working copies;
- success/failure reporting;
- prevent obvious accidental overwrite of wrong paths.

Manual verification:

- preserve a known original;
- finalize;
- compare resulting Word document visually and structurally in Word;
- confirm backups;
- run existing Letter Wizard against finalized test templates if practical.

---

## Phase 9 — Usability pass

Only after the full workflow works.

Possible improvements:

- remember last Word layout;
- keyboard shortcuts;
- faster subtype switching;
- confirmation/status bar;
- recent templates;
- better mismatch filtering;
- display bookmark text previews;
- optional unsaved-change indicator;
- cleanup abandoned session temp directories.

Avoid polishing features that do not materially speed template work.

---

## Future / optional

Not required for first functional release:

- multi-bookmark test scenarios;
- saved scenario JSON;
- side-by-side test/original comparison helpers;
- bulk validation across all configured subtypes;
- report/export of bookmark matrices;
- configuration import/export.

---

## Suggested first implementation session

Implement only Phase 1.

The first milestone should be:

> Launch app → choose two Word documents → safe working copies are created → Word opens both → user can switch Side by Side / English Focus / French Focus → closing app cleans up without modifying originals.

Once the user confirms that workflow feels right on their actual workstation/monitors, proceed to Template Type/Subtype configuration.
