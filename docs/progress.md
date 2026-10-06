# Development Progress

**Last updated:** 2026-10-06

This file is the primary handoff source for continuing development.

Any agent resuming work should read this file first, then `docs/design.md` and `docs/implementation-plan.md`.

---

## Current State

**Stage:** Design complete; implementation not started.

The repository has an agreed initial architecture and phased implementation plan.

No C# application code has been written yet.

### Architecture locked in

- Standalone C# WinForms application.
- Microsoft Word remains the renderer/editor.
- Word automation through `Microsoft.Office.Interop.Word`.
- Do not build a custom .docx renderer/editor.
- Do not reconstruct templates through Open XML/HTML/RTF as the routine editing path.
- Original templates are protected by session working copies.
- Test Mode uses an additional disposable copy layer.
- Bookmark placement is deterministic and user-controlled.
- No AI/fuzzy text matching for bookmark placement.
- English/French Word windows can be arranged:
  - Top / Bottom
  - Side by Side
  - English Focus
  - French Focus
- Template Type/Subtype configuration will initially use JSON.

### Development philosophy

This is a quick internal utility.

Prioritize:

1. working behavior;
2. document safety;
3. stability;
4. development speed.

Avoid:

- strict TDD;
- extensive automated test suites;
- unnecessary abstraction;
- framework overengineering.

The user will manually test features as development proceeds.

---

## Completed

### 2026-10-06 — Initial design

Completed project brainstorming and established:

- project purpose and scope;
- Template Type / Template Subtype model;
- English/French paired-template workflow;
- bookmark creation from direct user selections;
- INFO bookmark numbering model;
- bookmark auditing and discrepancy detection;
- safe working-copy architecture;
- Test Mode architecture;
- Word-native document-preservation strategy;
- switchable Word-window layout;
- phased implementation plan;
- cross-agent handoff process.

Documentation created:

- `README.md`
- `AGENTS.md`
- `docs/design.md`
- `docs/implementation-plan.md`
- `docs/progress.md`

---

## User-tested functionality

None yet. Implementation has not started.

---

## Known Issues / Risks

### Word range retention

Need to determine the most reliable way to retain English/French selections while switching Word windows and editing.

Potential fallback: explicit Capture English / Capture French actions if event-driven selection tracking is unreliable.

### Word window positioning

Top/Bottom and Side-by-Side behavior needs real testing on the user's workstation, including DPI/multiple-monitor behavior.

### COM lifecycle

Word Interop objects must be managed carefully enough that the utility does not routinely leave orphaned WINWORD.EXE processes.

Keep the approach straightforward; do not over-engineer COM cleanup until observed behavior requires it.

### Test semantics

Set Value/Delete behavior should eventually be compared with the existing Letter Wizard VBA so testing accurately represents final generation.

---

## Next Step

**Implement Phase 1 only.**

Create the C# WinForms application shell and prove the Word/document lifecycle:

1. launch manager;
2. select English/French source documents;
3. create safe temporary working copies;
4. launch/open both copies in Word;
5. implement Top/Bottom, Side-by-Side, English Focus, French Focus window commands;
6. close session cleanly;
7. confirm originals were never modified.

After the user manually tests Phase 1, update this file with the outcome before continuing.

---

## Handoff Notes

Do not begin with bookmark CRUD yet.

The first implementation session is intentionally a small architectural proof using real Word documents. If the external Word-window workflow feels awkward in actual use, adjust that interaction before building the rest of the application around it.

When committing code, keep commits small enough that the user can sync and manually test specific features.

At the end of every session, update this progress file with the current commit/branch if known and the exact next recommended action.
