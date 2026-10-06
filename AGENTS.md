# Agent Instructions

This repository is developed in short iterative sessions and may be worked on by different agents.

## Before making changes

Read these files in this order:

1. `docs/progress.md`
2. `docs/design.md`
3. `docs/implementation-plan.md`

Treat them as the authoritative project context.

## Development style

This is a quick internal C# utility, not a long-lived production product.

- Prefer simple, direct implementations.
- Avoid unnecessary abstractions and framework complexity.
- Do not introduce extensive automated testing unless specifically requested.
- Do not spend time on strict TDD.
- Keep the application stable enough for manual testing by the user.
- Make small, understandable commits.
- Do not redesign established architecture without recording the reason in `docs/progress.md`.

## Critical Word-document rule

This tool is a bookmark editor, not a Word document reformatter.

Do not introduce a document-processing path that reconstructs or rewrites the Word document through a custom renderer, HTML conversion, or general-purpose .docx regeneration.

Microsoft Word should remain responsible for document rendering and saving. Bookmark operations should use the Word object model against Word ranges/bookmarks.

## Bookmark placement

Bookmark placement must remain deterministic and user-controlled.

Do not use AI, fuzzy matching, language matching, or automatic text recognition to choose bookmark ranges.

## Progress tracking

At the end of every meaningful development session, update `docs/progress.md` with:

- what was completed;
- important files/classes changed;
- decisions made;
- known issues;
- what the user has manually tested;
- the exact recommended next step.

Keep the "Current State" section accurate enough that a new agent can resume without reading chat history.

If the architecture or planned phases change, update the corresponding design/implementation document as well.
