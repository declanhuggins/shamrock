# AI Contribution Guide

Guidance for AI agents building the multi-sheet Sheets/Forms HR system with Apps Script and TypeScript. Use this as the default workflow until project-specific requirements are provided.

This guide explains how to work in the repository. System invariants live in `docs/system/SYSTEM_SPEC.md`.

## Repository facts
- Target stack: Google Apps Script (V8) written in TypeScript, built locally in VS Code, deployed with `clasp`.
- Data domain: HR-facing Google Sheets workbooks plus Google Forms that trigger workflows.
- Safety: never commit secrets, personally identifiable information, or real sheet/form IDs without clearance. Keep IDs in a dedicated config module that can be gitignored if sensitive.

## Canonical documents (read these first)
- System invariants/spec (internal): `docs/system/SYSTEM_SPEC.md`
- Operator runbook (internal): `docs/runbooks/OPERATOR_RUNBOOK.md`
- AI feature-change checklist: `docs/ai/FEATURE_CHANGE_CHECKLIST.md`
- Public feature catalog: `docs/public/README.md`
- Public feature template: `docs/templates/FEATURE_PUBLIC_DOC_TEMPLATE.md`

## File and module layout
- `src/forms/`: form handlers (e.g., `onFormSubmit`), input validation, and mapping to domain objects.
- `src/sheets/`: sheet/table-specific helpers (read/write rows, column maps, formatting).
- `src/services/`: business logic that orchestrates forms, sheets, and external services.
- `src/triggers/`: installable trigger entry points (time-driven, onOpen, onEdit).
- `src/utils/`: generic helpers (dates, logging, guards).
- `src/config/`: sheet/form IDs, named ranges, feature flags; keep secrets out of source control.
- `src/types.ts`: shared types and enums for data contracts between modules.
- `docs/public/`: public-facing documentation that can be shared with non-developers.
- `docs/templates/`: reusable doc templates (e.g., feature pages).

Keep modules small: isolate form parsing from sheet writes, and keep reusable logic in `services/` or `utils/`. Prefer pure functions where possible to ease testing.

## System invariants (must always hold)
These are summarized here for convenience; the canonical version is in `docs/system/SYSTEM_SPEC.md`.

- Provisioning must be idempotent ensure-exists (safe to re-run; no duplicates).
- Tables are header-driven:
   - Row 1 contains machine-friendly stable headers.
   - Tables start at row 2.
   - Never assume column positions.
- Dropdown validations must be driven from Data Legend ranges.
- Frontend sheets are protected; users interact via forms and menus.
- All Google Forms must require verified responder emails.

## Standard flow for a new feature
1) **Clarify intent**: capture the requirement in a short note (who/what/when) before coding.
2) **Confirm invariants**: review `docs/system/SYSTEM_SPEC.md` and note any constraints that apply.
3) **Document first (public)**:
   - Add/extend a feature entry in `docs/public/README.md` using `docs/templates/FEATURE_PUBLIC_DOC_TEMPLATE.md`.
   - Include entry points (menus/triggers/forms), data touched (by header name), and a manual validation checklist.
4) **Plan the shape**: decide which sheet(s) and form(s) are touched; list triggers; define a data contract.
5) **Scaffold implementation** (later): segment parsing (forms), IO (sheets), orchestration (services), and entry points (triggers).
6) **Update runbooks/spec**: if operator steps or invariants changed, update the internal docs.

## How to update the README
- Keep the top-level description concise and aligned with the current scope.
- Maintain sections for:
  - **Getting Started**: install deps, build, push with `clasp`, and how to set script properties/IDs.
  - **Scripts**: explain the npm scripts and when to use them.
  - **Structure**: short overview of `src/` and `docs/` layout from this guide.
  - **Public Docs**: link to `docs/public/README.md`.
- When adding features, only surface high-level details (what the feature does, which sheet/form it touches). Deep technical steps stay in the public doc entry or inline code comments when essential.

## How to write public-facing feature docs
- Edit `docs/public/README.md`, copying the template from `docs/templates/FEATURE_PUBLIC_DOC_TEMPLATE.md`.
- Keep the tone operational: describe user flows, inputs/outputs, permissions, and how to run or troubleshoot the feature.
- Avoid internal jargon or secrets; use placeholders if an ID or URL should not be published.
- Include manual validation steps and a brief rollback note for operational safety.

Additions for this repo:
- Describe provisioning as “ensure-exists” and state that it is safe to re-run.
- Refer to columns by header name (row 1), not by column letter or index.
- Explicitly state which forms require verified responder emails.

## Defaults and guardrails
- Prefer declarative configs for sheet/form IDs so environments can change without code edits.
- Each feature should have a single entry point exported for triggers to call.
- Keep functions small and testable; avoid global state.
- When in doubt, write the public doc entry first; code to satisfy the documented behavior.

## To-do once requirements arrive
- Confirm actual sheet/form names, ID storage conventions, trigger strategy, and security requirements.
- Add linters/tests if introduced; extend this guide with the agreed patterns.
