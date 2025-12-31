# AI Contribution Guide

Guidance for AI agents building the multi-sheet Sheets/Forms HR system with Apps Script and TypeScript. Use this as the default workflow until project-specific requirements are provided.

## Repository facts
- Target stack: Google Apps Script (V8) written in TypeScript, built locally in VS Code, deployed with `clasp`.
- Data domain: HR-facing Google Sheets workbooks plus Google Forms that trigger workflows.
- Safety: never commit secrets, personally identifiable information, or real sheet/form IDs without clearance. Keep IDs in a dedicated config module that can be gitignored if sensitive.

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

## Standard flow for a new feature
1) **Clarify intent**: capture the requirement in a short note (who/what/when) before coding.
2) **Plan the shape**: decide which sheet(s) and form(s) are touched; list triggers; sketch the data contract in `src/types.ts`.
3) **Scaffold code**:
   - Add or update `src/config/<feature>.ts` for IDs and column names.
   - Create a form handler in `src/forms/` that maps inputs to typed objects.
   - Add sheet helpers in `src/sheets/` to read/write structured data.
   - Put orchestration/business logic in `src/services/`; keep side effects here.
   - Register triggers/entry points in `src/triggers/` and export them in `src/index.ts` (or the main entry file once created).
4) **Document immediately**:
   - Add a feature entry to `docs/public/README.md` using `docs/templates/FEATURE_PUBLIC_DOC_TEMPLATE.md`.
   - Ensure any new config knobs or deployment steps are described in the public entry if they are safe to share.
5) **Update README**: reflect new scripts, setup steps, or dependencies; keep the project overview accurate.
6) **Validate**: describe manual test steps in the feature doc; if automated tests are added later, link them there.

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

## Defaults and guardrails
- Prefer declarative configs for sheet/form IDs so environments can change without code edits.
- Each feature should have a single entry point exported for triggers to call.
- Keep functions small and testable; avoid global state.
- When in doubt, write the public doc entry first; code to satisfy the documented behavior.

## To-do once requirements arrive
- Confirm actual sheet/form names, ID storage conventions, trigger strategy, and security requirements.
- Add linters/tests if introduced; extend this guide with the agreed patterns.
