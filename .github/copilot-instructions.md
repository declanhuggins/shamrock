# Copilot Instructions for SHAMROCK

- **Stack**: Google Apps Script (V8) written in TypeScript, built locally in VS Code, deployed via `clasp`. See `appsscript.json` for runtime/services and `package.json` scripts.
- **Data domain**: HR-facing Google Sheets workbooks plus Google Forms; expect handlers for form submissions, sheet reads/writes, and triggers (onOpen/onEdit/time-driven).

## Repository layout conventions
- `src/forms/` form handlers (e.g., `onFormSubmit`), validation, mapping form responses to typed objects.
- `src/sheets/` helpers for specific tabs/column maps, read/write utilities.
- `src/services/` business logic orchestration; keep side effects here.
- `src/triggers/` installable trigger entry points; export these from the main entry (e.g., `src/index.ts` once added).
- `src/utils/` shared helpers; `src/config/` for sheet/form IDs and flags (no secrets in git); `src/types.ts` for shared contracts.
- `docs/AI_CONTRIBUTION_GUIDE.md` defines these patterns and workflows; follow it for new features.

## Key workflows
- Install deps: `npm install`.
- Build TS -> GAS: `npm run build` (outputs to `dist/`).
- Deploy to Apps Script: `npm run push` (requires `clasp` auth to the target script).
- Pull remote script: `npm run pull`; clean build output: `npm run clean`.
- When adding entry points, ensure they are exported for Apps Script to call (e.g., `onOpen`, `onEdit`, form handlers).

## Documentation expectations
- Public docs live in `docs/public/README.md`; add one section per feature using `docs/templates/FEATURE_PUBLIC_DOC_TEMPLATE.md`.
- Keep docs free of secrets/IDs; describe sheet/form names with friendly labels and include manual validation steps.
- Update `README.md` if scripts, setup steps, or structure change.

## Patterns and guardrails
- Keep modules small and separated: parse/validate in `forms/`, persistence in `sheets/`, orchestration in `services/`.
- Prefer declarative config for sheet/form IDs in `src/config/`; avoid hardcoding sensitive values.
- Make functions pure where possible; avoid global state. Log or describe manual validation steps in feature docs.
- Default advanced service: Google Sheets API v4 (see `appsscript.json`).
- Before coding, capture intent and data contracts (see AI guide); document the feature immediately after implementation.
