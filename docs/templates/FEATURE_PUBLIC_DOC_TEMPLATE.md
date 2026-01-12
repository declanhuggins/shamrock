# Feature Public Doc Template

Copy this template into `docs/public/README.md` (one section per feature) and fill in all placeholders. Keep it safe for public sharing—never include secrets or raw IDs.

---

## Feature Name
- **Owner/POC**: role or team (no emails by default)
- **Status**: draft/active/deprecated
- **Last updated**: YYYY-MM-DD

### Overview
- One-paragraph summary of what the feature does and who uses it.
- Primary sheet(s)/form(s) involved.

### User entry points
- Forms: name, purpose, and where responses land.
- Manual actions: which tab/button/menu item to use.
- Triggers: onOpen/onEdit/onFormSubmit/time-driven with frequency.

### Data touched
- Sheets/Tabs: list each tab, key columns (name, type, required), and derived formulas if relevant.
- Config: friendly names for IDs/URLs required (use placeholders if not public).
- External dependencies: other systems/services the feature calls.

### Workflow (happy path)
1) Input -> validation -> sheet write.
2) Downstream actions (notifications, rollups, formatting, summaries).
3) Any branching or retries.

### Error handling and safeguards
- How invalid inputs are handled.
- Logging/alerts (if any).
- Rollback guidance (e.g., revert rows, disable trigger).

### Deployment / configuration
- Prerequisites (sheet ownership, sharing, add-ons, advanced services).
- Setup steps (copy IDs into config, install triggers, run initializers).
- How to verify deployment succeeded.

### Validation checklist
- Manual steps to confirm the feature works (what to click, what to expect).
- Sample inputs and expected outputs.

### Known limits / open questions
- Edge cases, performance constraints, or TBD items.

### Change log
- YYYY-MM-DD: short description of the change.
