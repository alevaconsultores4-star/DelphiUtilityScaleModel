# AI Guardrails for this Repo

## Goals
- Keep financial calculations correct and stable.
- Allow UI/UX iteration without breaking the model.
- Prevent large unreviewed refactors.

## Hard Rules (Do Not Break)
1) Do NOT change financial logic (IRR, NPV, cashflow build, debt schedule, taxes) unless explicitly instructed.
2) All financial logic must live in /engine (or clearly labeled “MODEL” section). UI files should not contain business logic.
3) Do NOT introduce global mutable state. Use st.session_state only for UI state.
4) Do NOT change units/currency conversion logic without documenting it.
5) Any change affecting outputs must be validated with the Base Case check (below).

## Output Conventions
- Project IRR = unlevered cash flows (FCFF).
- Equity IRR = levered cash flows (after debt service).
- If debt is disabled, equity cash flow must revert to unlevered (or equity=project by definition).

## Base Case Regression Check
- Inputs file: /tests/base_case.json (or described below)
- Expected outputs:
  - Project IRR: __.__ %
  - Equity IRR: __.__ %
  - NPV @ WACC: ______

## Allowed Changes Without Extra Review
- Layout changes (tabs, columns, labels)
- Formatting (tables, charts, number formats)
- New UI controls that feed existing model inputs

## How to Work
- Make small commits.
- Prefer minimal diffs.
- When unsure, ask for confirmation before refactoring.
