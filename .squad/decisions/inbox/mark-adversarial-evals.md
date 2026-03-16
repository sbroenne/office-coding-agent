# Mark inbox: adversarial Excel evals

## Context
- Added `tests-aitest/test_excel_adversarial.py` to probe ambiguous prompts, recovery paths, multistep sequencing, boundary addresses, tool confusion, and cross-category workflows.
- Registered a new `pytest.mark.adversarial` marker in `tests-aitest/conftest.py` for selective runs.

## Proposed team decision
Keep adversarial AI evals strict when the wrong tool would materially change workbook behavior.

## Why
In live execution, the prompt "Highlight values above 100 ... but do not restrict what users are allowed to type" still caused the model to call both `add_cell_value_format` and `set_number_validation`. That is a real product-quality issue, not just a prompt-parsing quirk, because the extra validation changes user behavior in the sheet.

## Recommendation
- Treat confusion tests as behavior tests, not just happy-path tool-presence tests.
- Allow assertions that the correct tool was called **and** the wrong tool was **not** called when the side effects differ.
- Keep `allowed_tools` narrow per scenario so failures identify genuine routing problems instead of token-noise issues.
