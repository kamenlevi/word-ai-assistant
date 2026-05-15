# Eval Routine

## What this does

`eval/run.js` evaluates the Word AI assistant against the test cases in `cases.json` (plus any auto-generated cases in `generated-cases.json`). For each case:

1. Constructs a prompt with the document state, the user's request, and the current system prompt extracted from `server.js`.
2. Calls Llama 3.3 70B via OpenRouter.
3. Extracts the `CODE_JS::...::END_CODE` block (or `NEEDS_WEB_SEARCH::...::END` for live-data probes).
4. Checks `requiredPatterns` and `forbiddenPatterns` against the generated code.
5. Scores 0–100 via an LLM judge on four dimensions (API_CORRECTNESS, COMPLETENESS, WOULD_IT_WORK, APPROACH).

After all cases run:

- **Mastery promotion**: Any category averaging ≥95 gets 4 new harder cases generated at the next level.
- **New category**: When overall avg ≥95, a brand-new Word domain category gets 2 starter cases (≥90 promotes a level-2 category).
- **Stuck-category easing**: A category stuck on the same level for ≥3 runs gets 3 easier cases.
- **System prompt patch**: Any failures <70 trigger an LLM-generated rule/example appended to `eval/improvements.txt`, which the server prepends to the system prompt on next start.

## How to run

```
OPENROUTER_KEY=sk-or-v1-... npm run eval
```

Budget cap: $0.50 per session (hard limit, raises an exception if exceeded). Typical cost: $0.02-0.08.

## Files

- `cases.json` — base test cases (committed manually)
- `generated-cases.json` — auto-generated cases (CI commits these)
- `progress.json` — per-category levels & mastery timestamps
- `improvements.txt` — auto-generated system prompt additions
- `RESULTS.md` — latest scoreboard with trends
- `results/<timestamp>.json` — full result logs

## CI

`.github/workflows/eval.yml` runs `npm run eval` every 8 hours and commits any changes (improvements.txt, progress.json, generated-cases.json, RESULTS.md, last-run.txt).

`watchdog.js` is a local fallback for machines that should run the eval if CI missed its window — wire it to Task Scheduler via `watchdog.bat`.
