# Agent guide — 福岡 2026 travel handbook

`itinerary.xlsx` → `scripts/build_data.py` → `dist/itinerary_data.json` → `dist/index.html` (+ `sw.js`), published from `dist/` to GitHub Pages (jp.mikemomo.com). See README.md for the full pipeline.

## Regression testing (every change to dist/ or scripts/)

Run the existing suites. Do **not** write a new test harness from scratch — that is what made earlier verification rounds take 5–25 minutes; the suites below finish in about 10 seconds.

```bash
# once: browser test environment (uses the installed Microsoft Edge / Chrome, no browser download)
python3 -m venv .venv && .venv/bin/pip install playwright

# 1. static checks
for f in scripts/verify_*.py; do [ "$f" = scripts/verify_browser.py ] || python3 "$f"; done

# 2. browser regression (Playwright, 375x812, checks run concurrently)
.venv/bin/python scripts/verify_browser.py --shots <scratch dir>
.venv/bin/python scripts/verify_browser.py --only dark_contrast,theme_toggle   # a subset
```

`verify_browser.py` serves a temporary copy of `dist/`, so it never modifies project files, and it derives expectations from `itinerary_data.json`, so itinerary edits don't break it. Checks: `now_focus`, `tomorrow_reminder`, `offline_bar`, `offline_reload`, `single_shell_cache`, `manifest`, `theme_toggle`, `dark_contrast`, `data_update`.

### When you are the verifying agent
- Run both suites first and report their output verbatim; failures come with the failing values.
- A new feature or fix adds or extends a check in `scripts/verify_browser.py` (one async method on `Suite` plus its name in `CHECKS`); prove a new check can fail by breaking the feature once, then restore.
- Then review the diff (`git diff`) for real bugs only, with `file:line` and a concrete scenario.
- Put screenshots and reports in a scratch directory, never in the repo. Don't edit `dist/`, don't commit or push.
- For a larger review, split it: one agent runs the suites, others review separate concerns in parallel. Repetitive run-and-report work can use a fast model (e.g. `agy --model gemini-3.8-flash-medium`); keep a stronger model for code review.

## Project rules
- `dist/` is public: never publish booking or reservation numbers or the home address (`build_data.py` masks known ones).
- Edit `itinerary.xlsx` by rewriting the sheet XML inside the zip (back it up first). Saving with openpyxl drops the embedded images. The file is git-ignored, so only the regenerated JSON is committed.
- Timeline categories come from keyword matching in `build_data.py` (`CATEGORY_KEYWORDS`): the spot name decides first, remarks only when the name matches nothing.
