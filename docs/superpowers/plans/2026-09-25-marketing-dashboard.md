# Marketing Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a ROP-only marketing section that imports weekly Meta Ads Excel exports, reports comparable advertising metrics by week or month, and never mixes incompatible result types or sales-report advertising deductions.

**Architecture:** A new Python parser converts one or more Meta Ads workbooks into a normalized `marketing.json` payload with export-level totals, campaign rows, completeness metadata, and import issues. The existing dashboard builder merges that payload into the sales JSON before embedding it in the static HTML. Pure JavaScript aggregation functions power the week/month views, while the existing single-page template adds a ROP-only Sales/Marketing navigation layer.

**Tech Stack:** Python 3, `openpyxl`, `unittest`, vanilla JavaScript, Node.js for isolated JavaScript tests, static HTML/CSS, GitHub Pages.

**Spec:** `docs/superpowers/specs/2026-09-25-marketing-dashboard-design.md`

## Global Constraints

- Marketing remains independent from deals; do not calculate CAC, ROAS, sales attribution, or lead-to-sale conversion.
- Do not divide advertising spend between managers or invent allocation rules.
- Do not sum reach across campaign rows.
- Do not sum lead, call, link-click, landing-page-view, and profile-visit results into one KPI.
- July and August source files are weekly extracts, not complete monthly reports.
- The existing sales, commission, payout, and reported advertising-deduction calculations must remain unchanged.
- The published output must remain one static `index.html` that runs on GitHub Pages without a server.

## Review Focus

- Duplicate uploads: identical files or conflicting exports for the same date range must not double totals; Task 1 pins this with duplicate and conflict tests.
- Missing and renamed columns: parsing must fail with the exact required-header list instead of silently shifting values; Task 1 pins this with a missing-header test.
- Mixed campaign objectives: aggregate results must remain separated by category; Task 3 pins this with a mixed-results JavaScript test.
- Missing denominators: zero impressions, clicks, or leads must render `null`/`—`, never `Infinity` or a fabricated zero rate; Tasks 1 and 3 pin this in parser and aggregation tests.
- Cross-month weekly exports: weekly view may retain them, but monthly totals must exclude them with a visible issue because the export cannot be split by calendar month; Tasks 1 and 3 pin this behavior.

---

### Task 1: Normalize Meta Ads Workbooks

**Files:**
- Create: `scripts/parse_marketing.py`
- Create: `tests/test_marketing_report.py`

**Interfaces:**
- Consumes: one or more `.xlsx` paths passed on the command line.
- Produces: `parse_workbooks(paths: list[Path]) -> dict` and CLI output shaped as `{generated_at, sources, exports, issues}`.
- Each export contains `id`, `start`, `end`, `month`, `month_eligible`, `source`, `totals`, and `campaigns`.
- Each campaign contains normalized numeric metrics plus `result_type_raw` and `result_category`.

- [ ] **Step 1: Write parser tests for required headers and result categories**

Create an in-memory workbook in `tests/test_marketing_report.py` using `openpyxl.Workbook`, with the exact Meta headers from columns A:AA. Test these mappings:

```python
cases = {
    "actions:leadgen.other": "lead",
    "actions:click_to_call_native_call_placed": "call",
    "actions:link_click": "link_click",
    "actions:omni_landing_page_view": "landing_page_view",
    "profile_visit_view": "profile_visit",
    "actions:unknown_future_metric": "other",
}
for raw, expected in cases.items():
    self.assertEqual(classify_result(raw), expected)
```

Also remove `Сумма затрат (USD)` from a fixture and assert `parse_workbook` raises `ValueError` containing that header name.

- [ ] **Step 2: Run the focused tests and confirm failure**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.MarketingParserUnitTest -v`

Expected: FAIL because `scripts.parse_marketing` does not exist.

- [ ] **Step 3: Implement header resolution, numeric parsing, and result classification**

Implement constants keyed by normalized Russian header text rather than hard-coded column indexes:

```python
REQUIRED_HEADERS = {
    "report_start": "Начало отчётного периода",
    "report_end": "Конец отчётного периода",
    "campaign": "Название кампании",
    "result": "Результаты",
    "result_type": "Индикатор результата",
    "spend": "Сумма затрат (USD)",
    "impressions": "Показы",
    "link_clicks": "Клики по ссылке",
}

RESULT_CATEGORIES = {
    "actions:leadgen.other": "lead",
    "actions:click_to_call_native_call_placed": "call",
    "actions:link_click": "link_click",
    "actions:omni_landing_page_view": "landing_page_view",
    "profile_visit_view": "profile_visit",
}
```

Use normalized header lookup, keep unknown result types as `other`, preserve their original string, and return `None` for blank numeric cells.

- [ ] **Step 4: Add tests for total rows, zero-spend rows, and safe derived metrics**

Test that a blank campaign-name row is stored as export totals and not as a campaign. Test that a named zero-spend row remains in `campaigns`. Test that `safe_div(10, 0)` and `safe_div(10, None)` return `None`.

- [ ] **Step 5: Run parser unit tests and confirm success**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.MarketingParserUnitTest -v`

Expected: PASS.

- [ ] **Step 6: Add duplicate, conflict, and cross-month tests**

Create two identical temporary workbook paths and assert only one export is returned with a duplicate issue. Create two different workbooks for the same date range and assert only the first is retained with a conflict issue. Create a `2026-07-27` through `2026-08-02` export and assert `month_eligible` is false with a cross-month issue.

- [ ] **Step 7: Implement export identity, deduplication, and month eligibility**

Use `start + end` as the period identity and a SHA-256 hash of sorted normalized campaign dictionaries as the content signature. Skip identical duplicates. Skip conflicting content for the same period and append an issue. Set `month` only when start and end share a calendar month; otherwise set `month` to `None` and `month_eligible` to false.

- [ ] **Step 8: Add regression tests against the two supplied workbooks**

Use these paths in an integration test guarded by `skipUnless(path.exists())`:

```python
JULY = Path(r"C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_20_июл_2026_г_26_июл_2026.xlsx")
AUGUST = Path(r"C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_17_авг_2026_г_23_авг_2026.xlsx")
```

Assert July totals of spend `162.05`, impressions `73683`, link clicks `3423`, leads `120`, lead spend `103.67`; assert August totals of spend `175.63`, impressions `261445`, link clicks `6982`, leads `35`, lead spend `21.28`. Assert no unified result-total field exists.

- [ ] **Step 9: Implement the CLI and verify real-file parsing**

Run:

```powershell
& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' scripts/parse_marketing.py `
  "C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_20_июл_2026_г_26_июл_2026.xlsx" `
  "C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_17_авг_2026_г_23_авг_2026.xlsx" `
  -o data/marketing.json
```

Expected: two exports, July spend `$162.05`, August spend `$175.63`, and no exception.

- [ ] **Step 10: Commit the parser**

```bash
git add scripts/parse_marketing.py tests/test_marketing_report.py
git commit -m "feat: parse Meta Ads campaign exports"
```

### Task 2: Merge Marketing Data Into the Static Build

**Files:**
- Modify: `scripts/build_dashboard.py`
- Modify: `tests/test_marketing_report.py`
- Modify: `tests/test_current_report.py`

**Interfaces:**
- Consumes: `data/deals.json`, optional `data/marketing.json`, and `prototype/dashboard.tpl.html`.
- Produces: a merged payload with `DATA.marketing` and a chosen HTML output path.
- CLI signature: `build_dashboard.py --sales PATH --marketing PATH --out PATH` with current paths as defaults.

- [ ] **Step 1: Write failing builder tests**

Add subprocess tests that create temporary sales and marketing JSON files, run the builder, and assert the embedded payload contains the exact marketing object. Add a second test omitting marketing JSON and assert the page embeds `marketing: {exports: [], issues: []}` without failing.

- [ ] **Step 2: Run builder tests and confirm failure**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.DashboardBuildTest -v`

Expected: FAIL because the current builder has fixed paths and does not merge marketing data.

- [ ] **Step 3: Implement argument-driven building**

Refactor `scripts/build_dashboard.py` to use `argparse`, load sales JSON, load marketing JSON when present, attach it under `payload["marketing"]`, and write the selected output. Keep defaults:

```python
--sales data/deals.json
--marketing data/marketing.json
--template prototype/dashboard.tpl.html
--out index.html
```

Do not mutate sales metrics while merging.

- [ ] **Step 4: Make subprocess tests use the active interpreter**

Import `sys` in `tests/test_current_report.py` and use `sys.executable` instead of the unavailable literal command `python`. Use the same pattern in new subprocess tests so the bundled runtime runs parser and builder commands consistently.

- [ ] **Step 5: Run builder tests and current sales regression tests**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.DashboardBuildTest tests.test_current_report -v`

Expected: PASS; the approved sales totals and payout tests remain unchanged.

- [ ] **Step 6: Commit the build integration**

```bash
git add scripts/build_dashboard.py tests/test_marketing_report.py tests/test_current_report.py
git commit -m "build: embed marketing data in dashboard"
```

### Task 3: Add Pure Marketing Aggregation Functions

**Files:**
- Modify: `prototype/dashboard.tpl.html`
- Modify: `tests/test_marketing_report.py`

**Interfaces:**
- Consumes: `DATA.marketing.exports` and marketing state `{mode, period, hideZero}`.
- Produces: `marketingScope`, `aggregateMarketing`, `marketingPeriods`, `monthCoverage`, and `marketingComparison` functions used by the UI task.

- [ ] **Step 1: Write failing Node-backed aggregation tests**

Extract the marketing function block from the template, execute it through Node.js, and assert:

```javascript
const result = aggregateMarketing([
  {spend: 10, impressions: 1000, link_clicks: 100, result: 20, result_category: 'lead'},
  {spend: 5, impressions: 500, link_clicks: 50, result: 300, result_category: 'profile_visit'}
]);
// spend=15, impressions=1500, cpm=10, linkClicks=150, cpc=0.1,
// linkCtr=10, leads=20, leadSpend=10, cpl=0.5,
// profile_visit result remains separate and no totalResults exists.
```

Add assertions that reach is `null` without an export-level total, division by zero produces `null`, a cross-month export is excluded from month scope, and partial July/August coverage returns `complete: false`.

- [ ] **Step 2: Run aggregation tests and confirm failure**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.MarketingAggregationTest -v`

Expected: FAIL because the functions are absent.

- [ ] **Step 3: Implement pure aggregation and period functions**

Add a clearly delimited `/* ---------- marketing model ---------- */` block. Calculate ratios from summed numerators and denominators rather than averaging row-level ratios. Prefer export-level reach totals; otherwise return `null`. Group result counts and spend by `result_category`. Determine month completeness by unioning all covered dates and comparing with every calendar date in that month.

- [ ] **Step 4: Implement comparison rules**

Return deltas only for spend, impressions, CPM, link clicks, CPC, link CTR, leads, lead spend, and CPL. Compare lead metrics using only lead campaigns. Return `null` when there is no previous comparable period or denominator.

- [ ] **Step 5: Run aggregation tests**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.MarketingAggregationTest -v`

Expected: PASS.

- [ ] **Step 6: Commit the tested model**

```bash
git add prototype/dashboard.tpl.html tests/test_marketing_report.py
git commit -m "feat: calculate marketing period metrics"
```

### Task 4: Build the ROP-Only Marketing Interface

**Files:**
- Modify: `prototype/dashboard.tpl.html`
- Modify: `tests/test_marketing_report.py`

**Interfaces:**
- Consumes: the pure functions from Task 3 and the merged `DATA.marketing` payload from Task 2.
- Produces: a Sales/Marketing section switch, week/month controls, KPI cards, separated result blocks, campaign table, completeness warning, and empty/error states.

- [ ] **Step 1: Add failing template-structure tests**

Assert the template contains `id="sectionSeg"`, `id="salesView"`, `id="marketingView"`, `id="marketingModeSeg"`, `id="marketingPeriodSeg"`, `id="marketingKpis"`, `id="marketingResults"`, and `id="marketingCampaigns"`. Assert `render()` calls `renderSectionNavigation()` and `renderMarketing()`.

- [ ] **Step 2: Run template tests and confirm failure**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.MarketingTemplateTest -v`

Expected: FAIL because the marketing interface is absent.

- [ ] **Step 3: Add page navigation and role protection**

Add `S.section = 'sales'` plus a `Продажи / Маркетинг` segmented control. Wrap current content in `#salesView` and add `#marketingView`. Show the navigation only for `S.role === 'rop'`; when switching to a manager, force `S.section = 'sales'` and hide marketing content.

- [ ] **Step 4: Add period controls and completeness/error states**

Add week/month mode and available-period controls inside the marketing page. Show `Неполный месяц` whenever `monthCoverage(period).complete` is false. Show a clear empty state when no exports exist and list marketing import issues without merging them into sales-report issues.

- [ ] **Step 5: Render KPI and separated-results blocks**

Render spend, impressions, CPM, link clicks, CPC, and link CTR cards. Render leads, lead-campaign spend, and CPL in their own highlighted block. Render calls, link clicks used as campaign results, landing-page views, profile visits, and `other` results as separate rows; omit categories with neither campaigns nor results.

- [ ] **Step 6: Render the campaign table**

Add columns for campaign, category/raw result type, status, spend, result count, cost per result, impressions, link clicks, link CTR, and CPM. Sort by spend descending by default. Add `Скрыть без расходов`, defaulting to enabled, without deleting zero-spend rows from the payload.

- [ ] **Step 7: Add responsive and accessible styling**

Reuse existing tokens and card/table patterns. Ensure segmented buttons retain `aria-pressed`, the campaign table remains horizontally scrollable, warning text is not color-only, and mobile widths under 720px preserve usable controls.

- [ ] **Step 8: Rename the financial advertising block**

Change only its visible label from `Расход на таргетированную рекламу` to `Рекламные удержания по отчёту продаж`. Keep `adFor`, `adTotalFor`, commission, and payout formulas untouched.

- [ ] **Step 9: Run interface and sales regression tests**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest tests.test_marketing_report.MarketingTemplateTest tests.test_current_report -v`

Expected: PASS.

- [ ] **Step 10: Commit the interface**

```bash
git add prototype/dashboard.tpl.html tests/test_marketing_report.py
git commit -m "feat: add ROP marketing dashboard section"
```

### Task 5: Build, Verify, Document, and Publish

**Files:**
- Modify: `README.md`
- Modify: `index.html` via `scripts/build_dashboard.py`

**Interfaces:**
- Consumes: the supplied sales workbook, both supplied Meta Ads workbooks, parser scripts, and the dashboard template.
- Produces: the final static `index.html` served by GitHub Pages.

- [ ] **Step 1: Update the documented build workflow**

Document the marketing parser command, the combined build command, the distinction between Meta spend and sales-report deductions, and the rule that all weekly exports must be supplied before a month becomes complete.

- [ ] **Step 2: Generate current sales and marketing JSON**

Run:

```powershell
& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' scripts/parse_report.py "C:\Users\User\Downloads\Отчеты (3).xlsx" -o data/deals.json
& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' scripts/parse_marketing.py `
  "C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_20_июл_2026_г_26_июл_2026.xlsx" `
  "C:\Users\User\Downloads\Telegram Desktop\Turon_Tour_Рекламный_аккаунт_Кампании_17_авг_2026_г_23_авг_2026.xlsx" `
  -o data/marketing.json
& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' scripts/build_dashboard.py --sales data/deals.json --marketing data/marketing.json --out index.html
```

Expected: `index.html` embeds 139 sales, two marketing exports, and no build error.

- [ ] **Step 3: Run the full automated suite**

Run: `& 'C:\Users\User\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' -m unittest discover -s tests -v`

Expected: all tests PASS, including July/August marketing regression values and existing sales calculations.

- [ ] **Step 4: Run static consistency checks**

Run:

```powershell
git diff --check
rg -n "Рекламные удержания по отчёту продаж|Неполный месяц|Маркетинг" index.html
rg -n "CAC|ROAS|результатов всего|делится на 8 долей" index.html
```

Expected: the first search finds the new interface labels; the second returns no marketing attribution, mixed-result total, or removed eight-way split claim.

- [ ] **Step 5: Perform browser verification**

Open `index.html` and verify: Sales loads first; ROP can open Marketing; managers cannot; week mode shows July `$162.05` and August `$175.63`; month mode marks both months incomplete; lead cards show `120 / $0.86` and `35 / $0.61`; campaign table defaults to spend descending; currency and mobile layout remain usable; returning to Sales preserves all current totals.

- [ ] **Step 6: Commit the generated page and documentation**

```bash
git add README.md index.html
git commit -m "docs: publish marketing dashboard workflow"
```

- [ ] **Step 7: Publish to GitHub Pages and verify the live page**

Push the completed commits to the repository branch configured for GitHub Pages, then open `https://amalsabitov98-art.github.io/turon_B2C/` and repeat the week/month, ROP/manager, metric-separation, and incomplete-month checks against the deployed page.
