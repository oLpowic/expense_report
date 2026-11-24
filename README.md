# Expense report

🇵🇱 Polska wersja: [README.md](README.md)

The script reads bank exports (xlsx), assigns each expense to a category and builds
Excel reports with charts: monthly and yearly.

> Note: the reports, category names and column names are in Polish (they match the bank
> export). Polish names are given in quotes below, with an English gloss where needed.

## Requirements

- Python 3.10+
- libraries: `pip install pandas openpyxl`

## Quick start

1. Copy the configuration template (first run only):
   ```
   cp -r config.example config
   ```
   Then add your own shops and services to `config/rules.json` (see below).
2. Export your transaction history from the bank as xlsx and put the file in `data/exports/`.
   (You can keep adding newer exports next to it. Overlapping days are not counted twice.)
3. In the project folder run:
   ```
   python main.py
   ```
4. Your report is in `output/<year>/<yyyy-mm>.xlsx`.

## Commands

| Command | What it does |
|---|---|
| `python main.py` | monthly report for the latest month in the data |
| `python main.py 2026-11` | monthly report for the chosen month |
| `python main.py 2026-11 everyday` | monthly report with the `everyday` filter |
| `python main.py year` | yearly report for the year of the latest transaction |
| `python main.py year 2026` | yearly report for the chosen year |
| `python main.py year 2026 everyday` | yearly report with a filter |
| `python main.py help` | short help |

The yearly report is also created automatically: if you run `python main.py` in January and
`output/<previous year>_roczny.xlsx` does not exist yet, the script creates it.
You can also generate the yearly report earlier, at any time (an incomplete year is clearly
marked and gets overwritten on the next run).

## File structure

```
expense_report/
├── main.py                 ← the only file you run
├── src/
│   ├── loader.py           ← reads exports, merges them, removes duplicates
│   ├── categorize.py       ← categories, refunds, filters
│   ├── report_monthly.py   ← monthly report
│   └── report_yearly.py    ← yearly report
├── config.example/         ← configuration TEMPLATE (generic, safe to keep in git)
│   ├── settings.json
│   ├── categories.json
│   ├── rules.json
│   └── filters/
├── config/                 ← YOUR configuration (a copy of the template, not in git)
│   ├── settings.json       ← settings (column names, refunds, pending holds...)
│   ├── categories.json     ← list of categories and their type
│   ├── rules.json          ← keywords that map transactions to categories
│   └── filters/            ← filters (one JSON file each)
├── data/exports/           ← put your bank exports here (not in git)
└── output/
    ├── 2026/2026-11.xlsx   ← monthly reports for a given year (not in git)
    └── 2026_roczny.xlsx    ← yearly report ("roczny" = yearly)
```

## Privacy and git

Bank exports and reports contain real transactions, and `rules.json` reveals your habits
(shops, university, investments). Do not commit them. Suggested `.gitignore`:

```
# bank data and reports
data/
output/

# personal configuration (only config.example/ stays in git)
config/

# Python
__pycache__/
*.pyc
```

Keep only the `config.example/` template in the repository (generic rules, no personal data).
If you add something new and generic to your config (e.g. a new filter), copy it to
`config.example/` too. If you ever change the config format, update the template as well so
that a fresh copy still works.

## How categorisation works

The script joins the text of the `Odbiorca` (recipient) and `Opis` (description) columns,
lowercases it and goes through the rules in `config/rules.json` from top to bottom.
**The first rule whose keyword appears in the text wins.** If no rule matches, the
transaction goes to the "Inne" (Other) category.

### Adding a new shop
Add a keyword (lowercase) to the right category in `config/rules.json`:
```json
{"category": "Zakupy spożywcze", "keywords": ["biedronka", "lidl", "your new shop"]}
```
A keyword can be a fragment of the name. Avoid very short words (e.g. "bar"), they may match
something unrelated. Keep specific rules above general ones.

### Adding a new category
1. Add it to `config/categories.json` with a type: `fixed` (recurring expense), `variable`,
   or `savings` (investments, shown separately and excluded from total spending).
2. Add a rule for it in `config/rules.json`.

## What happens to each transaction

- **Expenses:** negative amounts.
- **Refunds:** positive amounts paid by card or BLIK are subtracted from the category matched
  by the description (e.g. a refund from Reserved reduces "Ubrania i zakupy", i.e. clothes
  and shopping).
- **Other incoming money** (salary, transfers from people): ignored.
- **Rejected and cancelled:** ignored.
- **Pending holds ("W trakcie realizacji" = in progress):** ignored by default, because they
  may show up a second time once booked. Change it in `settings.json` (`include_pending`).
- **Investments** (e.g. a deposit to a brokerage): counted separately, not part of total
  spending.

## Filters

A filter is a JSON file in `config/filters/`. All conditions in it must be met.
Fields (all optional):

| Field | Meaning |
|---|---|
| `type` | only type `fixed` or `variable` |
| `categories` | only these categories |
| `exclude_categories` | everything except these categories |
| `min_amount` | only single expenses from this amount up |

Available filters: `fixed_only`, `variable_only`, `big_expenses`, `everyday`, `food`,
`leisure`, `transport_housing`, `education`. A new filter is a new file, e.g.
`config/filters/mine.json`, then `python main.py 2026-11 mine`.

## What is in the report (xlsx)

- **"Podsumowanie" (Summary):** category table (total, share, change vs previous month),
  fixed vs variable, investments separately, top 5 expenses, chart.
- **"Porównanie" (Comparison, monthly) / "Miesiące" (Months, yearly):** categories × months,
  charts.
- **"Transakcje" (Transactions):** all rows with their category, with filters in the headers.
- **"Do przypisania" (To assign):** transactions in "Inne", i.e. what is worth adding to the
  rules.

Numbers in the report are calculated with formulas from the "Transakcje" sheet. If you change
a transaction's category there, totals and charts recalculate by themselves. In the monthly
report you can also change the yellow month cell (stored as `yyyymm`, e.g. `202610`).

Months for which the export does not cover every day are marked **"niepełny"** (incomplete).
Do not compare them 1:1 with complete months.

## Settings (`config/settings.json`)

| Key | Meaning |
|---|---|
| `input_folder` | folder with the exports |
| `columns` | column names in the bank file |
| `text_columns` | columns searched for keywords |
| `skip_status_contains` | statuses to skip (text fragments) |
| `include_pending` | `true` = also count pending holds |
| `pending_status_contains` | how to recognise pending holds in the Status column |
| `net_refunds` | `true` = subtract refunds from expenses |
| `refund_types` | transaction types where a positive amount is a refund |
| `default_category` | category for unrecognised transactions |
| `compare_months` | how many previous months in the comparison (default 6) |
| `auto_yearly_in_january` | `true` = create last year's yearly report in January |

## Common problems

- **"Brak folderu config/" (no config/ folder):** copy the template: `cp -r config.example config`.
- **"Brak kolumny ..." (missing column):** the bank changed its headers. Fix the names in
  `settings.json` → `columns`.
- **Too much in "Inne":** check the "Do przypisania" sheet and add keywords to `rules.json`.
- **"Nie ma filtru ..." (no such filter):** the filter name is the file name in
  `config/filters/` without `.json`.
- **Transaction in a currency other than PLN:** the script warns and counts the amount
  without conversion.
- **Empty chart or zeros after opening the file:** open it in Excel. Formulas recalculate on
  open (some previewers, e.g. on a phone, may show empty cells).