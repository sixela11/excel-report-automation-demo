# Case Study (Demo): Monthly Sales Consolidation

> **Portfolio demonstration built on synthetic data.** It is not a past client engagement. Figures come from the sample run.

## Scenario
A fictional Australian distributor receives a sales export from its system each month. Before each management report, a staff member merges the files by hand, removes duplicates and fixes broken customer IDs. Errors are found late, and nobody can confirm which rows changed.

## Challenges in the data
- A month re-exported with rows already reported, which inflates revenue.
- Customer IDs saved as numbers, so `000104` became `104` and broke lookups.
- Hand-typed errors: impossible dates (`31/02/2026`), "two" in a quantity field, negative prices, blank names.
- Inconsistent column headers and date formats between months.
- One month missing a whole column.

## Solution
A Python automation that, in one command:
1. Reads every monthly workbook without modifying it.
2. Standardises headers and keeps identifiers as text.
3. Validates every row and every file against clear rules.
4. Removes exact duplicates. For conflicting records (same invoice, different values), it keeps the first by file order and flags the other for a person to decide on.
5. Produces a formatted Excel report with a summary, clean data, duplicates and a full exception list.

## Results (sample run)
- 7 files, 243 rows processed in under a second.
- 3 duplicate rows caught, including one with conflicting values.
- 5 invalid rows excluded and listed with file, row and reason.
- 40 damaged IDs restored, each logged.
- 1 incomplete file rejected before it could corrupt the totals.
- Every final row traceable to its source file and Excel row.

## Business value (expected, not measured)
- Replaces a repetitive manual task with a repeatable, auditable process.
- Errors are surfaced *before* the report goes to management.
- Staff time moves from copying data to reviewing exceptions.
- No new software subscription; it runs on existing Excel files.

Actual time saved depends on each business's files and process. It would be measured during a real project, not assumed.

## Limitations
Fixed column layout (easily adjusted), `.xlsx` first sheet only, and Australian `dd/mm/yyyy` text dates. Conflicting duplicates keep the earlier file's row until a person reviews them.

## Fit for
Accounting and bookkeeping firms, property managers, distributors, trades businesses and anyone consolidating recurring spreadsheet exports.

**Starter projects are fixed-price.** Contact: [github.com/sixela11](https://github.com/sixela11)
