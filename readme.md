# AssetReconcile — Data Reconciliation & Automation Tool

Excel VBA automation that cross-checks a ServiceNow / eSMT asset export against a local Excel inventory, flags mismatches across 6 categories, and writes a one-click report. Built to replace a manual end-of-day reconciliation process during an IT Asset co-op placement.

## Why this exists

Asset records live in two places: the ServiceNow/eSMT ticketing system and a manually maintained Excel inventory. Over time the two drift apart — assets get reassigned, retired, or moved without both systems being updated. Catching that drift by eye across hundreds of rows is slow and error-prone. AssetReconcile automates the comparison.

## What it checks

| # | Category | Example |
|---|----------|---------|
| 1 | Missing in Inventory | Asset exists in ServiceNow but not in the local inventory |
| 2 | Missing in ServiceNow | Asset exists in inventory but has no ServiceNow record |
| 3 | Asset Name Mismatch | Different asset description between the two sources |
| 4 | Status Mismatch | e.g. ServiceNow says "Retired", Inventory says "Active" |
| 5 | Location Mismatch | Different floor/room recorded in each source |
| 6 | Assigned User Mismatch | Different owner recorded in each source |

## How it works

```
ServiceNow export (CSV)  ──┐
                            ├──► VBA macro compares row by row ──► Report sheet (flagged + color coded)
Local Excel Inventory    ──┘
```

1. Paste the ServiceNow/eSMT export into the **ServiceNow** sheet
2. Paste the current local inventory into the **Inventory** sheet
3. Run the macro (or click the button on the sheet)
4. Open the **Report** sheet — every mismatch is listed with asset tag, mismatch type, and the conflicting values from each source

## Tech stack

- **Excel VBA** — all reconciliation logic, no external dependencies
- **ServiceNow / eSMT CSV exports** — input data source

SQLite was considered for this project as a scaling path — if the data grew into the millions of rows, Excel/VBA would start to slow down and a local database would be a better fit for queries and history. At the scale this tool was used at (under a few hundred records per run), Excel/VBA handled it without issue, so SQLite wasn't necessary for the current scope.

## Project structure

```
assetreconcile/
├── inventory.xlsm          # Main workbook (ServiceNow, Inventory, Report sheets + macro)
├── servicenow_export.csv   # Sample ServiceNow export for testing
└── README.md
```

## Setup

1. Download `inventory.xlsm`
2. Enable macros when opening (Excel will prompt — click "Enable Content")
3. Paste your ServiceNow export into the **ServiceNow** sheet (starting at A1, with headers)
4. Paste your inventory data into the **Inventory** sheet (same column layout)
5. Run `ReconcileAssets` from the Developer tab → Macros, or click the button on the sheet
6. Check the **Report** sheet for results

### Expected column layout (both sheets)

| asset_tag | serial_number | asset_name | assigned_to | status | location | department |
|-----------|----------------|------------|--------------|--------|----------|-------------|

## Sample output

| Asset Tag | Mismatch Type | Details |
|-----------|----------------|---------|
| ASSET-002 | Location Mismatch | SN: Floor 5 \| INV: Floor 2 |
| ASSET-003 | Status Mismatch | SN: Retired \| INV: Active |
| ASSET-007 | User Mismatch | SN: Tom Brown \| INV: Lisa Wong |
| ASSET-016 | Missing in Inventory | Exists in ServiceNow only |
| ASSET-021 | Missing in ServiceNow | Exists in Inventory only |

## Known limitations

- Matching is done on exact `asset_tag` text match — case and whitespace differences will register as "missing" rather than a soft match
- Designed for single-run reconciliation rather than historical tracking across runs
- Tested at a scale of a few hundred rows; very large datasets (100k+ rows) would benefit from a database-backed approach instead of in-memory `Find()` loops

## Author

Movis Thakur — Software Development & Network Engineering, Sheridan College
[github.com/movisT12](https://github.com/movisT12)
