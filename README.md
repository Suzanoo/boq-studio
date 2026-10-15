# BOQ Studio — MS-1.5

Desktop application for converting badly formatted BOQ breakdown sheets into separate **OKD-compatible sheets**.

## MS-1.5 scope

- User chooses an `.xlsx` BOQ file.
- Program lists every sheet.
- User selects only breakdown sheets.
- Each selected sheet has its own column mapping.
- **Auto Detect** suggests the header row, first detail row, and OKD columns.
- Mapping templates can be saved, loaded, and automatically reused.
- A **Normalized Preview** shows WBS and ITEM output before export.
- Description can come from more than one source column.
- Program detects WBS/header rows and item rows using deterministic rules.
- WBS codes are generated separately inside each sheet.
- Original workbook is never modified.
- Output is a new `*_OKD.xlsx` file on the user's Desktop.
- Progress bar shows the current sheet and source row.
- A `Project` sheet combines every normalized breakdown sheet in the user-selected order.
- A `LOG` sheet records processed, skipped, warning, and error counts.

## Output columns

Every normalized sheet contains:

```text
WBS Code
WBS Level
Node Type
WBS-1
WBS-2
WBS-3
WBS-4
Description
Unit
Qty
Material
Labor
Amount
Source Sheet
Source Row
```

The required OKD columns remain present: `WBS-1` to `WBS-4`, `Description`, `Unit`, `Qty`, `Material`, `Labor`, and `Amount`.

## Install on Windows or macOS

```bash
cd boq-studio
python -m venv .venv
```

Activate the environment.

### Windows PowerShell

```powershell
.venv\Scripts\Activate.ps1
```

### macOS / Linux

```bash
source .venv/bin/activate
```

Install:

```bash
python -m pip install --upgrade pip
pip install -e .
```

Run:

```bash
boq-studio
```

or:

```bash
python -m boq_studio
```

## Workflow

1. Click **Choose Excel file**.
2. Select the breakdown sheets only.
3. Double-click each selected sheet.
4. Click **Auto Detect** and review its suggestions.
5. Adjust Header row, First data row, and mapped columns when needed.
6. Select one or more Description columns with Ctrl/Shift.
7. Open the **Normalized** tab and refresh the preview.
8. Optionally save or load a reusable `.json` mapping template.
9. Save the mapping and click **Normalize to Desktop**.
10. Open `<original-name>_OKD.xlsx` from Desktop.

## WBS rules in MS-1.2

A row is treated as an item when it contains a unit or a numeric Qty/Material/Labor/Amount value. Other descriptive rows become WBS nodes, except total and note rows, which are skipped.

WBS level is detected in this order:

1. Numbering such as `1`, `1.1`, `1.1.1`.
2. Numbering at the beginning of Description.
3. Position/indentation of the selected Description columns.

This is deliberately rule-based. MS-1 does not use AI and does not claim perfect automatic interpretation. `Source Sheet` and `Source Row` are retained so every normalized row can be checked against the original BOQ.

## Tests

```bash
pip install pytest
pytest
```

## Git

The repository is initialized. Run `git log --oneline` to see the milestone commits.


## MS-1.3 output rules

- Only rows with `Amount > 0` are exported.
- `WBS-1` is the placeholder `Project` for the user to rename later.
- `WBS-2` is the source sheet name.
- `WBS-3` and `WBS-4` are generated from detected BOQ headings.
- Preview and Export use the same normalization rules.


## MS-1.4 Project sheet

- The output workbook contains a `Project` sheet immediately after `LOG`.
- Its header appears once.
- Data rows from each successfully normalized breakdown sheet are appended in the same order selected by the user.
- `Source Sheet` and `Source Row` remain available for traceability.
- Use the `Project` sheet as the single consolidated input for OKD.

> Important: importing both `Project` and the individual normalized sheets into OKD will duplicate the amounts. Select only `Project` for the consolidated workflow.


## MS-1.5 amount checks

- A separate `SUMMARY` sheet shows the live `Amount` total for every normalized breakdown sheet.
- `Grand Total` sums all selected breakdown sheets.
- `Project Sheet Total` independently sums the consolidated `Project` sheet.
- `Difference` should remain zero and helps the user verify manual edits.
- Totals use Excel formulas and therefore update when the user edits amounts manually.
- No total rows are inserted into the OKD data sheets, preventing duplicate amounts during import.

## Standard entry-point checks (MS-1)

For development, install the optional test dependency and verify both supported launch methods:

```bash
pip install -e ".[dev]"
boq-studio
python -m boq_studio
pytest
```

The `.venv` directory is local to each computer and is intentionally excluded from Git and release ZIP files.

## S0 persistence safety

See [S0 session safety and recovery](docs/milestones/S0_PERSISTENCE_SAFETY.md) for Save/Resume/Discard behavior, local recovery files, and desktop acceptance steps. S0 retains the existing session format; portable Project storage is not included.
