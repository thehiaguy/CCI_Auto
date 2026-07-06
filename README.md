# CCI Daily → SXcoal Excel automation

Automates the daily transfer of data from the **CCI Daily PDF** (Sxcoal/Fenwei)
into **SXcoal data.xlsx**, replacing the manual copy-paste workflow.

## How it works

1. **PDF parsing.** The price/data tables in the PDF (cover page + back pages)
   are embedded as *images*, not text, so plain PDF parsers can't read them.
   The script renders those pages to PNG and sends them to the **Claude API**
   (vision + structured JSON output), which transcribes every table.
2. **Data Dump.** Extracted values are written into the workbook's
   `Data Dump` sheet at the exact fixed layout the other sheets reference
   (the extraction spec — row labels and columns — is read from the workbook
   itself at runtime, so it adapts if you edit the layout).
3. **Time-series sheets.** For each series sheet (Thermal Coal Daily,
   Coking Coal Daily, Met Coke Daily, Thermal Coal Weekly, Spot Futures,
   Thermal Coal Price (Port Mine), Truck Freight, Coastal/Seaborne Freight,
   Port Stockpile) the script replicates the manual daily step: it copies the
   bottom formula row down one row (relative references shift, exactly like
   dragging the fill handle), then freezes the old formula row's
   `'Data Dump'!` references and `WORKDAY` date cell to literal values.
   Day-over-day diff formulas, SUMs and the weekly AVERAGEIFS block are left
   live, as in the manual workflow.

The workbook is edited through Excel itself (xlwings/COM), so **charts,
tables, comments and formatting are fully preserved** and all formulas
recalculate. A timestamped copy is saved to a `backups\` folder next to the
workbook before every write.

Sheets **not** touched: Charts, Baltic, Indo-China Coal, Data (these come
from other sources), plus the news-summary sections 8–11 of Data Dump.
The manual cell `W` column in Coastal Coal Freight is also left to you.

## Setup (one-time)

```powershell
# Python 3.12 is installed at:
$py = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
& $py -m pip install -r requirements.txt

# API key (get one at https://platform.claude.com):
setx ANTHROPIC_API_KEY "sk-ant-..."     # persists; restart the terminal after
```

Requires desktop Excel (used via COM for saving).

## Daily use

```powershell
$py = "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe"
& $py cci_daily.py --pdf "C:\Users\res1\Downloads\CCI Daily (Jul 06, 2026).pdf"
```

By default the workbook is `C:\Users\res1\Downloads\SXcoal data.xlsx`
(override with `--xlsx`). The run is idempotent: if a sheet already has the
report date, it is skipped, so re-running the same PDF is safe.

Cost: one Claude API call per run (≈6 page images) — a few cents.

### Useful flags

| Flag | Purpose |
|---|---|
| `--dry-run --json out.json` | Extract only and save the JSON — inspect before writing |
| `--from-json out.json` | Write the workbook from a saved/corrected JSON (no API call) |
| `--force` | Append even if the date already exists (e.g. after a manual fix) |
| `--no-backup` | Skip the backup copy |

### Recommended first run

```powershell
& $py cci_daily.py --pdf "...pdf" --dry-run --json check.json   # eyeball check.json
& $py cci_daily.py --from-json check.json                        # then write
```

## Notes & maintenance

- **Weekly tables** (Thermal Coal Weekly) only append when the PDF's weekly
  table carries a new date — republished tables are skipped automatically.
- If Sxcoal **adds/removes a row** in a table (e.g. a new port), add/remove
  the matching row in `Data Dump` — the extraction spec follows the sheet.
  If a whole section moves to different rows, update `SECTIONS` at the top
  of `cci_daily.py`.
- Extraction accuracy: the model transcribes rendered tables; numbers are
  validated for shape but not against the PDF. For a critical workflow, use
  the `--dry-run` + `--from-json` two-step and spot-check a few values.
