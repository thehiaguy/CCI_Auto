# CCI Daily → SXcoal Excel automation

Turns the daily **CCI Daily PDF** (Sxcoal/Fenwei coal report) into an updated
**SXcoal data.xlsx** with one command — no more manual copy-paste. Claude
transcribes the PDF's table images, the script writes the workbook through
Excel itself (charts and formulas preserved), and a backup is taken before
every write.

---

## Setup (one-time, ~10 minutes)

You need a Windows PC with desktop **Excel** installed. Then:

**1. Install Python 3.12** — from [python.org](https://www.python.org/downloads/).
   In the installer, tick **"Add python.exe to PATH"**.

**2. Get this project** — clone it (or download the ZIP and extract):

```powershell
git clone https://github.com/thehiaguy/CCI_Auto.git
cd CCI_Auto
```

**3. Install the Python packages:**

```powershell
python -m pip install -r requirements.txt
```

**4. Install [Claude Code](https://claude.com/claude-code) and log in** —
   the script uses your Claude subscription to read the PDF, so no API bill.
   After installing, run `claude` once in a terminal and follow the login
   prompt. (Skip this only if you plan to use the paid API engine instead —
   see [Extraction engines](#extraction-engines).)

**5. Put the workbook in place.** By default the script looks for the Excel
   file in **your Downloads folder**, named exactly **`SXcoal data.xlsx`**:

   ```
   C:\Users\<your username>\Downloads\SXcoal data.xlsx
   ```

   (It uses whoever-is-logged-in's Downloads folder automatically — no need
   to edit any path in the script.) The name must match exactly, including
   the space: `SXcoal data.xlsx`.

   Keeping it somewhere else instead? You have two options:

   - pass the path on each run: `--xlsx "D:\path\to\SXcoal data.xlsx"`, or
   - change the default permanently: edit the `DEFAULT_XLSX` line near the
     top of `cci_daily.py`.

   The `backups\` folder is created next to the workbook automatically.

That's it. No API key, no config files needed for the default setup.

### Check it works

```powershell
python cci_daily.py --pdf "CCI Daily (Jul 07, 2026).pdf" --dry-run
```

`--dry-run` extracts and validates but writes nothing, so it's a safe
first test. If it prints the report date and section row counts, you're set.

---

## Daily use

Download the day's PDF, **close the workbook in Excel**, then:

```powershell
cd C:\Users\<you>\CCI_Auto
python cci_daily.py --pdf "C:\Users\<you>\Downloads\CCI Daily (Jul 07, 2026).pdf"
```

(Quote the path; if the PDF is in the current folder, just the filename works.)

That's the whole routine. The run:

1. backs up the workbook to `backups\` next to it,
2. extracts every table via your Claude subscription (takes a few minutes),
3. updates the `Data Dump` sheet and news sections 8–11,
4. appends the day's row to every time-series sheet,
5. fetches the SSE Composite Index from the web and saves.

The run is **idempotent** — sheets that already have the report date are
skipped, so re-running the same PDF is safe.

---

## All commands & flags

```powershell
# standard daily run
python cci_daily.py --pdf "CCI Daily (Jul 07, 2026).pdf"

# extract only — inspect the numbers before anything is written
python cci_daily.py --pdf "CCI Daily (Jul 07, 2026).pdf" --dry-run --json check.json

# write the workbook from a previously saved/corrected JSON (no Claude call)
python cci_daily.py --from-json check.json

# use the paid Anthropic API instead of the subscription
python cci_daily.py --pdf "CCI Daily (Jul 07, 2026).pdf" --engine api

# write to a different workbook (e.g. a test copy)
python cci_daily.py --pdf "CCI Daily (Jul 07, 2026).pdf" --xlsx ".\test copy.xlsx"

# re-append even if the date already exists (repairs only)
python cci_daily.py --from-json check.json --force

# skip the automatic backup copy
python cci_daily.py --pdf "CCI Daily (Jul 07, 2026).pdf" --no-backup

# skip the SSE website fetch (Composite Index column W stays manual that day)
python cci_daily.py --pdf "CCI Daily (Jul 07, 2026).pdf" --no-web
```

| Flag | Purpose |
|---|---|
| `--pdf PATH` | The CCI Daily PDF to process |
| `--xlsx PATH` | Workbook to update (default: `Downloads\SXcoal data.xlsx`) |
| `--engine cli\|api` | Extraction engine (default `cli` = your subscription) |
| `--dry-run` | Extract and validate only; workbook untouched |
| `--json FILE` | Save the extracted data as JSON (audit trail) |
| `--from-json FILE` | Skip extraction; write the workbook from a saved JSON |
| `--force` | Append rows even if the report date already exists |
| `--no-backup` | Don't copy the workbook to `backups\` first |
| `--no-web` | Don't fetch the SSE Composite Index (column W) |

## Extraction engines

| Engine | How | Cost |
|---|---|---|
| `cli` (default) | Runs the local `claude` CLI (Claude Code) — counts against your existing Claude subscription | No extra bill |
| `api` | Calls the Anthropic API directly; needs `ANTHROPIC_API_KEY` in the environment or in a `.env` file next to the script (one line: `ANTHROPIC_API_KEY=sk-ant-...`) | ~$0.55/run on Opus |

Both engines produce identical output (verified on 1,095/1,095 values).
The `cli` engine takes a few minutes per run; if you ever hit your plan's
usage limit, rerun later or fall back to `--engine api`.
**Never commit `.env`** — it's in `.gitignore`.

## How it works

1. **PDF parsing.** The price tables in the PDF (cover page + back pages)
   are embedded as *images*, not text, so plain PDF parsers can't read them.
   The script renders those pages to PNG and has Claude transcribe every
   table. Article pages are normal text and are read directly.
2. **Data Dump.** Extracted values are written into the workbook's
   `Data Dump` sheet at the exact fixed layout the other sheets reference.
   The extraction spec — row labels and column headers — is read from the
   sheet at runtime, so if Sxcoal reorders a table, matching is by *name*,
   not position, and nothing breaks.
3. **Time-series sheets.** For each series sheet the script replicates the
   manual daily step: copy the bottom formula row down one row (like dragging
   the fill handle), then freeze the old row's `'Data Dump'!` references and
   `WORKDAY` date to literal values. Diff formulas, SUMs and the weekly
   AVERAGEIFS block stay live.

The workbook is edited through Excel itself (xlwings/COM), so charts, tables
and formatting are fully preserved and all formulas recalculate.

### Extras handled automatically

- **Thermal Coal Weekly, AR:BK block** — dated from the PDF's own "Weekly
  Average" table date (section 2b, often a Sunday); updates on whichever
  issue carries a new date (normally Mondays).
- **Coastal Coal Freight, column W (Composite Index)** — fetched from the
  Shanghai Shipping Exchange (en.sse.net.cn) each run; current and previous
  day placed on their matching date rows.
- **Coastal Coal Freight, weekly block (col AA onwards)** — Monday's
  "Weekly: China's coastal coal freight..." article is parsed for the
  QHD-Guangzhou 60,000–70,000 DWT rate and a weekly row is appended.
- **Port Stockpile lag dates (Guangzhou)** — when a port row's own date is
  older than the report date, its value is placed on the row matching its
  own date, not the report-date row.
- **News sections 8–11 of Data Dump** — Claude reads the article pages and
  refreshes the summaries: section 8 (Shanxi mine suspensions) gains a new
  date column pair when fresh per-region figures appear; sections 9
  (Mongolian & import coking), 10 (Indonesia HBA, bi-weekly) and 11 (Global
  market notes) are rewritten with each issue's items. Sections without
  matching content are left untouched.

Sheets **not** touched: Charts, Baltic, Indo-China Coal, Data (these come
from other sources).

## Testing (dress rehearsal)

`test_run.py` simulates tomorrow's issue without touching the real workbook:
it shifts a saved extraction's dates forward one workday, runs the full
pipeline on a throwaway copy, and verifies the result. It needs a saved
extraction JSON (`check.json`) — create one with `--dry-run --json check.json`.

```powershell
python test_run.py            # offline, free — expect: ALL CHECKS PASSED
python test_run.py --keep     # keeps test_run.xlsx so you can inspect it in Excel
python test_run.py --live     # re-extract from the newest PDF first (uses subscription)
```

Run it after any change to the script or to the workbook's layout.

## Troubleshooting & maintenance

- **"PDF not found"** — check the path/filename; quote it, and use the full
  path if the PDF isn't in the current folder.
- **Excel save errors** — close the workbook in Excel before running; Excel
  can't save a file that's open elsewhere.
- **A run went wrong?** The pre-run workbook is in `backups\` next to the
  workbook — just copy it back.
- **Sxcoal adds/removes a table row** (e.g. a new port) — add/remove the
  matching row in `Data Dump`; the extraction spec follows the sheet. If a
  whole section moves, update `SECTIONS` at the top of `cci_daily.py`.
- **Weekly tables** only append when the PDF carries a new date —
  republished tables are skipped automatically.
- **News sections 8–11** are summaries, not transcriptions — wording varies
  run to run. They live at fixed row blocks (`NEWS_SECTIONS` in
  `cci_daily.py`); nothing else references those cells, so rewriting is
  safe. Don't add formulas pointing into rows 208+ of Data Dump.
- **Coastal AC (weekly article rate) / Coastal W (SSE index)** — if either
  source can't be read on a given day, the run warns and leaves the cell
  for manual entry; it never guesses.
- **Extraction accuracy** — numbers are validated for shape, not against the
  PDF. For extra assurance use the `--dry-run --json` + `--from-json`
  two-step and spot-check a few values.
