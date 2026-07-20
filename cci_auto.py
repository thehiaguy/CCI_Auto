r"""One-click updater: find the newest un-processed CCI Daily issue(s) and
write them into the workbook.

No arguments needed - this is the double-click entry point. It:

  1. reads the workbook's newest REAL date (ignoring the live WORKDAY formula
     row at the bottom of each sheet, which is a template guess, not data),
  2. scans the PDF archive for every CCI Daily issue newer than that,
  3. runs the normal pipeline (cci_daily.py) on each, oldest first.

If the workbook is already current it says so and does nothing. If an issue
fails it stops immediately, leaving the pre-run backup in backups\ so the
workbook is never left half-updated across issues.

The workbook and PDF archive default to your Downloads folder. To point this at
real locations, put them in a .env file next to this script (it is git-ignored,
so machine-specific paths stay off GitHub):

    CCI_XLSX=R:\some\path\SXcoal data.xlsx
    CCI_ARCHIVE=R:\some\path\PDF archive
"""

from __future__ import annotations

import datetime as dt
import os
import re
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
sys.path.insert(0, str(HERE))

from cci_daily import load_env_file  # noqa: E402  (reads .env next to this script)
from test_run import workbook_newest_date  # noqa: E402  (formula-aware reader)

load_env_file()

XLSX = Path(os.environ.get("CCI_XLSX") or Path.home() / "Downloads" / "SXcoal data.xlsx")
ARCHIVE = Path(os.environ.get("CCI_ARCHIVE") or Path.home() / "Downloads")

# Model for the extraction step. Sonnet transcribes these tables identically to
# Opus (verified cell-by-cell on the Jul 16 2026 issue: 1095 cells, 0 differences)
# for a fraction of the subscription quota. Without this the claude CLI would
# inherit whatever model is set in the user's global settings.
MODEL = "sonnet"

MONTHS = {
    m: i
    for i, m in enumerate(
        ["jan", "feb", "mar", "apr", "may", "jun",
         "jul", "aug", "sep", "oct", "nov", "dec"], 1)
}
# Matches "Jul 10, 2026", "July 10, 2026", "(May 18, 2026)" - both filename styles.
DATE_RE = re.compile(r"([A-Za-z]{3,9})\s+(\d{1,2}),\s*(\d{4})")


def pdf_date(name: str) -> dt.date | None:
    m = DATE_RE.search(name)
    if not m:
        return None
    mon = m.group(1)[:3].lower()
    if mon not in MONTHS:
        return None
    try:
        return dt.date(int(m.group(3)), MONTHS[mon], int(m.group(2)))
    except ValueError:
        return None


def main() -> int:
    print("=" * 64)
    print(" SXcoal workbook updater")
    print("=" * 64)

    if not XLSX.exists():
        print(f"\nERROR: workbook not found:\n  {XLSX}")
        print("Set CCI_XLSX in a .env file next to this script.")
        return 1
    if not ARCHIVE.exists():
        print(f"\nERROR: PDF archive folder not found:\n  {ARCHIVE}")
        print("Set CCI_ARCHIVE in a .env file next to this script.")
        return 1

    newest = workbook_newest_date(XLSX)
    print(f"\nWorkbook is current through: {newest}")

    todo = []
    for pdf in ARCHIVE.glob("CCI Daily*.pdf"):
        d = pdf_date(pdf.name)
        if d and d > newest:
            todo.append((d, pdf))
    todo.sort()

    if not todo:
        print("\nNothing to do - the workbook is already up to date. [OK]")
        return 0

    print(f"\n{len(todo)} new issue(s) to process:")
    for d, pdf in todo:
        print(f"    {d}   {pdf.name}")

    script = str(HERE / "cci_daily.py")
    for d, pdf in todo:
        print("\n" + "-" * 64)
        print(f" Processing {d}  ({pdf.name})")
        print("-" * 64)
        json_out = str(HERE / f"check_{d.strftime('%b%d').lower()}.json")
        result = subprocess.run(
            [sys.executable, script, "--pdf", str(pdf),
             "--xlsx", str(XLSX), "--json", json_out, "--model", MODEL]
        )
        if result.returncode != 0:
            print(f"\nERROR while processing {d}. Stopping here.")
            print("The pre-run backup is in the workbook's backups\\ folder if")
            print("you need to restore. Ask Claude in VS Code to help recover.")
            return 1

    print("\n" + "=" * 64)
    print(f" All done - workbook updated through {todo[-1][0]}. [OK]")
    print("=" * 64)
    return 0


if __name__ == "__main__":
    sys.exit(main())
