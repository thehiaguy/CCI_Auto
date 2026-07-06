"""CCI Daily PDF -> "SXcoal data.xlsx" automation.

Pipeline:
  1. Read the CCI Daily PDF. Article pages are text; the data tables
     (cover page + back pages) are embedded images, so they are rendered
     to PNG and transcribed by the Claude API (vision + structured output).
  2. Write the extracted values into the "Data Dump" sheet of the workbook,
     preserving its fixed layout (other sheets reference these cells by
     absolute address).
  3. For each time-series sheet, replicate the manual daily step:
     copy the bottom formula row down one row (Excel FillDown, so relative
     references shift), then freeze the old formula row's Data-Dump
     references and WORKDAY date cell to literal values.

The workbook is edited through Excel itself (xlwings/COM) so charts,
tables, comments and formats are fully preserved and formulas recalculate.

Usage:
  python cci_daily.py --pdf "C:\\path\\CCI Daily (Jul 06, 2026).pdf"
                      [--xlsx "C:\\path\\SXcoal data.xlsx"]
                      [--dry-run] [--json out.json] [--from-json in.json]
                      [--force] [--no-backup]

Requires ANTHROPIC_API_KEY in the environment (not needed with --from-json).
"""

from __future__ import annotations

import argparse
import base64
import datetime as dt
import json
import re
import shutil
import sys
from pathlib import Path

DEFAULT_XLSX = r"C:\Users\res1\Downloads\SXcoal data.xlsx"
MODEL = "claude-opus-4-8"
RENDER_SCALE = 2.2  # ~158 dpi; keeps pages under Claude's 2576px high-res limit

# Data Dump layout: (id, pdf table title, title_row, header_row, first_data, last_data, n_label_cols)
SECTIONS = [
    ("s1", "Fenwei CCI Price Index Daily", 3, 4, 5, 36, 1),
    ("s2", "Fenwei CCI Thermal Weekly Index", 38, 39, 40, 49, 1),
    ("s2b", "Fenwei CCI Thermal Index Weekly Average", 51, 52, 53, 57, 1),
    ("s3", "Fenwei CCI Coking Coal Index Daily", 59, 60, 61, 74, 1),
    ("s4", "China Spot-Futures Basis", 76, 77, 78, 80, 1),
    ("s5", "China Thermal Coal Prices", 83, 84, 85, 127, 1),
    ("s6a", "Fenwei CCI Truck Freight", 129, 130, 131, 141, 1),
    ("s6b", "China Coastal Coal Freight", 143, 144, 145, 151, 1),
    ("s6c", "Seaborne Coal Freight", 153, 154, 155, 158, 1),
    ("s7", "Chinese Coal Ports Roundup", 160, 161, 162, 207, 2),
]

# Time-series sheets driven by the Data Dump: (sheet name, gating section or None=report date)
SERIES_SHEETS = [
    ("Thermal Coal Daily", "s1"),
    ("Coking Coal Daily", "s1"),
    ("Met Coke Daily", "s1"),
    ("Thermal Coal Weekly", "s2"),
    ("Spot Futures", "s4"),
    ("Thermal Coal Price (Port Mine)", "s5"),
    ("Truck Freight", "s6a"),
    ("Coastal Coal Freight", "s6b"),
    ("Seaborne Coal Freight", "s6c"),
    ("Port Stockpile", None),
]

# Sheets with a trailing self-computing block that lags one row behind the
# value rows (extended by copying the previous value row's block down).
LAGGING_BLOCKS = {"Thermal Coal Weekly": ("AR", "BK")}

MONTH_DATE_RE = re.compile(
    r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|"
    r"Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{1,2},\s*\d{4}"
)
# Strings Excel would silently convert to numbers/dates; force them to stay text.
TEXTY_RE = re.compile(r"^[+\-]?[\d,.]+%?$|^\d{4}/\d{1,2}/\d{1,2}$")


def log(msg: str) -> None:
    print(msg, flush=True)


def col_letter(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# --------------------------------------------------------------------------
# 1. Build the extraction spec from the workbook's current Data Dump layout
# --------------------------------------------------------------------------

def build_spec(xlsx: Path) -> list[dict]:
    import openpyxl

    wb = openpyxl.load_workbook(xlsx, read_only=True, data_only=True)
    if "Data Dump" not in wb.sheetnames:
        sys.exit("ERROR: workbook has no 'Data Dump' sheet")
    ws = wb["Data Dump"]
    grid = {}
    for row in ws.iter_rows(min_row=1, max_row=220, max_col=12):
        for c in row:
            if c.value is not None:
                grid[(c.row, c.column)] = c.value
    wb.close()

    spec = []
    for sid, title, title_row, header_row, first, last, n_label in SECTIONS:
        first_val_col = n_label + 1
        max_col = max(
            (c for (r, c) in grid if first <= r <= last and c > n_label), default=first_val_col
        )
        columns = []
        for c in range(first_val_col, max_col + 1):
            name = grid.get((header_row, c))
            columns.append({"col": c, "name": str(name) if name is not None else f"col{c}"})
        rows = []
        for r in range(first, last + 1):
            a = grid.get((r, 1))
            if a is None:
                continue
            has_data = any((r, c) in grid for c in range(first_val_col, max_col + 1))
            if not has_data:
                continue  # group header ("Thermal Coal", "QHD Port") or footnote row
            if n_label == 2:
                label = f"{a} | {grid.get((r, 2), '')}"
            else:
                label = str(a)
            rows.append({"row": r, "label": label})
        spec.append(
            {
                "id": sid,
                "pdf_table_title": title,
                "title_row": title_row,
                "n_label_cols": n_label,
                "columns": columns,
                "rows": rows,
            }
        )
    return spec


# --------------------------------------------------------------------------
# 2. Render the PDF's data-table pages (they are embedded images)
# --------------------------------------------------------------------------

def render_table_pages(pdf_path: Path) -> list[tuple[int, bytes]]:
    import io

    import pdfplumber
    import pypdfium2 as pdfium

    table_pages = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for i, page in enumerate(pdf.pages):
            big = [im for im in page.images if im["width"] > 400 and im["height"] > 300]
            if big:
                table_pages.append(i)
    if not table_pages:
        sys.exit("ERROR: no data-table (image) pages found in the PDF")

    doc = pdfium.PdfDocument(str(pdf_path))
    out = []
    for i in table_pages:
        img = doc[i].render(scale=RENDER_SCALE).to_pil()
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        out.append((i + 1, buf.getvalue()))
    log(f"Rendered {len(out)} table page(s): {[p for p, _ in out]}")
    return out


# --------------------------------------------------------------------------
# 3. Extraction via the Claude API (vision + structured output)
# --------------------------------------------------------------------------

EXTRACTION_SCHEMA = {
    "type": "object",
    "properties": {
        "report_date": {"type": "string", "description": "Issue date, YYYY-MM-DD"},
        "issue": {"type": "string", "description": "Issue number, digits only"},
        "sections": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "id": {"type": "string"},
                    "date": {"type": "string", "description": "Table date, YYYY-MM-DD"},
                    "rows": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "properties": {
                                "label": {"type": "string"},
                                "values": {
                                    "type": "array",
                                    "items": {
                                        "anyOf": [
                                            {"type": "number"},
                                            {"type": "string"},
                                            {"type": "null"},
                                        ]
                                    },
                                },
                            },
                            "required": ["label", "values"],
                            "additionalProperties": False,
                        },
                    },
                },
                "required": ["id", "date", "rows"],
                "additionalProperties": False,
            },
        },
    },
    "required": ["report_date", "issue", "sections"],
    "additionalProperties": False,
}

PROMPT = """These are page images from one issue of the "CCI Daily" coal-market PDF (Sxcoal/Fenwei).
Transcribe its data tables into JSON following the extraction spec below.

The spec lists, for each section: the table title as it appears in the PDF (`pdf_table_title`),
the column headers to extract (`columns`, in order), and the exact row labels expected (`rows`).

Rules:
- Return one section object per spec section, using the spec `id`.
- Return one row object per spec row label, using the spec label VERBATIM, in spec order.
  For the ports roundup the label is "Port | Metric" (e.g. "Qinhuangdao | Stockpile").
- `values` must align 1:1 with the spec's `columns` list for that section.
- Plain numbers: JSON numbers. Strip thousands separators and a leading "+" but keep negative signs.
- Cells displayed with a % sign (e.g. -23.7%): return the display string, e.g. "-23.7%".
- Text cells (Basis, Vessel Size, etc.): return the text as shown.
- The Date column in the ports roundup: string "YYYY/MM/DD".
- Blank cells: null.
- `date` per section: the date in that table's own title ("... on Jul 03, 2026"); if the title
  has no date, use the report date.
- `report_date` and `issue`: from the cover page header ("CCI Daily, Issue NNNN, <date>").
- If a spec row is missing from the PDF, return it with all-null values.
- Transcribe carefully; do not round or infer values.

EXTRACTION SPEC:
"""


def extract_with_claude(images: list[tuple[int, bytes]], spec: list[dict]) -> dict:
    import anthropic

    client = anthropic.Anthropic()

    spec_for_model = [
        {
            "id": s["id"],
            "pdf_table_title": s["pdf_table_title"],
            "columns": [c["name"] for c in s["columns"]],
            "rows": [r["label"] for r in s["rows"]],
        }
        for s in spec
    ]

    content = []
    for page_no, png in images:
        content.append({"type": "text", "text": f"PDF page {page_no}:"})
        content.append(
            {
                "type": "image",
                "source": {
                    "type": "base64",
                    "media_type": "image/png",
                    "data": base64.standard_b64encode(png).decode(),
                },
            }
        )
    content.append({"type": "text", "text": PROMPT + json.dumps(spec_for_model, indent=1)})

    log(f"Calling Claude ({MODEL}) to transcribe the tables ...")
    with client.messages.stream(
        model=MODEL,
        max_tokens=32000,
        thinking={"type": "adaptive"},
        output_config={"format": {"type": "json_schema", "schema": EXTRACTION_SCHEMA}},
        messages=[{"role": "user", "content": content}],
    ) as stream:
        msg = stream.get_final_message()

    if msg.stop_reason == "max_tokens":
        sys.exit("ERROR: extraction output was truncated (max_tokens); rerun or raise the limit")
    text = next(b.text for b in msg.content if b.type == "text")
    data = json.loads(text)
    u = msg.usage
    log(f"Extraction done (input {u.input_tokens} tok, output {u.output_tokens} tok).")
    return data


def validate(data: dict, spec: list[dict]) -> None:
    by_id = {s["id"]: s for s in data.get("sections", [])}
    problems = []
    for s in spec:
        got = by_id.get(s["id"])
        if got is None:
            problems.append(f"section {s['id']} missing from extraction")
            continue
        want_labels = [r["label"] for r in s["rows"]]
        got_labels = {norm_label(r["label"]) for r in got["rows"]}
        for lbl in want_labels:
            if norm_label(lbl) not in got_labels:
                problems.append(f"{s['id']}: row not extracted: {lbl!r}")
        ncols = len(s["columns"])
        for r in got["rows"]:
            if len(r["values"]) != ncols:
                problems.append(
                    f"{s['id']} row {r['label']!r}: {len(r['values'])} values, expected {ncols}"
                )
    if problems:
        for p in problems:
            log(f"  WARNING: {p}")
    dt.date.fromisoformat(data["report_date"])  # raises if malformed


def norm_label(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip().lower()


# --------------------------------------------------------------------------
# 4. Write the workbook via Excel (xlwings) — preserves charts & formats
# --------------------------------------------------------------------------

def excel_safe(v):
    """Prefix Excel-parseable strings with an apostrophe so they stay text."""
    if isinstance(v, str) and TEXTY_RE.match(v.strip()):
        return "'" + v
    return v


def section_date(data: dict, sid: str) -> dt.date:
    for s in data["sections"]:
        if s["id"] == sid:
            try:
                return dt.date.fromisoformat(s["date"])
            except (ValueError, KeyError):
                break
    return dt.date.fromisoformat(data["report_date"])


def write_workbook(xlsx: Path, spec: list[dict], data: dict, force: bool, backup: bool) -> None:
    import xlwings as xw

    report_date = dt.date.fromisoformat(data["report_date"])

    if backup:
        bdir = xlsx.parent / "backups"
        bdir.mkdir(exist_ok=True)
        stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M%S")
        dest = bdir / f"{xlsx.stem} {stamp}{xlsx.suffix}"
        shutil.copy2(xlsx, dest)
        log(f"Backup: {dest}")

    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    try:
        wb = app.books.open(str(xlsx), update_links=False)
        dump = wb.sheets["Data Dump"]

        # --- 4a. Data Dump values -------------------------------------------------
        sections_by_id = {s["id"]: s for s in data["sections"]}
        for s in spec:
            got = sections_by_id.get(s["id"])
            if not got:
                log(f"  SKIP section {s['id']}: not in extraction")
                continue
            rows_by_label = {norm_label(r["label"]): r["values"] for r in got["rows"]}
            n_written = 0
            for r in s["rows"]:
                values = rows_by_label.get(norm_label(r["label"]))
                if values is None:
                    log(f"  WARNING: {s['id']}: no extracted values for {r['label']!r}")
                    continue
                if all(v is None for v in values):
                    continue  # row absent from this issue; keep existing values
                for c, v in zip(s["columns"], values):
                    dump.range((r["row"], c["col"])).value = excel_safe(v)
                n_written += 1
            # refresh the date in the section title
            sec_date = section_date(data, s["id"])
            tcell = dump.range((s["title_row"], 1))
            title = tcell.value
            if isinstance(title, str) and MONTH_DATE_RE.search(title):
                tcell.value = MONTH_DATE_RE.sub(sec_date.strftime("%B %d, %Y"), title, count=1)
            log(f"  Data Dump {s['id']}: wrote {n_written} rows")

        # header cell A1: issue + date
        a1 = dump.range("A1")
        if isinstance(a1.value, str):
            t = re.sub(r"Issue\s+\d+", f"Issue {data.get('issue', '?')}", a1.value)
            t = MONTH_DATE_RE.sub(report_date.strftime("%b %d, %Y"), t)
            a1.value = t

        app.api.CalculateFull()

        # --- 4b. Append a dated value row to each time-series sheet ---------------
        for sheet_name, sid in SERIES_SHEETS:
            target = section_date(data, sid) if sid else report_date
            try:
                sht = wb.sheets[sheet_name]
            except Exception:
                log(f"  SKIP {sheet_name!r}: sheet not found")
                continue
            append_series_row(sht, target, force)

        wb.save()
        wb.close()
        log(f"Saved: {xlsx}")
    finally:
        app.quit()


def append_series_row(sht, target_date: dt.date, force: bool) -> None:
    name = sht.name
    used = sht.used_range
    max_row = used.last_cell.row
    max_col = used.last_cell.column

    # locate the bottom formula row via column A
    col_a = sht.range((1, 1), (max_row, 1)).formula
    if isinstance(col_a, str):
        col_a = [[col_a]]
    f_row = None
    for i in range(len(col_a) - 1, -1, -1):
        v = col_a[i][0] if isinstance(col_a[i], (list, tuple)) else col_a[i]
        if isinstance(v, str) and v.startswith("="):
            f_row = i + 1
            break
    if f_row is None or f_row < 2:
        log(f"  SKIP {name}: no formula row found")
        return

    prev = sht.range((f_row - 1, 1)).value
    prev_date = prev.date() if isinstance(prev, dt.datetime) else prev
    if isinstance(prev_date, dt.date) and prev_date >= target_date and not force:
        log(f"  SKIP {name}: last row {prev_date} >= report {target_date} (already current)")
        return

    formulas = sht.range((f_row, 1), (f_row, max_col)).formula
    if isinstance(formulas, str):
        formulas = ((formulas,),)
    formulas = formulas[0] if isinstance(formulas[0], (list, tuple)) else formulas

    # copy the formula row down one row (FillDown adjusts relative references)
    last = col_letter(max_col)
    sht.range(f"A{f_row}:{last}{f_row + 1}").api.FillDown()

    # freeze the old formula row: Data Dump refs -> literal values, WORKDAY -> report date
    computed = sht.range((f_row, 1), (f_row, max_col)).value
    if not isinstance(computed, (list, tuple)):
        computed = [computed]
    n_frozen = 0
    for idx, f in enumerate(formulas):
        if not (isinstance(f, str) and f.startswith("=")):
            continue
        cell = sht.range((f_row, idx + 1))
        if "WORKDAY(" in f.upper():
            cell.value = target_date
            n_frozen += 1
        elif "Data Dump'!" in f or "'Data Dump'" in f:
            cell.value = computed[idx]
            n_frozen += 1
        # all other formulas (diffs, SUMs, AVERAGEIFS) stay live: they only
        # reference in-sheet cells and remain stable once those are literals

    # sheets with a self-computing block that lags one row behind value rows
    if name in LAGGING_BLOCKS:
        c1, c2 = LAGGING_BLOCKS[name]
        sht.range(f"{c1}{f_row - 1}:{c2}{f_row}").api.FillDown()

    log(f"  {name}: appended {target_date} (row {f_row}, froze {n_frozen} cells)")


# --------------------------------------------------------------------------

def main() -> None:
    ap = argparse.ArgumentParser(description="CCI Daily PDF -> SXcoal Excel")
    ap.add_argument("--pdf", help="path to the CCI Daily PDF")
    ap.add_argument("--xlsx", default=DEFAULT_XLSX, help="path to SXcoal data.xlsx")
    ap.add_argument("--dry-run", action="store_true", help="extract only; don't touch the workbook")
    ap.add_argument("--json", help="also save the extracted data to this JSON file")
    ap.add_argument("--from-json", help="skip the PDF/API and load extraction from this JSON file")
    ap.add_argument("--force", action="store_true", help="append rows even if the date already exists")
    ap.add_argument("--no-backup", action="store_true", help="don't copy the workbook to backups\\ first")
    args = ap.parse_args()

    xlsx = Path(args.xlsx)
    if not xlsx.exists():
        sys.exit(f"ERROR: workbook not found: {xlsx}")

    spec = build_spec(xlsx)
    log(f"Spec built from workbook: {sum(len(s['rows']) for s in spec)} data rows in {len(spec)} sections")

    if args.from_json:
        data = json.loads(Path(args.from_json).read_text(encoding="utf-8"))
    else:
        if not args.pdf:
            sys.exit("ERROR: --pdf is required (or use --from-json)")
        pdf = Path(args.pdf)
        if not pdf.exists():
            sys.exit(f"ERROR: PDF not found: {pdf}")
        images = render_table_pages(pdf)
        data = extract_with_claude(images, spec)

    validate(data, spec)
    log(f"Report date: {data['report_date']}  (Issue {data.get('issue')})")

    if args.json:
        Path(args.json).write_text(json.dumps(data, indent=1), encoding="utf-8")
        log(f"Extraction saved: {args.json}")

    if args.dry_run:
        log("Dry run: workbook not modified.")
        return

    write_workbook(xlsx, spec, data, force=args.force, backup=not args.no_backup)


if __name__ == "__main__":
    main()
