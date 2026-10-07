"""Write findings back as an annotated copy of the deliverable.

The copy is the original workbook with three columns appended to each checked
sheet - result, severity and the observations behind them - plus a summary
sheet. The original file is never touched.
"""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from dge.ded import workbook
from dge.ded.findings import CATEGORY_ORDER, Finding, Severity
from dge.ded.spec import Spec

# Structure findings - a missing column, a drifted guideline - belong to the
# sheet rather than to any row, so they get no per-row flag column.
ROW_CATEGORIES = tuple(name for name in CATEGORY_ORDER if name != "Structure")

# Result, then one flag per category so the sheet can be filtered down to a
# single kind of problem, then the full text last because it is the widest.
QC_COLUMNS = (
    "QC Result",
    "QC Severity",
    "QC Issues",
    *(f"QC {name}" for name in ROW_CATEGORIES),
    "QC Observations",
)
QC_WIDTHS = (12, 12, 8, *(13 for _ in ROW_CATEGORIES), 90)

HEADER_FILL = PatternFill("solid", fgColor="44546A")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FILLS = {
    Severity.ERROR: PatternFill("solid", fgColor="FFC7CE"),  # red
    Severity.WARNING: PatternFill("solid", fgColor="FFD9A3"),  # orange
    None: PatternFill("solid", fgColor="C6E0B4"),  # passed, green
}


def output_path(catalogue: Path, directory: Path | None = None) -> Path:
    """<timestamp>_<input file name>.xlsx in the output directory."""
    target = directory or catalogue.parent / "output"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return target / f"{stamp}_{catalogue.name}"


def _last_column(worksheet, header_row: int) -> int:
    """Rightmost header cell that actually holds a value.

    `max_column` is unreliable - the attribute sheet reports thousands of
    phantom columns left behind by formatting.
    """
    last = 0
    for cell in next(worksheet.iter_rows(min_row=header_row, max_row=header_row)):
        if cell.value is not None and str(cell.value).strip():
            last = cell.column
    return last


def write(
    catalogue: Path,
    destination: Path,
    findings: list[Finding],
    specs: dict[str, Spec],
    keep_comments: bool = False,
) -> Path:
    """Save an annotated copy of `catalogue` to `destination`."""
    # data_only: write the cached results rather than the formulas behind them.
    # Several columns are computed from the sheet's Excel Table (column W is
    # =Table1[[#This Row],[Domain Name]]), and replacing that table with a
    # plain autofilter leaves those formulas as #REF!. The report is a
    # snapshot for review, so values are what it should carry anyway.
    book = openpyxl.load_workbook(catalogue, data_only=True)

    by_sheet: dict[str, dict[int, list[Finding]]] = defaultdict(lambda: defaultdict(list))
    unplaced: list[Finding] = []
    for finding in findings:
        if finding.sheet and finding.row is not None:
            by_sheet[finding.sheet][finding.row].append(finding)
        else:
            unplaced.append(finding)

    evaluated = {spec.sheet_name for spec in specs.values()}
    for spec in specs.values():
        if spec.sheet_name not in book.sheetnames:
            continue
        _annotate_sheet(book[spec.sheet_name], spec, by_sheet.get(spec.sheet_name, {}))

    # The report carries only the sheets that were assessed.
    for name in [n for n in book.sheetnames if n not in evaluated]:
        del book[name]

    _strip_defined_names(book)
    _strip_legacy_parts(book, keep_comments)

    _write_findings(book, findings)
    _write_summary(book, findings, unplaced)

    destination.parent.mkdir(parents=True, exist_ok=True)
    book.save(destination)
    book.close()
    return destination


def _strip_defined_names(book) -> None:
    """Drop the workbook's legacy defined names.

    The deliverable carries thousands of them, inherited through years of
    copy-paste - many already #REF! or pointing at external files. Rewriting
    them produces a workbook Excel offers to repair, and none of them are
    scoped to the sheets the report keeps.
    """
    book.defined_names.clear()
    for worksheet in book.worksheets:
        worksheet.defined_names.clear()


def _strip_legacy_parts(book, keep_comments: bool = False) -> int:
    """Drop the workbook parts the report does not need and Excel chokes on.

    Links to 69 other workbooks, and the legacy comment drawing, which does
    not survive a round trip intact - the VML is what makes Excel offer to
    repair the file on open. Returns the number of comments removed; the
    original file keeps them either way.
    """
    book._external_links = []
    removed = 0
    for worksheet in book.worksheets:
        if keep_comments:
            continue
        worksheet.legacy_drawing = None
        # Comments hang off individual cells. Walking the sparse cell store
        # avoids materialising the sheet's thousands of phantom columns.
        for cell in list(worksheet._cells.values()):
            if cell.comment is not None:
                cell.comment = None
                removed += 1
    return removed


def _enable_filters(worksheet, spec: Spec, last_column: int) -> None:
    """Put one filter across every column, QC columns included.

    Each sheet arrives as an Excel Table whose fixed range stops short of the
    columns we append, and a table's filter cannot reach past it. Replacing
    the table with a plain autofilter covers the full width.
    """
    for name in list(worksheet.tables):
        del worksheet.tables[name]

    end = f"{get_column_letter(last_column)}{worksheet.max_row}"
    worksheet.auto_filter.ref = f"A{spec.header_row}:{end}"


def _reset_view(worksheet) -> None:
    """Open the sheet at A1, unfrozen.

    The deliverable stores wherever each sheet was last scrolled to - the
    attribute sheet reopens at K80 - which looks like a stuck freeze once the
    report is opened fresh.
    """
    worksheet.freeze_panes = None
    view = worksheet.sheet_view
    view.pane = None
    view.topLeftCell = None
    for selection in view.selection:
        selection.pane = None
        selection.activeCell = "A1"
        selection.sqref = "A1"


def _column_positions(worksheet, header_row: int) -> dict[str, int]:
    """Map each header to its column index, so a finding can reach its cell."""
    positions: dict[str, int] = {}
    for cell in next(worksheet.iter_rows(min_row=header_row, max_row=header_row)):
        if name := workbook.clean_header(cell.value):
            positions.setdefault(name, cell.column)
    return positions


def _highlight_cells(
    worksheet, found: list[Finding], positions: dict[str, int], number: int
) -> None:
    """Colour the individual cells a finding points at.

    The row verdict says something is wrong; this says which cell. Where a
    cell draws both an error and a warning, the error wins.
    """
    worst: dict[int, Severity] = {}
    for finding in found:
        column = positions.get(finding.column)
        if column is not None:
            worst[column] = max(worst.get(column, Severity.INFO), finding.severity)
    for column, severity in worst.items():
        if fill := FILLS.get(severity):
            worksheet.cell(row=number, column=column).fill = fill


def _annotate_sheet(worksheet, spec: Spec, rows: dict[int, list[Finding]]) -> None:
    start = _last_column(worksheet, spec.header_row) + 1
    positions = _column_positions(worksheet, spec.header_row)

    for offset, title in enumerate(QC_COLUMNS):
        cell = worksheet.cell(row=spec.header_row, column=start + offset, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")

    for offset, width in enumerate(QC_WIDTHS):
        worksheet.column_dimensions[get_column_letter(start + offset)].width = width

    # Only rows the checks actually looked at get a verdict; blank filler rows
    # below the data are left alone.
    for number in range(spec.first_data_row, worksheet.max_row + 1):
        found = rows.get(number)
        if found is None:
            if not _row_has_data(worksheet, number, start):
                continue
            values = ["Passed", "", "", *("" for _ in ROW_CATEGORIES), ""]
            fills: list[PatternFill | None] = [FILLS[None]] * len(values)
        else:
            worst = max(f.severity for f in found)
            by_category: dict[str, Severity] = {}
            for finding in found:
                current = by_category.get(finding.category)
                by_category[finding.category] = (
                    finding.severity if current is None else max(current, finding.severity)
                )
            row_fill = FILLS.get(worst, FILLS[None])
            values = [
                "Failed" if worst is Severity.ERROR else "Passed with warnings",
                str(worst),
                len(found),
                *(str(by_category[n]) if n in by_category else "" for n in ROW_CATEGORIES),
                "\n".join(_describe(f) for f in sorted(found, key=_order)),
            ]
            fills = [
                row_fill,
                row_fill,
                row_fill,
                *(FILLS.get(by_category[n]) if n in by_category else None for n in ROW_CATEGORIES),
                row_fill,
            ]
            _highlight_cells(worksheet, found, positions, number)

        last = len(values) - 1
        for offset, value in enumerate(values):
            cell = worksheet.cell(row=number, column=start + offset, value=value)
            if fills[offset] is not None:
                cell.fill = fills[offset]
            cell.alignment = Alignment(vertical="top", wrap_text=offset == last)

    _enable_filters(worksheet, spec, start + len(QC_COLUMNS) - 1)
    _reset_view(worksheet)


def _order(finding: Finding) -> tuple:
    return (-finding.severity, finding.column, finding.check)


def _describe(finding: Finding) -> str:
    where = f"{finding.column}: " if finding.column else ""
    return f"[{str(finding.severity).upper()}] {where}{finding.message}"


def _row_has_data(worksheet, number: int, before: int) -> bool:
    for cell in next(
        worksheet.iter_rows(min_row=number, max_row=number, max_col=before - 1)
    ):
        if cell.value is not None and str(cell.value).strip():
            return True
    return False


FINDINGS_HEADERS = (
    ("Severity", 11),
    ("Category", 15),
    ("Check", 26),
    ("Sheet", 24),
    ("Row", 8),
    ("Column", 30),
    ("Value", 38),
    ("Message", 95),
)


def _write_findings(book, findings: list[Finding]) -> None:
    """One row per finding, for working through a whole category at a time.

    The annotated sheets answer "what is wrong with this row"; this answers
    "where else does this problem occur", which is how the fixing is done.
    """
    if "QC Findings" in book.sheetnames:
        del book["QC Findings"]
    sheet = book.create_sheet("QC Findings", 0)

    for index, (title, width) in enumerate(FINDINGS_HEADERS, start=1):
        cell = sheet.cell(row=1, column=index, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        sheet.column_dimensions[get_column_letter(index)].width = width

    for number, finding in enumerate(findings, start=2):
        values = (
            str(finding.severity),
            finding.category,
            finding.check,
            finding.sheet,
            finding.row,
            finding.column,
            None if finding.value is None else str(finding.value)[:200],
            finding.message,
        )
        for index, value in enumerate(values, start=1):
            sheet.cell(row=number, column=index, value=value)
        if fill := FILLS.get(finding.severity):
            sheet.cell(row=number, column=1).fill = fill

    sheet.auto_filter.ref = f"A1:{get_column_letter(len(FINDINGS_HEADERS))}{len(findings) + 1}"
    sheet.freeze_panes = "A2"


def _write_summary(book, findings: list[Finding], unplaced: list[Finding]) -> None:
    """A QC Summary sheet: counts per check, then the sheet-level findings."""
    if "QC Summary" in book.sheetnames:
        del book["QC Summary"]
    sheet = book.create_sheet("QC Summary", 0)

    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["B"].width = 28
    sheet.column_dimensions["C"].width = 10
    sheet.column_dimensions["D"].width = 100

    sheet["A1"] = "Data Catalogue QC"
    sheet["A1"].font = Font(bold=True, size=14)
    sheet["A2"] = f"Run {datetime.now():%Y-%m-%d %H:%M:%S}"

    counts = Counter(f.severity for f in findings)
    sheet["A4"] = "Findings"
    sheet["A4"].font = Font(bold=True)
    for offset, severity in enumerate((Severity.ERROR, Severity.WARNING, Severity.INFO)):
        sheet.cell(row=5 + offset, column=1, value=str(severity))
        sheet.cell(row=5 + offset, column=2, value=counts[severity])

    sheet["D4"] = "By category"
    sheet["D4"].font = Font(bold=True)
    per_category = Counter(f.category for f in findings)
    line = 5
    for name in CATEGORY_ORDER:
        if not per_category[name]:
            continue
        sheet.cell(row=line, column=4, value=name)
        sheet.cell(row=line, column=5, value=per_category[name])
        line += 1

    row = max(10, line + 1)
    sheet.cell(row=row, column=1, value="Severity").font = Font(bold=True)
    sheet.cell(row=row, column=2, value="Check").font = Font(bold=True)
    sheet.cell(row=row, column=3, value="Count").font = Font(bold=True)
    sheet.cell(row=row, column=4, value="Sheet-level detail").font = Font(bold=True)

    grouped = Counter((f.severity, f.check) for f in findings)
    for (severity, check), count in sorted(grouped.items(), key=lambda kv: (-kv[0][0], kv[0][1])):
        row += 1
        sheet.cell(row=row, column=1, value=str(severity))
        sheet.cell(row=row, column=2, value=check)
        sheet.cell(row=row, column=3, value=count)
        if fill := FILLS.get(severity):
            sheet.cell(row=row, column=1).fill = fill

    if unplaced:
        row += 2
        sheet.cell(row=row, column=1, value="Not tied to a row").font = Font(bold=True)
        for finding in sorted(unplaced, key=_order):
            row += 1
            sheet.cell(row=row, column=1, value=str(finding.severity))
            sheet.cell(row=row, column=2, value=finding.check)
            cell = sheet.cell(row=row, column=4, value=finding.message)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
