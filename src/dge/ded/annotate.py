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

from dge.ded.findings import Finding, Severity
from dge.ded.spec import Spec

QC_COLUMNS = ("QC Result", "QC Severity", "QC Observations")

HEADER_FILL = PatternFill("solid", fgColor="44546A")
HEADER_FONT = Font(color="FFFFFF", bold=True)
FILLS = {
    Severity.ERROR: PatternFill("solid", fgColor="F8CBAD"),
    Severity.WARNING: PatternFill("solid", fgColor="FFE699"),
    None: PatternFill("solid", fgColor="C6E0B4"),  # passed
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


def _annotate_sheet(worksheet, spec: Spec, rows: dict[int, list[Finding]]) -> None:
    start = _last_column(worksheet, spec.header_row) + 1

    for offset, title in enumerate(QC_COLUMNS):
        cell = worksheet.cell(row=spec.header_row, column=start + offset, value=title)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical="center")

    widths = (12, 12, 90)
    for offset, width in enumerate(widths):
        worksheet.column_dimensions[get_column_letter(start + offset)].width = width

    # Only rows the checks actually looked at get a verdict; blank filler rows
    # below the data are left alone.
    for number in range(spec.first_data_row, worksheet.max_row + 1):
        found = rows.get(number)
        if found is None:
            if not _row_has_data(worksheet, number, start):
                continue
            result, severity, observations = "Passed", "", ""
            fill = FILLS[None]
        else:
            worst = max(f.severity for f in found)
            result = "Failed" if worst is Severity.ERROR else "Passed with warnings"
            severity = str(worst)
            observations = "\n".join(_describe(f) for f in sorted(found, key=_order))
            fill = FILLS.get(worst, FILLS[None])

        for offset, value in enumerate((result, severity, observations)):
            cell = worksheet.cell(row=number, column=start + offset, value=value)
            cell.fill = fill
            cell.alignment = Alignment(vertical="top", wrap_text=offset == 2)

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

    row = 10
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
