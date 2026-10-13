"""Fill the C.3.2 DQ Rule Execution Logs submission from the execution log.

`execution_log.xlsx` (sheet `execution_log`) holds one row per DQ rule run,
for both the Bronze and the Silver layer. Each layer becomes its own
submission: the template is loaded, its sample rows are cleared, and the
layer's rows are written in their place, one per rule.

The log identifies a rule only by its Purview GUID. The golden rulebook
(`DED_Golden_DQ_Rulebook.xlsx`, sheet `Golden Rulebook-CDErule level`) maps
that GUID - `Silver Technical ID` or `Bronze Technical ID`, by layer - to the
DQO `Rule ID` both layers share, the `Table ID`, the `Attribute ID` and the
rule's lifecycle. The match is on the GUID alone: a rule the rulebook does
not know by its GUID keeps those columns empty, even when a rulebook rule
of the same name exists, because a different GUID under the same name means
the rule was recreated and the rulebook has not caught up. Such rules are
listed, with any same-name candidate, in a gaps CSV beside the outputs.

Criticality and the created / modified dates have no source in either file
and are left empty for someone to fill.

`Threshold_Result` is derived from the score against the rule's threshold:
`Good` at or above it, `At Risk` within 5 percentage points below it, `Bad`
further below. A rule with no threshold gets no result.

Outputs land in `<input dir>/output/`, named
`DED_<YYYYMMDD>_Execution_Log_C.3.2_<Layer>.xlsx`, replacing any existing
file of that name. The inputs are only ever read.

    uv run python src/dge/ded/scripts/submissions/c3.2_execution_logs.py data/ded/submission_execution_log
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

LOG_FILE = "execution_log.xlsx"
LOG_SHEET = "execution_log"
TEMPLATE_FILE = "C3.2 Template.xlsx"
RULEBOOK_FILE = "DED_Golden_DQ_Rulebook.xlsx"
RULEBOOK_SHEET = "Golden Rulebook-CDErule level"
OUTPUT_NAME = "DED_{date:%Y%m%d}_Execution_Log_C.3.2_{layer}.xlsx"
GAPS_NAME = "DED_{date:%Y%m%d}_Execution_Log_C.3.2_rulebook_gaps.csv"
LAYERS = ("Bronze", "Silver")

#: How far below the threshold, in percentage points, a score is still
#: `At Risk` rather than `Bad`.
AT_RISK_BAND = 5

#: The rulebook column holding each layer's Purview GUID.
TECHNICAL_ID = {"Bronze": "Bronze Technical ID", "Silver": "Silver Technical ID"}

#: The rulebook's current rule name. Rulebook names are the silver ones, so
#: only a silver rule can find a same-name candidate.
RULEBOOK_NAME = "Rule_Name_Title(Updating as aligned with DQ)"

#: The rulebook marks a rule with no attribute in more than one way; the
#: template writes `TABLE_LEVEL`.
TABLE_LEVEL = "TABLE_LEVEL"
TABLE_LEVEL_SPELLINGS = frozenset(
    {"table_level", "rule applied at table level, attribute id not applicable"}
)
#: What a rule's technical id looks like: a Purview GUID, or a `DQP-CR-n`
#: id for a complex rule run outside Purview. Anything else in the rulebook's
#: technical id columns - `#NOTFOUND`, `NOTFOUND_SILVERdoc`, `#NAA`, `N/A` -
#: is a placeholder and must not pair rules up with each other.
TECHNICAL_ID_FORMAT = re.compile(r"[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|dqp-cr-\d+")
ATTRIBUTE_ID = re.compile(r"DED_Attribute_\d+")
TABLE_ID = re.compile(r"DED_Data_Table_\d+")

#: Template columns copied from an execution log column as they are.
FROM_LOG = {
    "Name": "Rule_Name_Title",
    "Description": "Description_Business_Definition",
    "Technical_Description": "Description_Business_Definition",
    "Failed_Rows": "num_of_failed_rows",
    "Frequency": "Frequency_Execution_Schedule",
    "Scanned_Time": "Date",
    "Data_Quality_Dimension": "Data_Quality_Dimension",
    "total_rows": "total_rows",
    "Technical_Rule_Reference": "Rule_ID",
    "Secondary Data Element": "Column_Name",
    "Stakeholder: Data Owner": "Business_Owner",
    "Stakeholder: Data Steward": "Technical_Owner",
    "Stakeholder: Governance Owner": "Business_Owner",
    "Stakeholder: Governance Administrator": "Technical_Owner",
}

#: Template columns holding a percentage. The log writes 95 for 95%, the
#: template formats the cell as a percentage and so wants 0.95.
PERCENT_FROM_LOG = {
    "dq_score": "dq_score",
    "Target": "Threshold",
    "Threshold": "Threshold",
}

#: Template columns that hold the same value on every row of a submission.
CONSTANTS = {
    "Input_Port_Name": "N/A",
    "Measuring_Method": "Purview",
    "Arabic_Description": "N/A",
    "Output Port Name": "N/A",
    "Stakeholder Details": (
        "Business Owner and Technical Steward details, including their roles "
        "and available contact information, have been documented to support "
        "governance, coordination, approvals, and issue resolution."
    ),
    "Stakeholders Type": "Business / Technical / Governance",
    "Operation": "Insert",
    "Created_By": "DED DQ Pipeline",
    "Modified By": "DED DQ Pipeline",
}

#: Template columns filled from the rulebook - see `RulebookEntry`.
FROM_RULEBOOK = ("Rule_ID", "Attribute_ID", "Data Table ID", "Lifecycle")

#: Template columns with no source in either file, left empty.
UNFILLED = ("Criticality", "Created_On", "Modified_On")


def collapse(header: object) -> str:
    """A header with its whitespace collapsed, for matching by name."""
    return " ".join(str(header).split()) if header is not None else ""


def text(value: object) -> str:
    return str(value).strip() if value is not None else ""


@dataclass(frozen=True)
class RulebookEntry:
    rule_id: str
    table_id: str | None
    attribute_id: str | None
    lifecycle: str | None
    name: str

    @property
    def active(self) -> bool:
        return (self.lifecycle or "").lower() == "active"


@dataclass
class Rulebook:
    #: Layer -> lowercased GUID -> every rulebook row carrying it.
    by_guid: dict[str, dict[str, list[RulebookEntry]]]
    #: Lowercased rule name -> rulebook rows, for gap candidates only.
    by_name: dict[str, list[tuple[RulebookEntry, str]]]

    def lookup(self, layer: str, guid: object) -> tuple[RulebookEntry | None, str | None]:
        """The rulebook row for a rule, or why there is none."""
        hits = self.by_guid.get(layer, {}).get(text(guid).lower(), [])
        if len(hits) > 1:
            # One GUID on several rows: trust the single active one, if any.
            hits = [h for h in hits if h.active] or hits
        if not hits:
            return None, "GUID not in rulebook"
        if len(hits) > 1:
            return None, "GUID on several rulebook rows: " + ", ".join(h.rule_id for h in hits)
        return hits[0], None


def load_rulebook(path: Path) -> Rulebook:
    """The rulebook's GUID mapping. Its header sits below a banner row."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = wb[RULEBOOK_SHEET].iter_rows(values_only=True)
        for row in rows:
            headers = [collapse(h) for h in row]
            if "Rule ID" in headers:
                break
        else:
            raise ValueError(f"no 'Rule ID' header in [{RULEBOOK_SHEET}]")
        missing = {"Rule ID", "Table ID", "Attribute ID", "Rule Lifecycle Status",
                   RULEBOOK_NAME, *TECHNICAL_ID.values()} - set(headers)
        if missing:
            raise ValueError(f"[{RULEBOOK_SHEET}] lacks columns: {', '.join(sorted(missing))}")
        # The first column of each name: the helper block repeats `Table ID`.
        col = {h: headers.index(h) for h in reversed(headers) if h}

        book = Rulebook(by_guid=defaultdict(lambda: defaultdict(list)), by_name=defaultdict(list))
        for row in rows:
            rule_id = text(row[col["Rule ID"]])
            if not rule_id:
                continue
            entry = RulebookEntry(
                rule_id=rule_id,
                table_id=clean_table_id(row[col["Table ID"]]),
                attribute_id=clean_attribute_id(row[col["Attribute ID"]]),
                lifecycle=text(row[col["Rule Lifecycle Status"]]).capitalize() or None,
                name=text(row[col[RULEBOOK_NAME]]),
            )
            for layer, header in TECHNICAL_ID.items():
                guid = text(row[col[header]]).lower()
                if TECHNICAL_ID_FORMAT.fullmatch(guid):
                    book.by_guid[layer][guid].append(entry)
                    book.by_name[entry.name.lower()].append((entry, f"{layer}:{guid}"))
        return book
    finally:
        wb.close()


def clean_table_id(value: object) -> str | None:
    """A `DED_Data_Table_n` id, or nothing for `Not found` and the like."""
    value = text(value)
    return value if TABLE_ID.fullmatch(value) else None


def clean_attribute_id(value: object) -> str | None:
    """A `DED_Attribute_n` id or `TABLE_LEVEL`, or nothing."""
    value = text(value)
    if ATTRIBUTE_ID.fullmatch(value):
        return value
    if value.lower() in TABLE_LEVEL_SPELLINGS:
        return TABLE_LEVEL
    return None


def load_log(path: Path) -> list[dict[str, object]]:
    """Every execution log row, keyed by header."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = wb[LOG_SHEET].iter_rows(values_only=True)
        headers = [collapse(h) for h in next(rows)]
        return [dict(zip(headers, row)) for row in rows if any(v is not None for v in row)]
    finally:
        wb.close()


def check_log_columns(rows: list[dict[str, object]]) -> list[str]:
    """The log columns this script reads that the log does not have."""
    needed = {"Layer", "Schema_Name", "Dataset_Name"}
    needed |= set(FROM_LOG.values()) | set(PERCENT_FROM_LOG.values())
    present = set(rows[0]) if rows else set()
    return sorted(needed - present)


def as_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def threshold_result(score: object, threshold: object) -> str | None:
    """Good / At Risk / Bad, compared in the log's 0-100 units."""
    if not isinstance(score, (int, float)) or not isinstance(threshold, (int, float)):
        return None
    if score >= threshold:
        return "Good"
    if score >= threshold - AT_RISK_BAND:
        return "At Risk"
    return "Bad"


def derive(row: dict[str, object]) -> dict[str, object]:
    scanned = as_date(row["Date"])
    return {
        "Batch_ID": f"{row['Layer']}_Batch_{scanned:%Y%m%d}" if scanned else None,
        "Threshold_Result": threshold_result(row["dq_score"], row["Threshold"]),
        "HierarchicalPath": f"{row['Schema_Name']}/{row['Dataset_Name']}",
    }


def to_record(row: dict[str, object], entry: RulebookEntry | None) -> dict[str, object]:
    """One template row, keyed by template header."""
    record: dict[str, object] = dict.fromkeys(UNFILLED + FROM_RULEBOOK)
    record.update(CONSTANTS)
    for target, source in FROM_LOG.items():
        record[target] = row[source]
    for target, source in PERCENT_FROM_LOG.items():
        value = row[source]
        record[target] = round(value / 100, 6) if isinstance(value, (int, float)) else None
    record.update(derive(row))
    if entry:
        record.update({
            "Rule_ID": entry.rule_id,
            "Attribute_ID": entry.attribute_id,
            "Data Table ID": entry.table_id,
            "Lifecycle": entry.lifecycle,
        })
    return record


def fill(template: Path, records: list[dict[str, object]], out: Path) -> list[str]:
    """Write `records` over the template's sample rows; return unknown headers."""
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    headers = [collapse(c.value) for c in ws[1]]
    unknown = [h for h in headers if h and h not in records[0]] if records else []

    # Each column keeps the look of the template's first sample row -
    # number formats included, which carry the percentages and dates.
    styles = [
        (copy(c.font), copy(c.fill), copy(c.border), copy(c.alignment), c.number_format)
        for c in ws[2]
    ]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)

    for r, record in enumerate(records, start=2):
        for col, (header, style) in enumerate(zip(headers, styles), start=1):
            cell = ws.cell(row=r, column=col, value=record.get(header))
            cell.font, cell.fill, cell.border, cell.alignment, cell.number_format = style

    last = get_column_letter(len(headers))
    ws.auto_filter.ref = f"A1:{last}{max(len(records) + 1, 2)}"
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return unknown


def gap_row(layer: str, row: dict[str, object], reason: str, book: Rulebook | None) -> dict[str, str]:
    """A rule the rulebook could not identify, with any same-name candidate."""
    candidates = book.by_name.get(text(row["Rule_Name_Title"]).lower(), []) if book else []
    return {
        "Layer": layer,
        "Technical_Rule_Reference": text(row["Rule_ID"]),
        "Name": text(row["Rule_Name_Title"]),
        "Table": text(row["Dataset_Name"]),
        "Dimension": text(row["Data_Quality_Dimension"]),
        "Reason": reason,
        "Same-name rulebook rule": "; ".join(e.rule_id for e, _ in candidates),
        "Its GUID": "; ".join(g for _, g in candidates),
    }


def summarise(layer: str, records: list[dict[str, object]], gaps: int, out: Path) -> None:
    results = Counter(r["Threshold_Result"] or "(no threshold)" for r in records)
    batches = Counter(r["Batch_ID"] for r in records)
    lifecycle = Counter(r["Lifecycle"] or "(unmatched)" for r in records)
    print(f"\n{layer}: {len(records)} rules -> {out}")
    print("  Threshold_Result: " + ", ".join(f"{k} {v}" for k, v in results.most_common()))
    print(f"  Batch_IDs: {len(batches)} ({', '.join(f'{k} {v}' for k, v in batches.most_common(3))}"
          + (", ..." if len(batches) > 3 else "") + ")")
    print(f"  Rulebook: {len(records) - gaps} matched, {gaps} not "
          f"(Rule_ID / Attribute_ID / Data Table ID / Lifecycle left empty)")
    print("  Lifecycle: " + ", ".join(f"{k} {v}" for k, v in lifecycle.most_common()))
    for column in ("Description", "Secondary Data Element", "Attribute_ID", "Data Table ID"):
        blank = sum(r[column] is None for r in records)
        if blank:
            print(f"  {column}: {blank} blank")
    print(f"  Left empty for manual fill: {', '.join(UNFILLED)}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("folder", type=Path,
                        help=f"folder holding {LOG_FILE}, {TEMPLATE_FILE} and {RULEBOOK_FILE}")
    parser.add_argument("--log", type=Path, help=f"execution log (default: <folder>/{LOG_FILE})")
    parser.add_argument("--template", type=Path, help=f"template (default: <folder>/{TEMPLATE_FILE})")
    parser.add_argument("--rulebook", type=Path, help=f"rulebook (default: <folder>/{RULEBOOK_FILE})")
    parser.add_argument("--out-dir", type=Path, help="output folder (default: <folder>/output)")
    parser.add_argument(
        "--date",
        type=lambda s: datetime.strptime(s, "%Y%m%d").date(),
        default=date.today(),
        help="submission date in the file name, YYYYMMDD (default: today)",
    )
    args = parser.parse_args(argv)

    log_path = args.log or args.folder / LOG_FILE
    template = args.template or args.folder / TEMPLATE_FILE
    rulebook_path = args.rulebook or args.folder / RULEBOOK_FILE
    out_dir = args.out_dir or args.folder / "output"
    for path in (log_path, template):
        if not path.is_file():
            print(f"error: {path} not found", file=sys.stderr)
            return 1

    rows = load_log(log_path)
    missing = check_log_columns(rows)
    if missing:
        print(f"error: {log_path} [{LOG_SHEET}] lacks columns: {', '.join(missing)}", file=sys.stderr)
        return 1

    # Without the rulebook the submission is still written, its ids empty.
    book: Rulebook | None = None
    try:
        book = load_rulebook(rulebook_path)
    except (OSError, KeyError, ValueError) as exc:
        print(f"warning: rulebook {rulebook_path} unusable ({exc}); "
              "Rule_ID, Attribute_ID, Data Table ID and Lifecycle left empty", file=sys.stderr)

    other = Counter(r["Layer"] for r in rows if r["Layer"] not in LAYERS)
    if other:
        print(f"warning: rows of unknown layers skipped: {dict(other)}", file=sys.stderr)

    gaps: list[dict[str, str]] = []
    for layer in LAYERS:
        records = []
        layer_gaps = 0
        for row in (r for r in rows if r["Layer"] == layer):
            entry, reason = book.lookup(layer, row["Rule_ID"]) if book else (None, "no rulebook")
            if reason:
                gaps.append(gap_row(layer, row, reason, book))
                layer_gaps += 1
            records.append(to_record(row, entry))
        if not records:
            print(f"warning: no {layer} rows in the log, no submission written", file=sys.stderr)
            continue
        out = out_dir / OUTPUT_NAME.format(date=args.date, layer=layer)
        unknown = fill(template, records, out)
        if unknown:
            print(f"warning: template columns not filled by this script: {', '.join(unknown)}",
                  file=sys.stderr)
        summarise(layer, records, layer_gaps, out)

    if gaps:
        gaps_out = out_dir / GAPS_NAME.format(date=args.date)
        with gaps_out.open("w", newline="", encoding="utf-8-sig") as fh:
            writer = csv.DictWriter(fh, fieldnames=list(gaps[0]))
            writer.writeheader()
            writer.writerows(gaps)
        print(f"\n{len(gaps)} rules without a rulebook match -> {gaps_out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
