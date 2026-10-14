"""Fill the C.3.2 DQ Rule Execution Logs submission from the execution log.

`execution_log.xlsx` (sheet `execution_log`) holds one row per DQ rule run,
for both the Bronze and the Silver layer. Each layer becomes its own
submission: the template is loaded, its headers renamed to the toolkit's
(`TOOLKIT_HEADERS`), its sample rows cleared, and the layer's rows written in
their place, one per rule. `Data Table ID` is not a toolkit column; it is
kept as an extra at the end.

The log identifies a rule only by its technical id - a Purview GUID, or a
`DQP-CR-n` id for a complex rule. The golden rulebook
(`DED_Golden_DQ_Rulebook.xlsx`, sheet `Golden Rulebook-CDErule level`) maps
that id - `Silver Technical ID` or `Bronze Technical ID`, by layer - to the
DQO `Rule ID` both layers share, the table, the attribute, the domain, the
lifecycle and the DGE guardrail. The match is on the technical id alone: a
rule the rulebook does not know by its id keeps those columns empty, even
when a rulebook rule of the same name exists, because a different id under
the same name means the rule was recreated and the rulebook has not caught
up. Such rules are kept out of the submissions and written instead to a gaps
workbook beside them - one sheet per layer, in the submission's columns,
with the reason and any same-name rulebook candidate - so a row can be
pasted back once its DQO is settled. Their table and domain are still known
from the other rules run against the same Purview dataset, and are taken
from there.

`Target` is the rule's own threshold from the log; `Threshold` is the DGE
guardrail from the rulebook. Where the rulebook gives no guardrail, the DGE
guardrail for the rule's dimension (`DGE_GUARDRAIL`) is assumed, and the
rule is listed on the gaps workbook's `Assumed guardrail` sheet.
`Threshold Result` compares the score with the guardrail: `Good` at or above
it, `At Risk` within 5 percentage points below it, `Bad` further below.

`Created On` and `Modified On` are the submission date. Criticality has no
source in either file and is left empty for someone to fill.

Outputs land in `<input dir>/output/`, named
`DED_<YYYYMMDD>_Execution_Log_C.3.2_<Layer>.xlsx`, replacing any existing
file of that name. The inputs are only ever read.

    uv run python src/dge/ded/scripts/submissions/c3.2_execution_logs.py data/ded/submission_execution_log
"""

from __future__ import annotations

import argparse
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
GAPS_NAME = "DED_{date:%Y%m%d}_Execution_Log_C.3.2_rulebook_gaps.xlsx"
LAYERS = ("Bronze", "Silver")

#: The template's headers that the toolkit names differently. Every other
#: template header is already the toolkit's.
TOOLKIT_HEADERS = {
    "Technical_Description": "Technical Description",
    "Failed_Rows": "Failed Rows",
    "Input_Port_Name": "Input Port Name",
    "Scanned_Time": "Scanned Time",
    "Measuring_Method": "Measuring Method",
    "Threshold_Result": "Threshold Result",
    "Data_Quality_Dimension": "DQ_Dimension",
    "total_rows": "Total Rows",
    "dq_score": "Score",
    "Technical_Rule_Reference": "Technical Rule Reference",
    "Arabic_Description": "Arabic Description",
    "Created_On": "Created On",
    "Created_By": "Created By",
    "Modified_On": "Modified On",
}

#: How far below the threshold, in percentage points, a score is still
#: `At Risk` rather than `Bad`.
AT_RISK_BAND = 5

#: The rulebook column holding each layer's technical id.
TECHNICAL_ID = {"Bronze": "Bronze Technical ID", "Silver": "Silver Technical ID"}

#: What a technical id looks like: a Purview GUID, or a `DQP-CR-n` id for a
#: complex rule run outside Purview. Anything else in the rulebook's
#: technical id columns - `#NOTFOUND`, `NOTFOUND_SILVERdoc`, `#NAA`, `N/A` -
#: is a placeholder and must not pair rules up with each other.
TECHNICAL_ID_FORMAT = re.compile(r"[0-9a-f]{8}(-[0-9a-f]{4}){3}-[0-9a-f]{12}|dqp-cr-\d+")

#: The rulebook's current rule name and expression - the `(Updating as
#: aligned with DQ)` columns, not their `(Reference, don't delete)` twins.
#: Rulebook names are the silver ones, so only a silver rule can find a
#: same-name candidate.
RULEBOOK_NAME = "Rule_Name_Title(Updating as aligned with DQ)"
RULEBOOK_EXPRESSION = "Rule_Expression_Logic (Updating as aligned with DQ)"
RULEBOOK_GUARDRAIL = "DGE Guardrail %"
RULEBOOK_DIMENSION = "DQ Dimension (Updating as aligned with DQ)"

#: The DGE guardrail per DQ dimension, as the rulebook applies it, assumed
#: for a rule whose rulebook row gives none. `check_guardrails` reports any
#: rulebook row that departs from it.
DGE_GUARDRAIL = {
    "completeness": 0.95,
    "consistency": 0.95,
    "timeliness": 0.90,
    "accuracy": 0.90,
    "uniqueness": 1.00,
    "validity": 1.00,
}

#: The rulebook's lifecycle statuses in the toolkit's vocabulary (Draft,
#: Published, Deprecated). A status not listed here is written as it is.
LIFECYCLE = {"active": "Published", "draft": "Draft", "inactive": "Deprecated"}

#: The rulebook marks a rule with no attribute in more than one way; the
#: template writes `TABLE_LEVEL`.
TABLE_LEVEL = "TABLE_LEVEL"
TABLE_LEVEL_SPELLINGS = frozenset(
    {"table_level", "rule applied at table level, attribute id not applicable"}
)
ATTRIBUTE_ID = re.compile(r"DED_Attribute_\d+")
TABLE_ID = re.compile(r"DED_Data_Table_\d+")
NOT_FOUND = frozenset({"", "not found", "n/a"})

#: Output columns copied from an execution log column as they are.
FROM_LOG = {
    "Name": "Rule_Name_Title",
    "Description": "Description_Business_Definition",
    "Failed Rows": "num_of_failed_rows",
    "Frequency": "Frequency_Execution_Schedule",
    "Scanned Time": "Date",
    "DQ_Dimension": "Data_Quality_Dimension",
    "Total Rows": "total_rows",
    "Technical Rule Reference": "Rule_ID",
    "Stakeholder: Data Owner": "Business_Owner",
    "Stakeholder: Data Steward": "Technical_Owner",
    "Stakeholder: Governance Owner": "Business_Owner",
    "Stakeholder: Governance Administrator": "Technical_Owner",
}

#: Output columns holding a percentage. The log writes 95 for 95%, the
#: template formats the cell as a percentage and so wants 0.95.
PERCENT_FROM_LOG = {
    "Score": "dq_score",
    "Target": "Threshold",
}

#: Output columns that hold the same value on every row of a submission.
CONSTANTS = {
    "Input Port Name": "N/A",
    "Measuring Method": "Purview",
    "Arabic Description": "N/A",
    "Output Port Name": "N/A",
    "Stakeholder Details": (
        "Business Owner and Technical Steward details, including their roles "
        "and available contact information, have been documented to support "
        "governance, coordination, approvals, and issue resolution."
    ),
    # Owners and stewards are the dataset's, not assigned per rule.
    "Stakeholders Type": "Inherited from higher level",
    "Operation": "Insert",
    "Created By": "DED DQ Pipeline",
    "Modified By": "DED DQ Pipeline",
}

#: Output columns with no source in either file, left empty. `Secondary
#: Data Element` is the *other* element of a cross-field check; the log
#: names only the rule's main column, which `Attribute_ID` already covers.
UNFILLED = ("Criticality", "Secondary Data Element")

#: Output columns filled from the rulebook, empty when the rule is not in it.
FROM_RULEBOOK = ("Rule_ID", "Attribute_ID", "Lifecycle")

#: Gaps workbook columns added after the submission's own.
GAP_COLUMNS = ("Reason", "Same-name rulebook rule", "Its technical id")


def collapse(header: object) -> str:
    """A header with its whitespace collapsed, for matching by name."""
    return " ".join(str(header).split()) if header is not None else ""


def text(value: object) -> str:
    return str(value).strip() if value is not None else ""


def number(value: object) -> float | None:
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(text(value))
    except ValueError:
        return None


@dataclass(frozen=True)
class RulebookEntry:
    rule_id: str
    table_id: str | None
    attribute_id: str | None
    domain: str | None
    lifecycle: str | None
    #: The DGE guardrail as a fraction - 0.95 for 95%.
    guardrail: float | None
    dimension: str
    expression: str | None
    name: str

    @property
    def active(self) -> bool:
        return (self.lifecycle or "").lower() == "active"


@dataclass
class Rulebook:
    #: Layer -> lowercased technical id -> every rulebook row carrying it.
    by_guid: dict[str, dict[str, list[RulebookEntry]]]
    #: Lowercased rule name -> rulebook rows, for gap candidates only.
    by_name: dict[str, list[tuple[RulebookEntry, str]]]
    entries: list[RulebookEntry]

    def lookup(self, layer: str, guid: object) -> tuple[RulebookEntry | None, str | None]:
        """The rulebook row for a rule, or why there is none."""
        hits = self.by_guid.get(layer, {}).get(text(guid).lower(), [])
        if len(hits) > 1:
            # One id on several rows: trust the single active one, if any.
            hits = [h for h in hits if h.active] or hits
        if not hits:
            return None, "technical id not in rulebook"
        if len(hits) > 1:
            return None, "technical id on several rulebook rows: " + ", ".join(h.rule_id for h in hits)
        return hits[0], None


def load_rulebook(path: Path) -> Rulebook:
    """The rulebook's technical id mapping. Its header sits below a banner row."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = wb[RULEBOOK_SHEET].iter_rows(values_only=True)
        for row in rows:
            headers = [collapse(h) for h in row]
            if "Rule ID" in headers:
                break
        else:
            raise ValueError(f"no 'Rule ID' header in [{RULEBOOK_SHEET}]")
        needed = {"Rule ID", "Table ID", "Attribute ID", "Domain", "Rule Lifecycle Status",
                  RULEBOOK_NAME, RULEBOOK_EXPRESSION, RULEBOOK_GUARDRAIL, RULEBOOK_DIMENSION,
                  *TECHNICAL_ID.values()}
        missing = needed - set(headers)
        if missing:
            raise ValueError(f"[{RULEBOOK_SHEET}] lacks columns: {', '.join(sorted(missing))}")
        # The first column of each name: the helper block repeats `Table ID`.
        col = {h: headers.index(h) for h in reversed(headers) if h}

        book = Rulebook(by_guid=defaultdict(lambda: defaultdict(list)), by_name=defaultdict(list),
                        entries=[])
        for row in rows:
            rule_id = text(row[col["Rule ID"]])
            if not rule_id:
                continue
            domain = text(row[col["Domain"]])
            entry = RulebookEntry(
                rule_id=rule_id,
                table_id=clean_table_id(row[col["Table ID"]]),
                attribute_id=clean_attribute_id(row[col["Attribute ID"]]),
                domain=None if domain.lower() in NOT_FOUND else domain,
                lifecycle=text(row[col["Rule Lifecycle Status"]]) or None,
                guardrail=number(row[col[RULEBOOK_GUARDRAIL]]),
                dimension=text(row[col[RULEBOOK_DIMENSION]]),
                expression=text(row[col[RULEBOOK_EXPRESSION]]) or None,
                name=text(row[col[RULEBOOK_NAME]]),
            )
            book.entries.append(entry)
            for layer, header in TECHNICAL_ID.items():
                guid = text(row[col[header]]).lower()
                if TECHNICAL_ID_FORMAT.fullmatch(guid):
                    book.by_guid[layer][guid].append(entry)
                    book.by_name[entry.name.lower()].append((entry, f"{layer}:{guid}"))
        return book
    finally:
        wb.close()


def check_guardrails(book: Rulebook) -> Counter[tuple[str, float]]:
    """Rulebook guardrails that depart from `DGE_GUARDRAIL`, by dimension
    and value - a sign the assumed guardrail may be wrong."""
    return Counter(
        (e.dimension, e.guardrail)
        for e in book.entries
        if e.guardrail is not None
        and DGE_GUARDRAIL.get(e.dimension.lower(), e.guardrail) != e.guardrail
    )


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


def toolkit_lifecycle(status: str | None) -> str | None:
    if not status:
        return None
    return LIFECYCLE.get(status.lower(), status)


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
    needed = {"Layer", "Schema_Name", "Dataset_Name", "Dataset_ID", "Source_System",
              "Column_Name", "Rule_Expression_Logic"}
    needed |= set(FROM_LOG.values()) | set(PERCENT_FROM_LOG.values())
    present = set(rows[0]) if rows else set()
    return sorted(needed - present)


@dataclass(frozen=True)
class DatasetInfo:
    """What the rulebook says about the table behind a Purview dataset."""
    table_id: str | None
    domain: str | None


def dataset_info(rows: list[dict[str, object]], entries: list[RulebookEntry | None]) -> dict[str, DatasetInfo]:
    """Table and domain per Purview dataset, from the rules the rulebook knows.

    Every rule run against a dataset is run against the same table, so a
    rule the rulebook does not know still has a known table and domain when
    another rule on its dataset is known. A dataset whose known rules
    disagree gets nothing.
    """
    tables: dict[str, set[str]] = defaultdict(set)
    domains: dict[str, set[str]] = defaultdict(set)
    for row, entry in zip(rows, entries):
        if entry:
            dataset = text(row["Dataset_ID"]).lower()
            if entry.table_id:
                tables[dataset].add(entry.table_id)
            if entry.domain:
                domains[dataset].add(entry.domain)
    one = lambda values: next(iter(values)) if len(values) == 1 else None
    return {
        d: DatasetInfo(one(tables.get(d, set())), one(domains.get(d, set())))
        for d in tables.keys() | domains.keys()
    }


def as_date(value: object) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def threshold_result(score: object, threshold: float | None) -> str | None:
    """Good / At Risk / Bad, compared in the log's 0-100 units."""
    if not isinstance(score, (int, float)) or threshold is None:
        return None
    threshold *= 100
    if score >= threshold:
        return "Good"
    if score >= threshold - AT_RISK_BAND:
        return "At Risk"
    return "Bad"


def hierarchical_path(row: dict[str, object], domain: str | None) -> str | None:
    """domain / system / schema / table / field - the field only when the
    rule has one. Without the domain the path is not written at all."""
    if not domain:
        return None
    parts = [domain, row["Source_System"], row["Schema_Name"], row["Dataset_Name"], row["Column_Name"]]
    return "/".join(text(p) for p in parts if text(p))


def to_record(row: dict[str, object], entry: RulebookEntry | None,
              dataset: DatasetInfo | None, submitted: date) -> dict[str, object]:
    """One output row, keyed by toolkit header. `ASSUMED` marks a row whose
    guardrail was assumed rather than read from the rulebook."""
    record: dict[str, object] = dict.fromkeys(UNFILLED + FROM_RULEBOOK)
    record.update(CONSTANTS)
    record["Created On"] = record["Modified On"] = submitted
    for target, source in FROM_LOG.items():
        record[target] = row[source]
    for target, source in PERCENT_FROM_LOG.items():
        value = row[source]
        record[target] = round(value / 100, 6) if isinstance(value, (int, float)) else None

    scanned = as_date(row["Date"])
    record["Batch_ID"] = f"{row['Layer']}_Batch_{scanned:%Y%m%d}" if scanned else None
    # What ran is what the log says; the rulebook's expression is the fallback.
    record["Technical Description"] = (
        text(row["Rule_Expression_Logic"]) or (entry.expression if entry else None) or None
    )
    table_id = (entry.table_id if entry else None) or (dataset.table_id if dataset else None)
    domain = (entry.domain if entry else None) or (dataset.domain if dataset else None)
    record["Data Table ID"] = table_id
    record["HierarchicalPath"] = hierarchical_path(row, domain)

    guardrail = entry.guardrail if entry else None
    record[ASSUMED] = guardrail is None
    if guardrail is None:
        guardrail = DGE_GUARDRAIL.get(text(row["Data_Quality_Dimension"]).lower())
    record["Threshold"] = guardrail
    record["Threshold Result"] = threshold_result(row["dq_score"], guardrail)

    if entry:
        record.update({
            "Rule_ID": entry.rule_id,
            "Attribute_ID": entry.attribute_id,
            "Lifecycle": toolkit_lifecycle(entry.lifecycle),
        })
    return record


#: The record key flagging an assumed guardrail; not an output column.
ASSUMED = "_guardrail_assumed"

Style = tuple


def open_template(template: Path) -> tuple[openpyxl.Workbook, list[str], list[Style]]:
    """The template with toolkit headers and no sample rows, its headers, and
    each column's cell style."""
    wb = openpyxl.load_workbook(template)
    ws = wb.worksheets[0]
    for cell in ws[1]:
        cell.value = TOOLKIT_HEADERS.get(collapse(cell.value), cell.value)
    headers = [collapse(c.value) for c in ws[1]]
    # Each column keeps the look of the template's first sample row -
    # number formats included, which carry the percentages and dates.
    styles = [
        (copy(c.font), copy(c.fill), copy(c.border), copy(c.alignment), c.number_format)
        for c in ws[2]
    ]
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row - 1)
    return wb, headers, styles


def write_rows(ws, headers: list[str], styles: list[Style], records: list[dict[str, object]]) -> None:
    for r, record in enumerate(records, start=2):
        for col, (header, style) in enumerate(zip(headers, styles), start=1):
            cell = ws.cell(row=r, column=col, value=record.get(header))
            cell.font, cell.fill, cell.border, cell.alignment, cell.number_format = style
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(len(records) + 1, 2)}"


def write_submission(template: Path, records: list[dict[str, object]], out: Path) -> list[str]:
    """Write `records` over the template's sample rows; return unknown headers."""
    wb, headers, styles = open_template(template)
    write_rows(wb.worksheets[0], headers, styles, records)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return [h for h in headers if h and records and h not in records[0]]


def write_gaps(template: Path, removed: dict[str, list[dict[str, object]]],
               assumed: list[dict[str, object]], out: Path) -> None:
    """One `<Layer> removed` sheet per layer, in the submission's columns plus
    `GAP_COLUMNS`, and an `Assumed guardrail` sheet."""
    wb, headers, styles = open_template(template)
    first = wb.worksheets[0]
    last = first.cell(row=1, column=len(headers))
    for i, name in enumerate(GAP_COLUMNS, start=len(headers) + 1):
        cell = first.cell(row=1, column=i, value=name)
        cell.font, cell.fill, cell.border, cell.alignment = (
            copy(last.font), copy(last.fill), copy(last.border), copy(last.alignment))
    headers = headers + list(GAP_COLUMNS)
    styles = styles + [styles[0]] * len(GAP_COLUMNS)

    # Renamed before copying: the template's title plus " Copy" is too long
    # for Excel.
    first.title = f"{LAYERS[0]} removed"
    sheets = [first] + [wb.copy_worksheet(first) for _ in LAYERS[1:]]
    for ws, layer in zip(sheets, LAYERS):
        ws.title = f"{layer} removed"
        write_rows(ws, headers, styles, removed.get(layer, []))

    ws = wb.create_sheet("Assumed guardrail")
    columns = ("Layer", "In submission", "Rule_ID", "Technical Rule Reference", "Name",
               "DQ_Dimension", "Threshold")
    ws.append(columns)
    for record in assumed:
        ws.append([record.get(c) for c in columns])
    for row in ws.iter_rows(min_row=2, min_col=len(columns), max_col=len(columns)):
        row[0].number_format = "0%"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{max(len(assumed) + 1, 2)}"

    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)


def gap_columns(layer: str, row: dict[str, object], reason: str,
                book: Rulebook | None) -> dict[str, str]:
    """Why a rule was removed, with any same-name rulebook rule of its layer."""
    candidates = book.by_name.get(text(row["Rule_Name_Title"]).lower(), []) if book else []
    candidates = [(e, g.split(":", 1)[1]) for e, g in candidates if g.startswith(f"{layer}:")]
    return {
        "Reason": reason,
        "Same-name rulebook rule": "; ".join(e.rule_id for e, _ in candidates),
        "Its technical id": "; ".join(g for _, g in candidates),
    }


def summarise(layer: str, records: list[dict[str, object]], removed: int, out: Path) -> None:
    def counts(column: str, blank: str) -> str:
        c = Counter(r[column] or blank for r in records)
        return ", ".join(f"{k} {v}" for k, v in c.most_common())

    batches = Counter(r["Batch_ID"] for r in records)
    print(f"\n{layer}: {len(records)} rules -> {out}")
    print(f"  Removed (not in rulebook): {removed}")
    print(f"  Threshold Result: {counts('Threshold Result', '(no guardrail)')}")
    print(f"  Threshold (DGE guardrail): {counts('Threshold', '(none)')}"
          f" - {sum(r[ASSUMED] for r in records)} assumed by dimension")
    print(f"  Lifecycle: {counts('Lifecycle', '(none)')}")
    print(f"  Batch_IDs: {len(batches)} ({', '.join(f'{k} {v}' for k, v in batches.most_common(3))}"
          + (", ..." if len(batches) > 3 else "") + ")")
    for column in ("Description", "Technical Description", "Attribute_ID", "Data Table ID", "HierarchicalPath"):
        blank = sum(r[column] is None for r in records)
        if blank:
            print(f"  {column}: {blank} blank")
    print(f"  Left empty by design: {', '.join(UNFILLED)}")


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
              f"{', '.join(FROM_RULEBOOK)}, Data Table ID and HierarchicalPath left empty",
              file=sys.stderr)

    if book:
        departures = check_guardrails(book)
        if departures:
            print("warning: rulebook guardrails departing from the assumed DGE guardrail: "
                  + ", ".join(f"{d} {g:.0%} x{n}" for (d, g), n in departures.most_common()),
                  file=sys.stderr)

    other = Counter(r["Layer"] for r in rows if r["Layer"] not in LAYERS)
    if other:
        print(f"warning: rows of unknown layers skipped: {dict(other)}", file=sys.stderr)

    removed: dict[str, list[dict[str, object]]] = {}
    assumed: list[dict[str, object]] = []
    for layer in LAYERS:
        layer_rows = [r for r in rows if r["Layer"] == layer]
        if not layer_rows:
            print(f"warning: no {layer} rows in the log, no submission written", file=sys.stderr)
            continue
        lookups = [
            book.lookup(layer, row["Rule_ID"]) if book else (None, "no rulebook")
            for row in layer_rows
        ]
        entries = [entry for entry, _ in lookups]
        datasets = dataset_info(layer_rows, entries)
        records, removed[layer] = [], []
        for row, (entry, reason) in zip(layer_rows, lookups):
            record = to_record(row, entry, datasets.get(text(row["Dataset_ID"]).lower()), args.date)
            if reason:
                # Without its DQO a rule cannot be submitted; it waits in the gaps.
                record.update(gap_columns(layer, row, reason, book))
                removed[layer].append(record)
            else:
                records.append(record)
            if record[ASSUMED]:
                assumed.append({**record, "Layer": layer, "In submission": "N" if reason else "Y"})

        out = out_dir / OUTPUT_NAME.format(date=args.date, layer=layer)
        unknown = write_submission(template, records, out)
        if unknown:
            print(f"warning: template columns not filled by this script: {', '.join(unknown)}",
                  file=sys.stderr)
        summarise(layer, records, len(removed[layer]), out)

    if any(removed.values()) or assumed:
        gaps_out = out_dir / GAPS_NAME.format(date=args.date)
        write_gaps(template, removed, assumed, gaps_out)
        print(f"\n{sum(map(len, removed.values()))} rules removed, "
              f"{len(assumed)} guardrails assumed -> {gaps_out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
