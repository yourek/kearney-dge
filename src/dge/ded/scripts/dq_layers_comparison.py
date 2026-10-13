"""Compare the bronze and silver DQ rule definitions, sprint by sprint.

The same logical rule is implemented twice - once over the bronze layer and
once over silver_cleansed - and the two should agree on everything but the
names. `SilverBronzeRuleMapping.xlsx` pairs them up: each row of its `S1`,
`S2`, `S3` sheets carries a bronze rule id (`Purview_B`), a silver rule id
(`Purview_S`) and the DQO label they share.

The report is one row per mapper row, holding the DQO, the two rule ids, and
for every compared column a verdict plus - where the two disagree - the
bronze and the silver value side by side. Read-only: the inputs are never
modified.

    uv run python -m dge.ded.scripts.dq_layers_comparison data/ded/dq_layers_comparison
"""

from __future__ import annotations

import argparse
import ast
import csv
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SPRINTS = (1, 2, 3)
MAPPING_FILE = "SilverBronzeRuleMapping.xlsx"
BRONZE_FILE = "dbo sprint{sprint}.xlsx"
SILVER_FILE = "silver_cleansed sprint{sprint}.xlsx"
RULE_SHEET = "Sheet1"

#: The columns worth comparing. The rest of the rule export - `name`, the
#: four ids, `dataset_name`, `description` - differs between the layers by
#: design (VIRTUAL_UMBRELLA against permit_umbrella) and says nothing.
COMPARED = (
    "status",
    "dimension",
    "score",
    "type",
    "column",
    "rule_expression_logic",
    "filter_Expression_logic",
    "null_Expression_logic",
    "typeProperties",
)

#: Columns holding a SQL fragment, compared as a token sequence rather than
#: as text - see `normalise_sql`.
SQL_COLUMNS = frozenset(
    {"rule_expression_logic", "filter_Expression_logic", "null_Expression_logic"}
)

#: Columns holding a Python dict repr, compared structurally - see
#: `normalise_properties`.
STRUCTURED_COLUMNS = frozenset({"typeProperties"})

#: Columns holding a comma-separated list of column names. These name the
#: participants in a rule, not a sequence, and the two layers order a
#: uniqueness key differently for the same rule, so they are compared as a
#: set - as the `columns` list inside `typeProperties` already is.
LIST_COLUMNS = frozenset({"column"})

#: Keys inside `typeProperties` whose value is a SQL fragment.
SQL_VALUED_KEYS = frozenset({"condition", "filterCriteria", "emptyCriteria"})

#: Keys inside `typeProperties` naming the dataset the rule runs against.
#: Bronze says `virtual_commercial_activities` where silver says
#: `commercial_activities` - the same disagreement as `dataset_name`, which
#: is excluded from the comparison for the same reason, so these are too.
LAYER_SPECIFIC_KEYS = frozenset({"referencedDatasetId", "referencedDatasetAlias"})

#: Punctuation that quotes or templates an identifier rather than forming
#: part of the expression: bronze writes `{REC_STATUS} == 'A'` in Purview's
#: templated dialect where silver writes `rec_status = 'A'` in plain
#: SparkSQL, and either layer may backtick a name.
QUOTING = frozenset({"{", "}", "`", '"'})

#: Splits a SQL fragment into string literals, words and single characters.
#: The literal alternative comes first so that a quoted string is taken
#: whole - its spacing and case are data and must survive normalisation,
#: unlike the spacing and case around it.
TOKEN = re.compile(r"('(?:[^']|'')*')|(\w+)|(\S)")

#: Values the mapper uses in place of a rule id or a DQO to mean "there is
#: none". Matched after casefolding and stripping non-alphanumerics, so
#: `N/A`, `n/a` and the typo `N?A` all land here.
SENTINELS = {"na", "none", "null", "tbd", ""}

#: A real DQO label looks like `DQO-475`; anything else is a placeholder,
#: including the literal "Need to add in golden rule book".
DQO_PATTERN = re.compile(r"^DQO-\d+$", re.IGNORECASE)

SAME = "same"
DIFFERENT = "different"
NOT_COMPARED = "not-compared"

MATCHED = "matched"
BRONZE_MISSING = "bronze-missing"
BRONZE_NOT_FOUND = "bronze-not-found"
SILVER_NOT_FOUND = "silver-not-found"
BRONZE_UNMAPPED = "bronze-unmapped"

GREEN = PatternFill("solid", fgColor="C6EFCE")
RED = PatternFill("solid", fgColor="FFC7CE")
AMBER = PatternFill("solid", fgColor="FFEB9C")
GREY = PatternFill("solid", fgColor="D9D9D9")
HEADER = PatternFill("solid", fgColor="DDEBF7")

STATUS_FILL = {
    MATCHED: None,
    BRONZE_MISSING: GREY,
    BRONZE_NOT_FOUND: AMBER,
    SILVER_NOT_FOUND: AMBER,
    BRONZE_UNMAPPED: GREY,
}


# --------------------------------------------------------------------------
# reading
# --------------------------------------------------------------------------


def read_sheet(path: Path, sheet: str) -> list[dict]:
    """Read a sheet into a list of row dicts, keyed by its header row.

    Headers are whitespace-collapsed, because the same column is spelt with
    one space in one workbook and two in another.
    """
    book = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet not in book.sheetnames:
            raise ValueError(f"{path.name} has no sheet named {sheet!r}")
        rows = book[sheet].iter_rows(values_only=True)
        try:
            header = [collapse(value) for value in next(rows)]
        except StopIteration:
            raise ValueError(f"{path.name}:{sheet} is empty") from None
        return [
            dict(zip(header, values))
            for values in rows
            if any(value is not None for value in values)
        ]
    finally:
        book.close()


def collapse(value) -> str:
    """Collapse a header's whitespace so the same column matches everywhere."""
    return re.sub(r"\s+", " ", str(value)).strip() if value is not None else ""


def text(value) -> str:
    return "" if value is None else str(value).strip()


def is_sentinel(value: str) -> bool:
    """True when a cell is a placeholder rather than a real identifier."""
    return re.sub(r"[^a-z0-9]", "", value.lower()) in SENTINELS


def load_mapping(path: Path, sprints=SPRINTS) -> tuple[list[dict], list[str]]:
    """Read the per-sprint mapper sheets into one list of pairings.

    The sheets disagree on their attribute columns - `S1` has a single
    `Attribute_Name`, `S2`/`S3` split it in two - so only the columns common
    to all three are carried through. Returns the pairings and any problems
    worth reporting.
    """
    problems: list[str] = []
    pairings: list[dict] = []
    seen: set[tuple[int, str, str]] = set()

    for sprint in sprints:
        sheet = f"S{sprint}"
        try:
            rows = read_sheet(path, sheet)
        except Exception as exc:
            problems.append(f"mapping sheet {sheet} could not be read: {exc}")
            continue

        for number, row in enumerate(rows, start=2):
            bronze = text(row.get("Purview_B"))
            silver = text(row.get("Purview_S"))
            key = (sprint, bronze.lower(), silver.lower())
            if key in seen:
                problems.append(
                    f"{sheet} row {number}: duplicate pairing "
                    f"{bronze or '(blank)'} / {silver or '(blank)'}, ignored"
                )
                continue
            seen.add(key)
            pairings.append(
                {
                    "sprint": sprint,
                    "mapping_sheet": sheet,
                    "mapping_row": number,
                    "DQO": text(row.get("DQO")),
                    "bronze_rule_id": bronze,
                    "silver_rule_id": silver,
                    "Rule_Title_B": text(row.get("Rule_Title_B")),
                    "Rule_Title_S": text(row.get("Rule_Title_S")),
                    "Asset_Name_B": text(row.get("Asset_Name_B")),
                    "Asset_Name_S": text(row.get("Asset_Name_S")),
                }
            )
    return pairings, problems


# --------------------------------------------------------------------------
# comparing
# --------------------------------------------------------------------------


def normalise(value):
    """Reduce a cell to the form the two layers are expected to agree on.

    Bronze names its tables and columns in upper case and silver in lower,
    which is convention rather than disagreement, so the comparison folds
    case and collapses whitespace. A blank and a `None` are the same thing.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return re.sub(r"\s+", " ", str(value)).strip().lower()


def normalise_sql(value) -> tuple:
    """Reduce a SQL fragment to the token sequence the layers should share.

    The two layers express the same logic in different dialects: bronze uses
    Purview's templated form, `{REC_STATUS} == 'A'`, and silver plain
    SparkSQL, `rec_status = 'A'`. Identifier case, the spacing between
    tokens, the braces and backticks that quote a name, and `==` against `=`
    are all dialect rather than meaning, so they are normalised away.

    What is *not* touched is the inside of a string literal. The condition
    `MACHINE_ADDRESS RLIKE '^[^ ]*( [^ ]+)*$'` depends on the spaces in that
    regex, and `rec_status = 'A'` on the case of the `A`, so a literal is
    kept verbatim and only the text around it is folded.
    """
    if value is None:
        return ()
    tokens: list[str] = []
    for literal, word, character in TOKEN.findall(str(value)):
        if literal:
            tokens.append(literal)
        elif word:
            tokens.append(word.lower())
        elif character in QUOTING:
            continue
        elif character == "=" and tokens and tokens[-1] == "=":
            continue  # `==` and `=` are the same comparison here
        else:
            tokens.append(character)
    return tuple(tokens)


def normalise_properties(value):
    """Reduce a `typeProperties` cell to a comparable structure.

    The cell holds the repr of a Python dict - condition, column list,
    filter, dialect. Comparing it as text reports every rule as different,
    because the column names inside it are quoted strings that follow each
    layer's casing (`'MACHINE_ADDRESS'` against `'machine_address'`) and the
    column list is not always in the same order.

    So it is parsed and compared structurally: the SQL-valued keys go
    through `normalise_sql`, other strings are folded as identifiers, lists
    are compared without regard to order, and a key that is absent is the
    same as one explicitly set to `None` - silver writes
    `'emptyCriteria': None` where bronze simply omits it. Anything that will
    not parse falls back to the SQL tokeniser, so a malformed cell is still
    compared rather than crashing the run.
    """
    if value is None:
        return ()
    try:
        return _structure(ast.literal_eval(str(value)))
    except (ValueError, SyntaxError, MemoryError, RecursionError):
        return ("unparsed", normalise_sql(value))


def _structure(value, key: str | None = None):
    """Normalise one node of a parsed `typeProperties` value."""
    if isinstance(value, dict):
        items = []
        for name, item in value.items():
            name = str(name)
            if name in LAYER_SPECIFIC_KEYS:
                continue
            normalised = _structure(item, name)
            if normalised in (None, "", ()):
                continue  # an absent key and an explicit None are the same
            items.append((name, normalised))
        return tuple(sorted(items, key=lambda pair: pair[0]))
    if isinstance(value, (list, tuple, set, frozenset)):
        # The column list is a set of participants, not a sequence; the two
        # layers order it differently for the same rule.
        return frozenset(_structure(item) for item in value)
    if isinstance(value, str):
        return normalise_sql(value) if key in SQL_VALUED_KEYS else value.strip().lower()
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def normalise_list(value) -> frozenset:
    """Reduce a comma-separated list of column names to a set."""
    if value is None:
        return frozenset()
    return frozenset(part.strip().lower() for part in str(value).split(",") if part.strip())


def normalise_for(column: str, value):
    """Normalise a cell the way its column demands."""
    if column in STRUCTURED_COLUMNS:
        return normalise_properties(value)
    if column in SQL_COLUMNS:
        return normalise_sql(value)
    if column in LIST_COLUMNS:
        return normalise_list(value)
    return normalise(value)


@dataclass
class Comparison:
    """One mapper row, resolved against both layers and compared."""

    pairing: dict
    status: str
    bronze: dict | None = None
    silver: dict | None = None
    verdicts: dict[str, str] = field(default_factory=dict)
    note: str = ""

    @property
    def differences(self) -> int:
        return sum(1 for verdict in self.verdicts.values() if verdict == DIFFERENT)


def compare(pairing: dict, bronze_rules: dict, silver_rules: dict) -> Comparison:
    """Resolve one pairing against the two layers and compare the columns."""
    bronze_id = pairing["bronze_rule_id"]
    silver_id = pairing["silver_rule_id"]

    if not bronze_id or is_sentinel(bronze_id):
        return Comparison(
            pairing,
            BRONZE_MISSING,
            silver=silver_rules.get(silver_id.lower()),
            note="the mapper records no bronze counterpart for this silver rule",
        )

    bronze = bronze_rules.get(bronze_id.lower())
    silver = silver_rules.get(silver_id.lower())
    if bronze is None:
        return Comparison(
            pairing,
            BRONZE_NOT_FOUND,
            silver=silver,
            note=f"bronze rule id {bronze_id} is not in the bronze export",
        )
    if silver is None:
        return Comparison(
            pairing,
            SILVER_NOT_FOUND,
            bronze=bronze,
            note=f"silver rule id {silver_id} is not in the silver export",
        )

    verdicts = {}
    for column in COMPARED:
        if column not in bronze or column not in silver:
            verdicts[column] = NOT_COMPARED
        else:
            verdicts[column] = (
                SAME
                if normalise_for(column, bronze[column])
                == normalise_for(column, silver[column])
                else DIFFERENT
            )
    return Comparison(pairing, MATCHED, bronze=bronze, silver=silver, verdicts=verdicts)


def index(rows: list[dict], label: str, problems: list[str]) -> dict[str, dict]:
    """Key rules by their lowercased rule id, reporting any duplicates."""
    indexed: dict[str, dict] = {}
    for row in rows:
        rule_id = text(row.get("rule_id")).lower()
        if not rule_id:
            continue
        if rule_id in indexed:
            problems.append(f"{label}: duplicate rule_id {rule_id}, the first one is used")
            continue
        indexed[rule_id] = row
    return indexed


def unmapped_pairing(sprint: int, rule: dict) -> dict:
    """A pairing standing in for a bronze rule the mapper never mentions."""
    return {
        "sprint": sprint,
        "mapping_sheet": "",
        "mapping_row": "",
        "DQO": "",
        "bronze_rule_id": text(rule.get("rule_id")),
        "silver_rule_id": "",
        "Rule_Title_B": text(rule.get("name")),
        "Rule_Title_S": "",
        "Asset_Name_B": text(rule.get("dataset_name")),
        "Asset_Name_S": "",
    }


def run(directory: Path, sprints=SPRINTS) -> tuple[list[Comparison], list[str]]:
    """Compare every sprint and return the comparisons and any problems."""
    problems: list[str] = []
    pairings, mapping_problems = load_mapping(directory / MAPPING_FILE, sprints)
    problems += mapping_problems

    comparisons: list[Comparison] = []
    for sprint in sprints:
        of_sprint = [p for p in pairings if p["sprint"] == sprint]
        try:
            bronze_rows = read_sheet(directory / BRONZE_FILE.format(sprint=sprint), RULE_SHEET)
            silver_rows = read_sheet(directory / SILVER_FILE.format(sprint=sprint), RULE_SHEET)
        except Exception as exc:
            problems.append(f"sprint {sprint} skipped: {exc}")
            continue

        bronze_rules = index(bronze_rows, f"sprint {sprint} bronze", problems)
        silver_rules = index(silver_rows, f"sprint {sprint} silver", problems)

        for column in COMPARED:
            for layer, rows in (("bronze", bronze_rows), ("silver", silver_rows)):
                if rows and column not in rows[0]:
                    problems.append(
                        f"sprint {sprint} {layer} export has no {column!r} column; "
                        "it is reported as not-compared"
                    )

        for pairing in of_sprint:
            comparisons.append(compare(pairing, bronze_rules, silver_rules))

        # Bronze rules the mapper never mentions: real rules with no recorded
        # silver counterpart, which is a gap in the mapping, not in a layer.
        mapped = {p["bronze_rule_id"].lower() for p in of_sprint if p["bronze_rule_id"]}
        for rule_id, rule in bronze_rules.items():
            if rule_id in mapped:
                continue
            comparisons.append(
                Comparison(
                    unmapped_pairing(sprint, rule),
                    BRONZE_UNMAPPED,
                    bronze=rule,
                    note="this bronze rule appears in no mapper row",
                )
            )

    return comparisons, problems


# --------------------------------------------------------------------------
# reporting
# --------------------------------------------------------------------------


def columns() -> list[str]:
    """The report's header row."""
    header = [
        "Sprint",
        "DQO",
        "DQO Mapped",
        "Status",
        "Differences",
        "Bronze Rule ID",
        "Silver Rule ID",
        "Bronze Rule Title",
        "Silver Rule Title",
        "Bronze Dataset",
        "Silver Dataset",
        "Note",
    ]
    for column in COMPARED:
        header += [column, f"{column} (bronze)", f"{column} (silver)"]
    return header


def cell(rule: dict | None, column: str) -> str:
    """A rule's value for a column, as the report shows it."""
    if rule is None:
        return ""
    value = rule.get(column)
    if value is None:
        return "(blank)"
    return str(value)


def report_row(comparison: Comparison) -> list:
    """Flatten one comparison into a report row.

    A column that agrees carries only its verdict; the two values are shown
    only where they differ, so the sheet can be read across without having
    to ignore hundreds of identical pairs.
    """
    pairing = comparison.pairing
    dqo = pairing["DQO"]
    row = [
        pairing["sprint"],
        dqo,
        "yes" if DQO_PATTERN.match(dqo) else "no",
        comparison.status,
        comparison.differences if comparison.status == MATCHED else "",
        pairing["bronze_rule_id"],
        pairing["silver_rule_id"],
        pairing["Rule_Title_B"] or text((comparison.bronze or {}).get("name")),
        pairing["Rule_Title_S"] or text((comparison.silver or {}).get("name")),
        text((comparison.bronze or {}).get("dataset_name")) or pairing["Asset_Name_B"],
        text((comparison.silver or {}).get("dataset_name")) or pairing["Asset_Name_S"],
        comparison.note,
    ]
    for column in COMPARED:
        verdict = comparison.verdicts.get(column, NOT_COMPARED)
        row.append(verdict)
        if verdict == DIFFERENT:
            row.append(cell(comparison.bronze, column))
            row.append(cell(comparison.silver, column))
        else:
            row += ["", ""]
    return row


def summarise(comparisons: list[Comparison], problems: list[str]) -> list[list]:
    """The rows of the summary sheet."""
    rows: list[list] = [["Section", "Item", "Count"]]
    for sprint in sorted({c.pairing["sprint"] for c in comparisons}):
        of_sprint = [c for c in comparisons if c.pairing["sprint"] == sprint]
        matched = [c for c in of_sprint if c.status == MATCHED]
        section = f"Sprint {sprint}"
        rows.append([section, "rule pairs compared", len(matched)])
        rows.append(
            [section, "pairs agreeing on every column",
             sum(1 for c in matched if c.differences == 0)]
        )
        rows.append(
            [section, "pairs with at least one difference",
             sum(1 for c in matched if c.differences > 0)]
        )
        for status in (BRONZE_MISSING, BRONZE_NOT_FOUND, SILVER_NOT_FOUND, BRONZE_UNMAPPED):
            rows.append([section, status, sum(1 for c in of_sprint if c.status == status)])

    rows.append(["Overall", "report rows", len(comparisons)])
    rows.append(
        ["Overall", "DQO placeholders (no golden rule book entry)",
         sum(1 for c in comparisons
             if c.pairing["DQO"] and not DQO_PATTERN.match(c.pairing["DQO"]))]
    )

    matched = [c for c in comparisons if c.status == MATCHED]
    for column in COMPARED:
        rows.append(
            ["Differences by column", column,
             sum(1 for c in matched if c.verdicts.get(column) == DIFFERENT)]
        )
    for problem in problems:
        rows.append(["Problem", problem, ""])
    return rows


def style_header(sheet, header) -> None:
    for position in range(1, len(header) + 1):
        styled = sheet.cell(row=1, column=position)
        styled.font = Font(bold=True)
        styled.fill = HEADER
        styled.alignment = Alignment(vertical="top", wrap_text=True)


def output_path(directory: Path, override: Path | None) -> Path:
    if override is not None:
        return override
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    destination = directory / "output"
    destination.mkdir(parents=True, exist_ok=True)
    return destination / f"{stamp}_dq_layers_comparison.xlsx"


def write_workbook(comparisons: list[Comparison], problems: list[str], destination: Path) -> None:
    """Write the comparison and summary sheets."""
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.title = "Comparison"

    header = columns()
    sheet.append(header)
    verdict_at = {header.index(column) + 1 for column in COMPARED}
    fill_for = {SAME: GREEN, DIFFERENT: RED, NOT_COMPARED: GREY}

    for comparison in comparisons:
        sheet.append(report_row(comparison))
        row = sheet.max_row
        fill = STATUS_FILL[comparison.status]
        if fill is not None:
            sheet.cell(row=row, column=header.index("Status") + 1).fill = fill
        if comparison.status != MATCHED:
            continue
        for column in COMPARED:
            verdict = comparison.verdicts.get(column, NOT_COMPARED)
            sheet.cell(row=row, column=header.index(column) + 1).fill = fill_for[verdict]

    style_header(sheet, header)
    sheet.freeze_panes = "F2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(header))}{sheet.max_row}"
    for position in range(1, len(header) + 1):
        letter = get_column_letter(position)
        sheet.column_dimensions[letter].width = 14 if position in verdict_at else 22

    summary = book.create_sheet("Summary")
    summary_rows = summarise(comparisons, problems)
    for row in summary_rows:
        summary.append(row)
    style_header(summary, summary_rows[0])
    summary.column_dimensions["A"].width = 24
    summary.column_dimensions["B"].width = 70
    summary.column_dimensions["C"].width = 10

    book.save(destination)


def write_csv(comparisons: list[Comparison], destination: Path) -> None:
    with destination.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(columns())
        for comparison in comparisons:
            writer.writerow(report_row(comparison))


def write_console(comparisons: list[Comparison], problems: list[str]) -> None:
    for sprint in sorted({c.pairing["sprint"] for c in comparisons}):
        of_sprint = [c for c in comparisons if c.pairing["sprint"] == sprint]
        matched = [c for c in of_sprint if c.status == MATCHED]
        differing = [c for c in matched if c.differences > 0]
        print(f"\nSprint {sprint}")
        print(f"  {len(matched)} rule pair(s) compared, {len(differing)} with a difference")
        counts: Counter = Counter()
        for comparison in matched:
            for column, verdict in comparison.verdicts.items():
                if verdict == DIFFERENT:
                    counts[column] += 1
        for column, count in counts.most_common():
            print(f"    {column:26s} {count:5d} differ")
        for status in (BRONZE_MISSING, BRONZE_NOT_FOUND, SILVER_NOT_FOUND, BRONZE_UNMAPPED):
            count = sum(1 for c in of_sprint if c.status == status)
            if count:
                print(f"    {status:26s} {count:5d}")
    if problems:
        print("\nProblems")
        for problem in problems:
            print(f"  - {problem}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Compare the bronze and silver DQ rule definitions sprint by sprint."
    )
    parser.add_argument(
        "directory",
        type=Path,
        help="folder holding the mapping, bronze and silver workbooks",
    )
    parser.add_argument(
        "--sprint",
        type=int,
        action="append",
        choices=SPRINTS,
        help="compare only this sprint; repeatable (default: all)",
    )
    parser.add_argument("--output", type=Path, help="write the report here instead")
    parser.add_argument("--csv", type=Path, help="also write the rows to this CSV")
    args = parser.parse_args(argv)

    if not args.directory.is_dir():
        parser.error(f"not a directory: {args.directory}")

    sprints = tuple(sorted(set(args.sprint))) if args.sprint else SPRINTS
    comparisons, problems = run(args.directory, sprints)

    print(f"DQ layer comparison: {args.directory}")
    write_console(comparisons, problems)

    destination = output_path(args.directory, args.output)
    write_workbook(comparisons, problems, destination)
    print(f"\nReport written to {destination}")

    if args.csv:
        write_csv(comparisons, args.csv)
        print(f"All {len(comparisons)} row(s) written to {args.csv}")

    differing = sum(1 for c in comparisons if c.status == MATCHED and c.differences > 0)
    unresolved = sum(
        1 for c in comparisons if c.status in (BRONZE_NOT_FOUND, SILVER_NOT_FOUND)
    )
    return 1 if differing or unresolved else 0


if __name__ == "__main__":
    sys.exit(main())
