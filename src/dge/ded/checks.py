"""The quality checks themselves.

Each check is a function over a loaded sheet (plus whatever context it needs)
returning a list of findings, so the set of checks can grow without reshaping
the script. Nothing here modifies the deliverable.
"""

from __future__ import annotations

import datetime as dt
import difflib
import re
from pathlib import Path
from typing import Any, Iterable

from dge.ded import workbook
from dge.ded.findings import Finding, Severity, error, info, warning
from dge.ded.spec import Column, Spec

EMAIL_RE = re.compile(r"^[^@\s,;]+@[^@\s,;]+\.[A-Za-z]{2,}$")
# An Excel date is acceptable when it *displays* in month-day-year order.
# mm/dd/yyyy, m/d/yy and mm-dd-yy all pass; yyyy-mm-dd and dd/mm/yyyy do not.
# The separator is not enforced - only the ordering the guideline is about.
MDY_FORMAT_RE = re.compile(r"m{1,2}([/.-])d{1,2}\1y{2,4}")
# Locale and colour prefixes: [$-409], [Red], and literal quoted text.
FORMAT_NOISE_RE = re.compile(r"\[[^\]]*\]|\"[^\"]*\"|\\.")
# varchar(60), decimal(10,2) - the base type is what we validate
PARAMETERISED_RE = re.compile(r"^([A-Za-z_][A-Za-z0-9_ ]*)\s*\((\s*\d+\s*(,\s*\d+\s*)?)\)$")


def separators_of(column: Column) -> list[str]:
    """The separators a list column may use.

    Deliverables built from the same template disagree - one writes stewards
    comma-separated, another semicolon-separated - so a column accepts any of
    the characters its spec lists.
    """
    separator = column.get("separator", ",")
    return [separator] if isinstance(separator, str) else list(separator)


def _parts(value: Any, separators: list[str]) -> tuple[list[str], list[str]]:
    """Split on any separator, keeping the ones used so text can be rebuilt."""
    pattern = "(" + "|".join(re.escape(s) for s in separators) + ")"
    tokens = re.split(pattern, str(value))
    return [t.strip() for t in tokens[0::2]], tokens[1::2]


def _split(value: Any, separators: list[str] | str) -> list[str]:
    if isinstance(separators, str):
        separators = [separators]
    values, _ = _parts(value, separators)
    return [v for v in values if v]


def _as_float(value: Any) -> float | None:
    """The cell as a number, or None when it does not hold one."""
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def _normalise(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def parse_terms(
    value: Any, known: dict[str, str], separators: list[str] | str = ","
) -> list[str]:
    """Split a list cell into terms, preferring the longest known term.

    Some glossary terms contain the separator themselves ("Address, Parcel or
    Owner Reference Number"), so splitting naively would shatter them into
    fragments that match nothing. Adjacent parts are rejoined whenever the
    result is a term the glossary actually defines.

    `known` maps a normalised term to its canonical spelling; pass an empty
    mapping to fall back to a plain split.
    """
    if isinstance(separators, str):
        separators = [separators]
    parts, used = _parts(value, separators)
    terms: list[str] = []
    index = 0
    while index < len(parts):
        if not parts[index]:
            index += 1
            continue
        for end in range(len(parts), index, -1):  # longest candidate first
            # Rebuild with the separators the cell actually used, so a term
            # containing a comma only matches where a comma was written.
            candidate = parts[index]
            for step in range(index + 1, end):
                candidate += f"{used[step - 1]} {parts[step]}"
            candidate = _normalise(candidate)
            if candidate in known:
                terms.append(known[candidate])
                index = end
                break
        else:
            terms.append(parts[index])
            index += 1
    return terms


def _is_exempt(row: workbook.Row, column: Column, sheet: workbook.Sheet) -> bool:
    """True when a column's rules are waived for this row.

    Some columns do not apply to every attribute - completeness figures are
    meaningless for something that is not a critical data element - and the
    deliverable records an agreed stand-in value instead of a number.
    """
    rule = column.get("exempt_if")
    if not rule or not sheet.has(rule["column"]):
        return False
    if row.text(rule["column"]).upper() != str(rule["equals"]).upper():
        return False
    # With no `value`, the column is waived however the cell is filled in.
    expected = rule.get("value")
    return expected is None or row.text(column.name) == expected


def _severity_from(rule: dict[str, Any], default: Severity = Severity.ERROR) -> Severity:
    name = str(rule.get("severity", "")).upper()
    return Severity[name] if name in Severity.__members__ else default


# --------------------------------------------------------------------------
# structure
# --------------------------------------------------------------------------

def check_guideline_drift(spec: Spec, guidelines_dir: Path) -> list[Finding]:
    """Verify the spec still matches the guideline workbook it was derived from.

    Without this, a reissued guideline would be silently checked against stale
    rules.
    """
    config = spec.guideline
    path = guidelines_dir / config["file"]
    if not path.exists():
        return [
            warning(
                "guideline-drift",
                f"guideline workbook not found: {path.name}; "
                "the spec could not be verified against it",
            )
        ]

    try:
        book = workbook.load_sheet(
            path, "Sheet1", config["header_row"], config["header_row"] + 1
        )
    except Exception as exc:  # unreadable, wrong sheet, no header row
        return [
            warning(
                "guideline-drift",
                f"could not read {path.name}: {exc}; "
                "the spec could not be verified against it",
            )
        ]
    published = [
        (row.text(config["property_column"]), row.text(config["requirement_column"]).lower())
        for row in book
        if row.text(config["property_column"])
    ]
    coded = {c.name: c for c in spec.columns}
    findings: list[Finding] = []

    for name, requirement in published:
        column = coded.get(name)
        if column is None:
            findings.append(
                error("guideline-drift", f"{name!r} is in the guidelines but not in the spec")
            )
        elif column.requirement != requirement and column.requirement != "conditional":
            findings.append(
                error(
                    "guideline-drift",
                    f"{name!r} is {requirement} in the guidelines "
                    f"but {column.requirement} in the spec",
                )
            )

    for name in coded.keys() - {n for n, _ in published}:
        findings.append(
            error("guideline-drift", f"{name!r} is in the spec but not in the guidelines")
        )
    return findings


def check_columns_present(sheet: workbook.Sheet, spec: Spec) -> list[Finding]:
    """Every guideline column must exist in the deliverable."""
    findings: list[Finding] = []
    for column in spec.columns:
        if not sheet.has(column.name):
            findings.append(
                error(
                    "missing-column",
                    f"guideline column {column.name!r} is not in the sheet",
                    sheet=sheet.name,
                    column=column.name,
                )
            )
    governed = {c.name for c in spec.columns}
    extra = [h for h in sheet.headers if h and h not in governed]
    if extra:
        findings.append(
            info(
                "extra-column",
                f"{len(extra)} column(s) not in the guidelines, ignored: "
                + ", ".join(repr(h) for h in extra),
                sheet=sheet.name,
            )
        )
    return findings


def report_unchecked(spec: Spec) -> list[Finding]:
    """Surface rules that are codified but cannot run yet."""
    return [
        info("not-checked", f"{c.name}: {' '.join(c.todo.split())}", column=c.name)
        for c in spec.columns
        if c.todo
    ]


# --------------------------------------------------------------------------
# completeness
# --------------------------------------------------------------------------

def check_mandatory(sheet: workbook.Sheet, spec: Spec) -> list[Finding]:
    """Mandatory columns must be filled; `N/A` and friends do not count."""
    findings: list[Finding] = []
    for column in spec.columns:
        if not sheet.has(column.name):
            continue
        conditional = column.get("mandatory_if")
        for row in sheet:
            value = row.get(column.name)
            if not spec.is_na(value) or _is_exempt(row, column, sheet):
                continue
            if column.is_mandatory:
                reason = "mandatory column is empty"
            elif conditional and row.text(conditional["column"]) == conditional["equals"]:
                reason = (
                    f"mandatory when {conditional['column']} "
                    f"is {conditional['equals']!r}, but empty"
                )
            else:
                continue
            findings.append(
                error(
                    "mandatory-empty",
                    reason,
                    sheet=sheet.name,
                    row=row.number,
                    column=column.name,
                    value=value,
                )
            )
    return findings


# --------------------------------------------------------------------------
# value validity
# --------------------------------------------------------------------------

def is_percent_format(number_format: str) -> bool:
    """True when a cell is formatted as a percentage, so 0.95 reads as 95%."""
    return "%" in FORMAT_NOISE_RE.sub("", number_format or "")


def displays_as_mdy(number_format: str) -> bool:
    """True when an Excel number format renders a date as MM/DD/YYYY."""
    cleaned = FORMAT_NOISE_RE.sub("", number_format or "").lower()
    # A format can hold separate positive/negative sections.
    return any(MDY_FORMAT_RE.search(part) for part in cleaned.split(";"))


def _check_value(value: Any, column: Column, spec: Spec, number_format: str = "") -> str | None:
    """Validate one cell against its column's type. Returns a message or None."""
    kind = column.type

    if kind == "date":
        if isinstance(value, (dt.datetime, dt.date)):
            # A real Excel date is fine as long as it displays as MM/DD/YYYY.
            if displays_as_mdy(number_format):
                return None
            return (
                f"an Excel date displayed as {number_format!r}; the guideline "
                "requires MM/DD/YYYY"
            )
        try:
            dt.datetime.strptime(str(value).strip(), spec.date_format)
        except ValueError:
            return "not a valid MM/DD/YYYY date"
        return None

    if kind == "enum":
        allowed = column.get("allowed", [])
        text = str(value).strip()
        if column.get("allow_parameterised") and (m := PARAMETERISED_RE.match(text)):
            text = m.group(1).strip()
        if not column.get("case_sensitive", True):
            if text.lower() in {a.lower() for a in allowed}:
                return None
        elif text in allowed:
            return None
        return f"not a permitted value; expected one of: {', '.join(allowed)}"

    if kind == "email_list":
        parts = _split(value, separators_of(column))
        if not parts:
            return "no email address found"
        bad = [p for p in parts if not EMAIL_RE.match(p)]
        return f"not a valid email address: {', '.join(bad)}" if bad else None

    if kind == "length":
        text = str(value).strip()
        if PARAMETERISED_RE.match(text):
            return None  # decimal(12,0); the figures are checked against the dump
        try:
            if float(text) != int(float(text)):
                return "not a whole number"
        except ValueError:
            return "not a length or a decimal(precision, scale)"
        low = column.get("min")
        if low is not None and float(text) < low:
            return f"below the minimum of {low}"
        return None

    if kind in ("integer", "number"):
        try:
            number = float(str(value).strip())
        except ValueError:
            return f"not a {kind}"
        if kind == "integer" and number != int(number):
            return "not a whole number"
        if column.get("percent_format_aware") and is_percent_format(number_format):
            # Excel stores 95% as 0.95; compare what the reader actually sees.
            number *= 100
        low, high = column.get("min"), column.get("max")
        if low is not None and number < low:
            return f"below the minimum of {low}"
        if high is not None and number > high:
            return f"above the maximum of {high}"
        return None

    return None


def check_values(sheet: workbook.Sheet, spec: Spec) -> list[Finding]:
    """Type, format and permitted-value checks over every filled cell."""
    findings: list[Finding] = []
    for column in spec.columns:
        if not sheet.has(column.name):
            continue
        for row in sheet:
            value = row.get(column.name)
            if spec.is_na(value):
                continue  # emptiness is check_mandatory's business
            # On an exempt row a stand-in is fine, but a figure that really
            # was supplied is still held to the column's rules.
            if _is_exempt(row, column, sheet) and _as_float(value) is None:
                continue
            if message := _check_value(value, column, spec, row.number_format(column.name)):
                findings.append(
                    error(
                        f"invalid-{column.type}",
                        message,
                        sheet=sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
            suspect = column.get("suspect_at_or_below")
            fmt = row.number_format(column.name)
            percent = column.get("percent_format_aware") and is_percent_format(fmt)

            # Text in a percentage-formatted cell is reported as an invalid
            # number above; there is nothing to scale here.
            if percent and column.get("percent_format_warning") and _as_float(value) is not None:
                # Displays correctly in Excel, but the stored value is on a
                # 0-1 scale, so anything reading the file outside Excel - a
                # CSV export, a database load - sees 0.95 rather than 95.
                findings.append(
                    warning(
                        "percent-formatted-value",
                        f"stored as {value} with percentage formatting, so it "
                        f"displays as {_as_float(value) * 100:g}%; the guideline "
                        "expects the value itself on a 0-100 scale",
                        sheet=sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )

            # A percentage-formatted cell is unambiguous: 0.95 displays as 95%.
            if (
                suspect is not None
                and column.type in ("integer", "number")
                and not percent
            ):
                try:
                    if 0 < float(str(value)) <= suspect:
                        findings.append(
                            warning(
                                "suspect-scale",
                                f"at or below {suspect} on a 0-100 percentage scale "
                                "and not formatted as a percentage; looks like a "
                                "0-1 fraction",
                                sheet=sheet.name,
                                row=row.number,
                                column=column.name,
                                value=value,
                            )
                        )
                except ValueError:
                    pass
    return findings


def check_date_order(sheet: workbook.Sheet, spec: Spec) -> list[Finding]:
    """`not_before` columns must not predate the column they reference."""
    findings: list[Finding] = []
    for column in spec.columns:
        other = column.get("not_before")
        if not other or not sheet.has(column.name) or not sheet.has(other):
            continue
        for row in sheet:
            later, earlier = _as_date(row.get(column.name), spec), _as_date(row.get(other), spec)
            if later and earlier and later < earlier:
                findings.append(
                    error(
                        "date-order",
                        f"is earlier than {other} ({earlier:%m/%d/%Y})",
                        sheet=sheet.name,
                        row=row.number,
                        column=column.name,
                        value=row.get(column.name),
                    )
                )
    return findings


def _as_date(value: Any, spec: Spec) -> dt.date | None:
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    try:
        return dt.datetime.strptime(str(value).strip(), spec.date_format).date()
    except (ValueError, TypeError):
        return None


# --------------------------------------------------------------------------
# uniqueness and references
# --------------------------------------------------------------------------

def check_uniqueness(sheet: workbook.Sheet, spec: Spec) -> list[Finding]:
    """`unique` and `unique_within` columns must not repeat."""
    findings: list[Finding] = []
    for column in spec.columns:
        if not sheet.has(column.name):
            continue
        scope = column.get("unique_within")
        if not column.get("unique") and not scope:
            continue
        seen: dict[tuple[str, str], int] = {}
        for row in sheet:
            value = row.text(column.name)
            if not value:
                continue
            key = (row.text(scope) if scope else "", value)
            if (first := seen.get(key)) is not None:
                where = f" within {scope} {key[0]!r}" if scope else ""
                findings.append(
                    error(
                        "duplicate-value",
                        f"duplicate{where}; first seen on row {first}",
                        sheet=sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
            else:
                seen[key] = row.number
    return findings


def check_foreign_keys(
    sheet: workbook.Sheet, spec: Spec, targets: dict[str, workbook.Sheet]
) -> list[Finding]:
    """Cross-sheet references must resolve to an existing row."""
    findings: list[Finding] = []
    for column in spec.columns:
        rule = column.get("foreign_key")
        if not rule or not sheet.has(column.name):
            continue
        target = targets.get(rule["spec"])
        if target is None or not target.has(rule["column"]):
            continue
        known = {row.text(rule["column"]) for row in target}
        for row in sheet:
            value = row.text(column.name)
            if value and value not in known:
                findings.append(
                    error(
                        "unresolved-reference",
                        f"no matching {rule['column']} in {target.name!r}",
                        sheet=sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
    return findings


def check_lookups(
    sheet: workbook.Sheet, spec: Spec, references: dict[str, list[dict[str, Any]]]
) -> list[Finding]:
    """Values must appear in their reference workbook, and pairs must agree."""
    findings: list[Finding] = []
    for column in spec.columns:
        if not sheet.has(column.name):
            continue

        if rule := column.get("lookup"):
            rows = references.get(rule["reference"])
            if rows is None:
                continue  # reference not supplied; reported by report_unchecked
            key = spec.references[rule["reference"]][rule["column"]]
            known = {str(r.get(key)).strip() for r in rows if r.get(key) is not None}
            for row in sheet:
                value = row.text(column.name)
                if value and value not in known:
                    findings.append(
                        error(
                            "unknown-reference-value",
                            f"not in the {rule['reference']} reference list",
                            sheet=sheet.name,
                            row=row.number,
                            column=column.name,
                            value=value,
                        )
                    )

        if rule := column.get("pairs_with"):
            rows = references.get(rule["reference"])
            if rows is None:
                continue
            config = spec.references[rule["reference"]]
            partner = spec.column(rule["column"])
            if partner is None or not sheet.has(partner.name):
                continue
            id_key = config[partner.get("lookup")["column"]]
            name_key = config[column.get("lookup")["column"]]
            expected = {
                str(r[id_key]).strip(): str(r[name_key]).strip()
                for r in rows
                if r.get(id_key) is not None
            }
            for row in sheet:
                partner_value, value = row.text(partner.name), row.text(column.name)
                want = expected.get(partner_value)
                if want and value and value != want:
                    findings.append(
                        error(
                            "mismatched-pair",
                            f"{partner.name} {partner_value!r} maps to {want!r}",
                            sheet=sheet.name,
                            row=row.number,
                            column=column.name,
                            value=value,
                        )
                    )
    return findings


def check_conditional_value(sheet: workbook.Sheet, spec: Spec) -> list[Finding]:
    """`must_equal_if` rules, e.g. a primary key cannot allow duplicates."""
    findings: list[Finding] = []
    for column in spec.columns:
        rule = column.get("must_equal_if")
        if not rule or not sheet.has(column.name) or not sheet.has(rule["column"]):
            continue
        for row in sheet:
            if row.text(rule["column"]) != rule["equals"]:
                continue
            value = row.text(column.name)
            if value and value != rule["value"]:
                findings.append(
                    error(
                        "inconsistent-value",
                        f"must be {rule['value']!r} when {rule['column']} "
                        f"is {rule['equals']!r}",
                        sheet=sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
    return findings


# --------------------------------------------------------------------------
# table / attribute consistency
# --------------------------------------------------------------------------

def glossary_index(rows: list[dict[str, Any]], config: dict[str, Any]) -> dict[str, str]:
    """Map every glossary term, normalised, to its canonical spelling."""
    key = config["term_column"]
    return {
        _normalise(str(row[key])): str(row[key]).strip()
        for row in rows
        if row.get(key) is not None
    }


def check_glossary(
    attribute_sheet: workbook.Sheet,
    attribute_spec: Spec,
    rows: list[dict[str, Any]],
) -> list[Finding]:
    """The glossary protocol: terms exist, are approved, and map to the table."""
    column = attribute_spec.column("Data Glossary")
    if column is None or not attribute_sheet.has(column.name):
        return []
    rule = column.get("glossary_terms")
    if not rule:
        return []

    config = attribute_spec.references[rule["reference"]]
    known = glossary_index(rows, config)
    separators = separators_of(column)
    findings: list[Finding] = []

    term_key, tables_key = config["term_column"], config["tables_column"]
    status_key, approved = config["status_column"], config["approved_value"]

    tables_for: dict[str, set[str]] = {}
    status_of: dict[str, str] = {}
    for row in rows:
        term = str(row.get(term_key, "")).strip()
        if not term:
            continue
        tables_for[term] = set(_split(row.get(tables_key), ","))
        status_of[term] = str(row.get(status_key, "")).strip()

    link = attribute_spec.column("Data Asset ID")
    used: set[str] = set()
    drifted: dict[tuple[str, str], int] = {}
    flagged_status: set[str] = set()

    for row in attribute_sheet:
        value = row.get(column.name)
        if attribute_spec.is_na(value):
            continue
        asset = row.text(link.name) if link else ""
        for term in parse_terms(value, known, separators):
            if _normalise(term) not in known:
                findings.append(
                    Finding(
                        _severity_from(rule["unknown_term"]),
                        "unknown-glossary-term",
                        f"{term!r} is not defined in the glossary",
                        sheet=attribute_sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
                continue
            used.add(term)
            if (status := status_of.get(term)) and status != approved and term not in flagged_status:
                flagged_status.add(term)
                findings.append(
                    Finding(
                        _severity_from(rule["approval"], Severity.WARNING),
                        "glossary-term-not-approved",
                        f"{term!r} has approval status {status!r}, expected {approved!r}",
                        sheet=attribute_sheet.name,
                        column=column.name,
                    )
                )
            if asset and asset not in tables_for.get(term, set()):
                drifted.setdefault((asset, term), row.number)

    for (asset, term), number in sorted(drifted.items(), key=lambda kv: kv[1]):
        findings.append(
            Finding(
                _severity_from(rule["reverse_coverage"], Severity.WARNING),
                "glossary-coverage-gap",
                f"the glossary entry for {term!r} does not list {asset}",
                sheet=attribute_sheet.name,
                row=number,
                column=column.name,
                value=term,
            )
        )

    if unused := sorted(set(known.values()) - used):
        shown = ", ".join(repr(t) for t in unused[:10])
        more = f" (and {len(unused) - 10} more)" if len(unused) > 10 else ""
        findings.append(
            Finding(
                _severity_from(rule["unused_terms"], Severity.INFO),
                "unused-glossary-term",
                f"{len(unused)} glossary term(s) are not used by any attribute: {shown}{more}",
            )
        )

    if ambiguous := [t for t in known.values() if any(s in t for s in separators)]:
        findings.append(
            info(
                "glossary-separator-in-term",
                f"{len(ambiguous)} glossary term(s) contain a list "
                f"separator, e.g. {ambiguous[0]!r}; matched longest-first so they "
                "are not split, but the template convention is ambiguous",
            )
        )
    return findings


# Schemas a warehouse creates for itself, never catalogue content.
SYSTEM_SCHEMAS = {"dbo", "sys", "information_schema", "queryinsights", "guest"}


def _names(wanted: str | list[str]) -> str:
    return ", ".join(repr(n) for n in ([wanted] if isinstance(wanted, str) else wanted))


def _first_present(sheet: workbook.Sheet, wanted: str | list[str]) -> str | None:
    """The first of the candidate column names the sheet actually has."""
    for name in [wanted] if isinstance(wanted, str) else wanted:
        if sheet.has(name):
            return name
    return None


def resolve_schema(
    rows: list[dict[str, Any]], preferred: str | None, override: str | None = None
) -> tuple[str | None, Finding | None]:
    """Decide which schema in the dump holds the catalogued tables.

    The spec names the schema it expects, but a second deliverable's dump
    uses its own - DED exports `silver_cleansed`, DCT exports `cln` - so a
    pinned name silently filters every row away. Returns the schema to use
    (None means "do not filter") and a finding explaining the choice.
    """
    present = {
        str(r["table_schema"]).strip()
        for r in rows
        if r.get("table_schema") and str(r["table_schema"]).strip()
    }
    if not present:
        return None, None

    if override:
        if override in present:
            return override, info("database-schema", f"using schema {override!r} as requested")
        return None, warning(
            "database-schema",
            f"requested schema {override!r} is not in the dump "
            f"(it holds {', '.join(sorted(present))}); comparing against every schema",
        )

    if preferred in present:
        return preferred, None

    candidates = sorted(s for s in present if s.lower() not in SYSTEM_SCHEMAS)
    if len(candidates) == 1:
        return candidates[0], info(
            "database-schema",
            f"the dump holds no {preferred!r} schema; using {candidates[0]!r}, "
            "the only non-system schema in it",
        )
    return None, warning(
        "database-schema",
        f"the dump holds no {preferred!r} schema and no single obvious "
        f"replacement ({', '.join(candidates) or 'none'}); comparing against "
        "every schema. Pass --schema to choose one",
    )


def check_database_schema(
    attribute_sheet: workbook.Sheet,
    attribute_spec: Spec,
    rows: list[dict[str, Any]],
    schema: str | None = None,
) -> list[Finding]:
    """Compare a catalogue sheet against the live database schema.

    The table sheet joins on the table alone; the attribute sheet joins on
    table and column, and compares type, length and nullability too. The
    database is ground truth. Cells still holding a placeholder are left to
    the ordinary checks rather than reported twice.
    """
    config = attribute_spec.database
    if not config:
        return []
    join = config["join"]
    by_column = "column" in join

    # Deliverables name the join columns differently - "Data Table Name (S)"
    # in one, "Data Table Name" in another - so the spec lists the candidates
    # and the first one present wins.
    table_column = _first_present(attribute_sheet, join["table"])
    name_column = _first_present(attribute_sheet, join["column"]) if by_column else None
    for wanted, found in ((join["table"], table_column), *([(join["column"], name_column)] if by_column else [])):
        if found is None:
            return [
                warning(
                    "database-join-column-missing",
                    f"{attribute_sheet.name!r} has none of the columns the "
                    f"database comparison joins on ({_names(wanted)}); the "
                    "comparison is skipped for this sheet",
                    sheet=attribute_sheet.name,
                )
            ]

    in_scope = [
        r
        for r in rows
        if r.get(join["to_table"])
        and (schema is None or str(r.get("table_schema", "")).strip() == schema)
    ]
    tables = {str(r[join["to_table"]]).strip().lower() for r in in_scope}
    index = {
        (str(r[join["to_table"]]).strip().lower(), str(r[join["to_column"]]).strip().lower()): r
        for r in in_scope
        if by_column and r.get(join["to_column"])
    }

    findings: list[Finding] = []
    for row in attribute_sheet:
        table = row.text(table_column)
        name = row.text(name_column) if by_column else ""
        if not table or (by_column and not name):
            continue
        match = index.get((table.lower(), name.lower())) if by_column else None
        if match is None:
            # Distinguish a misspelled table from a genuinely absent column,
            # otherwise every column of that table reads as missing.
            if table.lower() not in tables:
                suggestion = difflib.get_close_matches(table.lower(), tables, 1, 0.8)
                hint = f"; closest match is {suggestion[0]!r}" if suggestion else ""
                findings.append(
                    Finding(
                        _severity_from(config["unmatched"]),
                        "unknown-database-table",
                        f"no table {table!r} in the "
                        f"{schema or 'database'} schema{hint}",
                        sheet=attribute_sheet.name,
                        row=row.number,
                        column=table_column,
                        value=table,
                    )
                )
            elif not by_column:
                continue  # the table exists, which is all this sheet checks
            else:
                findings.append(
                    Finding(
                        _severity_from(config["unmatched"]),
                        "not-in-database",
                        f"no column {name!r} in database table {table!r}",
                        sheet=attribute_sheet.name,
                        row=row.number,
                        column=name_column,
                        value=name,
                    )
                )
            continue

        for rule in config["compare"]:
            column = attribute_spec.column(rule["column"])
            if column is None or not attribute_sheet.has(column.name):
                continue
            value = row.get(column.name)
            if attribute_spec.is_placeholder(value):
                continue
            if message := _compare_to_database(str(value).strip(), match, rule):
                findings.append(
                    Finding(
                        _severity_from(rule),
                        "database-mismatch",
                        message,
                        sheet=attribute_sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
    return findings


def _database_text(match: dict[str, Any], key: str) -> str:
    """A dump cell as text; SQL NULLs arrive as the string 'NULL'."""
    value = match.get(key)
    text = "" if value is None else str(value).strip()
    return "" if text.upper() == "NULL" else text


def _compare_to_database(value: str, match: dict[str, Any], rule: dict[str, Any]) -> str | None:
    """One catalogue cell against its database counterpart."""
    actual = _database_text(match, rule["to"])

    if mapping := rule.get("map"):
        actual = mapping.get(actual.upper(), actual)
        return None if value.upper() == actual.upper() else f"the database says {actual!r}"

    precision_key = rule.get("precision")
    if precision_key and not actual:
        # A numeric column: the catalogue writes decimal(precision, scale).
        precision = _database_text(match, precision_key)
        scale = _database_text(match, rule["scale"])
        if not precision:
            return None
        kind = _database_text(match, "data_type").lower()
        expected = f"{kind}({precision},{scale})" if scale else f"{kind}({precision})"
        if value.lower().replace(" ", "") != expected:
            return f"the database says {expected!r}"
        return None

    if not actual:
        return None
    return None if value.lower() == actual.lower() else f"the database says {actual!r}"


def check_derived_from_attributes(
    table_sheet: workbook.Sheet,
    table_spec: Spec,
    attribute_sheet: workbook.Sheet,
    attribute_spec: Spec,
) -> list[Finding]:
    """A table's classification should match the strictest of its attributes."""
    findings: list[Finding] = []
    link = attribute_spec.column("Data Asset ID")
    if link is None:
        return findings

    for column in table_spec.columns:
        rule = column.get("derive_from_attributes")
        if not rule or not table_sheet.has(column.name):
            continue
        source = rule["column"]
        if not attribute_sheet.has(source):
            continue
        order = rule["order"]
        rank = {name: i for i, name in enumerate(order)}
        severity = _severity_from(rule, Severity.WARNING)

        strictest: dict[str, str] = {}
        for row in attribute_sheet:
            asset, value = row.text(link.name), row.text(source)
            if asset and value in rank:
                current = strictest.get(asset)
                if current is None or rank[value] > rank[current]:
                    strictest[asset] = value

        key = table_spec.columns[0].name  # Data Asset ID (English)
        for row in table_sheet:
            asset, value = row.text(key), row.text(column.name)
            expected = strictest.get(asset)
            if expected and value and value != expected:
                findings.append(
                    Finding(
                        severity,
                        "classification-mismatch",
                        f"strictest attribute classification is {expected!r}",
                        sheet=table_sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
    return findings


def check_aggregated_from_attributes(
    table_sheet: workbook.Sheet,
    table_spec: Spec,
    attribute_sheet: workbook.Sheet,
    attribute_spec: Spec,
    known: dict[str, str] | None = None,
) -> list[Finding]:
    """A table's glossary must be exactly the union of its attributes' entries."""
    findings: list[Finding] = []
    link = attribute_spec.column("Data Asset ID")
    if link is None:
        return findings

    for column in table_spec.columns:
        rule = column.get("aggregates_from_attributes")
        if not rule or not table_sheet.has(column.name):
            continue
        source = rule["column"]
        if not attribute_sheet.has(source):
            continue
        separators = separators_of(column)
        severity = _severity_from(rule)

        terms = known or {}

        union: dict[str, set[str]] = {}
        for row in attribute_sheet:
            asset = row.text(link.name)
            value = row.get(source)
            if asset and not attribute_spec.is_na(value):
                union.setdefault(asset, set()).update(parse_terms(value, terms, separators))

        key = table_spec.columns[0].name
        for row in table_sheet:
            asset = row.text(key)
            if asset not in union:
                continue
            value = row.get(column.name)
            declared = (
                set(parse_terms(value, terms, separators))
                if not table_spec.is_na(value)
                else set()
            )
            expected = union[asset]
            if extra := declared - expected:
                findings.append(
                    Finding(
                        severity,
                        "glossary-mismatch",
                        "term(s) on the table but on none of its attributes: "
                        + ", ".join(sorted(extra)),
                        sheet=table_sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
            if missing := expected - declared:
                findings.append(
                    Finding(
                        severity,
                        "glossary-mismatch",
                        "term(s) on the attributes but missing from the table: "
                        + ", ".join(sorted(missing)),
                        sheet=table_sheet.name,
                        row=row.number,
                        column=column.name,
                        value=value,
                    )
                )
    return findings
