"""Findings produced by the QC checks, and how they are reported."""

from __future__ import annotations

import csv
import sys
from collections import Counter
from dataclasses import dataclass
from enum import IntEnum
from pathlib import Path
from typing import Any, Iterable


class Severity(IntEnum):
    INFO = 0
    WARNING = 1
    ERROR = 2

    def __str__(self) -> str:
        return self.name.lower()


# Findings are grouped into the buckets the fixing work actually divides
# into, so a reviewer can filter the report down to one kind of problem.
# The order here is the order the columns appear in the report.
CATEGORY_ORDER = (
    "Completeness",
    "Valid values",
    "Consistency",
    "Reference data",
    "Database",
    "Glossary",
    "Formatting",
    "Structure",
)

CATEGORIES = {
    "mandatory-empty": "Completeness",
    "invalid-enum": "Valid values",
    "invalid-date": "Valid values",
    "invalid-length": "Valid values",
    "invalid-integer": "Valid values",
    "invalid-number": "Valid values",
    "invalid-email_list": "Valid values",
    "invalid-text": "Valid values",
    "date-order": "Consistency",
    "duplicate-value": "Consistency",
    "inconsistent-value": "Consistency",
    "classification-mismatch": "Consistency",
    "unresolved-reference": "Reference data",
    "unknown-reference-value": "Reference data",
    "mismatched-pair": "Reference data",
    "database-mismatch": "Database",
    "unknown-database-table": "Database",
    "not-in-database": "Database",
    "unknown-glossary-term": "Glossary",
    "glossary-coverage-gap": "Glossary",
    "glossary-mismatch": "Glossary",
    "glossary-term-not-approved": "Glossary",
    "unused-glossary-term": "Glossary",
    "glossary-separator-in-term": "Glossary",
    "percent-formatted-value": "Formatting",
    "suspect-scale": "Formatting",
    "missing-column": "Structure",
    "extra-column": "Structure",
    "missing-sheet": "Structure",
    "guideline-drift": "Structure",
    "not-checked": "Structure",
    "missing-reference": "Structure",
}


def category_of(check: str) -> str:
    """The bucket a check belongs to; unmapped checks fall under Structure."""
    return CATEGORIES.get(check, "Structure")


@dataclass
class Finding:
    """One problem, located as precisely as the check can manage."""

    severity: Severity
    check: str
    message: str
    sheet: str = ""
    row: int | None = None
    column: str = ""
    value: Any = None

    @property
    def category(self) -> str:
        return category_of(self.check)

    @property
    def location(self) -> str:
        parts = [p for p in (self.sheet, self.column) if p]
        if self.row is not None:
            parts.append(f"row {self.row}")
        return " / ".join(parts)


def error(check: str, message: str, **kw: Any) -> Finding:
    return Finding(Severity.ERROR, check, message, **kw)


def warning(check: str, message: str, **kw: Any) -> Finding:
    return Finding(Severity.WARNING, check, message, **kw)


def info(check: str, message: str, **kw: Any) -> Finding:
    return Finding(Severity.INFO, check, message, **kw)


def _truncate(value: Any, limit: int = 60) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ")
    return text if len(text) <= limit else text[: limit - 1] + "…"


def write_console(
    findings: list[Finding],
    stream: Any = None,
    min_severity: Severity = Severity.INFO,
    max_per_check: int = 20,
) -> None:
    """Print findings grouped by check, worst first."""
    out = stream or sys.stdout
    shown = [f for f in findings if f.severity >= min_severity]

    if not shown:
        print("No findings.", file=out)
        return

    groups: dict[tuple[Severity, str], list[Finding]] = {}
    for f in shown:
        groups.setdefault((f.severity, f.check), []).append(f)

    for (severity, check), group in sorted(
        groups.items(), key=lambda kv: (-kv[0][0], kv[0][1])
    ):
        print(f"\n[{str(severity).upper()}] {check}  ({len(group)})", file=out)
        for f in group[:max_per_check]:
            location = f.location
            value = _truncate(f.value)
            line = f"  {location}: {f.message}" if location else f"  {f.message}"
            if value:
                line += f"  (value: {value!r})"
            print(line, file=out)
        if len(group) > max_per_check:
            print(f"  ... and {len(group) - max_per_check} more", file=out)

    counts = Counter(f.severity for f in findings)
    summary = ", ".join(
        f"{counts[s]} {s}" for s in (Severity.ERROR, Severity.WARNING, Severity.INFO)
    )
    print(f"\nTotal: {summary}", file=out)


def write_csv(findings: Iterable[Finding], path: Path) -> None:
    """Write every finding to CSV, for triage in Excel."""
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["severity", "check", "sheet", "row", "column", "value", "message"])
        for f in findings:
            writer.writerow(
                [str(f.severity), f.check, f.sheet, f.row or "", f.column,
                 "" if f.value is None else str(f.value), f.message]
            )
