"""Quality assessment of the Data Catalogue deliverable.

Checks the Data Table and Data Attribute sheets against the codified
guidelines in `dge/ded/specs/`. Read-only: the deliverable is never modified.

    uv run python -m dge.ded.scripts.data_catalogue_qc <catalogue.xlsx>
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from dge.ded import annotate, checks, spec as spec_module, workbook
from dge.ded.findings import Finding, Severity, warning, write_console, write_csv

SPEC_KEYS = ("catalogue_table", "catalogue_attribute")


def run(
    catalogue: Path,
    references_dir: Path,
    guidelines_dir: Path,
    specs: dict[str, spec_module.Spec] | None = None,
) -> list[Finding]:
    """Run every check and return the findings, worst first."""
    specs = specs if specs is not None else {k: spec_module.load(k) for k in SPEC_KEYS}
    findings: list[Finding] = []

    # Load the reference workbooks the lookups resolve against. A reference
    # that has not been supplied yet leaves its checks unrun, not failed.
    references: dict[str, list[dict]] = {}
    for current in specs.values():
        for name, config in current.references.items():
            if name in references:
                continue
            path = references_dir / config["file"]
            if path.exists():
                references[name] = workbook.load_reference(path, config)
            else:
                findings.append(
                    warning("missing-reference", f"reference workbook not found: {path.name}")
                )

    sheets: dict[str, workbook.Sheet] = {}
    for key, current in specs.items():
        try:
            sheets[key] = workbook.load_sheet(
                catalogue, current.sheet_name, current.header_row, current.first_data_row
            )
        except (KeyError, ValueError) as exc:
            findings.append(Finding(Severity.ERROR, "missing-sheet", str(exc)))

    for key, current in specs.items():
        findings += checks.check_guideline_drift(current, guidelines_dir)
        findings += checks.report_unchecked(current)
        sheet = sheets.get(key)
        if sheet is None:
            continue
        findings += checks.check_columns_present(sheet, current)
        findings += checks.check_mandatory(sheet, current)
        findings += checks.check_values(sheet, current)
        findings += checks.check_date_order(sheet, current)
        findings += checks.check_uniqueness(sheet, current)
        findings += checks.check_conditional_value(sheet, current)
        findings += checks.check_foreign_keys(sheet, current, sheets)
        findings += checks.check_lookups(sheet, current, references)

    table, attribute = sheets.get("catalogue_table"), sheets.get("catalogue_attribute")
    attribute_spec = specs["catalogue_attribute"]

    # Terms are matched longest-first, so both glossary checks need the index.
    known: dict[str, str] = {}
    if (glossary := references.get("glossary")) is not None:
        known = checks.glossary_index(glossary, attribute_spec.references["glossary"])
        if attribute:
            findings += checks.check_glossary(attribute, attribute_spec, glossary)

    if table and attribute:
        args = (table, specs["catalogue_table"], attribute, attribute_spec)
        findings += checks.check_derived_from_attributes(*args)
        findings += checks.check_aggregated_from_attributes(*args, known)

    findings.sort(key=lambda f: (-f.severity, f.check, f.sheet, f.row or 0))
    return findings


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("catalogue", type=Path, help="the Data Catalogue .xlsx to assess")
    parser.add_argument(
        "--references",
        type=Path,
        help="directory holding Domains.xlsx etc. (default: the catalogue's own directory)",
    )
    parser.add_argument(
        "--guidelines",
        type=Path,
        help="directory holding the B.1.2 / B.1.3 guideline workbooks "
        "(default: the catalogue's own directory)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="directory for the annotated copy "
        "(default: an 'output' directory beside the catalogue)",
    )
    parser.add_argument(
        "--keep-comments",
        action="store_true",
        help="keep the deliverable's cell comments in the report; they are "
        "dropped by default because the legacy drawing makes Excel repair the file",
    )
    parser.add_argument(
        "--no-annotate",
        action="store_true",
        help="report to the console only, do not write an annotated copy",
    )
    parser.add_argument("--csv", type=Path, help="also write every finding to this CSV")
    parser.add_argument(
        "--min-severity",
        choices=[str(s) for s in Severity],
        default="info",
        help="hide findings below this severity in the console output",
    )
    parser.add_argument(
        "--max-per-check", type=int, default=20, help="rows listed per check before truncating"
    )
    args = parser.parse_args(argv)

    if not args.catalogue.exists():
        parser.error(f"no such file: {args.catalogue}")

    directory = args.catalogue.parent
    specs = {key: spec_module.load(key) for key in SPEC_KEYS}
    findings = run(
        args.catalogue,
        args.references or directory,
        args.guidelines or directory,
        specs,
    )

    print(f"Data Catalogue QC: {args.catalogue.name}")
    write_console(
        findings,
        min_severity=Severity[args.min_severity.upper()],
        max_per_check=args.max_per_check,
    )

    if not args.no_annotate:
        destination = annotate.output_path(args.catalogue, args.output)
        annotate.write(args.catalogue, destination, findings, specs, args.keep_comments)
        print(f"\nAnnotated copy written to {destination}")

    if args.csv:
        write_csv(findings, args.csv)
        print(f"All {len(findings)} finding(s) written to {args.csv}")

    return 1 if any(f.severity is Severity.ERROR for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
