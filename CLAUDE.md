# CLAUDE.md

Guidance for Claude Code when working in this repository.

## What this repo is

A collection of purpose-built scripts that run processes for a client project —
primarily quality checks (QC) on project deliverables. It is not a library or a
service; each script is a self-contained process someone runs against a
deliverable file and reads the output of.

## Layout

```
src/dge/
  <entity>/          one package per deliverable entity
    scripts/         the runnable scripts for that entity
  ded/               currently the only entity
    scripts/
      data_catalogue_qc.py
```

New entities (e.g. `dct`) get their own folder under `src/dge/`, each with its
own `scripts/` package — one module per deliverable, named `<deliverable>_qc.py`
for quality checks. Shared helpers for an entity live alongside `scripts/` in
the entity folder, not inside it.

## Toolchain

- Python >= 3.14, managed with **uv** (`uv.lock` is committed).
- Add dependencies with `uv add <pkg>` — do not hand-edit `[project].dependencies`.
- Run scripts with `uv run python -m dge.ded.scripts.data_catalogue_qc` (or `uv run <script>`
  once an entry point exists in `[project.scripts]`).
- Windows host; the repo is used from PowerShell.

## Conventions

- Each QC script should be runnable standalone: a `main()` plus
  `if __name__ == "__main__":`, arguments parsed with `argparse`.
- Input paths (the deliverable under test) are passed as arguments, never
  hardcoded. Do not commit client deliverables or their contents to the repo.
- QC checks are read-only: a script reports on a deliverable, it never modifies it.
- Output is a report the user reads — make findings explicit (which check, which
  row/column/sheet, what was expected). Prefer reporting every failure over
  stopping at the first.
- Exit non-zero when checks fail, so a script can be used in automation later.
- Keep checks separable — one function per check — so the set of checks can grow
  without reshaping the script.

## Data Catalogue QC

The rules are codified in `src/dge/ded/specs/*.yaml`, one file per catalogue
sheet, derived by hand from the B.1.2 / B.1.3 guideline workbooks. The
guidelines remain the human source of record; the YAML is the executable one,
and `check_guideline_drift` fails the run if the two diverge.

Guidelines are superior to the deliverable: only the columns they define are
checked, extra columns in the deliverable are reported as info and ignored.

```
spec.py       loads the YAML into Spec/Column
workbook.py   reads sheets and lookup workbooks into Row/Sheet
checks.py     one function per check, each returning findings
findings.py   Finding, severity, console and CSV output
annotate.py   writes the annotated copy of the deliverable
scripts/data_catalogue_qc.py   CLI that wires them together
```

The primary output is an annotated copy of the deliverable, written to
`<catalogue dir>/output/<YYYYMMDD_HHMMSS>_<input name>.xlsx`. Each checked
sheet gains `QC Result` / `QC Severity` / `QC Observations` columns, colour
coded, with every finding for that row listed in one cell; a `QC Summary`
sheet carries the counts and the findings that are not tied to a row. The
input file is only ever read.

The deliverable carries a lot of inherited baggage that does not survive an
openpyxl round trip - Excel offers to repair the result. `annotate.py` strips
it from the report: 4557 legacy defined names (437 already `#REF!`), 69
external workbook links, and the legacy comment drawing. It also replaces
each sheet's Excel Table with a plain autofilter, because a table's filter
cannot reach the appended QC columns. Pass `--keep-comments` to retain the
deliverable's cell comments at the risk of the repair prompt returning.

Adding a check: add the rule to the YAML, then a function in `checks.py`, then
call it from `run()`. A rule that cannot run yet (its reference file is
missing) carries a `todo:` and is reported as "not-checked" rather than
failing.

    uv run python -m dge.ded.scripts.data_catalogue_qc <catalogue.xlsx> --csv out.csv

Exit code is 1 when any error-severity finding is raised.

## Cell formats carry meaning

Two checks read the cell's Excel number format, not just its value, so
`workbook.load_sheet` keeps it:

- a date passes if it *displays* in month-day-year order, whatever its
  separator (the deliverable uses `mm-dd-yy`);
- a percentage-formatted cell holds 0.95 for 95%, so `percent_format_aware`
  columns are scaled up before their 0-100 range is checked. Only an
  unformatted value at or below 1 is ambiguous and warns.

## The glossary protocol

`B2.1_Glossary.xlsx` is a deliverable in its own right, used here as a lookup.
Two of its 1122 terms contain a comma - the same character the catalogue uses
to separate list entries - so `checks.parse_terms` matches terms longest-first
instead of splitting naively. Any check reading a `Data Glossary` cell must go
through it, or those terms shatter into fragments that match nothing.

The glossary also names the tables each term applies to, which drives the
reverse-coverage check (a warning: the two deliverables are on separate
update cycles).

## Status

`data_catalogue_qc.py` runs with every codified rule active. `Domains.xlsx` is
the canonical source for domain names, even where the catalogue and the
glossary agree with each other against it.
