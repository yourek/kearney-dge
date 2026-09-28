"""Loading of the codified catalogue rules in `specs/`."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

SPECS_DIR = Path(__file__).parent / "specs"


@dataclass
class Column:
    """One column's rules, as written in the spec YAML."""

    n: int
    name: str
    requirement: str
    type: str
    rules: dict[str, Any] = field(default_factory=dict)

    @property
    def is_mandatory(self) -> bool:
        return self.requirement == "mandatory"

    @property
    def todo(self) -> str | None:
        return self.rules.get("todo")

    def get(self, key: str, default: Any = None) -> Any:
        return self.rules.get(key, default)


@dataclass
class Spec:
    """A codified sheet: where to find it, and the rules for its columns."""

    key: str
    sheet_name: str
    header_row: int
    first_data_row: int
    na_tokens: list[str]
    date_format: str
    columns: list[Column]
    guideline: dict[str, Any]
    references: dict[str, Any]

    def column(self, name: str) -> Column | None:
        return next((c for c in self.columns if c.name == name), None)

    def is_na(self, value: Any) -> bool:
        """True when a cell counts as not filled in."""
        if value is None:
            return True
        return str(value).strip() in self.na_tokens


_RESERVED = {"n", "name", "requirement", "type"}


def load(key: str) -> Spec:
    """Load `specs/<key>.yaml`."""
    path = SPECS_DIR / f"{key}.yaml"
    raw = yaml.safe_load(path.read_text(encoding="utf-8"))
    sheet = raw["sheet"]
    return Spec(
        key=key,
        sheet_name=sheet["name"],
        header_row=sheet["header_row"],
        first_data_row=sheet["first_data_row"],
        na_tokens=raw["na_tokens"],
        date_format=raw["date_format"],
        guideline=raw["guideline"],
        references=raw.get("references", {}),
        columns=[
            Column(
                n=c["n"],
                name=c["name"],
                requirement=c["requirement"],
                type=c["type"],
                rules={k: v for k, v in c.items() if k not in _RESERVED},
            )
            for c in raw["columns"]
        ],
    )
