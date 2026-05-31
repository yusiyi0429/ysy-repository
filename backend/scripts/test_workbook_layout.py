#!/usr/bin/env python3
"""Unit tests for revision row alignment."""

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from workbook_layout import normalize_revision_notes  # noqa: E402


def test_data_relative_row():
    layouts = {"Sheet1": {"header_rows": 2, "data_start_row": 3, "max_row": 20, "max_column": 10}}
    notes = [{"sheet": "Sheet1", "row": 1, "col": 5, "action": "modify", "new_value": "x"}]
    out, stats = normalize_revision_notes(notes, layouts)
    assert out[0]["row"] == 3, out
    assert stats["adjusted"] >= 1


def test_skip_header_target():
    layouts = {"Sheet1": {"header_rows": 2, "data_start_row": 3, "max_row": 20, "max_column": 10}}
    notes = [{"sheet": "Sheet1", "row": 1, "col": 1, "action": "modify", "new_value": "bad"}]
    out, stats = normalize_revision_notes(notes, layouts)
    # row 1 < data_start 3 → converted to 3, not skipped
    assert len(out) == 1


if __name__ == "__main__":
    test_data_relative_row()
    test_skip_header_target()
    print("workbook_layout tests passed.")
