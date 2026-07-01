# xlextractor — Project Context for Claude Code

## Purpose

`xlextractor` is a Python library for declaratively parsing structured data
out of Excel worksheets that have no consistent tabular layout. The primary
use case is accounting report exports, where a single worksheet contains
interleaved header rows, vendor rows, multi-row data tables, and totals
rows — each with a different shape.

---

## Design Philosophy

The core idea is that adding a new row type should require **only**:
1. A `RowParser` entry with a `signature` dict and a `parse` callable
2. A `_parse_*` function built with `make_field_parser` or `make_table_parser`

Nothing in the main loop, accumulator, or dispatch logic should ever need to
change. All behaviour is driven by data in `RowParser`.

---

## File Layout

```
xlextractor/
├── src/
│   └── xlextractor/
│       ├── __init__.py       # Public API re-exports
│       ├── _types.py         # Dataclasses, enums, TypedDicts
│       ├── columns.py        # CellIndexes* IntEnums (A..CZ column indexes)
│       ├── transforms.py     # Pure value-conversion functions
│       ├── parsers.py        # Parser-building machinery
│       └── worksheet.py      # Workbook I/O
├── tests/
│   ├── test_transforms.py
│   ├── test_parsers.py
│   └── test_worksheet.py
├── pyproject.toml
└── README.md
```

---

## Key Design Decisions

**`RowParser.signature` values and `loop_breakers` must be pre-casefolded.**
The comparison functions casefold the cell value at runtime but compare against
the stored strings as-is. This avoids redundant work on hot paths.
Convention: write them lowercase in source (they usually already are).

**`ParserType.FIXED_ROWS`** covers both single-row and multi-row fixed parsers.
`row_count=1` and `row_count=2` are both valid. The enum value carries no
additional behaviour beyond signalling that `row_count` is authoritative.

**`ParserType.LOOP_UNTIL_BROKEN`** parsers do not `add` their header row.
The signature row is the column-header row of the table, not a data row.
The accumulator starts collecting on the row *after* the signature match.

**`make_table_parser` ignores `row_offset`** in `FieldDef.sources`.
Each row in the table is self-contained. Use `0` by convention so the
`sources` format stays consistent with `make_field_parser`.

**`Accumulator` is internal** (prefixed `_`). Consumers only interact with
`RowParser`, `FieldDef`, `parse_worksheet`, and the result list.

**`columns.py` defines four `IntEnum`s covering Excel columns A through CZ**
(0-based, `CellIndexes.A = 0`). Single-letter columns are `CellIndexes`
(A-Z, 0-25); two-letter columns are one class per leading letter —
`CellIndexesA` (AA-AZ, 26-51), `CellIndexesB` (BA-BZ, 52-77), `CellIndexesC`
(CA-CZ, 78-103) — so Excel column "AA" is `CellIndexesA.A`, "BC" is
`CellIndexesB.C`, etc. Numbering is continuous across the four classes.
Members are plain ints, so they work as dict keys/tuple elements with no
other code changes, e.g. `FieldDef("VendorID", [(0, CellIndexes.B)])`.
