# xlextractor

A small Python library for declaratively parsing structured data out of
Excel worksheets that have no consistent tabular layout — interleaved
header rows, key/value rows, multi-row data tables, and totals rows, all
in the same sheet.

The motivating use case is Great Plains (GP) accounting report exports,
where a single worksheet mixes several different row shapes with no
fixed schema.

## Installation

Not yet published to PyPI. Add it as a local or git dependency with `uv`:

```bash
uv add --editable ../path/to/xlextractor
# or
uv add git+https://github.com/bradleydonmorris/xlExtractor
```

Requires Python 3.14+. The only runtime dependency is [`openpyxl`](https://foss.heptapod.net/openpyxl/openpyxl).

## How it works

You describe every row shape your worksheet can contain as a `RowParser`,
then hand the list of `RowParser`s to `parse_worksheet`. Nothing about the
parsing loop changes as you add row types — only the `RowParser` list
grows.

Each `RowParser` has:

- **`signature`** — a `{column_index: expected_text}` dict used to detect
  the start of this row type. Values must be lowercase (they're compared
  against the casefolded cell text).
- **`parser_type`** — `ParserType.FIXED_ROWS` for a row (or a fixed number
  of rows) whose shape you know in advance, or
  `ParserType.LOOP_UNTIL_BROKEN` for a variable-length table that reads
  until a breaker row (e.g. `"total"`) is seen in column A.
- **`parse`** — a callable built with `make_field_parser` (for
  `FIXED_ROWS`) or `make_table_parser` (for `LOOP_UNTIL_BROKEN`) from a
  list of `FieldDef`s.

Each `FieldDef` maps one output key to one or more source cells, with an
optional `transform` to convert the raw cell value(s) into the type you
want. Source columns are 0-based; use the `CellIndexes` enum from
`xlextractor` instead of raw indexes so `sources` and `signature` read
like the worksheet's own column letters. For sheets wider than column Z,
`CellIndexesA` (AA-AZ), `CellIndexesB` (BA-BZ), and `CellIndexesC` (CA-CZ)
extend the same scheme — Excel column "AA" is `CellIndexesA.A`, "BC" is
`CellIndexesB.C`, and so on.

## Example

```python
from pathlib import Path

from xlextractor import (
	CellIndexes as ci,
	FieldDef, ParserType, RowParser,
	as_float, as_string,
	make_field_parser, make_table_parser,
	get_worksheet_name, parse_worksheet,
)

row_parsers: list[RowParser] = [
	RowParser(
		row_type="VENDOR",
		parser_type=ParserType.FIXED_ROWS,
		signature={ci.A: "vendor id:", ci.C: "name:"},
		row_count=1,
		loop_breakers=None,
		parse=make_field_parser([
			FieldDef("VendorID", [(0, ci.B)], as_string),
			FieldDef("Name",     [(0, ci.D)], as_string),
		]),
	),
	RowParser(
		row_type="INVOICE",
		parser_type=ParserType.LOOP_UNTIL_BROKEN,
		signature={ci.A: "invoice number"},
		row_count=None,
		loop_breakers=["total"],
		parse=make_table_parser([
			FieldDef("InvoiceNumber", [(0, ci.A)], as_string),
			FieldDef("Amount",       [(0, ci.B)], as_float),
		]),
	),
]

file_path = Path("that-not-so-awesome-spreadsheet.xlsx")
sheet_name = get_worksheet_name(file_path, "ungly data")
results = parse_worksheet(file_path, sheet_name, row_parsers)
```

`results` is a list of `ParseResult` dicts, each with `row_type`,
`row_indexes` (the worksheet rows the block came from), and `data` (a
dict for `FIXED_ROWS`, a list of dicts for `LOOP_UNTIL_BROKEN`).

## Debugging your `RowParser`s

While you're writing and testing `FieldDef`/`RowParser` definitions for a
new worksheet, `report_data_types` prints the Python type(s) actually
observed for every `(row_type, field)` pair across a set of results:

```python
from xlextractor import parse_worksheet, report_data_types

results = parse_worksheet(file_path, sheet_name, row_parsers)
report_data_types(results)
```

```
Row Type                  Field Name                Types
------------------------- ------------------------- -------------------------
VENDOR                    VendorID                  str
INVOICE                   Amount                    float | NoneType
```

Useful for catching a `transform` that isn't handling every cell shape in
the sheet (e.g. a column you expected to always be numeric but sometimes
comes through blank). It's a development/debugging aid, not something to
run in production pipelines.

## Development

```bash
uv sync
uv run pytest
```
