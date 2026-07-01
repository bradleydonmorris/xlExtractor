from __future__ import annotations

from collections.abc import Callable

import openpyxl
import pytest
from openpyxl.cell.cell import Cell
from openpyxl.worksheet.worksheet import Worksheet

from xlextractor._types import FieldDef, ParserType, RowParser
from xlextractor.parsers import (
	cell_str,
	icase_compare,
	icase_compare_list,
	identify_parser,
	is_empty_row,
	make_field_parser,
	make_table_parser,
)

MakeRow = Callable[[list[object]], tuple[Cell, ...]]


@pytest.fixture
def make_row() -> MakeRow:
	worksheet: Worksheet = openpyxl.Workbook().active
	next_row_number = [1]

	def _make_row(values: list[object]) -> tuple[Cell, ...]:
		row_number = next_row_number[0]
		next_row_number[0] += 1
		for col_offset, value in enumerate(values, start=1):
			worksheet.cell(row=row_number, column=col_offset, value=value)
		return tuple(worksheet[row_number])

	return _make_row


def dummy_row_parser(row_type: str, signature: dict[int, str]) -> RowParser:
	return RowParser(
		row_type=row_type,
		parser_type=ParserType.FIXED_ROWS,
		signature=signature,
		row_count=1,
		loop_breakers=None,
		parse=None,
	)


class TestCellStr:
	def test_strips_and_casefolds_cell_value(self, make_row: MakeRow) -> None:
		row = make_row(["  Hello World  "])
		assert cell_str(row[0]) == "hello world"

	def test_none_cell_value_becomes_literal_none_string(self, make_row: MakeRow) -> None:
		row = make_row([None])
		assert cell_str(row[0]) == "none"

	def test_non_cell_raw_string_passthrough(self) -> None:
		assert cell_str("Hello") == "hello"

	def test_non_cell_non_string_is_stringified(self) -> None:
		assert cell_str(42) == "42"


class TestIcaseCompare:
	def test_matches_case_insensitively(self, make_row: MakeRow) -> None:
		row = make_row(["Vendor ID:"])
		assert icase_compare(row[0], "vendor id:") is True

	def test_does_not_match_different_text(self, make_row: MakeRow) -> None:
		row = make_row(["Something Else"])
		assert icase_compare(row[0], "vendor id:") is False


class TestIcaseCompareList:
	def test_matches_one_of_the_list(self, make_row: MakeRow) -> None:
		row = make_row(["Total"])
		assert icase_compare_list(row[0], ["grand total", "total"]) is True

	def test_matches_none_of_the_list(self, make_row: MakeRow) -> None:
		row = make_row(["Foo"])
		assert icase_compare_list(row[0], ["grand total", "total"]) is False


class TestIdentifyParser:
	def test_returns_matching_parser(self, make_row: MakeRow) -> None:
		vendor_parser = dummy_row_parser("VENDOR", {0: "vendor id:", 2: "name:"})
		total_parser = dummy_row_parser("TOTAL", {0: "total"})
		row = make_row(["Vendor ID:", "12345", "Name:"])
		assert identify_parser([vendor_parser, total_parser], row) is vendor_parser

	def test_returns_none_when_no_signature_matches(self, make_row: MakeRow) -> None:
		vendor_parser = dummy_row_parser("VENDOR", {0: "vendor id:", 2: "name:"})
		row = make_row(["Unrelated", "12345", "Row"])
		assert identify_parser([vendor_parser], row) is None

	def test_returns_first_match_in_list_order(self, make_row: MakeRow) -> None:
		first = dummy_row_parser("FIRST", {0: "total"})
		second = dummy_row_parser("SECOND", {0: "total"})
		row = make_row(["Total"])
		assert identify_parser([first, second], row) is first


class TestIsEmptyRow:
	def test_true_when_all_referenced_columns_are_none(self, make_row: MakeRow) -> None:
		row = make_row([None, None, "ignored"])
		fields = [FieldDef("A", [(0, 0)]), FieldDef("B", [(0, 1)])]
		assert is_empty_row(row, fields) is True

	def test_false_when_any_referenced_column_has_a_value(self, make_row: MakeRow) -> None:
		row = make_row([None, "value"])
		fields = [FieldDef("A", [(0, 0)]), FieldDef("B", [(0, 1)])]
		assert is_empty_row(row, fields) is False

	def test_ignores_columns_not_referenced_by_fields(self, make_row: MakeRow) -> None:
		row = make_row([None, "unreferenced value"])
		fields = [FieldDef("A", [(0, 0)])]
		assert is_empty_row(row, fields) is True


class TestMakeFieldParser:
	def test_single_cell_passthrough_without_transform(self, make_row: MakeRow) -> None:
		row = make_row(["12345"])
		parse = make_field_parser([FieldDef("VendorID", [(0, 0)])])
		assert parse([row]) == {"VendorID": "12345"}

	def test_single_cell_with_transform(self, make_row: MakeRow) -> None:
		row = make_row(["12345"])
		parse = make_field_parser([FieldDef("VendorID", [(0, 0)], int)])
		assert parse([row]) == {"VendorID": 12345}

	def test_multi_cell_transform_receives_unpacked_values(self, make_row: MakeRow) -> None:
		row = make_row(["Jane", "Doe"])
		fields = [FieldDef("FullName", [(0, 0), (0, 1)], lambda first, last: f"{first} {last}")]
		parse = make_field_parser(fields)
		assert parse([row]) == {"FullName": "Jane Doe"}

	def test_multi_row_uses_row_offset(self, make_row: MakeRow) -> None:
		row0 = make_row(["header value"])
		row1 = make_row(["detail value"])
		fields = [
			FieldDef("Header", [(0, 0)]),
			FieldDef("Detail", [(1, 0)]),
		]
		parse = make_field_parser(fields)
		assert parse([row0, row1]) == {"Header": "header value", "Detail": "detail value"}


class TestMakeTableParser:
	def test_parses_each_row_with_transform(self, make_row: MakeRow) -> None:
		rows = [make_row(["10"]), make_row(["20"])]
		fields = [FieldDef("Amount", [(0, 0)], int)]
		parse = make_table_parser(fields)
		assert parse(rows) == [{"Amount": 10}, {"Amount": 20}]

	def test_passthrough_without_transform(self, make_row: MakeRow) -> None:
		rows = [make_row(["foo"]), make_row(["bar"])]
		fields = [FieldDef("Value", [(0, 0)])]
		parse = make_table_parser(fields)
		assert parse(rows) == [{"Value": "foo"}, {"Value": "bar"}]

	def test_skips_empty_rows_by_default(self, make_row: MakeRow) -> None:
		rows = [make_row(["foo"]), make_row([None]), make_row(["bar"])]
		fields = [FieldDef("Value", [(0, 0)])]
		parse = make_table_parser(fields)
		assert parse(rows) == [{"Value": "foo"}, {"Value": "bar"}]

	def test_keeps_empty_rows_when_disabled(self, make_row: MakeRow) -> None:
		rows = [make_row(["foo"]), make_row([None])]
		fields = [FieldDef("Value", [(0, 0)])]
		parse = make_table_parser(fields, skip_empty_rows=False)
		assert parse(rows) == [{"Value": "foo"}, {"Value": None}]

	def test_row_offset_is_ignored(self, make_row: MakeRow) -> None:
		rows = [make_row(["foo"])]
		fields = [FieldDef("Value", [(5, 0)])]
		parse = make_table_parser(fields)
		assert parse(rows) == [{"Value": "foo"}]
