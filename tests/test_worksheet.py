from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from xlextractor._types import FieldDef, ParserType, RowParser
from xlextractor.parsers import make_field_parser, make_table_parser
from xlextractor.transforms import as_float, as_string
from xlextractor.worksheet import get_worksheet_name, parse_worksheet, report_data_types

VENDOR_PARSER = RowParser(
	row_type="VENDOR",
	parser_type=ParserType.FIXED_ROWS,
	signature={0: "vendor id:"},
	row_count=1,
	loop_breakers=None,
	parse=make_field_parser([FieldDef("VendorID", [(0, 1)], as_string)]),
)

ADDRESS_PARSER = RowParser(
	row_type="ADDRESS",
	parser_type=ParserType.FIXED_ROWS,
	signature={0: "address:"},
	row_count=2,
	loop_breakers=None,
	parse=make_field_parser(
		[
			FieldDef("Line1", [(0, 1)], as_string),
			FieldDef("Line2", [(1, 1)], as_string),
		]
	),
)

LINE_ITEM_PARSER = RowParser(
	row_type="LINE_ITEM",
	parser_type=ParserType.LOOP_UNTIL_BROKEN,
	signature={0: "invoice number"},
	row_count=None,
	loop_breakers=["total"],
	parse=make_table_parser([FieldDef("Amount", [(0, 0)], as_float)]),
)


def build_xlsx(tmp_path: Path, sheet_name: str, rows: list[list[object]]) -> Path:
	workbook = openpyxl.Workbook()
	worksheet = workbook.active
	worksheet.title = sheet_name
	for row in rows:
		worksheet.append(row)
	file_path = tmp_path / "test.xlsx"
	workbook.save(file_path)
	return file_path


class TestGetWorksheetName:
	def test_returns_first_sheet_when_name_is_none(self, tmp_path: Path) -> None:
		file_path = build_xlsx(tmp_path, "Original GP Download", [["a"]])
		assert get_worksheet_name(file_path, None) == "Original GP Download"

	def test_matches_case_insensitively(self, tmp_path: Path) -> None:
		file_path = build_xlsx(tmp_path, "Original GP Download", [["a"]])
		assert get_worksheet_name(file_path, "original gp download") == "Original GP Download"

	def test_returns_none_when_sheet_not_found(self, tmp_path: Path) -> None:
		file_path = build_xlsx(tmp_path, "Original GP Download", [["a"]])
		assert get_worksheet_name(file_path, "does not exist") is None


class TestParseWorksheet:
	def test_raises_when_sheet_not_found(self, tmp_path: Path) -> None:
		file_path = build_xlsx(tmp_path, "Sheet1", [["a"]])
		with pytest.raises(ValueError, match="does not exist"):
			parse_worksheet(file_path, "does not exist", [VENDOR_PARSER])

	def test_returns_empty_list_when_no_rows_match(self, tmp_path: Path) -> None:
		file_path = build_xlsx(tmp_path, "Sheet1", [["unrelated"], ["also unrelated"]])
		assert parse_worksheet(file_path, "Sheet1", [VENDOR_PARSER]) == []

	def test_fixed_rows_single_row_parser(self, tmp_path: Path) -> None:
		file_path = build_xlsx(
			tmp_path,
			"Sheet1",
			[
				["Some Header"],
				["Vendor ID:", "12345"],
				["trailing"],
			],
		)
		results = parse_worksheet(file_path, "Sheet1", [VENDOR_PARSER])
		assert results == [
			{
				"row_type": "VENDOR",
				"row_indexes": [1],
				"data": {"VendorID": "12345"},
			}
		]

	def test_fixed_rows_multi_row_parser_uses_row_offset(self, tmp_path: Path) -> None:
		file_path = build_xlsx(
			tmp_path,
			"Sheet1",
			[
				["Address:", "123 Main St"],
				[None, "Suite 400"],
			],
		)
		results = parse_worksheet(file_path, "Sheet1", [ADDRESS_PARSER])
		assert results == [
			{
				"row_type": "ADDRESS",
				"row_indexes": [0, 1],
				"data": {"Line1": "123 Main St", "Line2": "Suite 400"},
			}
		]

	def test_loop_until_broken_collects_rows_between_header_and_breaker(self, tmp_path: Path) -> None:
		file_path = build_xlsx(
			tmp_path,
			"Sheet1",
			[
				["Some Header"],
				["Vendor ID:", "12345"],
				["Invoice Number"],
				[100.0],
				[200.5],
				["Total"],
				["trailing", "row"],
			],
		)
		results = parse_worksheet(file_path, "Sheet1", [VENDOR_PARSER, LINE_ITEM_PARSER])
		assert len(results) == 2
		line_item_result = results[1]
		assert line_item_result["row_type"] == "LINE_ITEM"
		assert line_item_result["data"] == [{"Amount": 100.0}, {"Amount": 200.5}]
		# Header row (index 2) is not itself a data row; the two collected
		# rows are at indexes 3 and 4.
		assert line_item_result["row_indexes"] == [3, 4]

	def test_loop_until_broken_without_breaker_runs_to_end_of_sheet(self, tmp_path: Path) -> None:
		file_path = build_xlsx(
			tmp_path,
			"Sheet1",
			[
				["Invoice Number"],
				[100.0],
				[200.5],
			],
		)
		results = parse_worksheet(file_path, "Sheet1", [LINE_ITEM_PARSER])
		assert results == []

	def test_skips_empty_rows_within_table(self, tmp_path: Path) -> None:
		file_path = build_xlsx(
			tmp_path,
			"Sheet1",
			[
				["Invoice Number"],
				[100.0],
				[None],
				[200.5],
				["Total"],
			],
		)
		results = parse_worksheet(file_path, "Sheet1", [LINE_ITEM_PARSER])
		assert results[0]["data"] == [{"Amount": 100.0}, {"Amount": 200.5}]


class TestReportDataTypes:
	def test_prints_row_type_field_and_types(self, capsys: pytest.CaptureFixture[str]) -> None:
		results = [
			{"row_type": "VENDOR", "row_indexes": [0], "data": {"VendorID": "12345"}},
			{
				"row_type": "LINE_ITEM",
				"row_indexes": [1, 2],
				"data": [{"Amount": 100.0}, {"Amount": None}],
			},
		]
		report_data_types(results)
		out = capsys.readouterr().out
		assert "VENDOR" in out
		assert "VendorID" in out
		assert "str" in out
		assert "LINE_ITEM" in out
		assert "Amount" in out
		assert "float" in out
		assert "NoneType" in out

	def test_handles_empty_results(self, capsys: pytest.CaptureFixture[str]) -> None:
		report_data_types([])
		out = capsys.readouterr().out
		assert "Row Type" in out
