from __future__ import annotations

from typing import Any
from collections.abc import Callable
from openpyxl.cell.cell import Cell

from ._types import FieldDef, RowParser

def is_empty_row(row: tuple[Cell, ...], fields: list[FieldDef]) -> bool:
	return all(row[col].value is None for f in fields for _, col in f.sources)


def make_field_parser(
	fields: list[FieldDef],
) -> Callable[[list[tuple[Cell, ...]]], dict[str, Any]]:
	def _parse(rows: list[tuple[Cell, ...]]) -> dict[str, Any]:
		result: dict[str, Any] = {}
		for f in fields:
			values = [rows[row_offset][col].value for row_offset, col in f.sources]
			result[f.key] = f.transform(*values) if f.transform is not None else values[0]
		return result
	return _parse


def make_table_parser(
	fields: list[FieldDef],
	skip_empty_rows: bool = True,
) -> Callable[[list[tuple[Cell, ...]]], list[dict[str, Any]]]:
	# row_offset is intentionally ignored — each row is self-contained.
	cols = [(f.key, [col for _, col in f.sources], f.transform) for f in fields]

	def _parse(rows: list[tuple[Cell, ...]]) -> list[dict[str, Any]]:
		results: list[dict[str, Any]] = []
		for row in rows:
			if skip_empty_rows and is_empty_row(row, fields):
				continue
			result: dict[str, Any] = {
				key: transform(*[row[col].value for col in src_cols]) if transform is not None else row[src_cols[0]].value
				for key, src_cols, transform in cols
			}
			results.append(result)
		return results
	return _parse


def cell_str(value: Cell | object) -> str:
	raw = value.value if isinstance(value, Cell) else value
	return str(raw).strip().casefold()


def icase_compare(compare_value: Cell | object, compare_to: str) -> bool:
	return cell_str(compare_value) == compare_to


def icase_compare_list(compare_value: Cell | object, compare_to_list: list[str]) -> bool:
	return cell_str(compare_value) in compare_to_list


def identify_parser(row_parsers: list[RowParser], row: tuple[Cell, ...]) -> RowParser | None:
	return next(
		(
			parser
			for parser in row_parsers
			if all(icase_compare(row[col], val) for col, val in parser.signature.items())
		),
		None,
	)
