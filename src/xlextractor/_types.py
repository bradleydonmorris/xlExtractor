from __future__ import annotations

from typing import Any, TypedDict
from enum import Enum
from collections.abc import Callable
from dataclasses import dataclass, field
from openpyxl.cell.cell import Cell


class ParserType(Enum):
	FIXED_ROWS = 0        # row_count determines when complete
	LOOP_UNTIL_BROKEN = 1


class ParseResult(TypedDict):
	row_type: str
	row_indexes: list[int]
	data: dict[str, Any] | list[dict[str, Any]]


@dataclass
class RowParser:
	row_type: str
	parser_type: ParserType
	signature: dict[int, str]
	row_count: int | None  # None = loop until breaker
	loop_breakers: list[str] | None
	parse: Callable[[list[tuple[Cell, ...]]], dict[str, Any]] | None


@dataclass
class Accumulator:
	parser: RowParser
	start_index: int
	rows: list[tuple[Cell, ...]] = field(default_factory=list)
	_force_complete: bool = field(default=False, init=False)

	def add(self, row: tuple[Cell, ...]) -> None:
		self.rows.append(row)

	@property
	def is_complete(self) -> bool:
		if self._force_complete:
			return True
		return (
			self.parser.row_count is not None
			and len(self.rows) >= self.parser.row_count
		)

	def force_complete(self) -> None:
		self._force_complete = True

	def finalize(self) -> ParseResult:
		# LOOP_UNTIL_BROKEN parsers don't add their header row, so the first
		# collected row is one past start_index. FIXED_ROWS parsers add their
		# first row at start_index itself.
		first_row_index = (
			self.start_index + 1
			if self.parser.parser_type == ParserType.LOOP_UNTIL_BROKEN
			else self.start_index
		)
		return ParseResult(
			row_type=self.parser.row_type,
			row_indexes=list(range(first_row_index, first_row_index + len(self.rows))),
			data=self.parser.parse(self.rows) if self.parser.parse else None,
		)


@dataclass
class FieldDef:
	"""Declarative field: output key, one or more cell sources, optional transform.

	sources: list of (row_offset, col_index) tuples.
	  - Single-cell:  [(0, cell_b)]
	  - Multi-cell:   [(0, cell_c), (0, cell_d)]  → transform receives both values
	  - Table rows:   row_offset is ignored by make_table_parser; use 0 by convention.
	transform: called with unpacked source values. Omit for single-cell pass-through.
	"""
	key: str
	sources: list[tuple[int, int]]
	transform: Callable[..., Any] | None = None
