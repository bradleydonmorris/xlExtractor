from __future__ import annotations

from datetime import datetime, time

import pytest

from xlextractor.transforms import (
	as_datetime,
	as_float,
	as_int,
	as_string,
	combine_date_time,
)


class TestAsString:
	def test_none_returns_none(self) -> None:
		assert as_string(None) is None

	def test_string_passthrough(self) -> None:
		assert as_string("hello") == "hello"

	@pytest.mark.parametrize(
		"value, expected",
		[
			(42, "42"),
			(3.14, "3.14"),
			(True, "True"),
		],
	)
	def test_non_string_is_stringified(self, value: object, expected: str) -> None:
		assert as_string(value) == expected


class TestAsInt:
	def test_none_returns_none(self) -> None:
		assert as_int(None) is None

	def test_int_passthrough(self) -> None:
		assert as_int(7) == 7

	def test_valid_numeric_string(self) -> None:
		assert as_int("42") == 42

	def test_invalid_string_returns_none(self) -> None:
		assert as_int("not a number") is None

	def test_float_string_is_rejected(self) -> None:
		assert as_int("3.14") is None

	def test_float_value_returns_none(self) -> None:
		assert as_int(3.14) is None

	def test_unsupported_type_returns_none(self) -> None:
		assert as_int([1, 2, 3]) is None


class TestAsFloat:
	def test_none_returns_none(self) -> None:
		assert as_float(None) is None

	def test_float_passthrough(self) -> None:
		assert as_float(3.14) == 3.14

	def test_int_is_converted(self) -> None:
		assert as_float(7) == 7.0

	def test_valid_numeric_string(self) -> None:
		assert as_float("3.14") == 3.14

	def test_invalid_string_returns_none(self) -> None:
		assert as_float("not a number") is None

	def test_unsupported_type_returns_none(self) -> None:
		assert as_float([1, 2, 3]) is None


class TestAsDatetime:
	def test_none_returns_none(self) -> None:
		assert as_datetime(None) is None

	def test_datetime_passthrough(self) -> None:
		dt = datetime(2026, 1, 1, 12, 30)
		assert as_datetime(dt) is dt

	def test_valid_iso_string(self) -> None:
		assert as_datetime("2026-01-01T12:30:00") == datetime(2026, 1, 1, 12, 30)

	def test_invalid_string_returns_none(self) -> None:
		assert as_datetime("not a date") is None

	def test_unsupported_type_returns_none(self) -> None:
		assert as_datetime(42) is None


class TestCombineDateTime:
	def test_valid_date_and_time(self) -> None:
		d = datetime(2026, 1, 1, 9, 0)
		t = time(14, 30)
		assert combine_date_time(d, t) == datetime(2026, 1, 1, 14, 30)

	def test_wrong_date_type_returns_none(self) -> None:
		assert combine_date_time("2026-01-01", time(14, 30)) is None

	def test_wrong_time_type_returns_none(self) -> None:
		assert combine_date_time(datetime(2026, 1, 1), "14:30") is None

	def test_both_none_returns_none(self) -> None:
		assert combine_date_time(None, None) is None
