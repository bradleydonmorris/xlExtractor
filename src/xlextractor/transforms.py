from __future__ import annotations

from typing import Any
from datetime import datetime, time


def as_string(v: Any) -> str | None:
	if v is None:
		return None
	return v if isinstance(v, str) else str(v)


def as_int(v: Any) -> int | None:
	if v is None:
		return None
	if isinstance(v, int):
		return v
	if isinstance(v, str):
		try:
			return int(v)
		except (ValueError, TypeError):
			return None
	return None


def as_float(v: Any) -> float | None:
	if v is None:
		return None
	if isinstance(v, float):
		return v
	if isinstance(v, (int, str)):
		try:
			return float(v)
		except (ValueError, TypeError):
			return None
	return None


def as_datetime(v: Any) -> datetime | None:
	if v is None:
		return None
	if isinstance(v, datetime):
		return v
	if isinstance(v, str):
		try:
			return datetime.fromisoformat(v)
		except ValueError:
			return None
	return None


def combine_date_time(d: Any, t: Any) -> datetime | None:
	return datetime.combine(d.date(), t) if isinstance(d, datetime) and isinstance(t, time) else None
