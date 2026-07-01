from __future__ import annotations

import string

import pytest

from xlextractor import CellIndexes, CellIndexesA, CellIndexesB, CellIndexesC


def test_members_are_zero_based_in_alphabetical_order() -> None:
	for index, letter in enumerate(string.ascii_uppercase):
		assert getattr(CellIndexes, letter) == index


def test_members_behave_as_plain_ints() -> None:
	assert CellIndexes.A + 1 == CellIndexes.B
	assert {CellIndexes.B: "value"}[1] == "value"


@pytest.mark.parametrize(
	"enum_cls, base",
	[
		(CellIndexesA, 26),
		(CellIndexesB, 52),
		(CellIndexesC, 78),
	],
)
def test_two_letter_classes_are_offset_and_alphabetical(
	enum_cls: type[CellIndexesA | CellIndexesB | CellIndexesC], base: int
) -> None:
	for index, letter in enumerate(string.ascii_uppercase):
		assert getattr(enum_cls, letter) == base + index


def test_numbering_is_continuous_across_classes() -> None:
	assert CellIndexesA.A == CellIndexes.Z + 1
	assert CellIndexesB.A == CellIndexesA.Z + 1
	assert CellIndexesC.A == CellIndexesB.Z + 1
	assert CellIndexesC.Z == 103
