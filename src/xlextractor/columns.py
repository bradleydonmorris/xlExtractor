"""Named 0-based column indexes for ``FieldDef.sources`` and
``RowParser.signature``, so callers can write ``CellIndexes.B`` instead of
the magic number ``1``. Members are plain ints (``IntEnum``), so they work
directly as dict keys and tuple elements anywhere a column index is expected.

Single-letter Excel columns (A-Z) are covered by ``CellIndexes``. Two-letter
columns are split one class per leading letter — ``CellIndexesA`` for AA-AZ,
``CellIndexesB`` for BA-BZ, ``CellIndexesC`` for CA-CZ — so a two-letter
column maps to ``CellIndexes<LeadingLetter>.<TrailingLetter>``, e.g. Excel
column "AA" is ``CellIndexesA.A``. Numbering is continuous across classes:
``CellIndexes.Z`` is 25, ``CellIndexesA.A`` is 26, and so on through
``CellIndexesC.Z`` (Excel column CZ) at 103.
"""

from enum import IntEnum


class CellIndexes(IntEnum):
	"""Represents cell indexes A (0) through Z (25)"""
	A: int = 0
	B: int = 1
	C: int = 2
	D: int = 3
	E: int = 4
	F: int = 5
	G: int = 6
	H: int = 7
	I: int = 8
	J: int = 9
	K: int = 10
	L: int = 11
	M: int = 12
	N: int = 13
	O: int = 14
	P: int = 15
	Q: int = 16
	R: int = 17
	S: int = 18
	T: int = 19
	U: int = 20
	V: int = 21
	W: int = 22
	X: int = 23
	Y: int = 24
	Z: int = 25

class CellIndexesA(IntEnum):
	"""Represents cell indexes AA (26) through AZ (51)"""
	A: int = 26
	B: int = 27
	C: int = 28
	D: int = 29
	E: int = 30
	F: int = 31
	G: int = 32
	H: int = 33
	I: int = 34
	J: int = 35
	K: int = 36
	L: int = 37
	M: int = 38
	N: int = 39
	O: int = 40
	P: int = 41
	Q: int = 42
	R: int = 43
	S: int = 44
	T: int = 45
	U: int = 46
	V: int = 47
	W: int = 48
	X: int = 49
	Y: int = 50
	Z: int = 51

class CellIndexesB(IntEnum):
	"""Represents cell indexes BA (52) through BZ (77)"""
	A: int = 52
	B: int = 53
	C: int = 54
	D: int = 55
	E: int = 56
	F: int = 57
	G: int = 58
	H: int = 59
	I: int = 60
	J: int = 61
	K: int = 62
	L: int = 63
	M: int = 64
	N: int = 65
	O: int = 66
	P: int = 67
	Q: int = 68
	R: int = 69
	S: int = 70
	T: int = 71
	U: int = 72
	V: int = 73
	W: int = 74
	X: int = 75
	Y: int = 76
	Z: int = 77


class CellIndexesC(IntEnum):
	"""Represents cell indexes CA (78) through CZ (103)"""
	A: int = 78
	B: int = 79
	C: int = 80
	D: int = 81
	E: int = 82
	F: int = 83
	G: int = 84
	H: int = 85
	I: int = 86
	J: int = 87
	K: int = 88
	L: int = 89
	M: int = 90
	N: int = 91
	O: int = 92
	P: int = 93
	Q: int = 94
	R: int = 95
	S: int = 96
	T: int = 97
	U: int = 98
	V: int = 99
	W: int = 100
	X: int = 101
	Y: int = 102
	Z: int = 103
