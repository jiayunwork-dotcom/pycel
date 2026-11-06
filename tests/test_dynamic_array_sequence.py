# -*- coding: UTF-8 -*-
#
# Copyright 2011-2024 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""Tests for the SEQUENCE dynamic array function and its spill behavior."""

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import MAX_COL, MAX_ROW, SPILL_ERROR, VALUE_ERROR


SHEET = 'S'


def make_compiler(layout, extra_sheets=()):
    """Build a compiler from ``{address: formula_or_value}`` mappings."""
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = SHEET
    for address, value in layout.items():
        worksheet[address] = value

    for name, sheet_layout in extra_sheets:
        other = workbook.create_sheet(name)
        for address, value in sheet_layout.items():
            other[address] = value

    return ExcelCompiler(excel=workbook)


def s(address):
    return f'{SHEET}!{address}'


# ----------------------------------------------------------------------
# basic SEQUENCE values
# ----------------------------------------------------------------------

def test_sequence_one_argument_is_a_column_of_ones():
    compiler = make_compiler({'A1': '=SEQUENCE(3)'})
    assert compiler.evaluate(s('A1')) == 1
    assert compiler.evaluate(s('A2')) == 2
    assert compiler.evaluate(s('A3')) == 3

    # nothing spills to the side or past the end
    assert compiler.evaluate(s('B1')) is None
    assert compiler.evaluate(s('A4')) is None


def test_sequence_omitted_optional_arguments_default_to_one():
    compiler = make_compiler({
        'A1': '=SEQUENCE(3,,)',
        'C1': '=SEQUENCE(2,2,,3)',
    })
    assert [compiler.evaluate(s(a)) for a in ('A1', 'A2', 'A3')] == [1, 2, 3]
    assert compiler.evaluate(s('B1')) is None

    # omitted start defaults to 1, step is 3
    assert [compiler.evaluate(s(a)) for a in ('C1', 'D1', 'C2', 'D2')] == \
        [1, 4, 7, 10]


def test_sequence_start_step_and_negative_step():
    compiler = make_compiler({
        'A1': '=SEQUENCE(2,2,10,1)',
        'D1': '=SEQUENCE(1,3,5,10)',
        'A5': '=SEQUENCE(4,1,1,2)',
        'D5': '=SEQUENCE(2,2,4,-1)',
    })
    assert [compiler.evaluate(s(a)) for a in ('A1', 'B1', 'A2', 'B2')] == \
        [10, 11, 12, 13]
    assert [compiler.evaluate(s(a)) for a in ('D1', 'E1', 'F1')] == \
        [5, 15, 25]
    assert [compiler.evaluate(s(a)) for a in ('A5', 'A6', 'A7', 'A8')] == \
        [1, 3, 5, 7]
    assert [compiler.evaluate(s(a)) for a in ('D5', 'E5', 'D6', 'E6')] == \
        [4, 3, 2, 1]


def test_sequence_fractional_sizes_are_truncated():
    compiler = make_compiler({
        'J1': '=SEQUENCE(2.7)',
        'M1': '=SEQUENCE(0.5)',
    })
    assert compiler.evaluate(s('J1')) == 1
    assert compiler.evaluate(s('J2')) == 2
    assert compiler.evaluate(s('J3')) is None
    assert compiler.evaluate(s('M1')) == VALUE_ERROR


def test_sequence_4x5_anchor_example():
    # B2 holds SEQUENCE(4,5), it must occupy B2:F5, row major
    compiler = make_compiler({
        'B2': '=SEQUENCE(4,5)',
        # a formula summing the whole spill region, read before the anchor
        'H1': '=SUM(B2:F5)',
        'H2': '=C3',
    })

    # reading the sum first forces the region to be discovered
    assert compiler.evaluate(s('H1')) == sum(range(1, 21))  # 210
    assert compiler.evaluate(s('H2')) == 7

    anchor_value = 1
    for row in range(2, 6):
        for column in range(2, 7):
            assert compiler.evaluate(
                s(f'{get_column_letter(column)}{row}')) == anchor_value
            anchor_value += 1
    assert anchor_value == 21

    # the anchor itself holds the first element
    assert compiler.evaluate(s('B2')) == 1
    assert compiler.evaluate(s('C3')) == 7

    # outside of the spill region stays empty
    assert compiler.evaluate(s('G2')) is None
    assert compiler.evaluate(s('B6')) is None


# ----------------------------------------------------------------------
# #VALUE! validation
# ----------------------------------------------------------------------

def test_sequence_value_errors():
    compiler = make_compiler({
        'A1': '=SEQUENCE(0)',
        'A3': '=SEQUENCE(1,0)',
        'A5': '=SEQUENCE(-1)',
        'A7': '=SEQUENCE(2,"x")',
        'A9': '=SEQUENCE("x")',
        'A11': '=SEQUENCE(2,2,"nope")',
        'A13': '=SEQUENCE(2,2,1,"z")',
    })
    for anchor in ('A1', 'A3', 'A5', 'A7', 'A9', 'A11', 'A13'):
        assert compiler.evaluate(s(anchor)) == VALUE_ERROR

    # a failed anchor writes nothing next to itself
    assert compiler.evaluate(s('B1')) is None
    assert compiler.evaluate(s('A2')) is None
    assert compiler.evaluate(s('B7')) is None
    assert compiler.evaluate(s('A8')) is None


def test_sequence_propagates_error_arguments():
    compiler = make_compiler({
        'G1': '=1/0',
        'H1': '=SEQUENCE(G1)',
    })
    assert compiler.evaluate(s('H1')) == '#DIV/0!'
    assert compiler.evaluate(s('I1')) is None


# ----------------------------------------------------------------------
# direct reads, formulas and cross sheet references over spills
# ----------------------------------------------------------------------

def test_spilled_cells_readable_without_touching_anchor():
    compiler = make_compiler({'A1': '=SEQUENCE(2,2)'})

    # fresh compiler, read a spilled cell before the anchor was evaluated
    assert compiler.evaluate(s('B2')) == 4
    assert compiler.evaluate(s('B1')) == 2
    assert compiler.evaluate(s('A2')) == 3
    assert compiler.evaluate(s('A1')) == 1


def test_formula_over_spilled_cell_only_reads_d1():
    compiler = make_compiler({
        'A1': '=SEQUENCE(2,2)',
        'D1': '=B2+10',
    })
    assert compiler.evaluate(s('D1')) == 14


def test_spilled_cell_referenced_from_other_sheet():
    compiler = make_compiler(
        {'A1': '=SEQUENCE(2,2)'},
        extra_sheets=(('T', {'A1': f'={SHEET}!B2'}),),
    )
    assert compiler.evaluate('T!A1') == 4


def test_array_result_participates_in_arithmetic_and_spills():
    compiler = make_compiler({'K1': '=SEQUENCE(2,2)+10'})
    assert [compiler.evaluate(s(a)) for a in ('K1', 'L1', 'K2', 'L2')] == \
        [11, 12, 13, 14]


def test_array_aggregated_inside_sum_returns_scalar():
    compiler = make_compiler({
        'K4': '=SUM(SEQUENCE(2,2))',
        'K6': '=SUM(SEQUENCE(4,5))',
    })
    assert compiler.evaluate(s('K4')) == 10
    assert compiler.evaluate(s('K6')) == 210

    # an aggregated array spills nothing
    assert compiler.evaluate(s('L4')) is None
    assert compiler.evaluate(s('K5')) is None


# ----------------------------------------------------------------------
# #SPILL!
# ----------------------------------------------------------------------

def test_spill_blocked_by_constant_propagates_spill_error():
    compiler = make_compiler({
        'A1': '=SEQUENCE(2,2)',
        'B2': 99,
        'D1': '=A1+1',
        'D2': '=SUM(A1:B2)',
    })
    assert compiler.evaluate(s('A1')) == SPILL_ERROR

    # nothing other than the anchor is written
    assert compiler.evaluate(s('B1')) is None
    assert compiler.evaluate(s('A2')) is None
    assert compiler.evaluate(s('B2')) == 99

    assert compiler.evaluate(s('D1')) == SPILL_ERROR
    assert compiler.evaluate(s('D2')) == SPILL_ERROR


def test_spill_blocked_by_formula_cell():
    compiler = make_compiler({
        'A1': '=SEQUENCE(2,2)',
        'B2': '=1+1',
    })
    assert compiler.evaluate(s('A1')) == SPILL_ERROR
    assert compiler.evaluate(s('B1')) is None
    assert compiler.evaluate(s('A2')) is None

    # the blocking formula keeps working
    assert compiler.evaluate(s('B2')) == 2


def test_spill_off_the_sheet_edge():
    last_col = get_column_letter(MAX_COL)
    compiler = make_compiler({
        f'{last_col}1': '=SEQUENCE(1,2)',
        f'A{MAX_ROW}': '=SEQUENCE(2)',
    })
    assert compiler.evaluate(s(f'{last_col}1')) == SPILL_ERROR
    assert compiler.evaluate(s(f'A{MAX_ROW}')) == SPILL_ERROR


# ----------------------------------------------------------------------
# dynamic updates through set_value
# ----------------------------------------------------------------------

def _parametric_compiler():
    return make_compiler({
        'H1': 2,
        'A1': '=SEQUENCE(H1,H1)',
        'D5': '=B2+100',
        'D6': '=C3+1000',
    })


def test_spill_grows_with_changed_parameter():
    compiler = _parametric_compiler()
    assert [compiler.evaluate(s(a)) for a in ('A1', 'B1', 'A2', 'B2')] == \
        [1, 2, 3, 4]
    assert compiler.evaluate(s('D5')) == 104
    assert compiler.evaluate(s('D6')) == 1000

    compiler.set_value(s('H1'), 3)
    assert [compiler.evaluate(s(f'{col}{row}'))
            for row in (1, 2, 3) for col in 'ABC'] == \
        [1, 2, 3, 4, 5, 6, 7, 8, 9]
    assert compiler.evaluate(s('C3')) == 9
    assert compiler.evaluate(s('D5')) == 105
    assert compiler.evaluate(s('D6')) == 1009


def test_spill_shrinks_and_extra_cells_are_cleared():
    compiler = _parametric_compiler()
    compiler.evaluate(s('C3'))  # force the 2x2 area (C3 is empty at 2x2)
    compiler.set_value(s('H1'), 3)
    assert compiler.evaluate(s('C3')) == 9

    compiler.set_value(s('H1'), 1)
    assert compiler.evaluate(s('A1')) == 1
    for address in ('B1', 'A2', 'B2', 'C3'):
        assert compiler.evaluate(s(address)) is None
    assert compiler.evaluate(s('D5')) == 100
    assert compiler.evaluate(s('D6')) == 1000

    compiler.set_value(s('H1'), 2)
    assert [compiler.evaluate(s(a)) for a in ('A1', 'B1', 'A2', 'B2')] == \
        [1, 2, 3, 4]
    assert compiler.evaluate(s('D5')) == 104


def test_writing_into_spill_area_blocks_then_clearing_restores():
    compiler = _parametric_compiler()
    assert compiler.evaluate(s('A1')) == 1
    assert compiler.evaluate(s('B2')) == 4

    compiler.set_value(s('B1'), 77)
    assert compiler.evaluate(s('A1')) == SPILL_ERROR
    assert compiler.evaluate(s('B1')) == 77
    # all previously spilled siblings, other than the written cell, clear
    assert compiler.evaluate(s('A2')) is None
    assert compiler.evaluate(s('B2')) is None
    assert compiler.evaluate(s('D5')) == 100

    # clearing the occupant restores the spill
    compiler.set_value(s('B1'), None)
    assert [compiler.evaluate(s(a)) for a in ('A1', 'B1', 'A2', 'B2')] == \
        [1, 2, 3, 4]
    assert compiler.evaluate(s('D5')) == 104


def test_overwriting_anchor_with_a_number_clears_spill():
    compiler = _parametric_compiler()
    assert compiler.evaluate(s('B2')) == 4

    compiler.set_value(s('A1'), 42)
    assert compiler.evaluate(s('A1')) == 42
    for address in ('B1', 'A2', 'B2'):
        assert compiler.evaluate(s(address)) is None


def test_cross_sheet_formula_follows_parameter_change():
    compiler = make_compiler(
        {
            'H1': 2,
            'A1': '=SEQUENCE(H1,H1)',
        },
        extra_sheets=(('T', {'A1': f'={SHEET}!B2'}),),
    )
    assert compiler.evaluate('T!A1') == 4
    compiler.set_value(s('H1'), 3)
    assert compiler.evaluate('T!A1') == 5
    compiler.set_value(s('H1'), 1)
    # referencing the now empty spilled cell gives 0, as in Excel
    assert compiler.evaluate('T!A1') == 0


# ----------------------------------------------------------------------
# recalculate and serialization
# ----------------------------------------------------------------------

def test_recalculate_rebuilds_spill():
    compiler = make_compiler({'A1': '=SEQUENCE(2,2)'})
    assert compiler.evaluate(s('B2')) == 4
    compiler.recalculate()
    assert [compiler.evaluate(s(a)) for a in ('A1', 'B1', 'A2', 'B2')] == \
        [1, 2, 3, 4]


def test_spill_survives_json_round_trip(tmp_path):
    compiler = make_compiler({
        'H1': 2,
        'A1': '=SEQUENCE(H1,H1)',
        'D5': '=B2+100',
    })
    assert compiler.evaluate(s('B2')) == 4
    assert compiler.evaluate(s('D5')) == 104

    filename = tmp_path / 'spill.json'
    compiler.to_file(str(filename))

    reloaded = ExcelCompiler.from_file(str(filename))
    assert [reloaded.evaluate(s(a)) for a in ('A1', 'B1', 'A2', 'B2')] == \
        [1, 2, 3, 4]
    assert reloaded.evaluate(s('D5')) == 104

    # the dynamic behavior itself survives the round trip
    reloaded.set_value(s('H1'), 3)
    assert reloaded.evaluate(s('B2')) == 5
    assert reloaded.evaluate(s('C3')) == 9
    assert reloaded.evaluate(s('D5')) == 105

    reloaded.set_value(s('H1'), 1)
    assert reloaded.evaluate(s('B2')) is None
    assert reloaded.evaluate(s('C3')) is None
    assert reloaded.evaluate(s('D5')) == 100

    reloaded.set_value(s('B1'), 5)
    reloaded.set_value(s('H1'), 2)
    assert reloaded.evaluate(s('A1')) == SPILL_ERROR
    reloaded.set_value(s('B1'), None)
    assert reloaded.evaluate(s('B2')) == 4
