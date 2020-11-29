# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import pytest

import pycel.excellib
from pycel.excelutil import (
    DIV0,
    NA_ERROR,
    NAME_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import load_to_test_module
from pycel.lib.text import (
    concat,
    concatenate,
    find,
    left,
    lower,
    mid,
    replace,
    right,
    text,
    trim,
    upper,
    value,
    x_len,
)

# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.text, __name__)


@pytest.mark.parametrize(
    'args, result', (
        ('a 1 abc'.split(), 'a1abc'),
        ('a Jan-00 abc'.split(), 'aJan-00abc'),
        ('a	#DIV/0! abc'.split(), DIV0),
        ('a	1 #DIV/0!'.split(), DIV0),
        ('a #NAME? abc'.split(), NAME_ERROR),
        (('a', True, 'abc'), 'aTRUEabc'),
        (('a', False, 'abc'), 'aFALSEabc'),
        (('a', 2, 'abc'), 'a2abc'),
    )
)
def test_concatenate(args, result):
    assert concat(*args) == result
    assert concatenate(*args) == result
    assert concat(args) == result
    assert concatenate(args) == VALUE_ERROR


@pytest.mark.parametrize(
    'to_find, find_in, result', (
        (2, 2.5, 1),
        ('.', 2.5, 2),
        (5, 2.5, 3),
        ('2', 2.5, 1),
        ('.', 2.5, 2),
        ('5', 2.5, 3),
        ('2', '2.5', 1),
        ('.', '2.5', 2),
        ('T', True, 1),
        ('U', True, 3),
        ('u', True, VALUE_ERROR),
        (DIV0, 'x' + DIV0, DIV0),
        ('V', DIV0, DIV0),
    )
)
def test_find(to_find, find_in, result):
    assert find(to_find, find_in) == result


@pytest.mark.parametrize(
    'text, num_chars, expected', (
        ('abcd', 5, 'abcd'),
        ('abcd', 4, 'abcd'),
        ('abcd', 3, 'abc'),
        ('abcd', 2, 'ab'),
        ('abcd', 1, 'a'),
        ('abcd', 0, ''),

        (1.234, 3, '1.2'),

        ('abcd', -1, VALUE_ERROR),
        ('abcd', 'x', VALUE_ERROR),
        (DIV0, 1, DIV0),
        ('abcd', NAME_ERROR, NAME_ERROR),
    )
)
def test_left(text, num_chars, expected):
    assert left(text, num_chars) == expected


@pytest.mark.parametrize(
    'text, expected', (
        ('aBcD', 'abcd'),
        (1.234, '1.234'),
        (1, '1'),
        (True, 'true'),
        (False, 'false'),
        ('TRUe', 'true'),
        (DIV0, DIV0),
    )
)
def test_lower(text, expected):
    assert lower(text) == expected


class TestMid:

    def test_invalid_parameters(self):
        assert mid(VALUE_ERROR, 2, 2) == VALUE_ERROR
        assert mid('Romain', VALUE_ERROR, 2) == VALUE_ERROR
        assert mid('Romain', 2, VALUE_ERROR) == VALUE_ERROR
        assert mid(DIV0, 2, 2) == DIV0
        assert mid('Romain', DIV0, 2) == DIV0
        assert mid('Romain', 2, DIV0) == DIV0

        assert mid('Romain', 'x', 2) == VALUE_ERROR
        assert mid('Romain', 2, 'x') == VALUE_ERROR

    def test_num_chars_must_be_integer(self):
        assert 'Ro' == mid('Romain', 1, 2.1)

    def test_start_num_must_be_superior_or_equal_to_1(self):
        assert VALUE_ERROR == mid('Romain', 0, 3)

    def test_num_chars_must_be_positive(self):
        assert VALUE_ERROR == mid('Romain', 1, -1)

    def test_from_not_str(self):
        assert '23' == mid(1234, 2, 2)

    def test_mid(self):
        assert 'omain' == mid('Romain', 2, 9)
        assert 'om' == mid('Romain', 2.1, 2)
        assert 'om' == mid('Romain', 2, 2.1)


@pytest.mark.parametrize(
    'expected, old_text, start_num, num_chars, new_text', (
        ('AB CD_X_', 'AB CD', 7, 2, '_X_'),
        ('AB CD_X_', 'AB CD', 6, 2, '_X_'),
        ('AB C_X_', 'AB CD', 5, 2, '_X_'),
        ('AB _X_', 'AB CD', 4, 2, '_X_'),
        ('AB_X_D', 'AB CD', 3, 2, '_X_'),
        ('A_X_CD', 'AB CD', 2, 2, '_X_'),
        ('_X_ CD', 'AB CD', 1, 2, '_X_'),
        (VALUE_ERROR, 'AB CD', 0, 2, '_X_'),
        ('_X_', 'AB CD', 1, 6, '_X_'),
        ('_X_', 'AB CD', 1, 5, '_X_'),
        ('_X_D', 'AB CD', 1, 4, '_X_'),
        ('AB C_X_', 'AB CD', 5, 1, '_X_'),
        ('AB C_X_', 'AB CD', 5, 2, '_X_'),
        ('AB _X_D', 'AB CD', 4, 1, '_X_'),
        ('AB _X_', 'AB CD', 4, 2, '_X_'),
        ('AB_X_ CD', 'AB CD', 3, 0, '_X_'),
        (VALUE_ERROR, 'AB CD', 3, -1, '_X_'),
        ('_X_ CD', 'AB CD', True, 2, '_X_'),
        (VALUE_ERROR, 'AB CD', False, 2, '_X_'),
        ('AB_X_CD', 'AB CD', 3, True, '_X_'),
        ('AB_X_ CD', 'AB CD', 3, False, '_X_'),
        ('_X_ CD', 'AB CD', 1, 2, '_X_'),
        (VALUE_ERROR, 'AB CD', 0, 2, '_X_'),
        (DIV0, DIV0, 2, 2, '_X_'),
        (DIV0, 'AB CD', DIV0, 2, '_X_'),
        (DIV0, 'AB CD', 2, DIV0, '_X_'),
        (DIV0, 'AB CD', 2, 2, DIV0),
        ('A0CD', 'AB CD', 2, 2, '0'),
        ('AFALSECD', 'AB CD', 2, 2, 'FALSE'),
        ('T_X_E', 'TRUE', 2, 2, '_X_'),
        ('F_X_SE', 'FALSE', 2, 2, '_X_'),
        ('A_X_', 'A', 2, 2, '_X_'),
        ('1_X_1', '1.1', 2, 1, '_X_'),
        (VALUE_ERROR, '1.1', 'A', 1, '_X_'),
        (VALUE_ERROR, '1.1', 2, 'A', '_X_'),
        ('1_X_1', '1.1', 2.2, 1, '_X_'),
        ('1_X_1', '1.1', 2.9, 1, '_X_'),
        ('1._X_', '1.1', 3, 1, '_X_'),
        ('1_X_1', '1.1', 2, 1.5, '_X_'),
        ('1.0', '1.1', 3, 1, 0),
    )
)
def test_replace(expected, old_text, start_num, num_chars, new_text):
    assert replace(old_text, start_num, num_chars, new_text) == expected


@pytest.mark.parametrize(
    'text, num_chars, expected', (
        ('abcd', 5, 'abcd'),
        ('abcd', 4, 'abcd'),
        ('abcd', 3, 'bcd'),
        ('abcd', 2, 'cd'),
        ('abcd', 1, 'd'),
        ('abcd', 0, ''),

        (1234.1, 2, '.1'),

        ('abcd', -1, VALUE_ERROR),
        ('abcd', 'x', VALUE_ERROR),
        (VALUE_ERROR, 1, VALUE_ERROR),
        ('abcd', VALUE_ERROR, VALUE_ERROR),
    )
)
def test_right(text, num_chars, expected):
    assert right(text, num_chars) == expected


@pytest.mark.parametrize(
    'text_value, value_format, result', (

        # Thousand separator
        ("12200000", "#,###", "12,200,000"),
        ("12200000", "0,000.00", "12,200,000.00"),

        # Number, currency, accounting
        ("1234.56", "0.00", "1234.56"),
        ("1234.56", "#,##0", "1,235"),
        ("1234.56", "#,##0.00", "1,234.56"),
        ("1234.56", "$#,##0", "$1,235"),
        ("1234.56", "$#,##0.00", "$1,234.56"),
        ("1234.56", "$ * #,##0", "$1,235"),
        ("1234.56", "$ * #,##0.00", "$1,234.56"),

        # Months, days, years
        ('15/01/2021', "m", '01'),  # Excel returns 1
        ('15/01/2021', "mm", '01'),
        ('15/01/2021', "mmm", 'Jan'),
        ('15/01/2021', "mmmm", 'January'),
        ('15/01/2021', "mmmmm", 'Jan'),  # Excel returns J
        ('15/01/2021', "d", '15'),
        ('15/01/2021', "dd", '15'),
        ('15/01/2021', "ddd", 'Fri'),
        ('15/01/2021', "dddd", 'Friday'),
        ('15/01/2021', "yy", '21'),
        ('15/01/2021', "yyyy", '2021'),
        ('2021-01-15', 'yyyy', '2021'),

        # Hours, minutes and seconds
        ('3:33 pm', "h", '15'),
        ('3:33 pm', "hh", '15'),
        ('3:33 pm', "m", '01'),  # Excel returns 1
        ('3:33 pm', "mm", '01'),
        ('3:33:30 pm', "s", '30'),
        ('3:33:30 pm', "ss", '30'),
        ('3:33 pm', "h AM/PM", '03 pm'),
        ('3:33 am', "h AM/PM", '03 am'),
        ('3:33 pm', "h:mm AM/PM", '03:33 pm'),
        ('3:33:30 pm', "h:mm:ss A/P", '15:33:30 A/P'),
        ('3:33 pm', "h:mm:ss.00", '15:33:00.00'),
        ('99:99', '', '99:99'),
        # not supported
        # ('3:33 pm', "[h]:mm", '1:02'),
        # ('3:33 pm', "[mm]:ss", '62:16'),
        # ('3:33 pm', "[ss].00", '3735.80'),

        # Date & Time
        ("31/12/1989 15:30:00", "MM/DD/YYYY", "12/31/1989"),
        ("31/12/89 15:30:00", "MM/DD/YY", "12/31/89"),
        ("1989-12-31", "YYYY-MM-DD", "1989-12-31"),
        ("1989/12/31", "YYYY/MM/DD", "1989/12/31"),

        ("31/12/1989 15:30:00",
         "MM/DD/YYYY hh:mm AM/PM", "12/31/1989 03:30 pm"),

        # Percentage
        ('0.244740088392962', '0%', '24%'),
        ('0.244740088392962', '0.0%', '24.5%'),
        ('0.244740088392962', '0.00%', '24.47%'),

        # text without formatting - returned as-is
        ('test', '', 'test'),
        (55, '', '55'),

    )
)
def test_text(text_value, value_format, result):
    assert text(text_value, value_format).lower() == result.lower()


@pytest.mark.parametrize(
    'text, expected', (
        ('ABCD', 'ABCD'),
        ('AB CD', 'AB CD'),
        ('AB  CD', 'AB CD'),
        ('AB   CD   EF', 'AB CD EF'),
        (1.234, '1.234'),
        (1, '1'),
        (True, 'TRUE'),
        (False, 'FALSE'),
        ('tRUe', 'tRUe'),
        (DIV0, DIV0),
    )
)
def test_trim(text, expected):
    assert trim(text) == expected


@pytest.mark.parametrize(
    'text, expected', (
        ('aBcD', 'ABCD'),
        (1.234, '1.234'),
        (1, '1'),
        (True, 'TRUE'),
        (False, 'FALSE'),
        ('tRUe', 'TRUE'),
        (DIV0, DIV0),
    )
)
def test_upper(text, expected):
    assert upper(text) == expected


@pytest.mark.parametrize(
    'param, result', (
        (0, 0),
        (2, 2),
        (2.1, 2.1),
        (-2.1, -2.1),
        ('-2.1', -2.1),
        ('3', 3),
        ('3.', 3),
        ('3.0', 3),
        ('.01', 0.01),
        ('1E5', 100000),
        ('X', VALUE_ERROR),
        ('`1', VALUE_ERROR),
        (False, VALUE_ERROR),
        (True, VALUE_ERROR),
        (NA_ERROR, NA_ERROR),
        (DIV0, DIV0),
    )
)
def test_value(param, result):
    assert value(param) == result


@pytest.mark.parametrize(
    'param, result', (
        ('A', 1),
        ('BB', 2),
        (3.0, 3),
        (True, 4),
        (False, 5),
        (None, 0),
        (NA_ERROR, NA_ERROR),
        (DIV0, DIV0),
    )
)
def test_x_len(param, result):
    assert x_len(param) == result
