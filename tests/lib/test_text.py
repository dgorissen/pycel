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
from pycel.excelcompiler import ExcelCompiler
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
    len_,
    lower,
    mid,
    replace,
    right,
    substitute,
    text as text_func,
    trim,
    upper,
    value,
)

# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.text, __name__)


def test_text_ws(fixture_xls_copy):
    compiler = ExcelCompiler(fixture_xls_copy('text.xlsx'))
    result = compiler.validate_serialized()
    assert result == {}


@pytest.mark.parametrize(
    'args, expected', (
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
def test_concatenate(args, expected):
    assert concat(*args) == expected
    assert concatenate(*args) == expected
    assert concat(args) == expected
    assert concatenate(args) == VALUE_ERROR


@pytest.mark.parametrize(
    'to_find, find_in, expected', (
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
def test_find(to_find, find_in, expected):
    assert find(to_find, find_in) == expected


@pytest.mark.parametrize(
    'text, num_chars, expected', (
        ('abcd', 5, 'abcd'),
        ('abcd', 4, 'abcd'),
        ('abcd', 3, 'abc'),
        ('abcd', 2, 'ab'),
        ('abcd', 1, 'a'),
        ('abcd', 0, ''),

        (1.234, 3, '1.2'),

        (True, 3, 'TRU'),
        (False, 2, 'FA'),

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


@pytest.mark.parametrize(
    'text, start, count, expected', (
        (VALUE_ERROR, 2, 2, VALUE_ERROR),
        ('Romain', VALUE_ERROR, 2, VALUE_ERROR),
        ('Romain', 2, VALUE_ERROR, VALUE_ERROR),
        (DIV0, 2, 2, DIV0),
        ('Romain', DIV0, 2, DIV0),
        ('Romain', 2, DIV0, DIV0),

        ('Romain', 'x', 2, VALUE_ERROR),
        ('Romain', 2, 'x', VALUE_ERROR),

        ('Romain', 1, 2.1, 'Ro'),

        ('Romain', 0, 3, VALUE_ERROR),
        ('Romain', 1, -1, VALUE_ERROR),

        (1234, 2, 2, '23'),
        (12.34, 2, 2, '2.'),

        (True, 2, 2, 'RU'),
        (False, 2, 2, 'AL'),
        (None, 2, 2, ''),

        ('Romain', 2, 9, 'omain'),
        ('Romain', 2.1, 2, 'om'),
        ('Romain', 2, 2.1, 'om'),
    )
)
def test_mid(text, start, count, expected):
    assert mid(text, start, count) == expected


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

        (True, 3, 'RUE'),
        (False, 2, 'SE'),

        ('abcd', -1, VALUE_ERROR),
        ('abcd', 'x', VALUE_ERROR),
        (VALUE_ERROR, 1, VALUE_ERROR),
        ('abcd', VALUE_ERROR, VALUE_ERROR),
    )
)
def test_right(text, num_chars, expected):
    assert right(text, num_chars) == expected


@pytest.mark.parametrize(
    'text, old_text, new_text, instance_num, expected', (
        ('abcdef', 'cd', '', None, 'abef'),
        ('abcdef', 'cd', 'X', None, 'abXef'),
        ('abcdef', 'cd', 'XY', None, 'abXYef'),
        ('abcdef', 'cd', '', True, VALUE_ERROR),
        ('abcdef', 'cd', '', 'PLUGH', VALUE_ERROR),

        ('abcdabcdab', 'a', 'X', 1, 'Xbcdabcdab'),
        ('abcdabcdab', 'a', 'X', 2, 'abcdXbcdab'),
        ('abcdabcdab', 'a', 'X', 3, 'abcdabcdXb'),
        ('abcdabcdab', 'ab', 'X', None, 'XcdXcdX'),
        ('abcdabcdab', 'ab', 'X', 0, VALUE_ERROR),
        ('abcdabcdab', 'ab', 'X', 1, 'Xcdabcdab'),
        ('abcdabcdab', 'ab', 'X', 2, 'abcdXcdab'),
        ('abcdabcdab', 'ab', 'X', 3, 'abcdabcdX'),
        ('abcdabcdab', 'ab', 'X', 4, 'abcdabcdab'),
        ('abcdabcdab', 'abc', 'X', 1, 'Xdabcdab'),
        ('abcdabcdab', 'abc', 'X', 2, 'abcdXdab'),

        ('abcdabcdab', 'cd', 'X', None, 'abXabXab'),
        ('abcdabcdab', 'cd', 'X', -1, VALUE_ERROR),
        ('abcdabcdab', 'cd', 'X', 0, VALUE_ERROR),
        ('abcdabcdab', 'cd', 'X', 1, 'abXabcdab'),
        ('abcdabcdab', 'cd', 'X', 2, 'abcdabXab'),
        ('abcdabcdab', 'cd', 'X', 3, 'abcdabcdab'),

        (VALUE_ERROR, 'ab', 'X', None, VALUE_ERROR),
        ('abcdabcdab', DIV0, 'X', None, DIV0),
        ('abcdabcdab', 'ab', NAME_ERROR, None, NAME_ERROR),
        ('abcdabcdab', 'ab', 'X', NA_ERROR, NA_ERROR),

        (True, 'R', '', None, 'TUE'),
        (False, 'AL', '^', None, 'F^SE'),
        (False, 'AL', 1.2, None, 'F1.2SE'),
        (321.245, 21, 1.2, None, '31.2.245'),
    )
)
def test_substitute(text, old_text, new_text, instance_num, expected):
    assert substitute(text, old_text, new_text, instance_num) == expected


@pytest.mark.parametrize(
    'text_value, value_format, expected', (

        # Thousand separator
        ('12200000', '#,###', '12,200,000'),
        ('12200000', '0,000.00', '12,200,000.00'),

        # Number, currency, accounting
        ('1234.56', '0.00', '1234.56'),
        ('1234.56', '#,##0', '1,235'),
        ('1234.56', '#,##0.00', '1,234.56'),
        ('1234.56', '$#,##0', '$1,235'),
        ('1234.56', '$#,##0.00', '$1,234.56'),
        ('1234.56', '$ * #,##0', '$ 1,235'),
        ('1234.56', '$ * #,##0.00', '$ 1,234.56'),

        # Months, days, years
        ('2021-01-05', 'm', '1'),
        ('2021-01-05', 'mm', '01'),
        ('2021-01-05', 'mmm', 'Jan'),
        ('2021-01-05', 'mmmm', 'January'),
        ('2021-01-05', 'mmmmm', 'J'),
        ('2021-01-05', 'd', '5'),
        ('2021-01-05', 'dd', '05'),
        ('2021-01-05', 'ddd', 'Tue'),
        ('2021-01-05', 'dddd', 'Tuesday'),
        ('2021-01-05', 'ddddd', 'Tuesday'),
        ('2021-01-05', 'y', '21'),
        ('2021-01-05', 'yy', '21'),
        ('2021-01-05', 'yyy', '2021'),
        ('2021-01-05', 'yyyy', '2021'),
        ('2021-01-05', 'yyyyy', '2021'),

        # Hours, minutes and seconds
        ('3:33 am', 'h', '3'),
        ('3:33 am', 'hh', '03'),
        ('3:33 pm', 'h', '15'),
        ('3:33 pm', 'hh', '15'),
        ('3:33 pm', 'm', '1'),
        ('3:33 pm', 'mm', '01'),
        ('3:33:03 pm', 's', '3'),
        ('3:33:30 am', 's', '30'),
        ('3:33:30 pm', 'ss', '30'),
        ('3:33 pm', 'h AM/PM', '3 pm'),
        ('3:33 am', 'h AM/PM', '3 am'),
        ('3:33 pm', 'h:mm AM/PM', '3:33 PM'),
        ('3:33:30 pm', 'h:mm:ss A/P', '3:33:30 P'),
        ('3:33 pm', 'h:mm:ss.00', '15:33:00.00'),
        ('3:22:33.67 pm', 'mm:ss.00', '22:33.67'),
        ('99:99', '', '99:99'),

        # Elapsed Time
        ('3:33 pm', '[h]:mm', '15:33'),
        ('3:33:14 pm', '[mm]:ss', '933:14'),
        ('3:33:14.78 pm', '[ss].00', '55994.78'),
        (39815.17021, '[hh]:mm', '955564:05'),
        ('-1', '[hh]', '-24'),
        ('-1', '[hh]:mm', VALUE_ERROR),
        ('-1', '[mm]', '-1440'),
        ('-1', '[mm]:ss', VALUE_ERROR),
        ('-1', '[ss]', '-86400'),
        ('-1', '[ss].000', VALUE_ERROR),
        ('2958466', '[hh]', '71003184'),  # > 9999-12-31 ok w/ elapsed
        ('2958466', 'hh', VALUE_ERROR),
        ('23:59.012345', '#.#######', '.0166552'),
        ('59:9999.0123', '#.########', '.15670153'),
        ('60:9999.012345', '#.###########', '60:9999.012345'),

        # Date & Time
        ('1989-12-31 15:30:00', 'MM/DD/YYYY', '12/31/1989'),
        ('1989-12-31', 'YYYY-MM-DD', '1989-12-31'),
        ('1989-12-31', 'YYYY/MM/DD', '1989/12/31'),
        (39815.17, 'dddd, mmmm dd, yyyy  hh:mm:ss', 'Friday, January 02, 2009  04:04:48'),
        (39815.17, 'dddd, mmmm dd, yyyy', 'Friday, January 02, 2009'),
        ('1989-12-31 15:30:00', 'MM/DD/YYYY hh:mm AM/PM', '12/31/1989 03:30 pm'),

        # Percentage
        ('0.244740088392962', '0%', '24%'),
        ('0.244740088392962', '0.0%', '24.5%'),
        ('0.244740088392962', '0.00%', '24.47%'),
        ('0.244740088392962', '0#,#%%%%%%%', '24,474,008,839,296%%%%%%%'),

        # text without formatting - returned as-is
        ('test', '', 'test'),
        (55, '', ''),
        (0, '', ''),
        ('-55', '', '-'),

        # non-numerics
        ('FALSE', '#.#', 'FALSE'),
        ('TRUE', '#.#', 'TRUE'),
        ('PLUGH', '#.#', 'PLUGH'),
        (None, 'hh', '00'),
        (None, '#', ''),
        (None, '#.##', '.'),

        # m is month
        ('2000-12-30 15:35', 'am/p', 'a12/p'),
        # dates are converted to serial numbers
        ('2021-01-01 10:3', '#.00000', '44197.41875'),

        # multiple fields
        ('-1', '##.00;##.0;"ZERO";->@<-', '1.0'),
        ('0', '##.00;##.0;"ZERO";->@<-', 'ZERO'),
        ('1', '##.00;##.0;"ZERO";->@<-', '1.00'),
        ('X', '##.00;##.0;"ZERO";->@<-', '->X<-'),

        # mixed fields
        (-1, '#,#ZZ.##', '-1ZZ.'),
        (None, '#,#"ZZ"#.0z00', 'ZZ.0z00'),
        ('0', '#,#"ZZ"#.0z00', 'ZZ.0z00'),
        ('', '#,#"ZZ"#.0z00', ''),
        (-1, 'w;x@', '-w'),
        (890123.456789, ',#.00.0', ',890123.45.7'),
        (890123.456789, '.#', '890123.5'),
        (890123.456789, '%#', '%89012346'),
        (890123.456789, '%#%#%', '%89012345678%9%'),
        (890123.456789, '%#%#%#%', '%890123456789%0%0%'),
        ('1234.56', '.#', '1234.6'),
        ('1234.56', '.##', '1234.56'),
        ('1234.56', '.##0', '1234.560'),

        # format parse errors
        (0, '\\', VALUE_ERROR),
        (0, 'a\\a', 'aa'),
        (0, '[', VALUE_ERROR),
        (0, '[h', VALUE_ERROR),
        (0, '[hm]', VALUE_ERROR),
        (0, '#@', VALUE_ERROR),
        (0, '#m', VALUE_ERROR),
        (0, 'm@', VALUE_ERROR),
        (0, '@;@', VALUE_ERROR),
        (0, '@;#', VALUE_ERROR),
        ('', '0.00*', VALUE_ERROR),
    )
)
def test_text(text_value, value_format, expected):
    assert text_func(text_value, value_format).lower() == expected.lower()


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
    'param, expected', (
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
        (None, 0),
        ('X', VALUE_ERROR),
        ('`1', VALUE_ERROR),
        (False, VALUE_ERROR),
        (True, VALUE_ERROR),
        (NA_ERROR, NA_ERROR),
        (DIV0, DIV0),
    )
)
def test_value(param, expected):
    assert value(param) == expected


@pytest.mark.parametrize(
    'param, expected', (
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
def test_len_(param, expected):
    assert len_(param) == expected
