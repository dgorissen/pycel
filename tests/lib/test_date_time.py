# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import datetime as dt

import pytest

import pycel.lib.date_time
from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import (
    DIV0,
    NUM_ERROR,
    VALUE_ERROR,
)
from pycel.lib.date_time import (
    date,
    date_from_int,
    DATE_MAX_INT,
    DATE_ZERO,
    DateTimeFormatter,
    datevalue,
    is_leap_year,
    max_days_in_month,
    MICROSECOND,
    normalize_year,
    now,
    SECOND,
    time_from_serialnumber,
    time_from_serialnumber_with_microseconds,
    timevalue,
    today,
    yearfrac,
)
from pycel.lib.function_helpers import load_to_test_module


# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.date_time, __name__)


@pytest.mark.parametrize(
    'result, value', (
        ((1900, 1, 1), 1),
        ((1900, 1, 31), 31),
        ((1900, 2, 29), 60),
        ((1900, 3, 1), 61),
        ((2009, 7, 6), 40000),
    )
)
def test_date_from_int(result, value):
    assert date_from_int(value) == result


@pytest.mark.parametrize(
    'result, value', (
        ((23, 59, 59, 999999), 0 - (MICROSECOND)),
        ((0, 0, 0, 0), 0),
        ((23, 58, 33, 600000), 0.999),
        ((23, 59, 51, 360000), 0.9999),
        ((23, 59, 59, 0), 1 - (MICROSECOND * 1e6)),
        ((23, 59, 59, 136000), 0.99999),
        ((23, 59, 59, 900000), 1 - (MICROSECOND * 1e5)),
        ((23, 59, 59, 913600), 0.999999),
        ((23, 59, 59, 990000), 1 - (MICROSECOND * 1e4)),
        ((23, 59, 59, 999000), 1 - (MICROSECOND * 1e3)),
        ((23, 59, 59, 999900), 1 - (MICROSECOND * 1e2)),
        ((23, 59, 59, 999990), 1 - (MICROSECOND * 1e1)),
        ((23, 59, 59, 999999), 1 - (MICROSECOND)),
        ((0, 0, 0, 0), 1),
        ((0, 0, 0, 1), 1 + (MICROSECOND)),
        ((0, 0, 0, 10), 1 + (MICROSECOND * 10)),
        ((23, 59, 59, 500000), 0 - MICROSECOND * 5e5),
        ((23, 59, 59, 500000), 1 - MICROSECOND * 5e5),
        ((2, 24, 0, 0), 1.1),
        ((23, 59, 59, 999999), 1234 - (MICROSECOND)),
        ((0, 0, 0, 0), 1234),
        ((0, 0, 0, 1), 1234 + (MICROSECOND)),
    )
)
def test_time_from_serialnumber_microsecond(result, value):
    assert time_from_serialnumber_with_microseconds(value) == result


@pytest.mark.parametrize(
    'result, value', (
        ((0, 0, 0), 1),
        ((23, 58, 34), 0.999),
        ((23, 59, 51), 0.9999),
        ((23, 59, 59), 1 - (MICROSECOND * 1e6)),
        ((0, 0, 0), 0),
        ((23, 59, 59), 0 - MICROSECOND * 5e5),
        ((23, 59, 59), 1 - MICROSECOND * 5e5),
        ((2, 24, 0), 1.1),
    )
)
def test_time_from_serialnumber(result, value):
    assert time_from_serialnumber(value) == result


@pytest.mark.parametrize(
    'value, result', (
        (1900, True),
        (1904, True),
        (2000, True),
        (2104, True),

        (1, False),
        (2100, False),
        (2101, False),
        (2103, False),
        (2102, False),

        ('x', TypeError),
        (-1, TypeError),
        (0, TypeError),
    )
)
def test_is_leap_year(value, result):
    if result == TypeError:
        with pytest.raises(result):
            is_leap_year(value)
    else:
        assert is_leap_year(value) == result


def test_get_max_days_in_month():
    assert 31 == max_days_in_month(1, 2000)
    assert 29 == max_days_in_month(2, 2000)
    assert 28 == max_days_in_month(2, 2001)
    assert 31 == max_days_in_month(3, 2000)
    assert 30 == max_days_in_month(4, 2000)
    assert 31 == max_days_in_month(5, 2000)
    assert 30 == max_days_in_month(6, 2000)
    assert 31 == max_days_in_month(7, 2000)
    assert 31 == max_days_in_month(8, 2000)
    assert 30 == max_days_in_month(9, 2000)
    assert 31 == max_days_in_month(10, 2000)
    assert 30 == max_days_in_month(11, 2000)
    assert 31 == max_days_in_month(12, 2000)

    # excel thinks 1900 is a leap year
    assert 29 == max_days_in_month(2, 1900)


@pytest.mark.parametrize(
    'result, value', (
        ((1900, 1, 1), (1900, 1, 1)),
        ((1900, 2, 1), (1900, 1, 32)),
        ((1900, 3, 1), (1900, 1, 61)),
        ((1900, 4, 1), (1900, 1, 92)),
        ((1900, 5, 1), (1900, 1, 122)),
        ((1900, 4, 1), (1900, 0, 123)),
        ((1900, 3, 1), (1900, -1, 122)),

        ((1899, 12, 1), (1900, 1, -31)),
        ((1899, 12, 1), (1900, 0, 1)),
        ((1899, 11, 1), (1900, -1, 1)),

        ((1918, 12, 1), (1920, -12, 1)),
        ((1919, 1, 1), (1920, -11, 1)),
        ((1919, 11, 1), (1920, -1, 1)),
        ((1919, 12, 1), (1920, 0, 1)),
        ((1920, 1, 1), (1920, 1, 1)),
        ((1920, 11, 1), (1920, 11, 1)),
        ((1920, 12, 1), (1920, 12, 1)),
        ((1921, 1, 1), (1920, 13, 1)),
        ((1921, 11, 1), (1920, 23, 1)),
        ((1921, 12, 1), (1920, 24, 1)),
        ((1922, 1, 1), (1920, 25, 1)),
    )
)
def test_normalize_year(result, value):
    assert normalize_year(*value) == result


@pytest.mark.parametrize(
    'year, month, day, expected', (
        ('2016', 1, 1, date(2016, 1, 1)),
        (2016, '1', 1, date(2016, 1, 1)),
        (2016, 1, '1', date(2016, 1, 1)),
        (-1, 1, 1, NUM_ERROR),
        (10000, 1, 1, NUM_ERROR),
        (1900, 1, -1, NUM_ERROR),
        (1900, 1, 0, 0),
        (1900, 1, 1, 1),
        (1900, 2, 28, 59),
        (1900, 2, 29, 60),
        (1900, 3, 1, 61),
        (2009, -1, 1, date(2008, 11, 1)),
        (2009, 1, -1, date(2008, 12, 30)),
        (2009, 14, 1, date(2010, 2, 1)),
        (2009, 1, 400, date(2010, 2, 4)),
        (2008, 2, 29, 39507),
        (2008, 11, 3, 39755),
    )
)
def test_date(year, month, day, expected):
    assert date(year, month, day) == expected


def test_date_year_offset():
    zero = dt.datetime(1900, 1, 1) - dt.timedelta(2)
    assert (dt.datetime(1900, 1, 1) - zero).days == date(0, 1, 2)
    assert (dt.datetime(1900 + 1899, 1, 1) - zero).days == date(1899, 1, 1)
    assert (dt.datetime(1900 + 1899, 1, 1) - zero).days == date(1899, 1, 1)


@pytest.mark.parametrize(
    'value, expected', (
        ('1:00', 1 / 24),
        ('1:30', 1.5 / 24),
        ('1:30:30', (1.5 + 1 / 120) / 24),
        ('2:00', 2 / 24),
        ('2:00 PM', 14 / 24),
        ('1:00:00', 1 / 24),
        ('2:00:00 AM', 2 / 24),
        ('1:00:00 PM', 13 / 24),
        ('12:59:59 A', 0.041655093),
        ('12:59:59.99 A', 0.041666551),
        ('12.1:59:59.99 A', VALUE_ERROR),
        ('12:59.1:59.99 A', VALUE_ERROR),
        ('12:XX:59 AM', VALUE_ERROR),
        (DIV0, DIV0),
        (VALUE_ERROR, VALUE_ERROR),
        ('1', VALUE_ERROR),
        ('1 AM', VALUE_ERROR),
        ('1:00 ZM', VALUE_ERROR),
        ('1:00:00 ZM', VALUE_ERROR),
        (0, VALUE_ERROR),
        (1.0, VALUE_ERROR),
        (True, VALUE_ERROR),

        ('23:9999', 23 / 24 + 9999 / 24 / 60),
        ('24:9999', VALUE_ERROR),
        ('23:100000', VALUE_ERROR),
        ('23:59:9999', 23 / 24 + 59 / 24 / 60 + 9999 / 24 / 60 / 60),
        ('23:59:10000', VALUE_ERROR),
        ('59:9999.0', 59 / 24 / 60 + 9999 / 24 / 60 / 60),
        ('60:9999.0', VALUE_ERROR),
        ('59:9999.123456789', 59 / 24 / 60 + 9999.123456789 / 24 / 60 / 60),

        ('23.:59:9999', VALUE_ERROR),
        ('23:59.:9999', VALUE_ERROR),
        ('23:59:9999.', 23 / 24 + 59 / 24 / 60 + 9999 / 24 / 60 / 60),
        ('23:59', 23 / 24 + 59 / 24 / 60),
        ('23:59.', 23 / 24 + 59 / 24 / 60),
        ('23:59.0', 23 / 24 / 60 + 59 / 24 / 60 / 60),
        ('23:59.0.', VALUE_ERROR),
    )
)
def test_timevalue(value, expected):
    result = timevalue(value)
    if isinstance(result, str) or isinstance(expected, str):
        assert result == expected
    else:
        assert result == pytest.approx(expected)


@pytest.mark.parametrize(
    'value, expected', (
        ('12/31/1899', VALUE_ERROR),
        ('1/1/1900', 1),
        ('2/28/1900', 59),
        ('2/29/1900', 60),
        ('2/30/1900', VALUE_ERROR),
        ('12/31/1900 12:00', 366),
        ('3/1/1900', 61),
        ('1/1/1950', 18264),
        ('1/1/2000', 36526),
        ('xyzzy', VALUE_ERROR),
        (1, VALUE_ERROR),
        ('TRUE', VALUE_ERROR),
        ('1.1', VALUE_ERROR),
        (True, VALUE_ERROR),
        (DIV0, DIV0),
    )
)
def test_datevalue(value, expected):
    assert datevalue(value) == expected


def test_today_now():
    before = dt.date.today()
    a_today = today()
    after = dt.date.today()
    assert before <= DATE_ZERO.date() + dt.timedelta(days=a_today) <= after

    before = dt.datetime.now()
    a_now = now()
    after = dt.datetime.now()

    days = int(a_now)
    seconds = int((a_now - days) / SECOND + 1e-6)
    now_dt = dt.timedelta(days=days, seconds=seconds)

    before -= dt.timedelta(microseconds=before.microsecond)
    after -= dt.timedelta(microseconds=after.microsecond)
    assert before <= DATE_ZERO + now_dt <= after


class TestYearfrac:

    def test_start_date_must_be_number(self):
        assert VALUE_ERROR == yearfrac('not a number', 1)

    def test_end_date_must_be_number(self):
        assert VALUE_ERROR == yearfrac(1, 'not a number')

    def test_start_date_must_be_positive(self):
        assert NUM_ERROR == yearfrac(-1, 0)

    def test_end_date_must_be_positive(self):
        assert NUM_ERROR == yearfrac(0, -1)

    def test_basis_must_be_between_0_and_4(self):
        assert NUM_ERROR == yearfrac(1, 2, 5)

    def test_yearfrac_basis_0(self):
        assert 7.30277777777778 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20)))

    def test_yearfrac_basis_1(self):
        assert 7.299110198 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 1))

    def test_yearfrac_basis_2(self):
        assert 7.405555556 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 2))

    def test_yearfrac_basis_3(self):
        assert 7.304109589 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 3))

    def test_yearfrac_basis_4(self):
        assert 7.302777778 == pytest.approx(
            yearfrac(date(2008, 1, 1), date(2015, 4, 20), 4))

    def test_yearfrac_inverted(self):
        assert yearfrac(date(2008, 1, 1), date(2015, 4, 20)) == pytest.approx(
            yearfrac(date(2015, 4, 20), date(2008, 1, 1)))

    def test_yearfrac_basis_1_sub_year(self):
        assert 11 / 365 == pytest.approx(
            yearfrac(date(2015, 4, 20), date(2015, 5, 1), 1))

        assert 11 / 366 == pytest.approx(
            yearfrac(date(2016, 4, 20), date(2016, 5, 1), 1))

        assert 316 / 366 == pytest.approx(
            yearfrac(date(2016, 2, 20), date(2017, 1, 1), 1))

        assert 61 / 366 == pytest.approx(
            yearfrac(date(2015, 12, 31), date(2016, 3, 1), 1))

    @pytest.mark.parametrize(
        'start, end, expected', (
            (date(2007, 2, 28), date(2007, 3, 31), 0.086111111),
            (date(2007, 2, 28), date(2007, 8, 31), 0.502777778),
            (date(2008, 2, 29), date(2008, 3, 31), 0.086111111),
            (date(2008, 2, 29), date(2008, 8, 31), 0.502777778),
        )
    )
    def test_yearfrac_basis_0_feb_eom(self, start, end, expected):
        assert expected == pytest.approx(yearfrac(start, end, 0))


def test_yearfrac_ws(fixture_xls_copy):
    excel_compiler = ExcelCompiler(fixture_xls_copy('yearfrac.xlsx'))

    failed_cells = excel_compiler.validate_serialized()
    assert failed_cells == {}


@pytest.mark.parametrize(
    'datetime, format, expected', (
        ('2007-02-03 9:08:07.0123', 'e', '2007'),
        ('2007-02-03 9:08:07.0123', 'yyyyy', '2007'),
        ('2007-02-03 9:08:07.0123', 'yyyy', '2007'),
        ('2007-02-03 9:08:07.0123', 'yyy', '2007'),
        ('2007-02-03 9:08:07.0123', 'yy', '07'),
        ('2007-02-03 9:08:07.0123', 'y', '07'),
        ('2007-02-03 9:08:07.0123', 'mmmmmm', 'February'),
        ('2007-02-03 9:08:07.0123', 'mmmmm', 'F'),
        ('2007-02-03 9:08:07.0123', 'mmmm', 'February'),
        ('2007-02-03 9:08:07.0123', 'mmm', 'Feb'),
        ('2007-02-03 9:08:07.0123', 'mm', '02'),
        ('2007-02-03 9:08:07.0123', 'm', '2'),
        ('2007-02-03 9:08:07.0123', 'ddddd', 'Saturday'),
        ('2007-02-03 9:08:07.0123', 'dddd', 'Saturday'),
        ('2007-02-03 9:08:07.0123', 'ddd', 'Sat'),
        ('2007-02-03 9:08:07.0123', 'dd', '03'),
        ('2007-02-03 9:08:07.0123', 'd', '3'),
        ('2007-02-03 9:08:07.0123', 'hhh', '09'),
        ('2007-02-03 9:08:07.0123', 'hh', '09'),
        ('2007-02-03 9:08:07.0123', 'h', '9'),
        ('2007-02-03 9:08:07.0123', 'HHH', '09'),
        ('2007-02-03 9:08:07.0123', 'HH', '09'),
        ('2007-02-03 9:08:07.0123', 'H', '9'),
        ('2007-02-03 9:08:07.0123', 'MMM', '08'),
        ('2007-02-03 9:08:07.0123', 'MM', '08'),
        ('2007-02-03 9:08:07.0123', 'M', '8'),
        ('2007-02-03 9:08:07.0123', 'sss', '07'),
        ('2007-02-03 9:08:07.0123', 'ss', '07'),
        ('2007-02-03 9:08:07.0123', 's', '7'),
        ('2007-02-03 9:08:07.0123', '.000', '.012'),
        ('2007-02-03 9:08:07.0123', '.00', '.01'),
        ('2007-02-03 9:08:07.0123', '.0', '.0'),
        ('2007-02-03 9:08:07.0123', '.', '.'),
        ('2007-02-03 9:08:07.0123', 'am/pm', 'AM'),
        ('2007-02-03 9:08:07.0123', 'a/p', 'a'),
        ('2007-02-03 9:08:07.0123', 'A/P', 'A'),
        ('2007-02-03 9:08:07.0123', 'A/p', 'A'),
        ('2007-02-03 9:08:07.0123', 'a/P', 'a'),
        ('2007-02-03 19:08:07.0123', 'am/pm', 'PM'),
        ('2007-02-03 19:08:07.0123', 'a/p', 'p'),
        ('2007-02-03 19:08:07.0123', 'A/P', 'P'),
        ('2007-02-03 19:08:07.0123', 'A/p', 'p'),
        ('2007-02-03 19:08:07.0123', 'a/P', 'P'),
        ('2007-02-03 19:08:07.0123', '[h]', '938803'),
        ('2007-02-03 19:08:07.0123', '[m]', '56328188'),
        ('2007-02-03 19:08:07.0123', '[s]', '3379691287'),
        ('1907-02-03 9:08:07.0123', 'yy', '07'),
        ('19:08:07.0123', 'd', '0'),
        ('19:08:07.0123', 'h', '19'),
        ('19:08:07.0123', '.000', '.012'),
        ('1907-02-03', 'yy', '07'),
        ('1907-02-03', 'h', '0'),
        ('1907-02-03', 's', '0'),
        ('1907-02-03', '.000', '.000'),
        ('10:08:07 pm', 'd', '0'),
        (39116.79730338, 'yyyy', '2007'),
        ('60', 'yyyy', '1900'),
        ('60', 'mm', '02'),
        ('60', 'dd', '29'),
        (60, 'yyyy', '1900'),
        (60, 'mm', '02'),
        (60, 'dd', '29'),
        (0, 'yyyy', '1900'),
        (0, 'mm', '01'),
        (0, 'dd', '00'),
        (2591.38063671296, 'yyyy', '1907'),
        (2591.38063671296, 'mm', '02'),
        (2591.38063671296, 'dd', '03'),
        (2591.38063671296, 'hh', '09'),
        (2591.38063671296, 'MM', '08'),
        (2591.38063671296, 'ss', '07'),
        (2591.38063671296, '.000', '.012'),

        (True, '', None),
        (False, '', None),
        ({}, '', None),
        ('1907-02-xx', '', None),
        (-.1, '', None),
        (DATE_MAX_INT, '', None),
        (DATE_MAX_INT - 1, 'yyyy', '9999'),
        (DATE_MAX_INT - 1, 'mm', '12'),
        (DATE_MAX_INT - 1, 'dd', '31'),
        (0, '[hh]', VALUE_ERROR),
    )
)
def test_date_time_formatter_new(datetime, format, expected):
    obj = DateTimeFormatter.new(datetime)
    if expected is None:
        assert obj is expected
    else:
        assert obj.format(format) == expected


@pytest.mark.parametrize(
    'serial_number, format, expected', (
        (39116.79730338, 'yyyy', '2007'),
        (60, 'yyyy', '1900'),
        (60, 'mm', '02'),
        (60, 'dd', '29'),
        (0, 'yyyy', '1900'),
        (0, 'mm', '01'),
        (0, 'dd', '00'),
        (2591.38063671296, 'yyyy', '1907'),
        (2591.38063671296, 'mm', '02'),
        (2591.38063671296, 'dd', '03'),
        (2591.38063671296, 'hh', '09'),
        (2591.38063671296, 'MM', '08'),
        (2591.38063671296, 'ss', '07'),
        (2591.38063671296, '.000', '.012'),
        (DATE_MAX_INT, 'yyyy', VALUE_ERROR),
        (DATE_MAX_INT, '[h]', str(71003184)),
        (DATE_MAX_INT, '[m]', str(71003184 * 60)),
        (DATE_MAX_INT, '[s]', str(71003184 * 60 * 60)),
        (-1.25, 'yyyy', VALUE_ERROR),
        (-1.25, '[h]', str(-30)),
        (-1.25, '[m]', str(-30 * 60)),
        (-1.25, '[s]', str(-30 * 60 * 60)),
    )
)
def test_date_time_formatter_init(serial_number, format, expected):
    assert DateTimeFormatter(serial_number).format(format) == expected


def test_with_spreadsheet(fixture_xls_copy):
    excel_compiler = ExcelCompiler(fixture_xls_copy('date-time.xlsx'))

    failed_cells = excel_compiler.validate_serialized()
    assert failed_cells == {}
