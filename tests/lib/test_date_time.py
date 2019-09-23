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
    DATE_ZERO,
    datevalue,
    is_leap_year,
    max_days_in_month,
    MICROSECOND,
    normalize_year,
    now,
    SECOND,
    time_from_serialnumber,
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


class TestDate:

    def test_values_can_str(self):
        assert date('2016', 1, 1) == date(2016, '1', 1) == date(2016, 1, '1')

    def test_year_must_be_positive(self):
        assert NUM_ERROR == date(-1, 1, 1)

    def test_year_must_have_less_than_10000(self):
        assert NUM_ERROR == date(10000, 1, 1)

    def test_result_must_be_positive(self):
        assert NUM_ERROR == date(1900, 1, -1)

    def test_not_stricly_positive_month_substracts(self):
        assert date(2009, -1, 1) == date(2008, 11, 1)

    def test_not_stricly_positive_day_substracts(self):
        assert date(2009, 1, -1) == date(2008, 12, 30)

    def test_month_superior_to_12_change_year(self):
        assert date(2009, 14, 1) == date(2010, 2, 1)

    def test_day_superior_to_365_change_year(self):
        assert date(2009, 1, 400) == date(2010, 2, 4)

    def test_year_for_29_feb(self):
        assert 39507 == date(2008, 2, 29)

    def test_year_regular(self):
        assert 39755 == date(2008, 11, 3)

    def test_year_offset(self):
        zero = dt.datetime(1900, 1, 1) - dt.timedelta(2)
        assert (dt.datetime(1900, 1, 1) - zero).days == date(0, 1, 1)
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
    )
)
def test_timevalue(value, expected):
    if isinstance(expected, str):
        assert timevalue(value) == expected
    else:
        assert timevalue(value) == pytest.approx(expected)


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


def test_with_spreadsheet(fixture_xls_copy):
    excel_compiler = ExcelCompiler(fixture_xls_copy('date-time.xlsx'))

    failed_cells = excel_compiler.validate_calcs()
    assert failed_cells == {}
