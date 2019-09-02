import datetime as dt

import pytest

import pycel.lib.date_time
from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import (
    DATE_ZERO,
    DIV0,
    NUM_ERROR,
    SECOND,
    VALUE_ERROR,
)
from pycel.lib.date_time import (
    date,
    now,
    timevalue,
    today,
    yearfrac,
)
from pycel.lib.function_helpers import load_to_test_module


# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.lib.date_time, __name__)


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
