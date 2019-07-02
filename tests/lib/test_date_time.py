import datetime as dt
import pytest

import pycel.lib.date_time
from pycel.lib.date_time import (
    date,
    yearfrac
)

from pycel.excelutil import (
    NUM_ERROR,
    VALUE_ERROR,
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
