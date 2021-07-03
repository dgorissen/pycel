# -*- coding: UTF-8 -*-
#
# Copyright 2011-2021 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import pytest

import pycel.excellib
from pycel.excelcompiler import ExcelCompiler
from pycel.excellib import (
    sumif,
    sumifs,
)
from pycel.excelutil import (
    DIV0,
    EMPTY,
    ERROR_CODES,
    flatten,
    NA_ERROR,
    NAME_ERROR,
    NUM_ERROR,
    REF_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import load_to_test_module
from pycel.lib.stats import (
    average,
    averageif,
    averageifs,
    count,
    countif,
    countifs,
    forecast,
    intercept,
    large,
    linest,
    max_,
    maxifs,
    min_,
    minifs,
    slope,
    small,
    trend,
)

# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.excellib, __name__)
load_to_test_module(pycel.lib.stats, __name__)


def test_stats_ws(fixture_xls_copy):
    compiler = ExcelCompiler(fixture_xls_copy('stats.xlsx'))
    result = compiler.validate_calcs(tolerance=1e-6)
    assert result == {}


def assert_np_close(result, expected):
    for res, exp in zip(flatten(result), flatten(expected)):
        if res not in ERROR_CODES and exp not in ERROR_CODES:
            exp = pytest.approx(exp)
        if res != exp:
            assert result == expected


@pytest.mark.parametrize(
    'data, expected', (
        ((1, '3', 2.0, pytest, 3, 'x'), 2),
        (((1, '3', (2.0, pytest, 3), 'x'),), 2),
        (((-0.1, None, 'x', True),), -0.1),
        ((['x'],), DIV0),
        ((VALUE_ERROR,), VALUE_ERROR),
        (((2, VALUE_ERROR),), VALUE_ERROR),
        ((DIV0,), DIV0),
        (((2, DIV0),), DIV0),
    )
)
def test_average(data, expected):
    assert average(*data) == expected


@pytest.mark.parametrize(
    'data, expected', (
        ((12, 12), AssertionError),
        ((12, 12, 12), 12),
        ((((1, 1, 2, 2, 2), ), ((1, 1, 2, 2, 2), ), 2), 2),
        ((((1, 1, 2, 2, 2), ), ((1, 1, 2, 2, 2), ), 3), DIV0),
        ((((1, 2, 3, 4, 5), ), ((1, 2, 3, 4, 5), ), ">=3"), 4),
        ((((100, 123, 12, 23, 634), ),
          ((1, 2, 3, 4, 5), ), ">=3"), 223),
        ((((100, 123), (12, 23)), ((1, 2), (3, 4)), ">=3"), 35 / 2),
        ((((100, 123, 12, 23, None), ),
          ((1, 2, 3, 4, 5), ), ">=3"), 35 / 2),
        (('JUNK', ((), ), ((), ), ), VALUE_ERROR),
        ((((1, 2), ), ((1,), ), '', ((1, 2), ), ''), VALUE_ERROR),
        ((((1, 2, 3, 4, 5), ),
          ((1, 2, 3, 4, 5), ), ">=3",
          ((1, 2, 3, 4, 5), ), "<=4"), 7 / 2),
    )
)
def test_averageifs(data, expected):
    if isinstance(expected, type(Exception)):
        with pytest.raises(expected):
            averageifs(*data)
    else:
        assert averageifs(*data) == expected


def test_count():
    data = (
        0,
        1,
        1.1,
        '1.1',
        True,
        False,
        'A',
        'TRUE',
        'FALSE',
    )
    assert count(data, data[3], data[5], data[7])


@pytest.mark.parametrize(
    'value, criteria, expected', (
        (((7, 25, 13, 25), ), '>10', 3),
        (((7, 25, 13, 25), ), '<10', 1),
        (((7, 10, 13, 25), ), '>=10', 3),
        (((7, 10, 13, 25), ), '<=10', 2),
        (((7, 10, 13, 25), ), '<>10', 3),
        (((7, 'e', 13, 'e'), ), 'e', 2),
        (((7, 'e', 13, 'f'), ), '>e', 1),
        (((7, 25, 13, 25), ), 25, 2),
        (((7, 25, None, 25),), '<10', 1),
        (((7, 25, None, 25),), '>10', 2),
    )
)
def test_countif(value, criteria, expected):
    assert countif(value, criteria) == expected


class TestCountIfs:
    # more tests might be welcomed

    def test_countifs_regular(self):
        assert 1 == countifs(((7, 25, 13, 25), ), 25,
                             ((100, 102, 201, 20), ), ">100")

    def test_countifs_odd_args_len(self):
        with pytest.raises(AssertionError):
            countifs(((7, 25, 13, 25), ), 25, ((100, 102, 201, 20), ))


@pytest.mark.parametrize(
    'Y, X, expected_slope, expected_intercept, expected_fit, input_x', (
        ([[1, 2, 3, 4]], [[2, 3, 4, 5]], 1, -1, 1.5, 2.5),
        ([[1, 2, 3, 4]], [[-2, -3, -4, -5]], -1, -1, 1.5, -2.5),
        ([[-1, -2, -3, -4]], [[2, 3, 4, 5]], -1, 1, -1.5, 2.5),
        ([[1, 2, 3, 'a']], [[2, 3, 4, 5]], VALUE_ERROR, VALUE_ERROR, VALUE_ERROR, 2.5),
        ([[1, 2, 3, 4]], [[2, 3, 4, 'a']], VALUE_ERROR, VALUE_ERROR, VALUE_ERROR, 2.5),
        ([[1, 2], [3, 4]], [[2, 3, 4, 5]], NA_ERROR, NA_ERROR, NA_ERROR, 2.5),
        (NUM_ERROR, [[2, 3, 4, 5]], NUM_ERROR, NUM_ERROR, NUM_ERROR, 2.5),
        ([[1, 2, 3, 4]], NAME_ERROR, NAME_ERROR, NAME_ERROR, NAME_ERROR, 2.5),
        ([[1, 2, 3, 4]], [[2, 3, 4, 5]], None, None, REF_ERROR, REF_ERROR),
        ([[1, 2, 3, 4]], [[2, 2, 2, 2]], DIV0, DIV0, DIV0, 2.5),
        ([[1, 2, 3, 4]], [[2, 3, 4, 5], [1, 2, 4, 8]], NA_ERROR, NA_ERROR, 1.5, ((2.5, 2),)),
        ([[1, 2, 3, 4]], [[2, 3, 4, 5], [1, 2, 4, 8]], NA_ERROR, NA_ERROR, REF_ERROR, ((2.5,),)),
    )
)
def test_forecast_intercept_slope_trend(
        Y, X, expected_slope, expected_intercept, expected_fit, input_x):
    def approx_with_error(result):
        if result in ERROR_CODES:
            return result
        else:
            return pytest.approx(result)

    if expected_fit is not None:
        expected = NA_ERROR if isinstance(input_x, tuple) else expected_fit
        assert forecast(input_x, Y, X) == approx_with_error(expected)
    if expected_intercept is not None:
        assert intercept(Y, X) == approx_with_error(expected_intercept)
    if expected_slope is not None:
        assert slope(Y, X) == approx_with_error(expected_slope)

    if expected_fit not in ERROR_CODES and not isinstance(input_x, tuple):
        assert trend(Y, X, [[input_x, input_x]])[0][0] == approx_with_error(expected_fit)
        assert trend(Y, X, [[input_x, input_x]])[0][1] == approx_with_error(expected_fit)

    if expected_fit == NA_ERROR:
        expected_fit = REF_ERROR
    if expected_fit == DIV0:
        expected_fit = sum(Y[0]) / len(Y[0])
    assert trend(Y, X, input_x) == approx_with_error(expected_fit)


class TestVariousIfsSizing:

    test_vector = tuple(range(1, 7)) + tuple('abcdef')
    test_vectors = ((test_vector, ), ) * 4 + (test_vector[0],) * 4

    conditions = '>3', '<=2', '<=c', '>d'
    data_columns = ('averageif', 'countif', 'sumif', 'averageifs',
                    'countifs', 'maxifs', 'minifs', 'sumifs')

    responses_list = (
        (5, 3, 15, 5, 3, 6, 4, 15),
        (1.5, 2, 3, 1.5, 2, 2, 1, 3),
        (DIV0, 3, 0, DIV0, 3, 0, 0, 0),
        (DIV0, 2, 0, DIV0, 2, 0, 0, 0),

        (DIV0, 0, 0, DIV0, 0, 0, 0, 0),
        (1, 1, 1, 1, 1, 1, 1, 1),
        (DIV0, 0, 0, DIV0, 0, 0, 0, 0),
        (DIV0, 0, 0, DIV0, 0, 0, 0, 0),
    )

    responses = dict(
        (dc, tuple((r, cond, tv) for r, cond, tv in zip(resp, conds, tvs)))
        for dc, resp, tvs, conds in zip(
            data_columns, zip(*responses_list), (test_vectors, ) * 8,
            ((conditions + conditions), ) * 8
        ))

    params = 'expected, criteria, values'

    @staticmethod
    @pytest.mark.parametrize(params, responses['averageif'])
    def test_averageif(expected, criteria, values):
        assert averageif(values, criteria) == expected
        assert averageif(values, criteria, values) == expected

    @staticmethod
    @pytest.mark.parametrize(params, responses['countif'])
    def test_countif(expected, criteria, values):
        assert countif(values, criteria) == expected

    @staticmethod
    @pytest.mark.parametrize(params, responses['sumif'])
    def test_sumif(expected, criteria, values):
        assert sumif(values, criteria) == expected
        assert sumif(values, criteria, values) == expected

    @staticmethod
    @pytest.mark.parametrize(params, responses['averageifs'])
    def test_averageifs(expected, criteria, values):
        assert averageifs(values, values, criteria) == expected

    @staticmethod
    @pytest.mark.parametrize(params, responses['countifs'])
    def test_countifs(expected, criteria, values):
        assert countifs(values, criteria) == expected

    @staticmethod
    @pytest.mark.parametrize(params, responses['maxifs'])
    def test_maxifs(expected, criteria, values):
        assert maxifs(values, values, criteria) == expected

    @staticmethod
    @pytest.mark.parametrize(params, responses['minifs'])
    def test_minifs(expected, criteria, values):
        assert minifs(values, values, criteria) == expected

    @staticmethod
    @pytest.mark.parametrize(params, responses['sumifs'])
    def test_sumifs(expected, criteria, values):
        assert sumifs(values, values, criteria) == expected

    def test_ifs_size_errors(self):
        criteria, v1 = self.responses['sumifs'][0][1:]
        v2 = (v1[0][:-1], )
        assert countifs(v1, criteria, v2, criteria) == VALUE_ERROR
        assert sumifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR
        assert maxifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR
        assert minifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR
        assert averageifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR


@pytest.mark.parametrize(
    'data, k, expected', (
        ([3, 1, 2], 0, NUM_ERROR),
        ([3, 1, 2], 1, 3),
        ([3, 1, 2], 2, 2),
        ([3, 1, 2], 3, 1),
        ([3, 1, 2], 4, NUM_ERROR),
        ([3, 1, 2], '2', 2),
        ([3, 1, 2], 1.1, 2),
        ([3, 1, 2], '1.1', 2),
        ([3, 1, 2], 0.1, NUM_ERROR),
        ([3, 1, 2], 3.1, NUM_ERROR),
        ([3, 1, 2], 'abc', VALUE_ERROR),
        ([3, 1, 2], True, 3),
        ([3, 1, 2], False, NUM_ERROR),
        ([3, 1, 2], 'True', VALUE_ERROR),
        ([3, 1, 2], REF_ERROR, REF_ERROR),
        ([3, 1, 2], EMPTY, VALUE_ERROR),
        (REF_ERROR, 2, REF_ERROR),
        (None, 2, NUM_ERROR),
        ('abc', 2, NUM_ERROR),
        (99, 1, 99),
        ('99', 1, 99),
        ('99.9', 1, 99.9),
        (['99', 9], 1, 99),
        (['99.9', 9], 1, 99.9),
        ([3, 1, 2], None, NUM_ERROR),
        ([3, 1, 2], 0, NUM_ERROR),
        ([3, 1, 2], 4, NUM_ERROR),
        ([3, 1, 'aa'], 2, 1),
        ([3, 1, 'aa'], 3, NUM_ERROR),
        ([3, 1, True], 1, 3),
        ([3, 1, True], 3, NUM_ERROR),
        ([3, 1, '2'], 2, 2),
        ([3, 1, REF_ERROR], 1, REF_ERROR),
    )
)
def test_large(data, k, expected):
    assert large(data, k) == expected


@pytest.mark.parametrize(
    'X, Y, const, stats, expected', (
        ([[1, 2, 3]], [[2, 3, 4]], None, None, ((1, -1),)),
        ([[1, 2, 3]], [[2, 3, 4]], True, None, ((1, -1),)),
        ([[1, 2, 3]], [[2, 3, 4]], False, None, ((0.6896551724137928, 0),)),
        ([[1, 2, 3]], [[2, 3, 4]], True, False, ((1, -1),)),
        ([[1, 2, 3]], [[2, 3, 4]], False, False, ((0.6896551724137928, 0),)),
        ([[1, 2, 3]], [[2, 3, 4]], None, True, (
            (1, -1.0),
            (0, 0),
            (1.0, 0),
            ('#NUM!', 1),
            (2, 0),
        )),
        ([[1, 2, 3]], [[2, 3, 4]], True, True, (
            (1, -1.0),
            (0, 0),
            (1.0, 0),
            ('#NUM!', 1),
            (2, 0),
        )),
        ([[1, 2, 3]], [[2, 3, 4]], False, True, (
            (0.6896551724137928, 0),
            (0.05972588991616818, NA_ERROR),
            (0.9852216748768473, 0.32163376045133846),
            (133.33333333333348, 2),
            (13.79310344827585, 0.20689655172413793),
        )),
        ([[True, 2, 3]], [[2, 3, 4]], None, None, ((1, -1),)),
        ([['1', 2, 3]], [[2, 3, 4]], None, None, VALUE_ERROR),
        ([[1, 2, 3]], [[2, 3, '4']], None, None, VALUE_ERROR),
        ([[1, 2, 3]], [[2, 3, VALUE_ERROR]], None, None, VALUE_ERROR),
        ([[1, 2, 3]], [[2, 3, DIV0]], None, None, VALUE_ERROR),
        ([[NAME_ERROR, 2, 3]], [[2, 3, 4]], None, None, VALUE_ERROR),
        ([[1, 2, 3]], [[2, 3]], None, None, REF_ERROR),
        ([[1, 2, 3]], [[2, 2, 2]], None, None, ((0, 2),)),
        ([[1, 2, 3]], [[2, 2, 2], [3, 3, 3]], None, None, ((0, 0, 2),)),
        ([[1, 2, 3]], [[2, 2, 2], [3, 3, 3]], None, True, (
            (0, 0, 2),
            (0, 0, 0),
            (1, 0, NA_ERROR),
            (NUM_ERROR, 0, NA_ERROR),
            (2, 0, NA_ERROR),
        )),
    )
)
def test_linest(X, Y, const, stats, expected):
    assert_np_close(linest(X, Y, const, stats), expected)


@pytest.mark.parametrize(
    'data, max_expected, min_expected', (
        ('abcd', 0, 0),
        ((2, None, 'x', 3), 3, 2),
        ((-0.1, None, 'x', True), -0.1, -0.1),
        (VALUE_ERROR, VALUE_ERROR, VALUE_ERROR),
        ((2, VALUE_ERROR), VALUE_ERROR, VALUE_ERROR),
        (DIV0, DIV0, DIV0),
        ((2, DIV0), DIV0, DIV0),
    )
)
def test_max_min(data, max_expected, min_expected):
    assert max_(data) == max_expected
    assert min_(data) == min_expected


@pytest.mark.parametrize(
    'data, k, expected', (
        ([3, 1, 2], 0, NUM_ERROR),
        ([3, 1, 2], 1, 1),
        ([3, 1, 2], 2, 2),
        ([3, 1, 2], 3, 3),
        ([3, 1, 2], 4, NUM_ERROR),
        ([3, 1, 2], '2', 2),
        ([3, 1, 2], 1.1, 2),
        ([3, 1, 2], '1.1', 2),
        ([3, 1, 2], 0.1, NUM_ERROR),
        ([3, 1, 2], 3.1, NUM_ERROR),
        ([3, 1, 2], 'abc', VALUE_ERROR),
        ([3, 1, 2], True, 1),
        ([3, 1, 2], False, NUM_ERROR),
        ([3, 1, 2], 'True', VALUE_ERROR),
        ([3, 1, 2], REF_ERROR, REF_ERROR),
        ([3, 1, 2], EMPTY, VALUE_ERROR),
        (REF_ERROR, 2, REF_ERROR),
        (None, 2, NUM_ERROR),
        ('abc', 2, NUM_ERROR),
        (99, 1, 99),
        ('99', 1, 99),
        ('99.9', 1, 99.9),
        (['99', 999], 1, 99),
        (['99.9', 999], 1, 99.9),
        ([3, 1, 2], None, NUM_ERROR),
        ([3, 1, 2], 0, NUM_ERROR),
        ([3, 1, 2], 4, NUM_ERROR),
        ([3, 1, 'aa'], 2, 3),
        ([3, 1, 'aa'], 3, NUM_ERROR),
        ([3, 1, True], 1, 1),
        ([3, 1, True], 3, NUM_ERROR),
        ([3, 1, '2'], 2, 2),
        ([3, 1, REF_ERROR], 1, REF_ERROR),
    )
)
def test_small(data, k, expected):
    assert small(data, k) == expected


@pytest.mark.parametrize(
    'Y, X, new_X, expected', (
        ([[1, 2, 3, 4]], None, None, [[1, 2, 3, 4]]),
        ([[1, 2, 3, 4]], [[2, 3, 4, 5]], None, [[1, 2, 3, 4]]),
        ([[1, 2, 3, 4]], [[2, 3, 4, 5]], [[1, 2]], [[0, 1]]),
        ([[1, 2, 3, 4]], None, [[2, 3, 4, 5]], [[2, 3, 4, 5]]),
        ([[1, 2, 3, 4]], None, 3, 3),
        ([[1, 2, 3]], [[2, 2, 2], [3, 3, 3]], 1, REF_ERROR),
        ([[1, 2, 3]], [[2, 2, 2], [3, 3, 3]], None, ((2, 2, 2),)),
    )
)
def test_trend_shapes(Y, X, new_X, expected):
    import numpy as np

    expected = tuple(flatten(expected))
    result = np.array(trend(Y, X, new_X, True)).ravel()
    assert_np_close(result, expected)

    result = np.array(trend([[x] for x in Y[0]], X, new_X)).ravel()
    assert_np_close(result, expected)

    if X is not None and new_X is None:
        result = np.array(trend([[x] for x in Y[0]], np.array(X).transpose(), None)).ravel()
        assert_np_close(result, expected)
