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
    large,
    # ::TODO:: finish test cases for remainder of functions
    # linest,
    max_,
    maxifs,
    min_,
    minifs,
    small,
)

# dynamic load the lib functions from excellib and apply metadata
load_to_test_module(pycel.excellib, __name__)
load_to_test_module(pycel.lib.stats, __name__)


def x_test_stats_ws(fixture_xls_copy):
    compiler = ExcelCompiler(fixture_xls_copy('stats.xlsx'))
    result = compiler.validate_calcs()
    assert result == {}


@pytest.mark.parametrize(
    'data, result', (
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
def test_average(data, result):
    assert average(*data) == result


@pytest.mark.parametrize(
    'data, result', (
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
def test_averageifs(data, result):
    if isinstance(result, type(Exception)):
        with pytest.raises(result):
            averageifs(*data)
    else:
        assert averageifs(*data) == result


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
    'value, criteria, result', (
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
def test_countif(value, criteria, result):
    assert countif(value, criteria) == result


class TestCountIfs:
    # more tests might be welcomed

    def test_countifs_regular(self):
        assert 1 == countifs(((7, 25, 13, 25), ), 25,
                             ((100, 102, 201, 20), ), ">100")

    def test_countifs_odd_args_len(self):
        with pytest.raises(AssertionError):
            countifs(((7, 25, 13, 25), ), 25, ((100, 102, 201, 20), ))


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

    params = 'result, criteria, values'

    @staticmethod
    @pytest.mark.parametrize(params, responses['averageif'])
    def test_averageif(result, criteria, values):
        assert averageif(values, criteria) == result
        assert averageif(values, criteria, values) == result

    @staticmethod
    @pytest.mark.parametrize(params, responses['countif'])
    def test_countif(result, criteria, values):
        assert countif(values, criteria) == result

    @staticmethod
    @pytest.mark.parametrize(params, responses['sumif'])
    def test_sumif(result, criteria, values):
        assert sumif(values, criteria) == result
        assert sumif(values, criteria, values) == result

    @staticmethod
    @pytest.mark.parametrize(params, responses['averageifs'])
    def test_averageifs(result, criteria, values):
        assert averageifs(values, values, criteria) == result

    @staticmethod
    @pytest.mark.parametrize(params, responses['countifs'])
    def test_countifs(result, criteria, values):
        assert countifs(values, criteria) == result

    @staticmethod
    @pytest.mark.parametrize(params, responses['maxifs'])
    def test_maxifs(result, criteria, values):
        assert maxifs(values, values, criteria) == result

    @staticmethod
    @pytest.mark.parametrize(params, responses['minifs'])
    def test_minifs(result, criteria, values):
        assert minifs(values, values, criteria) == result

    @staticmethod
    @pytest.mark.parametrize(params, responses['sumifs'])
    def test_sumifs(result, criteria, values):
        assert sumifs(values, values, criteria) == result

    def test_ifs_size_errors(self):
        criteria, v1 = self.responses['sumifs'][0][1:]
        v2 = (v1[0][:-1], )
        assert countifs(v1, criteria, v2, criteria) == VALUE_ERROR
        assert sumifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR
        assert maxifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR
        assert minifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR
        assert averageifs(v1, v1, criteria, v2, criteria) == VALUE_ERROR


@pytest.mark.parametrize(
    'data, k, result', (
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
def test_large(data, k, result):
    assert result == large(data, k)


@pytest.mark.parametrize(
    'data, max_result, min_result', (
        ('abcd', 0, 0),
        ((2, None, 'x', 3), 3, 2),
        ((-0.1, None, 'x', True), -0.1, -0.1),
        (VALUE_ERROR, VALUE_ERROR, VALUE_ERROR),
        ((2, VALUE_ERROR), VALUE_ERROR, VALUE_ERROR),
        (DIV0, DIV0, DIV0),
        ((2, DIV0), DIV0, DIV0),
    )
)
def test_max_min(data, max_result, min_result):
    assert max_(data) == max_result
    assert min_(data) == min_result


@pytest.mark.parametrize(
    'data, k, result', (
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
def test_small(data, k, result):
    assert result == small(data, k)
