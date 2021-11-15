# -*- coding: UTF-8 -*-
#
# Copyright 2011-2021 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import math
import pytest

from pycel.lib.financial import (
    irr,
    npv,
    pmt,
    ppmt
)


@pytest.mark.parametrize(
    'values, guess, expected',
    (
        ((-100, -50, 100, 200, 400), None, 0.671269),
        ((-70000, 12000, 15000, 18000, 21000, 26000), None, 0.086631),
        ((-70000, 12000, 15000, 18000, 21000), None, -0.021245),
        ((-70000, 12000, 15000), 0.10, -0.443507)
    )
)
def test_irr(values, guess, expected):
    assert math.isclose(irr(values, guess=guess), expected, abs_tol=1e-4)


@pytest.mark.parametrize(
    'rate, values, expected',
    (
        (0.1, (-10000, 3000, 4200, 6800), 1188.443412),
        (0.08, (-40000, 8000, 9200, 10000, 12000, 14500), 1779.686625)
    )
)
def test_npv(rate, values, expected):
    assert math.isclose(npv(rate, *values), expected, abs_tol=1e-2)


@pytest.mark.parametrize(
    'rate, nper, pv, fv, when, expected',
    (
        (0.05, 12, 100, 400, 0, -36.412705),
        (0.00667, 10, 10000, 0, 0, -1037.050788),
        (0.00667, 10, 10000, 0, 1, -1030.179490),
        (0.005, 216, 0, 50000, 0, -129.0811609)
    )
)
def test_pmt(rate, nper, pv, fv, when, expected):
    assert math.isclose(pmt(rate, nper, pv, fv, when), expected, abs_tol=1e-4)


@pytest.mark.parametrize(
    'rate, per, nper, pv, fv, when, expected',
    (
        (0.05, 12, 100, 400, 0, 0, -0.262118),
        (0.00833, 1, 24, 2000, 0, 0, -75.626160),
        (0.08, 10, 10, 200000, 0, 0, -27598.053460)
    )
)
def test_ppmt(rate, per, nper, pv, fv, when, expected):
    assert math.isclose(ppmt(rate, per, nper, pv, fv, when), expected, abs_tol=1e-4)
