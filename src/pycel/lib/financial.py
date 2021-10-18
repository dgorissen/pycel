# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of Excel financial functions
"""
import numpy_financial as npf

from pycel.excelutil import flatten

def irr(values, guess=None):
    # Returns the internal rate of return for the cash flow 'values'
    return npf.irr(list(flatten(values)))


def pmt(rate, nper, pv, fv=0, when=0):
    # Returns the payment for a loan given a constant interest 'rate', total
    # number of payments 'nper', and the present value 'pv'
    return npf.pmt(rate, nper, pv, fv=fv, when=when)


def ppmt(rate, per, nper, pv, fv=0, when=0):  # pylint: disable=too-many-arguments
    # Returns the payment on the principal for a loan for 'per' pay periods
    return npf.ppmt(rate, per, nper, pv, fv=fv, when=when)