# -*- coding: UTF-8 -*-
#
# Copyright 2011-2021 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

"""
Python equivalents of Statistics Excel functions
"""
import math
from heapq import nlargest, nsmallest

import numpy as np

from pycel.excellib import _numerics
from pycel.excelutil import (
    coerce_to_number,
    DIV0,
    ERROR_CODES,
    find_corresponding_index,
    flatten,
    handle_ifs,
    list_like,
    NA_ERROR,
    NUM_ERROR,
    REF_ERROR,
    VALUE_ERROR,
)
from pycel.lib.function_helpers import (
    excel_helper,
)


# Boolean, unsigned integer, signed integer, float, complex.
_NP_NUMERIC_KINDS = set('buifc')


def _slope_intercept(Y, X):
    """Groom linest results for SLOPE(), INTERCEPT() and FORECAST()"""
    try:
        coefs, full_rank = linest_helper(Y, X)
    except AssertionError:
        return NA_ERROR
    except ValueError:
        return VALUE_ERROR

    if len(coefs) != 2:
        return NA_ERROR
    if not full_rank:
        return DIV0
    return coefs


# def avedev(value):
#     # Excel reference: https://support.microsoft.com/en-us/office/
#     #   avedev-function-58fe8d65-2a84-4dc7-8052-f3f87b5c6639


def average(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   average-function-047bac88-d466-426c-a32b-8f33eb960cf6
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data
    elif len(data) == 0:
        return DIV0
    else:
        return sum(data) / len(data)


# def averagea(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   averagea-function-f5f84098-d453-4f4c-bbba-3d2c66356091


def averageif(rng, criteria, average_range=None):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   averageif-function-faec8e2e-0dec-4308-af69-f5576d8ac642

    # WARNING:
    # - The following is not currently implemented:
    #  The average_range argument does not have to be the same size and shape
    #  as the range argument. The actual cells that are added are determined by
    #  using the upper leftmost cell in the average_range argument as the
    #  beginning cell, and then including cells that correspond in size and
    #  shape to the range argument.
    if average_range is None:
        average_range = rng
    return averageifs(average_range, rng, criteria)


def averageifs(average_range, *args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   AVERAGEIFS-function-48910C45-1FC0-4389-A028-F7C5C3001690
    if not list_like(average_range):
        average_range = ((average_range, ), )

    coords = handle_ifs(args, average_range)

    # A returned string is an error code
    if isinstance(coords, str):
        return coords

    data = _numerics((average_range[r][c] for r, c in coords), keep_bools=True)
    if len(data) == 0:
        return DIV0
    return sum(data) / len(data)


# def beta.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   beta-dist-function-11188c9c-780a-42c7-ba43-9ecb5a878d31


# def beta.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   beta-inv-function-e84cb8aa-8df0-4cf6-9892-83a341d252eb


# def binom.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   binom-dist-function-c5ae37b6-f39c-4be2-94c2-509a1480770c


# def binom.dist.range(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   binom-dist-range-function-17331329-74c7-4053-bb4c-6653a7421595


# def binom.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   binom-inv-function-80a0370c-ada6-49b4-83e7-05a91ba77ac9


# def chisq.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   chisq-dist-function-8486b05e-5c05-4942-a9ea-f6b341518732


# def chisq.dist.rt(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   chisq-dist-rt-function-dc4832e8-ed2b-49ae-8d7c-b28d5804c0f2


# def chisq.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   chisq-inv-function-400db556-62b3-472d-80b3-254723e7092f


# def chisq.inv.rt(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   chisq-inv-rt-function-435b5ed8-98d5-4da6-823f-293e2cbc94fe


# def chisq.test(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   chisq-test-function-2e8a7861-b14a-4985-aa93-fb88de3f260f


# def confidence.norm(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   confidence-norm-function-7cec58a6-85bb-488d-91c3-63828d4fbfd4


# def confidence.t(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   confidence-t-function-e8eca395-6c3a-4ba9-9003-79ccc61d3c53


# def correl(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   correl-function-995dcef7-0c0a-4bed-a3fb-239d7b68ca92


def count(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   COUNT-function-a59cd7fc-b623-4d93-87a4-d23bf411294c

    return sum(1 for x in flatten(args)
               if isinstance(x, (int, float)) and not isinstance(x, bool))


# def counta(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   counta-function-7dc98875-d5c1-46f1-9a82-53f3219e2509


# def countblank(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   countblank-function-6a92d772-675c-4bee-b346-24af6bd3ac22


def countif(rng, criteria):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   COUNTIF-function-e0de10c6-f885-4e71-abb4-1f464816df34
    if not list_like(rng):
        rng = ((rng, ), )
    valid = find_corresponding_index(rng, criteria)
    return len(valid)


def countifs(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   COUNTIFS-function-dda3dc6e-f74e-4aee-88bc-aa8c2a866842
    coords = handle_ifs(args)

    # A returned string is an error code
    if isinstance(coords, str):
        return coords

    return len(coords)


# def covariance.p(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   covariance-p-function-6f0e1e6d-956d-4e4b-9943-cfef0bf9edfc


# def covariance.s(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   covariance-s-function-0a539b74-7371-42aa-a18f-1f5320314977


# def devsq(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   devsq-function-8b739616-8376-4df5-8bd0-cfe0a6caf444


# def expon.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   expon-dist-function-4c12ae24-e563-4155-bf3e-8b78b6ae140e


# def f.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   f-dist-function-a887efdc-7c8e-46cb-a74a-f884cd29b25d


# def f.dist.rt(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   f-dist-rt-function-d74cbb00-6017-4ac9-b7d7-6049badc0520


# def f.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   f-inv-function-0dda0cf9-4ea0-42fd-8c3c-417a1ff30dbe


# def f.inv.rt(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   f-inv-rt-function-d371aa8f-b0b1-40ef-9cc2-496f0693ac00


# def fisher(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   fisher-function-d656523c-5076-4f95-b87b-7741bf236c69


# def fisherinv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   fisherinv-function-62504b39-415a-4284-a285-19c8e82f86bb


@excel_helper(number_params=(0))
def forecast(x, Y, X):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   forecasting-functions-reference-897a2fe9-6595-4680-a0b0-93e0308d5f6e
    coefs = _slope_intercept(Y, X)
    if coefs in ERROR_CODES:
        return coefs
    return coefs[0] * x + coefs[1]


# def forecast.ets(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   forecasting-functions-reference-897a2fe9-6595-4680-a0b0-93e0308d5f6e#_forecast.ets


# def forecast.ets.confint(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   forecasting-functions-reference-897a2fe9-6595-4680-a0b0-93e0308d5f6e#_forecast.ets.confint


# def forecast.ets.seasonality(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   forecasting-functions-reference-897a2fe9-6595-4680-a0b0-93e0308d5f6e#_forecast.ets.seasonality


# def forecast.ets.stat(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   forecasting-functions-reference-897a2fe9-6595-4680-a0b0-93e0308d5f6e#_forecast.ets.stat


# def forecast.linear(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   forecasting-functions-reference-897a2fe9-6595-4680-a0b0-93e0308d5f6e#_forecast.linear


# def frequency(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   frequency-function-44e3be2b-eca0-42cd-a3f7-fd9ea898fdb9


# def f.test(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   f-test-function-100a59e7-4108-46f8-8443-78ffacb6c0a7


# def gamma(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   gamma-function-ce1702b1-cf55-471d-8307-f83be0fc5297


# def gamma.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   gamma-dist-function-9b6f1538-d11c-4d5f-8966-21f6a2201def


# def gamma.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   gamma-inv-function-74991443-c2b0-4be5-aaab-1aa4d71fbb18


# def gammaln(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   gammaln-function-b838c48b-c65f-484f-9e1d-141c55470eb9


# def gammaln.precise(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   gammaln-precise-function-5cdfe601-4e1e-4189-9d74-241ef1caa599


# def gauss(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   gauss-function-069f1b4e-7dee-4d6a-a71f-4b69044a6b33


# def geomean(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   geomean-function-db1ac48d-25a5-40a0-ab83-0b38980e40d5


# def growth(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   growth-function-541a91dc-3d5e-437d-b156-21324e68b80d


# def harmean(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   harmean-function-5efd9184-fab5-42f9-b1d3-57883a1d3bc6


# def hypgeom.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   hypgeom-dist-function-6dbd547f-1d12-4b1f-8ae5-b0d9e3d22fbf


@excel_helper()
def intercept(Y, X):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   intercept-function-2a9b74e2-9d47-4772-b663-3bca70bf63ef
    coefs = _slope_intercept(Y, X)
    if coefs in ERROR_CODES:
        return coefs
    return coefs[1]


# def kurt(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   kurt-function-bc3a265c-5da4-4dcb-b7fd-c237789095ab


@excel_helper()
def large(array, k):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   large-function-3af0af19-1190-42bb-bb8b-01672ec00a64
    data = _numerics(array, to_number=coerce_to_number)
    if isinstance(data, str):
        return data

    k = coerce_to_number(k)
    if isinstance(k, str):
        return VALUE_ERROR

    if not data or k is None or k < 1 or k > len(data):
        return NUM_ERROR

    k = math.ceil(k)
    return nlargest(k, data)[-1]


def linest_helper(Y, X=None, const=True, stats=False):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   linest-function-84d7d0d9-6e50-4101-977a-fa7abf772b6d
    """Perform an OLS model fir

    :param Y: Vector of output data
    :param X: Input Data
    :param const: force the intercept to zero
    :param stats: Out extended statistics
    :return:  numpy.linalg.lstsq
        https://numpy.org/doc/stable/reference/generated/numpy.linalg.lstsq.html
    """
    Y = np.array(Y)
    assert 1 in Y.shape
    Y = Y.ravel()

    if X is None:
        length = len(Y)
        X = np.resize(np.repeat(np.arange(1, length + 1), 1), (length, 1))
    else:
        X = np.array(X)
        assert len(Y) in X.shape
        if X.shape[0] != len(Y):
            X = X.transpose()

    for data in (X, Y):
        if data.dtype.kind not in _NP_NUMERIC_KINDS:
            raise ValueError

    if const:
        # add a constant column
        A = np.hstack((np.ones((len(Y), 1)), X))
    else:
        # force the intercept to zero if no const desired
        A = X

    # perform the fit
    coefs, residuals, rank, sing_vals = np.linalg.lstsq(A, Y, rcond=None)
    full_rank = (rank == len(coefs))
    result_coefs = tuple(reversed(coefs if const else (0,) + tuple(coefs)))
    if not full_rank:
        result_coefs = (0,) * (len(result_coefs) - 1) + (np.sum(Y) / len(Y),)

    if stats:
        # Compute some extended stats as Excel does
        Y_predicted = A @ coefs
        if const:
            sum_sq_regression = np.sum((Y_predicted - np.sum(Y) / len(Y)) ** 2)
            sum_sq_total = (len(Y) - 1) * np.var(Y, ddof=1)
        else:
            # https://stats.stackexchange.com/a/26205/154402
            sum_sq_regression = np.sum(Y_predicted ** 2)
            sum_sq_total = np.sum(Y ** 2)
        sum_sq_resid = np.sum((Y - Y_predicted) ** 2)
        r2_score = 1 - (sum_sq_resid / sum_sq_total)

        # standard error stats
        try:
            stderr_y_2 = (1 / (len(Y) - len(coefs))) * (Y_predicted - Y) @ (Y_predicted - Y).T
            stderr_y = np.sqrt(stderr_y_2)
            std_err = tuple(reversed(np.sqrt((stderr_y_2 * np.linalg.inv(A.T @ A)).diagonal())))
        except (ZeroDivisionError, np.linalg.LinAlgError):
            stderr_y = 0
            std_err = (0,) * len(result_coefs)
            r2_score = 1
            sum_sq_regression = result_coefs[-1]
            sum_sq_resid = 0

        if len(std_err) < len(result_coefs):
            std_err += (NA_ERROR,) * (len(result_coefs) - len(std_err))

        # F and dof stats
        dof = len(Y) - len(coefs)
        denom = (len(result_coefs) - 1) * (1 - r2_score)
        f_score = NUM_ERROR if denom == 0 else r2_score * dof / denom

        na_filler = (NA_ERROR,) * max(0, len(result_coefs) - 2)
        return (
            result_coefs,
            std_err,
            (r2_score, stderr_y, *na_filler),
            (f_score, dof, *na_filler),
            (sum_sq_regression, sum_sq_resid, *na_filler),
        ), full_rank
    else:
        return result_coefs, full_rank


def linest(Y, X=None, const=None, stats=None):
    kwargs = {}
    if const is not None:
        kwargs['const'] = const
    if stats is not None:
        kwargs['stats'] = stats

    try:
        coefs, full_rank = linest_helper(Y, X, **kwargs)
    except AssertionError:
        return REF_ERROR
    except ValueError:
        return VALUE_ERROR

    if stats:
        return coefs
    else:
        return (coefs,)


# def logest(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   logest-function-f27462d8-3657-4030-866b-a272c1d18b4b


# def lognorm.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   lognorm-dist-function-eb60d00b-48a9-4217-be2b-6074aee6b070


# def lognorm.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   lognorm-inv-function-fe79751a-f1f2-4af8-a0a1-e151b2d4f600


def max_(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   max-function-e0012414-9ac8-4b34-9a47-73e662c08098
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data

    # however, if no non numeric cells, return zero (is what excel does)
    elif len(data) < 1:
        return 0
    else:
        return max(data)


# def maxa(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   maxa-function-814bda1e-3840-4bff-9365-2f59ac2ee62d


def maxifs(max_range, *args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   maxifs-function-dfd611e6-da2c-488a-919b-9b6376b28883
    if not list_like(max_range):
        max_range = ((max_range, ), )

    try:
        coords = handle_ifs(args, max_range)

        # A returned string is an error code
        if isinstance(coords, str):
            return coords

        return max(_numerics(
            (max_range[r][c] for r, c in coords),
            keep_bools=True
        ))
    except ValueError:
        return 0


# def median(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   median-function-d0916313-4753-414c-8537-ce85bdd967d2


def min_(*args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   min-function-61635d12-920f-4ce2-a70f-96f202dcc152
    data = _numerics(*args)

    # A returned string is an error code
    if isinstance(data, str):
        return data

    # however, if no non numeric cells, return zero (is what excel does)
    elif len(data) < 1:
        return 0
    else:
        return min(data)


def minifs(min_range, *args):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   minifs-function-6ca1ddaa-079b-4e74-80cc-72eef32e6599
    if not list_like(min_range):
        min_range = ((min_range, ), )

    try:
        coords = handle_ifs(args, min_range)

        # A returned string is an error code
        if isinstance(coords, str):
            return coords

        return min(_numerics(
            (min_range[r][c] for r, c in coords),
            keep_bools=True
        ))
    except ValueError:
        return 0


# def mina(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   mina-function-245a6f46-7ca5-4dc7-ab49-805341bc31d3


# def mode.mult(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   mode-mult-function-50fd9464-b2ba-4191-b57a-39446689ae8c


# def mode.sngl(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   mode-sngl-function-f1267c16-66c6-4386-959f-8fba5f8bb7f8


# def negbinom.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   negbinom-dist-function-c8239f89-c2d0-45bd-b6af-172e570f8599


# def norm.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   norm-dist-function-edb1cc14-a21c-4e53-839d-8082074c9f8d


# def norminv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   norminv-function-87981ab8-2de0-4cb0-b1aa-e21d4cb879b8


# def norm.s.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   norm-s-dist-function-1e787282-3832-4520-a9ae-bd2a8d99ba88


# def norm.s.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   norm-s-inv-function-d6d556b4-ab7f-49cd-b526-5a20918452b1


# def pearson(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   pearson-function-0c3e30fc-e5af-49c4-808a-3ef66e034c18


# def percentile.exc(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   percentile-exc-function-bbaa7204-e9e1-4010-85bf-c31dc5dce4ba


# def percentile.inc(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   percentile-inc-function-680f9539-45eb-410b-9a5e-c1355e5fe2ed


# def percentrank.exc(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   percentrank-exc-function-d8afee96-b7e2-4a2f-8c01-8fcdedaa6314


# def percentrank.inc(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   percentrank-inc-function-149592c9-00c0-49ba-86c1-c1f45b80463a


# def permut(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   permut-function-3bd1cb9a-2880-41ab-a197-f246a7a602d3


# def permutationa(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   permutationa-function-6c7d7fdc-d657-44e6-aa19-2857b25cae4e


# def phi(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   phi-function-23e49bc6-a8e8-402d-98d3-9ded87f6295c


# def poisson.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   poisson-dist-function-8fe148ff-39a2-46cb-abf3-7772695d9636


# def prob(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   prob-function-9ac30561-c81c-4259-8253-34f0a238fc49


# def quartile.exc(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   quartile-exc-function-5a355b7a-840b-4a01-b0f1-f538c2864cad


# def quartile.inc(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   quartile-inc-function-1bbacc80-5075-42f1-aed6-47d735c4819d


# def rank.avg(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   rank-avg-function-bd406a6f-eb38-4d73-aa8e-6d1c3c72e83a


# def rank.eq(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   rank-eq-function-284858ce-8ef6-450e-b662-26245be04a40


# def rsq(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   rsq-function-d7161715-250d-4a01-b80d-a8364f2be08f


# def skew(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   skew-function-bdf49d86-b1ef-4804-a046-28eaea69c9fa


# def skew.p(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   skew-p-function-76530a5c-99b9-48a1-8392-26632d542fcb


@excel_helper()
def slope(Y, X):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   slope-function-11fb8f97-3117-4813-98aa-61d7e01276b9
    coefs = _slope_intercept(Y, X)
    if coefs in ERROR_CODES:
        return coefs
    return coefs[0]


@excel_helper()
def small(array, k):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   small-function-17da8222-7c82-42b2-961b-14c45384df07
    data = _numerics(array, to_number=coerce_to_number)
    if isinstance(data, str):
        return data

    k = coerce_to_number(k)
    if isinstance(k, str):
        return VALUE_ERROR

    if not data or k is None or k < 1 or k > len(data):
        return NUM_ERROR

    k = math.ceil(k)
    return nsmallest(k, data)[-1]


# def standardize(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   standardize-function-81d66554-2d54-40ec-ba83-6437108ee775


# def stdev.p(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   stdev-p-function-6e917c05-31a0-496f-ade7-4f4e7462f285


# def stdev.s(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   stdev-s-function-7d69cf97-0c1f-4acf-be27-f3e83904cc23


# def stdeva(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   stdeva-function-5ff38888-7ea5-48de-9a6d-11ed73b29e9d


# def stdevpa(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   stdevpa-function-5578d4d6-455a-4308-9991-d405afe2c28c


# def steyx(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   steyx-function-6ce74b2c-449d-4a6e-b9ac-f9cef5ba48ab


# def t.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   t-dist-function-4329459f-ae91-48c2-bba8-1ead1c6c21b2


# def t.dist.2t(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   t-dist-2t-function-198e9340-e360-4230-bd21-f52f22ff5c28


# def t.dist.rt(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   t-dist-rt-function-20a30020-86f9-4b35-af1f-7ef6ae683eda


# def t.inv(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   t-inv-function-2908272b-4e61-4942-9df9-a25fec9b0e2e


# def t.inv.2t(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   t-inv-2t-function-ce72ea19-ec6c-4be7-bed2-b9baf2264f17


@excel_helper(bool_params=3)
def trend(Y, X=None, new_X=None, const=None):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   trend-function-e2f135f0-8827-4096-9873-9a7cf7b51ef1
    kwargs = {}
    if const is not None:
        kwargs['const'] = const

    try:
        coefs, full_rank = linest_helper(Y, X, **kwargs)
    except AssertionError:
        return REF_ERROR
    except ValueError:
        return VALUE_ERROR

    if new_X is None:
        if X is not None:
            new_X = np.array(X)
        else:
            length = max(len(Y), len(Y[0]))
            width = len(coefs) - 1
            new_X = np.resize(np.repeat(np.arange(1, length + 1), width), (length, width))

    if list_like(new_X):
        new_X = np.array(new_X)
        if len(coefs) - 1 not in new_X.shape:
            return REF_ERROR

        if new_X.shape[1] != len(coefs) - 1:
            if full_rank:
                result = np.array(coefs[-2::-1]).transpose() @ new_X + coefs[-1]
            else:
                result = (coefs[-1],) * new_X.shape[1]
            return result[0] if len(result) == 1 else (tuple(result),)
        else:
            if full_rank:
                result = new_X @ np.array(coefs[-2::-1]) + coefs[-1]
            else:
                result = (coefs[-1],) * new_X.shape[0]
            return result[0] if len(result) == 1 else tuple((x,) for x in result)

    elif len(coefs) != 2:
        # new_X is a scaler, this needs to be a single coef fit
        return REF_ERROR
    elif not full_rank:
        return coefs[-1]
    else:
        return coefs[0] * new_X + coefs[1]


# def trimmean(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   trimmean-function-d90c9878-a119-4746-88fa-63d988f511d3


# def t.test(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   t-test-function-d4e08ec3-c545-485f-962e-276f7cbed055


# def var.p(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   var-p-function-73d1285c-108c-4843-ba5d-a51f90656f3a


# def var.s(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   var-s-function-913633de-136b-449d-813e-65a00b2b990b


# def vara(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   vara-function-3de77469-fa3a-47b4-85fd-81758a1e1d07


# def varpa(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   varpa-function-59a62635-4e89-4fad-88ac-ce4dc0513b96


# def weibull.dist(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   weibull-dist-function-4e783c39-9325-49be-bbc9-a83ef82b45db


# def z.test(value):
    # Excel reference: https://support.microsoft.com/en-us/office/
    #   z-test-function-d633d5a3-2031-4614-a016-92180ad82bee


# Older mappings for excel functions that match Python built-in and keywords
xmax = max_
xmin = min_
