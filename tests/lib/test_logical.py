import pytest

from pycel.excelutil import (
    DIV0,
    ERROR_CODES,
    NA_ERROR,
    VALUE_ERROR,
)

from pycel.lib.logical import (
    iferror,
    x_if,
)


def test_iferror():
    assert 'A' == iferror('A', 2)

    for error in ERROR_CODES:
        assert 2 == iferror(error, 2)


@pytest.mark.parametrize(
    'test_value, true_value, false_value, result', (
        ('xyzzy', 3, 2, VALUE_ERROR),
        ('0', 2, 1, VALUE_ERROR),
        (True, 2, 1, 2),
        (False, 2, 1, 1),
        ('True', 2, 1, 2),
        ('False', 2, 1, 1),
        (None, 2, 1, 1),
        (NA_ERROR, 0, 0, NA_ERROR),
        (DIV0, 0, 0, DIV0),
        (1, VALUE_ERROR, 1, VALUE_ERROR),
        (0, VALUE_ERROR, 1, 1),
        (0, 1, VALUE_ERROR, VALUE_ERROR),
        (1, 1, VALUE_ERROR, 1),
    )
)
def test_x_if(test_value, true_value, false_value, result):
    assert x_if(test_value, true_value, false_value) == result
