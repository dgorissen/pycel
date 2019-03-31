from unittest import mock
import importlib
import math
import pytest


from pycel.excelutil import DIV0, NUM_ERROR, VALUE_ERROR

from pycel.lib.function_helpers import (
    cse_array_wrapper,
    error_string_wrapper,
    load_functions,
    math_wrapper,
)


DATA = (
    (1, 2),
    (3, 4),
)


@pytest.mark.parametrize(
    'arg_num, f_args, result', (
        (0, (1,), 2),
        (1, (0, 1), 2),
        (0, (DATA, ), ((2, 3), (4, 5))),
        (1, (1, DATA), ((2, 3), (4, 5))),
    )
)
def test_cse_array_wrapper(arg_num, f_args, result):

    def f_test(*args):
        return args[arg_num] + 1

    assert cse_array_wrapper(f_test, arg_num)(*f_args) == result


@pytest.mark.parametrize(
    'arg_nums, f_args, result', (
        ((0, 1), (DIV0, 1), DIV0),
        ((0, 1), (1, DIV0), DIV0),
        ((0, 1), (NUM_ERROR, DIV0), NUM_ERROR),
        ((0, 1), (DIV0, NUM_ERROR), DIV0),
        ((0,), (1, DIV0), "args: (1, '#DIV/0!')"),
        ((1,), (1, DIV0), DIV0),
    )
)
def test_error_string_wrapper(arg_nums, f_args, result):

    def f_test(*args):
        return 'args: {}'.format(args)

    esf_name = 'pycel.lib.function_helpers.error_string_functions'
    with mock.patch(esf_name, {'f_test': arg_nums}):
        assert error_string_wrapper(f_test)(*f_args) == result

    assert error_string_wrapper(f_test, arg_nums)(*f_args) == result


@pytest.mark.parametrize(
    'value, result', (
        (1, 1),
        (DIV0, DIV0),
        (None, 0),
        ('1.1', 1.1),
        ('xyzzy', VALUE_ERROR),
    )
)
def test_math_wrap(value, result):
    assert math_wrapper(lambda x: x)(value) == result


def test_math_wrap_domain_error():
    assert math_wrapper(math.log)(-1) == NUM_ERROR


def test_load_functions():

    modules = (
        importlib.import_module('pycel.excellib'),
        importlib.import_module('pycel.lib.logical'),
        importlib.import_module('math'),
    )

    namespace = locals()

    names = 'degrees x_if junk'.split()
    missing = load_functions(names, namespace, modules)
    assert missing == {'junk'}
    assert 'degrees' in namespace
    assert 'x_if' in namespace

    names = 'radians x_if junk'.split()
    missing = load_functions(names, namespace, modules)
    assert missing == {'junk'}
    assert 'radians' in namespace

    assert namespace['radians'](180) == math.pi
    assert namespace['radians'](((180, 360),)) == ((math.pi, 2 * math.pi),)

    assert namespace['x_if'](0, 'Y', 'N') == 'N'
    assert namespace['x_if'](((0, 1),), 'Y', 'N') == (('N', 'Y'),)

    missing = load_functions(['index'], namespace, modules)
    assert namespace['index'](DIV0, 1, 1) == DIV0
