import importlib
import math
import pytest

from pycel.lib.function_helpers import cse_array_wrapper, load_functions


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


def test_load_functions():

    modules = (
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
