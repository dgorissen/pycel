import pytest

from pycel.lib.function_info import (
    all_excel_functions,
    function_version,
    function_category,
    func_status_msg,
)


def test_function_info():
    assert 'INDEX' in all_excel_functions
    assert function_version['INDEX'] == ''
    assert function_category['INDEX'] == 'Lookup and reference'


@pytest.mark.parametrize(
    'function, known, group, introduced', (
        ('ACOS', True, 'Math and trigonometry', ''),
        ('ACOT', True, 'Math and trigonometry', 'Excel 2013'),
        ('ACOU', False, '', ''),
    )
)
def test_func_status_msg(function, known, group, introduced):
    is_known, msg = func_status_msg(function)
    assert known == is_known
    assert group in msg
    assert ('not a known' in msg) != (function in all_excel_functions)

    if introduced:
        assert introduced in msg
    else:
        assert 'introduced' not in msg
