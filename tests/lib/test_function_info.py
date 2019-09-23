# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import pytest

from pycel.lib.function_info import (
    all_excel_functions,
    func_status_msg,
    function_category,
    function_version,
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
