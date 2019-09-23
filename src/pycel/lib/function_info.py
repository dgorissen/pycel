# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import os

from pycel.lib.function_info_data import function_info

# set of all function names
all_excel_functions = frozenset(f.name for f in function_info)

# function name to excel version and function category
function_version = {f.name: f.version for f in function_info}
function_category = {f.name: f.category for f in function_info}
base_url = 'https://support.office.com/en-us/article/'


def func_status_msg(name):
    """Return a string with info about an excel function"""
    name = name.upper()
    known = name in all_excel_functions
    if known:
        msg = '{} is in the "{}" group'.format(name, function_category[name])
        version = function_version[name]
        if version:
            msg += ', and was introduced in {}'.format(version)
    else:
        msg = '{} is not a known Excel function'.format(name)
    return known, msg


def scrape_function_list():  # pragma: no cover
    """Development Code to scrape web for list of excel functions
        builds: function_info_data.py
    """
    import requests
    from bs4 import BeautifulSoup

    base_dir = os.path.dirname(__file__)
    tmp_data_name = 'tmp_function_list_page'

    from_web = True
    if from_web:
        url = base_url + 'Excel-functions-alphabetical-' \
                         'b3944572-255d-4efb-bb96-c6d90033e188'

        page = requests.get(url)
        soup = BeautifulSoup(page.text, 'html.parser')

        # temporarily save page for further testing
        tmp_data_py = os.path.join(base_dir, tmp_data_name + '.py')
        with open(tmp_data_py, 'wb') as f:
            f.write('page_html = """{}"""'.format(page.text).encode('utf-8'))

    else:
        import importlib
        web_data = importlib.import_module('pycel.lib.{}'.format(tmp_data_name))
        soup = BeautifulSoup(web_data.page_html, 'html.parser')

    table = max(((len(table.find_all('tr')), table)
                 for table in soup.find_all('table')))[1]

    rows_data = []
    for row in table.find_all('tr'):
        p = tuple(p.text for p in row.find_all('p'))
        if ':' not in p[1]:
            continue
        href = tuple(a.get('href') for a in row.find_all('a'))
        row_url = href[-1]
        name = p[0].strip().replace(' function', '')
        category, description = (x.strip() for x in p[1].split(':', 1))
        version = ''
        for img in row.find_all('img'):
            version = img['alt'].replace(' button', '')
        rows_data.append((name, category, version, row_url, description))

    with open(os.path.join(base_dir, 'function_info_data.py'), 'w') as f:
        f.write("import collections\n\n")
        f.write("FunctionInfo = collections.namedtuple(\n")
        f.write("    'FunctionInfo', 'name category version uri')\n\n")
        f.write('function_info = (\n')
        for row_data in rows_data:
            for name in row_data[0].split(','):
                f.write("    FunctionInfo('{}', '{}', '{}', '{}'),\n".format(
                    name.strip().rstrip('s'), *row_data[1:]))
        f.write(')\n')


def print_function_header():  # pragma: no cover
    """Development Code to generate sample function header stubs"""
    from .function_info_data import function_info
    print()
    for row in function_info:
        if row[1].startswith('Math'):
            print("# def {}(value):".format(row.name.lower()))
            print("    # Excel reference: {}".format(base_url))
            print("    #   {}".format(row.uri))
            print()
            print()
