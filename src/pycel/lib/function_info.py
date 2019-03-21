from pycel.lib.function_info_data import function_info

# set of all function names
all_excel_functions = frozenset(f.name for f in function_info)

# function name to excel version and function category
function_version = {f.name: f.version for f in function_info}
function_category = {f.name: f.category for f in function_info}


def func_status_msg(name):
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
    """Code to scrape web for list of functions"""
    import requests
    from bs4 import BeautifulSoup

    url = 'https://support.office.com/en-us/article/Excel-functions'\
          '-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188'

    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')

    # from .function_list_page import page_html
    # soup = BeautifulSoup(page_html, 'html.parser')

    table = max(((len(table.find_all('tr')), table)
                 for table in soup.find_all('table')))[1]

    rows_data = []
    for row in table.find_all('tr'):
        p = tuple(p.text for p in row.find_all('p'))
        if ':' not in p[1]:
            continue
        name = p[0].strip().replace(' function', '')
        category, description = (x.strip() for x in p[1].split(':', 1))
        version = ''
        for img in row.find_all('img'):
            version = img['alt'].replace(' button', '')
        rows_data.append((name, category, version, description))

    print()
    print('function_info = (')
    for row in rows_data:
        print("    FunctionInfo('{}', '{}', '{}'),".format(*row))
    print(')')
