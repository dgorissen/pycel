import calendar
import collections
import datetime as dt
import re

from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    quote_sheetname,
)


def is_range(address):
    return ':' in address


def has_sheet(address):
    return '!' in address


def get_sheet(address, sheet=''):
    sh = ''
    if has_sheet(address):
        sh, address = address.split('!')
        if sh.startswith("'") and sh.endswith("'"):
            sh = sh[1:-1]
        sh = sh.replace("''", "'")

    if sh and sheet:
        if sh != sheet:
            raise Exception("Mismatched sheets %s and %s" % (sh, sheet))

    if sh or sheet:
        sheet = str(sh or sheet)

    return sheet, address


def split_range(rng, sheet=''):
    sheet, rng = get_sheet(rng, sheet=sheet)

    if is_range(rng):
        start, end = rng.split(':')
    else:
        start, end = rng, None

    return sheet, start, end


def split_address(address, sheet=''):
    sheet, start, end = split_range(address, sheet=sheet)

    if end is not None:
        raise Exception('Found range {} expected address'.format(address))

    # ignore case
    address = start.upper()

    # regular <col><row> format
    if re.match(r'^[A-Z\$]+[\d\$]+$', address):
        col, row = [_f for _f in re.split(r'([A-Z\$]+)', address) if _f]

    # R<row>C<col> format
    elif re.match(r'^R\d+C\d+$', address):
        row, col = address.split('C')
        row = str(row[1:])
        col = num2col(int(col))

    # R[<row>]C[<col>] format
    elif re.match(r'^R\[\d+\]C\[\d+\]$', address):
        row, col = address.split('C')
        row = str(row[2:-1])
        col = num2col(int(col[1:-1]))

    else:
        raise Exception('Invalid address format: {}'.format(address))

    return sheet, col, row


def resolve_range(rng, sheet=''):
    sheet, start, end = split_range(rng, sheet=sheet)

    if sheet:
        sheet += '!'
        
    # single cell, no range
    if not is_range(rng):
        return [sheet + start], 1, 1

    sh, start_col, start_row = split_address(start)
    sh, end_col, end_row = split_address(end)
    start_col_idx = col2num(start_col)
    end_col_idx = col2num(end_col)

    start_row = int(start_row)
    end_row = int(end_row)

    # single column
    if start_col == end_col:
        nrows = end_row - start_row + 1
        data = [index2address(c, r, s) for (s, c, r) in
                zip([sheet] * nrows, [start_col] * nrows,
                    list(range(start_row, end_row + 1)))]
        return data, len(data), 1

    # single row
    elif start_row == end_row:
        ncols = end_col_idx - start_col_idx + 1
        data = [index2address(c, r, s) for (s, c, r) in
                zip([sheet] * ncols,
                    list(range(start_col_idx, end_col_idx + 1)),
                    [start_row] * ncols)]
        return data, 1, len(data)

    # rectangular range
    else:
        cells = []
        for r in range(start_row, end_row + 1):
            row = []
            for c in range(start_col_idx, end_col_idx + 1):
                row.append(index2address(c, r, sheet))

            cells.append(row)

        return cells, len(cells), len(cells[0])


def col2num(column):
    # e.g., convert BA -> 53
    return column_index_from_string(column)


def num2col(column_number):
    # convert back 53 -> BA
    return get_column_letter(column_number)


def address2index(a, sheet=''):
    sh, c, r = split_address(a, sheet=sheet)
    return col2num(c), int(r)


def index2address(c, r, sheet=''):
    if isinstance(c, int):
        c = get_column_letter(c)

    if sheet:
        sheet = '{}!'.format(quote_sheetname(sheet.strip('!')))

    return "{}{}{}".format(sheet, c, r)


def get_linest_degree(cell):
    # TODO: assumes a row or column of linest formulas &
    # that all coefficients are needed

    sh, c, r, ci = cell.address_parts()
    # figure out where we are in the row

    # to the left
    i = ci - 1
    while i > 0:
        f = cell.excel.get_formula_from_range(index2address(i, r))
        if f is None or f != cell.formula:
            break
        else:
            i = i - 1

    # to the right
    j = ci + 1
    while True:
        f = cell.excel.get_formula_from_range(index2address(j, r))
        if f is None or f != cell.formula:
            break
        else:
            j = j + 1

    # assume the degree is the number of linest's
    # last -1 is because an n degree polynomial has n+1 coefs
    degree = (j - i - 1) - 1

    # which coef are we (left most coef is the coef for the highest power)
    coef = ci - i

    # no linests left or right, try looking up/down
    if degree == 0:
        # up
        i = r - 1
        while i > 0:
            f = cell.excel.get_formula_from_range("%s%s" % (c, i))
            if f is None or f != cell.formula:
                break
            else:
                i = i - 1

        # down
        j = r + 1
        while True:
            f = cell.excel.get_formula_from_range("%s%s" % (c, j))
            if f is None or f != cell.formula:
                break
            else:
                j = j + 1

        degree = (j - i - 1) - 1
        coef = r - i

    # if degree is zero -> only one linest formula
    # linear regression -> degree should be one
    return max(degree, 1), coef


def flatten(items):
    for item in items:
        if isinstance(item, collections.Iterable) and not isinstance(item, str):
            yield from flatten(item)
        else:
            yield item


def uniqueify(seq):
    seen = set()
    return tuple(x for x in seq if x not in seen and not seen.add(x))


def is_number(s):
    try:
        float(s)
        return True
    except (ValueError, TypeError):
        return False


def is_leap_year(year):
    if not is_number(year):
        raise TypeError("%s must be a number" % str(year))
    if year <= 0:
        raise TypeError("%s must be strictly positive" % str(year))

    # Watch out, 1900 is a leap according to Excel =>
    # https://support.microsoft.com/en-us/kb/214326
    return year % 4 == 0 and year % 100 != 0 or year % 400 == 0 or year == 1900


def get_max_days_in_month(month, year):
    if month == 2 and is_leap_year(year):
        return 29

    return calendar.monthrange(year, month)[1]


def normalize_year(y, m, d):
    """taking into account negative month and day values"""
    if m <= 0:
        y -= int(abs(m) / 12 + 1)
        m = 12 - (abs(m) % 12)
        normalize_year(y, m, d)
    elif m > 12:
        y += int(m / 12)
        m = m % 12

    if d <= 0:
        d += get_max_days_in_month(m, y)
        m -= 1
        y, m, d = normalize_year(y, m, d)

    else:
        days_in_month = get_max_days_in_month(m, y)
        if d > days_in_month:
            m += 1
            d -= days_in_month
            y, m, d = normalize_year(y, m, d)

    return y, m, d


def date_from_int(datestamp):

    if datestamp == 31 + 29:
        # excel thinks 1900 is a leap year
        return 1900, 2, 29

    date = dt.datetime(1899, 12, 30) + dt.timedelta(days=datestamp)
    if datestamp < 31 + 29:
        date += dt.timedelta(days=1)

    return date.year, date.month, date.day


def criteria_parser(criteria):
    if is_number(criteria):
        def check(x):
            return x == float(criteria)

    elif type(criteria) == str:

        search = re.search(r'(\W*)(.*)', criteria.lower()).group
        operator = search(1)
        value = search(2)
        value = float(value) if is_number(value) else str(value)

        def test_is_number(x):
            if not is_number(x):
                raise TypeError('excellib.countif() doesnt\'t work for checking'
                                ' non number items against non equality')

        if operator == '<':
            def check(x):
                test_is_number(x)
                return x < value
        elif operator == '>':
            def check(x):
                test_is_number(x)
                return x > value
        elif operator == '>=':
            def check(x):
                test_is_number(x)
                return x >= value
        elif operator == '<=':
            def check(x):
                test_is_number(x)
                return x <= value
        elif operator == '<>':
            def check(x):
                test_is_number(x)
                return x != value
        else:
            def check(x):
                return x == criteria
    else:
        raise ValueError("Couldn't parse criteria: {}".format(criteria))

    return check


def find_corresponding_index(rng, criteria):
    # parse criteria
    check = criteria_parser(criteria)

    valid = []

    for index, item in enumerate(rng):
        if check(item):
            valid.append(index)

    return valid
