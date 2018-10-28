import collections
import re
import string


def is_range(address):
    return ':' in address


def has_sheet(address):
    return '!' in address


def split_range(rng):
    rng = rng.split("!")
    if len(rng) == 1:
        sheet, rng = None, rng[0]
    else:
        sheet, rng = rng

    if is_range(rng):
        start, end = rng.split(':')
    else:
        start, end = rng, None

    return sheet, start, end


def split_address(address):
    sheet = None
    if has_sheet(address):
        sheet, address = address.split('!')

    # ignore case
    address = address.upper()

    # regular <col><row> format
    if re.match(r'^[A-Z\$]+[\d\$]+$', address):
        col, row = [_f for _f in re.split(r'([A-Z\$]+)', address) if _f]

    # R<row>C<col> format
    elif re.match(r'^R\d+C\d+$', address):
        row, col = address.split('C')
        row = row[1:]

    # R[<row>]C[<col>] format
    elif re.match(r'^R\[\d+\]C\[\d+\]$', address):
        row, col = address.split('C')
        row = row[2:-1]
        col = col[2:-1]

    else:
        raise Exception('Invalid address format: {}'.format(address))

    return sheet, col, row


def resolve_range(rng, flatten=False, sheet=''):
    sh, start, end = split_range(rng)

    if sh and sheet:
        if sh != sheet:
            raise Exception("Mismatched sheets %s and %s" % (sh, sheet))
        else:
            sheet += '!'
    elif sh and not sheet:
        sheet = sh + "!"
    elif sheet and not sh:
        sheet += "!"
    else:
        pass

    # single cell, no range
    if not is_range(rng):
        return [sheet + rng], 1, 1

    sh, start_col, start_row = split_address(start)
    sh, end_col, end_row = split_address(end)
    start_col_idx = col2num(start_col)
    end_col_idx = col2num(end_col)

    start_row = int(start_row)
    end_row = int(end_row)

    # single column
    if start_col == end_col:
        nrows = end_row - start_row + 1
        data = ["%s%s%s" % (s, c, r) for (s, c, r) in
                zip([sheet] * nrows, [start_col] * nrows,
                    list(range(start_row, end_row + 1)))]
        return data, len(data), 1

    # single row
    elif start_row == end_row:
        ncols = end_col_idx - start_col_idx + 1
        data = ["%s%s%s" % (s, num2col(c), r) for (s, c, r) in
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
                row.append(sheet + num2col(c) + str(r))

            cells.append(row)

        if flatten:
            # flatten into one list
            l = flatten(cells)
            return l, 1, len(l)
        else:
            return cells, len(cells), len(cells[0])

        # e.g., convert BA -> 53


def col2num(col):
    if not col:
        raise Exception("Column may not be empty")

    tot = 0
    for i, c in enumerate(c for c in col[::-1] if c != "$"):
        if c == '$':
            continue
        tot += (ord(c) - 64) * 26 ** i
    return tot


# convert back
def num2col(num):
    if num < 1:
        raise Exception("Number must be larger than 0: %s" % num)

    s = ''
    q = num
    while q > 0:
        (q, r) = divmod(q, 26)
        if r == 0:
            q = q - 1
            r = 26
        s = string.ascii_uppercase[r - 1] + s
    return s


def address2index(a):
    sh, c, r = split_address(a)
    return col2num(c), int(r)


def index2addres(c, r, sheet=None):
    return "%s%s%s" % (sheet + "!" if sheet else "", num2col(c), r)


def get_linest_degree(excel, cl):
    # TODO: assumes a row or column of linest formulas &
    # that all coefficients are needed

    sh, c, r, ci = cl.address_parts()
    # figure out where we are in the row

    # to the left
    i = ci - 1
    while i > 0:
        f = excel.get_formula_from_range(index2addres(i, r))
        if f is None or f != cl.formula:
            break
        else:
            i = i - 1

    # to the right
    j = ci + 1
    while True:
        f = excel.get_formula_from_range(index2addres(j, r))
        if f is None or f != cl.formula:
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
            f = excel.get_formula_from_range("%s%s" % (c, i))
            if f is None or f != cl.formula:
                break
            else:
                i = i - 1

        # down
        j = r + 1
        while True:
            f = excel.get_formula_from_range("%s%s" % (c, j))
            if f is None or f != cl.formula:
                break
            else:
                j = j + 1

        degree = (j - i - 1) - 1
        coef = r - i

    # if degree is zero -> only one linest formula
    # linear regression -> degree should be one
    return max(degree, 1), coef


def flatten(items):
    for el in items:
        if isinstance(el, collections.Iterable) and not isinstance(el, str):
            yield from flatten(el)
        else:
            yield el


def uniqueify(seq):
    seen = set()
    return [x for x in seq if x not in seen and not seen.add(x)]


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
    if not is_number(year) or not is_number(month):
        raise TypeError("All inputs must be a number")
    if year <= 0 or month <= 0:
        raise TypeError("All inputs must be strictly positive")

    if month in (4, 6, 9, 11):
        return 30
    elif month == 2:
        if is_leap_year(year):
            return 29
        else:
            return 28
    else:
        return 31


def normalize_year(y, m, d):
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
        if m in (4, 6, 9, 11) and d > 30:
            m += 1
            d -= 30
            y, m, d = normalize_year(y, m, d)
        elif m == 2:
            if (is_leap_year(y)) and d > 29:
                m += 1
                d -= 29
                y, m, d = normalize_year(y, m, d)
            elif (not is_leap_year(y)) and d > 28:
                m += 1
                d -= 28
                y, m, d = normalize_year(y, m, d)
        elif d > 31:
            m += 1
            d -= 31
            y, m, d = normalize_year(y, m, d)

    return y, m, d


def date_from_int(nb):
    if not is_number(nb):
        raise TypeError("%s is not a number" % str(nb))

    # origin of the Excel date system
    current_year = 1900
    current_month = 0
    current_day = 0

    while nb > 0:
        if not is_leap_year(current_year) and nb > 365:
            current_year += 1
            nb -= 365
        elif is_leap_year(current_year) and nb > 366:
            current_year += 1
            nb -= 366
        else:
            current_month += 1
            max_days = get_max_days_in_month(current_month, current_year)

            if nb > max_days:
                nb -= max_days
            else:
                current_day = nb
                nb = 0

    return current_year, current_month, current_day


def criteria_parser(criteria):
    if is_number(criteria):
        def check(x):
            return x == criteria  # and type(x) == type(criteria)

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
        raise Exception("Couldn't parse criteria: {}".format(criteria))

    return check


def find_corresponding_index(rng, criteria):
    # parse criteria
    check = criteria_parser(criteria)

    valid = []

    for index, item in enumerate(rng):
        if check(item):
            valid.append(index)

    return valid
