import calendar
import collections
import datetime as dt
import re

from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)

R1C1_ROW_RE_STR = r"R(\[-?\d+\]|\d+)?"
R1C1_COL_RE_STR = r"C(\[-?\d+\]|\d+)?"
R1C1_COORD_RE_STR = "(?P<row>{0})?(?P<col>{1})?".format(
    R1C1_ROW_RE_STR, R1C1_COL_RE_STR)

R1C1_COORDINATE_RE = re.compile('^' + R1C1_COORD_RE_STR + '$', re.VERBOSE)

R1C1_RANGE_EXPR = """
(?P<min_row>{0})?
(?P<min_col>{1})?
(:(?P<max_row>{0})?
(?P<max_col>{1})?)?
""".format(R1C1_ROW_RE_STR, R1C1_COL_RE_STR)

R1C1_RANGE_RE = re.compile('^' + R1C1_RANGE_EXPR + '$', re.VERBOSE)

MAX_COL = 18278
MAX_ROW = 1048576

VALID_R1C1_RANGE_ITEM_COMBOS = {
    (0, 1, 0, 0),
    (1, 0, 0, 0),
    (1, 1, 0, 0),
    (0, 1, 0, 1),
    (1, 0, 1, 0),
    (1, 1, 1, 1),
}


#::TODO:: test if case is insensitive
# ::TODO:: validate that A:A and 1:1 produce a range with correct size

class AddressRange(collections.namedtuple(
        'Address', 'sheet start end coordinate address')):

    def __new__(cls, address, sheet=''):

        if isinstance(address, str):
            return cls.create(address, sheet=sheet)

        elif isinstance(address, AddressRange):
            if not sheet or sheet == address.sheet:
                return address

            elif not address.sheet:
                start = AddressCell(address.start.coordinate, sheet=sheet)
                end = AddressCell(address.end.coordinate, sheet=sheet)

            else:
                raise ValueError("Mismatched sheets '{}' and '{}'".format(
                    address, sheet))

        else:
            assert (isinstance(address, tuple) and
                    4 == len(address) and
                    None in address or address[0:2] != address[2:]), \
                "AddressRange expected a range '{}'".format(address)

            start_col, start_row, end_col, end_row = address
            start = AddressCell(
                (start_col, start_row, start_col, start_row), sheet)
            end = AddressCell(
                (end_col, end_row, end_col, end_row), sheet)

        coordinate = '{0}:{1}'.format(start.coordinate, end.coordinate)

        format_str = '{0}!{1}' if sheet else '{1}'
        return super(AddressRange, cls).__new__(
            cls, sheet, start, end, coordinate,
            format_str.format(sheet, coordinate))

    def __str__(self):
        return self.address

    @property
    def is_range(self):
        return True

    @property
    def size(self):
        if 0 in (self.end.row, self.start.row):
            height = MAX_ROW
        else:
            height = self.end.row - self.start.row + 1

        if 0 in (self.end.col_idx, self.start.col_idx):
            width = MAX_COL
        else:
            width = self.end.col_idx - self.start.col_idx + 1

        return AddressSize(height, width)

    @property
    def has_sheet(self):
        return bool(self.sheet)

    @property
    def rows(self):
        """Get each addresses for every cell, yields one row at a time."""
        col_range = self.start.col_idx, self.end.col_idx + 1
        for row in range(self.start.row, self.end.row + 1):
            yield (AddressCell((col, row, col, row), self.sheet)
                   for col in range(*col_range))

    @property
    def cols(self):
        """Get each addresses for every cell, yields one column at a time."""
        col_range = self.start.col_idx, self.end.col_idx + 1
        for col in range(*col_range):
            yield (AddressCell((col, row, col, row), self.sheet)
                   for row in range(self.start.row, self.end.row + 1))

    @classmethod
    def create(cls, address, sheet='', cell=None):
        if isinstance(address, AddressRange):
            return AddressRange(address, sheet=sheet)

        elif isinstance(address, AddressCell):
            return AddressCell(address, sheet=sheet)

        sheetname, addr = split_sheetname(address, sheet=sheet)
        addr_tuple = extended_range_boundaries(addr, cell=cell)

        if None in addr_tuple or addr_tuple[0:2] != addr_tuple[2:]:
            return AddressRange(addr_tuple, sheet=sheetname)
        else:
            return AddressCell(addr_tuple, sheet=sheetname)


class AddressCell(collections.namedtuple(
        'AddressCell', 'sheet column row coordinate address')):

    def __new__(cls, address, sheet=None):

        if isinstance(address, str):
            return cls.create(address, sheet=sheet)

        elif isinstance(address, AddressCell):
            if not sheet or sheet == address.sheet:
                return address

            elif not address.sheet:
                column, row, coordinate = address[1:4]

            else:
                raise ValueError("Mismatched sheets '{}' and '{}'".format(
                    address, sheet))

        else:
            assert (isinstance(address, tuple) and
                    4 == len(address) and
                    None not in address or address[0:2] == address[2:]), \
                "AddressCell expected a cell '{}'".format(address)

            column, row = (a or '' for a in address[:2])
            column = column and get_column_letter(column)
            coordinate = '{0}{1}'.format(column, row)

        if sheet:
            format_str = '{0}!{1}'
        else:
            format_str = '{1}'

        return super(AddressCell, cls).__new__(
            cls, sheet, column, row or 0, coordinate,
            format_str.format(sheet, coordinate))

    def __str__(self):
        return self.address

    def __contains__(self, item):
        return item in self.address

    @property
    def is_range(self):
        return False

    @property
    def size(self):
        return AddressSize(1, 1)

    @property
    def has_sheet(self):
        return bool(self.sheet)

    @property
    def col_idx(self):
        return column_index_from_string(self.column) if self.column else 0

    def inc_col(self, inc):
        return (self.col_idx + inc - 1) % MAX_COL + 1

    def inc_row(self, inc):
        return (self.row + inc - 1) % MAX_ROW + 1

    def address_at_offset(self, row_inc=0, col_inc=0):
        new_col = self.inc_col(col_inc)
        new_row = self.inc_row(row_inc)
        return AddressCell((new_col, new_row, new_col, new_row),
                           sheet=self.sheet)

    @classmethod
    def create(cls, address, sheet='', cell=None):
        addr = AddressRange.create(address, sheet=sheet, cell=cell)
        if not isinstance(addr, AddressCell):
            raise ValueError(
                "{0} is not a valid coordinate".format(address))
        return addr


AddressSize = collections.namedtuple('AddressSize', 'height width')


def unquote_sheetname(sheetname):
    """
    Remove quotes from around, and embedded "''" in, quoted sheetnames

    sheetnames with special characters are quoted in formulas
    This is the inverse of openpyxl.utils.quote_sheetname
    """
    if sheetname.startswith("'") and sheetname.endswith("'"):
        sheetname = sheetname[1:-1].replace("''", "'")
    return sheetname


def split_sheetname(address, sheet=''):
    sh = ''
    if '!' in address:
        sh, address_part = address.split('!', maxsplit=1)
        assert '!' not in address_part, \
            "Only rectangular formulas are supported {}".format(address)
        sh = unquote_sheetname(sh)
        address = address_part

        if sh and sheet and sh != sheet:
            raise ValueError("Mismatched sheets '{}' and '{}'".format(
                sh, sheet))

    return sheet or sh, address


def extended_range_boundaries(address, cell=None):
    """
    R1C1 reference style
    
    You can also use a reference style where both the rows and the columns on
    the worksheet are numbered. The R1C1 reference style is useful for
    computing row and column positions in macros. In the R1C1 style, Excel
    indicates the location of a cell with an "R" followed by a row number
    and a "C" followed by a column number.

    Reference   Meaning

    R[-2]C      A relative reference to the cell two rows up and in
                the same column

    R[2]C[2]    A relative reference to the cell two rows down and
                two columns to the right

    R2C2        An absolute reference to the cell in the second row and
                in the second column

    R[-1]       A relative reference to the entire row above the active cell

    R           An absolute reference to the current row

    """
    try:
        # if this is normal reference then just use the openpyxl converter
        return range_boundaries(address)
    except ValueError:
        pass

    m = R1C1_RANGE_RE.match(address)
    if not m:
        raise ValueError(
            "{0} is not a valid coordinate or range".format(address))

    def from_relative_to_absolute(r1_or_c1):
        def require_cell():
            assert cell is not None, \
                "Must pass a cell to decode a relative address {}".format(
                    address)

        if not r1_or_c1.endswith(']'):
            if len(r1_or_c1) > 1:
                return int(r1_or_c1[1:])

            else:
                require_cell()
                if r1_or_c1[0].upper() == 'R':
                    return cell.row
                else:
                    return cell.col

        else:
            require_cell()
            if r1_or_c1[0].lower() == 'r':
                return (cell.row + int(r1_or_c1[2:-1]) - 1) % MAX_ROW + 1
            else:
                return (cell.col_idx + int(r1_or_c1[2:-1]) - 1) % MAX_COL + 1

    min_col, min_row, max_col, max_row = (
        g if g is None else from_relative_to_absolute(g) for g in (
        m.group(n) for n in ('min_col', 'min_row', 'max_col', 'max_row'))
    )

    items_present = (min_col is not None, min_row is not None,
                     max_col is not None, max_row is not None)

    if items_present not in VALID_R1C1_RANGE_ITEM_COMBOS:
        raise ValueError(
            "{0} is not a valid coordinate or range".format(address))

    if min_col is not None:
        min_col = min_col

    if min_row is not None:
        min_row = min_row

    if max_col is not None:
        max_col = max_col
    else:
        max_col = min_col

    if max_row is not None:
        max_row = max_row
    else:
        max_row = min_row

    return min_col, min_row, max_col, max_row


def resolve_range(address):
    """Return a list or nested lists with AddressCell for each element"""
    
    # ::TODO:: look at removing the assert
    assert isinstance(address, (AddressRange, AddressCell))

    # single cell, no range
    if not address.is_range:
        data = [address]

    else:

        start = address.start
        end = address.end

        # single column
        if start.column == end.column:
            data = list(next(address.cols))

        # single row
        elif start.row == end.row:
            data = list(next(address.rows))

        # rectangular range
        else:
            data = list(list(row) for row in address.rows)

    return data


def get_linest_degree(cell):
    # TODO: assumes a row or column of linest formulas &
    # that all coefficients are needed

    address = cell.address
    # figure out where we are in the row

    # to the left
    i = 0
    while True:
        i -= 1
        f = cell.excel.get_formula_from_range(
            address.address_at_offset(row_inc=0, col_inc=i))
        if not f or f != cell.formula:
            break

    # to the right
    j = 0
    while True:
        j += 1
        f = cell.excel.get_formula_from_range(
            address.address_at_offset(row_inc=0, col_inc=j))
        if not f or f != cell.formula:
            break

    # assume the degree is the number of linest's
    # last -1 is because an n degree polynomial has n+1 coefs
    degree = (j - i - 1) - 1

    # which coef are we (left most coef is the coef for the highest power)
    coef = -i

    # no linests left or right, try looking up/down
    if degree == 0:
        # up
        i = 0
        while True:
            i -= 1
            f = cell.excel.get_formula_from_range(
                address.address_at_offset(row_inc=i, col_inc=0))
            if not f or f != cell.formula:
                break

        # down
        j = 0
        while True:
            j += 1
            f = cell.excel.get_formula_from_range(
                address.address_at_offset(row_inc=j, col_inc=0))
            if not f or f != cell.formula:
                break

        degree = (j - i - 1) - 1
        coef = -i

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
