import calendar
import collections
import datetime as dt
import operator
import re

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import (
    get_column_letter,
    range_boundaries as openpyxl_range_boundaries
)


ERROR_CODES = frozenset(Tokenizer.ERROR_CODES)
DIV0 = '#DIV/0!'

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

MAX_COL = 16384
MAX_ROW = 1048576

VALID_R1C1_RANGE_ITEM_COMBOS = {
    (0, 1, 0, 1),
    (1, 0, 1, 0),
    (1, 1, 1, 1),
}

OPERATORS = {
    '': operator.eq,
    '<': operator.lt,
    '>': operator.gt,
    '<=': operator.le,
    '>=': operator.ge,
    '<>': operator.ne,
}


class AddressRange(collections.namedtuple(
        'Address', 'address sheet start end coordinate')):

    def __new__(cls, address, *args, sheet=''):
        if args:
            return super(AddressRange, cls).__new__(cls, address, *args)

        if isinstance(address, str):
            return cls.create(address, sheet=sheet)

        elif isinstance(address, AddressCell):
            return AddressCell(address, sheet=sheet)

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
                (start_col, start_row, start_col, start_row), sheet=sheet)
            end = AddressCell(
                (end_col, end_row, end_col, end_row), sheet=sheet)

        coordinate = '{0}:{1}'.format(start.coordinate, end.coordinate)

        format_str = '{0}!{1}' if sheet else '{1}'
        return super(AddressRange, cls).__new__(
            cls, format_str.format(sheet, coordinate),
            sheet, start, end, coordinate)

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
    def sort_key(self):
        return self.sheet, self.start.col_idx, self.start.row

    @property
    def rows(self):
        """Get each addresses for every cell, yields one row at a time."""
        col_range = self.start.col_idx, self.end.col_idx + 1
        for row in range(self.start.row, self.end.row + 1):
            yield (AddressCell((col, row, col, row), sheet=self.sheet)
                   for col in range(*col_range))

    @property
    def cols(self):
        """Get each addresses for every cell, yields one column at a time."""
        col_range = self.start.col_idx, self.end.col_idx + 1
        for col in range(*col_range):
            yield (AddressCell((col, row, col, row), sheet=self.sheet)
                   for row in range(self.start.row, self.end.row + 1))

    @classmethod
    def create(cls, address, sheet='', cell=None):
        if isinstance(address, AddressRange):
            return AddressRange(address, sheet=sheet)

        elif isinstance(address, AddressCell):
            return AddressCell(address, sheet=sheet)

        sheetname, addr = split_sheetname(address, sheet=sheet)
        addr_tuple, sheetname = range_boundaries(
            addr, sheet=sheetname, cell=cell)

        if None in addr_tuple or addr_tuple[0:2] != addr_tuple[2:]:
            return AddressRange(addr_tuple, sheet=sheetname)
        else:
            return AddressCell(addr_tuple, sheet=sheetname)


class AddressCell(collections.namedtuple(
        'AddressCell', 'address sheet col_idx row coordinate')):

    def __new__(cls, address, *args, sheet=''):
        if args:
            return super(AddressCell, cls).__new__(cls, address, *args)

        if isinstance(address, str):
            return cls.create(address, sheet=sheet)

        elif isinstance(address, AddressCell):
            if not sheet or sheet == address.sheet:
                return address

            elif not address.sheet:
                row, col_idx, coordinate = address[2:5]

            else:
                raise ValueError("Mismatched sheets '{}' and '{}'".format(
                    address, sheet))

        else:
            assert (isinstance(address, tuple) and
                    4 == len(address) and
                    None not in address or address[0:2] == address[2:]), \
                "AddressCell expected a cell '{}'".format(address)

            col_idx, row = (a or 0 for a in address[:2])
            column = (col_idx or '') and get_column_letter(col_idx)
            coordinate = '{0}{1}'.format(column, row or '')

        if sheet:
            format_str = '{0}!{1}'
        else:
            format_str = '{1}'

        return super(AddressCell, cls).__new__(
            cls, format_str.format(sheet, coordinate),
            sheet, col_idx, row, coordinate)

    def __str__(self):
        return self.address

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
    def sort_key(self):
        return self.sheet, self.col_idx, self.row

    @property
    def column(self):
        return (self.col_idx or '') and get_column_letter(self.col_idx)

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


def range_boundaries(address, cell=None, sheet=None):
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

    R           An absolute reference to the current row as part of a range

    """
    try:
        # if this is normal reference then just use the openpyxl converter
        boundaries = openpyxl_range_boundaries(address)
        if None not in boundaries or ':' in address:
            return boundaries, sheet
    except ValueError:
        pass

    m = R1C1_RANGE_RE.match(address)
    if not m:
        name_addr = (cell and cell.excel and
                     cell.excel.defined_names.get(address))
        if name_addr:
            return openpyxl_range_boundaries(name_addr[0]), name_addr[1]

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
                    return cell.col_idx

        else:
            require_cell()
            if r1_or_c1[0].lower() == 'r':
                return (cell.row + int(r1_or_c1[2:-1]) - 1) % MAX_ROW + 1
            else:
                return (cell.col_idx + int(r1_or_c1[2:-1]) - 1) % MAX_COL + 1

    min_col, min_row, max_col, max_row = (
        g if g is None else from_relative_to_absolute(g) for g in (
            m.group(n) for n in ('min_col', 'min_row', 'max_col', 'max_row')
        )
    )

    items_present = (min_col is not None, min_row is not None,
                     max_col is not None, max_row is not None)

    is_range = ':' in address
    if (is_range and items_present not in VALID_R1C1_RANGE_ITEM_COMBOS or
            not is_range and sum(items_present) < 2):
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

    return (min_col, min_row, max_col, max_row), sheet


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


def coerce_to_number(value):
    if not isinstance(value, str):
        if is_number(value) and int(value) == float(value):
            return int(value)
        return value

    try:
        if value == DIV0:
            return 1 / 0
        elif '.' not in value:
            return int(value)
    except (ValueError, TypeError):
        pass

    try:
        return float(value)
    except (ValueError, TypeError):
        return value


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
        # numeric equals comparision
        def check(x):
            return is_number(x) and x == float(criteria)

    elif type(criteria) == str:

        search = re.search(r'(\W*)(.*)', criteria).group
        criteria_operator = search(1)
        op = OPERATORS[criteria_operator]
        value = search(2)

        # all operators except == (blank) are numeric
        numeric_compare = bool(criteria_operator) or is_number(value)

        def validate_number(x):
            if is_number(x):
                return True
            else:
                if numeric_compare:
                    raise TypeError(
                        'excellib.countif() doesnt\'t work for checking'
                        ' non number items against non equality')
                return False

        value = float(value) if validate_number(value) else str(value).lower()

        def check(x):
            if is_number(x):
                return op(x, value)
            else:
                return x.lower() == value

    else:
        raise ValueError("Couldn't parse criteria: {}".format(criteria))

    return check


def find_corresponding_index(rng, criteria):
    """This does not parse all of the patterns available to countif, etc"""
    # parse criteria
    check = criteria_parser(criteria)

    return [index for index, item in enumerate(rng) if check(item)]
