import calendar
import collections
import datetime as dt
import operator
import re

import numpy as np
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils import (
    get_column_letter,
    range_boundaries as openpyxl_range_boundaries
)


ERROR_CODES = frozenset(Tokenizer.ERROR_CODES)
DIV0 = '#DIV/0!'
EMPTY = '#EMPTY!'
VALUE_ERROR = '#VALUE!'
NUM_ERROR = '#NUM!'
NA_ERROR = '#N/A'

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

TABLE_REF_RE = re.compile(r"^(?P<table_name>[^[]+)\[(?P<table_selector>.*)\]$")

TABLE_SELECTOR_RE = re.compile(
    r"^(?P<row_or_column>[^[]+)$|"
    r"^@\[(?P<this_row_column>[^[]*)\]$|"
    r"^ *(?P<rows>(\[([^\]]+)\] *, *)*)"
    r"(\[(?P<start_col>[^\]]+)\] *: *)?"
    r"(\[(?P<end_col>.+)\] *)?$")

QUESTION_MARK_RE = re.compile(r'\?(?<!~)')
STAR_RE = re.compile(r'\*(?<!~)')

MAX_COL = 16384
MAX_ROW = 1048576

VALID_R1C1_RANGE_ITEM_COMBOS = {
    (0, 1, 0, 1),
    (1, 0, 1, 0),
    (1, 1, 1, 1),
}

OPERATORS = {
    '': operator.eq,
    '=': operator.eq,
    '<': operator.lt,
    '>': operator.gt,
    '<=': operator.le,
    '>=': operator.ge,
    '<>': operator.ne,
}

OPERATORS_RE = re.compile('^(?P<oper>(=|<>|<=?|>=?))?(?P<value>.*)$')

PYTHON_AST_OPERATORS = {
    'Eq': operator.eq,
    'Lt': operator.lt,
    'Gt': operator.gt,
    'LtE': operator.le,
    'GtE': operator.ge,
    'NotEq': operator.ne,
    'Add': operator.add,
    'Sub': operator.sub,
    'UAdd': operator.pos,
    'USub': operator.neg,
    'Mult': operator.mul,
    'Div': operator.truediv,
    'FloorDiv': operator.floordiv,
    'Mod': operator.mod,
    'Pow': operator.pow,
    'LShift': operator.lshift,
    'RShift': operator.rshift,
    'BitOr': operator.or_,
    'BitXor': operator.xor,
    'BitAnd': operator.and_,
    # 'MatMult': operator.matmul,  # not supported on py34
}

COMPARISION_OPS = frozenset(('Eq', 'Lt', 'Gt', 'LtE', 'GtE', 'NotEq'))


class PyCelException(Exception):
    """Base class for PyCel errors"""


class AddressRange(collections.namedtuple(
        'Address', 'address sheet start end coordinate')):
    """ Helper class for constructing, validating and accessing Range Addresses

    **Tuple Attributes:**

    .. py:attribute:: address

        `AddressRange` as a string

    .. py:attribute:: sheet

        Sheet name

    .. py:attribute:: start

        `AddressCell` for upper left corner of `AddressRange`

    .. py:attribute:: end

        `AddressCell` for lower right corner of `AddressRange`

    .. py:attribute:: coordinate

        Address without the sheetname

    **Non-tuple Attributes:**

    """

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

    def __add__(self, other):
        """Assumes rectangular only"""
        other = AddressRange.create(other)
        min_col_idx = min(self.col_idx, other.col_idx)
        min_row = min(self.row, other.row)

        max_col_idx = max(self.col_idx + self.size.width,
                          other.col_idx + other.size.width) - 1
        max_row = max(self.row + self.size.height,
                      other.row + other.size.height) - 1

        return AddressRange((min_col_idx, min_row, max_col_idx, max_row))

    @property
    def col_idx(self):
        """col_idx for left column"""
        return self.start.col_idx

    @property
    def row(self):
        """top row"""
        return self.start.row

    @property
    def is_range(self):
        """Is this address a range?"""
        return True

    @property
    def size(self):
        """Range dimensions"""
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
        """Does the address have a sheet?"""
        return bool(self.sheet)

    @property
    def sort_key(self):
        return self.sheet, self.start.col_idx, self.start.row

    @property
    def rows(self):
        """Get each address for every cell, yields one row at a time."""
        col_range = self.start.col_idx, self.end.col_idx + 1
        for row in range(self.start.row, self.end.row + 1):
            yield (AddressCell((col, row, col, row), sheet=self.sheet)
                   for col in range(*col_range))

    @property
    def cols(self):
        """Get each address for every cell, yields one column at a time."""
        col_range = self.start.col_idx, self.end.col_idx + 1
        for col in range(*col_range):
            yield (AddressCell((col, row, col, row), sheet=self.sheet)
                   for row in range(self.start.row, self.end.row + 1))

    @classmethod
    def create(cls, address, sheet='', cell=None):
        """ Factory method.

        Able to construct R1C1, defined names, and structured references
        style addresses, if passed a `excelcomppiler._Cell`.

        :param address: str, AddressRange, AddressCell
        :param sheet: sheet for address, if not included
        :param cell: `excelcompiler._Cell` reference
        :return: `AddressRange or AddressCell`
        """

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
    """ Helper class for constructing, validating and accessing Cell Addresses

    **Tuple Attributes:**

    .. py:attribute:: address

        `AddressRange` as a string

    .. py:attribute:: sheet

        Sheet name

    .. py:attribute:: col_idx

        Column number as a 1 based index

    .. py:attribute:: row

        Row number as a 1 based index

    .. py:attribute:: coordinate

        Address without the sheetname

    **Non-tuple Attributes:**

    """

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
        """Is this address a range?"""
        return False

    @property
    def size(self):
        """Range dimensions"""
        return AddressSize(1, 1)

    @property
    def has_sheet(self):
        """Does the address have a sheet?"""
        return bool(self.sheet)

    @property
    def sort_key(self):
        return self.sheet, self.col_idx, self.row

    @property
    def column(self):
        """column letter"""
        return (self.col_idx or '') and get_column_letter(self.col_idx)

    def inc_col(self, inc):
        """ Generate an address offset by `inc` columns.

        :param inc: integer number of columns to offset by
        """
        return (self.col_idx + inc - 1) % MAX_COL + 1

    def inc_row(self, inc):
        """ Generate an address offset by `inc` rows.

        :param inc: integer number of rows to offset by
        """
        return (self.row + inc - 1) % MAX_ROW + 1

    def address_at_offset(self, row_inc=0, col_inc=0):
        """ Construct an `AddressCell` offset from the address

        :param row_inc: Number of rows to offset.
        :param col_inc: Number of columns to offset
        :return: `AddressCell`
        """
        new_col = self.inc_col(col_inc)
        new_row = self.inc_row(row_inc)
        return AddressCell((new_col, new_row, new_col, new_row),
                           sheet=self.sheet)

    @classmethod
    def create(cls, address, sheet='', cell=None):
        """ Factory method.

        Able to construct R1C1, defined names, and structured references
        style addresses, if passed a `excelcomppiler._Cell`.

        :param address: str, AddressRange, AddressCell
        :param sheet: sheet for address, if not included
        :param cell: `excelcompiler._Cell` reference
        :return: `AddressCell`
        """
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


def structured_reference_boundaries(address, cell=None):
    # Excel reference: https://support.office.com/en-us/article/
    #   Using-structured-references-with-Excel-tables-
    #   F5ED2452-2337-4F71-BED3-C8AE6D2B276E

    match = TABLE_REF_RE.match(address)
    if not match:
        return None

    if cell is None:
        raise PyCelException(
            "Must pass cell for Structured Reference {}".format(address))

    name = match.group('table_name')
    table, sheet = cell.excel.table(name)

    if table is None:
        raise PyCelException(
            "Table {} not found for Structured Reference: {}".format(
                name, address))

    boundaries = openpyxl_range_boundaries(table.ref)
    assert None not in boundaries

    selector = match.group('table_selector')

    if not selector:
        # all columns and the data rows
        rows, start_col, end_col = None, None, None

    else:
        selector_match = TABLE_SELECTOR_RE.match(selector)
        if selector_match is None:
            raise PyCelException(
                "Unknown Structured Reference Selector: {}".format(selector))

        row_or_column = selector_match.group('row_or_column')
        this_row_column = selector_match.group('this_row_column')

        if row_or_column:
            rows = start_col = None
            end_col = row_or_column

        elif this_row_column:
            rows = '#This Row'
            start_col = None
            end_col = this_row_column

        else:
            rows = selector_match.group('rows')
            start_col = selector_match.group('start_col')
            end_col = selector_match.group('end_col')

            if not rows:
                rows = None

            else:
                assert '[' in rows
                rows = [r.split(']')[0] for r in rows.split('[')[1:]]
                if len(rows) != 1:
                    # not currently supporting multiple row selects
                    raise PyCelException(
                        "Unknown Structured Reference Rows: {}".format(
                            address))

                rows = rows[0]

        if end_col.startswith('#'):
            # end_col collects the single field case
            assert rows is None and start_col is None
            rows = end_col
            end_col = None

        elif end_col.startswith('@'):
            rows = '#This Row'
            end_col = end_col[1:]
            if len(end_col) == 0:
                end_col = start_col

    if rows is None:
        # skip the headers and footers
        min_row = boundaries[1] + (
            table.headerRowCount if table.headerRowCount else 0)
        max_row = boundaries[3] - (
            table.totalsRowCount if table.totalsRowCount else 0)

    else:
        if rows == '#All':
            min_row, max_row = boundaries[1], boundaries[3]

        elif rows == '#Data':
            min_row = boundaries[1] + (
                table.headerRowCount if table.headerRowCount else 0)
            max_row = boundaries[3] - (
                table.totalsRowCount if table.totalsRowCount else 0)

        elif rows == '#Headers':
            min_row = boundaries[1]
            max_row = boundaries[1] + (
                table.headerRowCount if table.headerRowCount else 0) - 1

        elif rows == '#Totals':
            min_row = boundaries[3] - (
                table.totalsRowCount if table.totalsRowCount else 0) + 1
            max_row = boundaries[3]

        elif rows == '#This Row':
            # ::TODO:: If not in a data row, return #VALUE! How to do this?
            min_row = max_row = cell.address.row

        else:
            raise PyCelException(
                "Unknown Structured Reference Rows: {}".format(rows))

    if end_col is None:
        # all columns
        min_col_idx, max_col_idx = boundaries[0], boundaries[2]

    else:
        # a specific column
        column_idx = next((idx for idx, c in enumerate(table.tableColumns)
                           if c.name == end_col), None)
        if column_idx is None:
            raise PyCelException(
                "Column {} not found for Structured Reference: {}".format(
                    end_col, address))
        max_col_idx = boundaries[0] + column_idx

        if start_col is None:
            min_col_idx = max_col_idx

        else:
            column_idx = next((idx for idx, c in enumerate(table.tableColumns)
                               if c.name == start_col), None)
            if column_idx is None:
                raise PyCelException(
                    "Column {} not found for Structured Reference: {}".format(
                        start_col, address))
            min_col_idx = boundaries[0] + column_idx

    if min_row > max_row or min_col_idx > max_col_idx:
        raise PyCelException("Columns out of order : {}".format(address))

    return (min_col_idx, min_row, max_col_idx, max_row), sheet


def range_boundaries(address, cell=None, sheet=None):
    try:
        # if this is normal reference then just use the openpyxl converter
        boundaries = openpyxl_range_boundaries(address)
        if None not in boundaries or ':' in address:
            return boundaries, sheet
    except ValueError:
        pass

    # test for R1C1 style address
    boundaries = r1c1_boundaries(address, cell=cell, sheet=sheet)
    if boundaries:
        return boundaries

    # Try to see if the is a structured table reference
    boundaries = structured_reference_boundaries(address, cell=cell)
    if boundaries:
        return boundaries

    # Try to see if this is a defined name
    name_addr = cell and cell.excel and cell.excel.defined_names.get(address)
    if name_addr:
        return openpyxl_range_boundaries(name_addr[0]), name_addr[1]

    if len(address.split(':')) > 2:
        raise NotImplementedError("Multiple Colon Ranges not implemented")

    raise ValueError(
        "{0} is not a valid coordinate or range".format(address))


def r1c1_boundaries(address, cell=None, sheet=None):
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

    # test for R1C1 style address
    m = R1C1_RANGE_RE.match(address)

    if not m:
        return None

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


def flatten(data, coerce=lambda x: x):
    """ flatten items, converting top level items as needed

    :param data: data to flatten
    :param coerce: apply coercion to top level, but not to sub ranges
    :return: flattened (coerced) items
    """
    if isinstance(data, collections.Iterable) and not isinstance(data, str):
        for item in data:
            yield from flatten(coerce(item))
    else:
        yield coerce(data)


def uniqueify(seq):
    seen = set()
    return tuple(x for x in seq if x not in seen and not seen.add(x))


def is_number(value):
    try:
        float(value)
        return True
    except (ValueError, TypeError):
        return False


def coerce_to_number(value, raise_div0=True):
    if not isinstance(value, str):
        if isinstance(value, int):
            return value
        if is_number(value) and int(value) == float(value):
            return int(value)
        return value

    try:
        if value == DIV0 and raise_div0:
            return 1 / 0
        elif '.' not in value:
            return int(value)
    except (ValueError, TypeError):
        pass

    try:
        return float(value)
    except (ValueError, TypeError):
        return value


def math_wrap(bare_func):
    """wrapper for functions that take numbers to handle errors"""

    def func(*args):
        # this is a bit of a ::HACK:: to quickly address the most common cases
        # for reasonable math function parameters
        for arg in args:
            if arg in ERROR_CODES:
                return arg
        if not (is_number(args[0]) or args[0] in (None, EMPTY)):
            return VALUE_ERROR
        try:
            return bare_func(*(0 if a in (None, EMPTY)
                               else coerce_to_number(a) for a in args))
        except ValueError as exc:
            if "math domain error" in str(exc):
                return NUM_ERROR
            raise  # pragma: no cover
    return func


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


def build_wildcard_re(lookup_value):
    regex = QUESTION_MARK_RE.sub('.', STAR_RE.sub('.*', lookup_value))
    if regex != lookup_value:
        # this will be a regex match"""
        compiled = re.compile('^{}$'.format(regex.lower()))
        return lambda x: compiled.match(x.lower()) is not None
    else:
        return None


def criteria_parser(criteria):
    """
    General rules:

        Criteria will be coerced to numbers,
        For equality comparisions, values will be coerced to numbers
        < and > will always be False when comparing strings to numbers
        <> will always be True when comparing strings to numbers

       You can use the wildcard characters—the question mark (?) and
       asterisk (*)—as the criteria argument. A question mark matches
       any single character; an asterisk matches any sequence of
       characters. If you want to find an actual question mark or
       asterisk, type a tilde (~) preceding the character.
    """

    if is_number(criteria):
        # numeric equals comparision
        criteria = coerce_to_number(criteria)

        def check(x):
            return is_number(x) and coerce_to_number(x) == criteria

    elif isinstance(criteria, str):
        match = OPERATORS_RE.match(criteria)
        criteria_operator = match.group('oper') or ''
        value = match.group('value')
        op = OPERATORS[criteria_operator]

        if op == operator.eq:

            if is_number(value):
                return criteria_parser(value)

            check = build_wildcard_re(value)
            if check is not None:
                return check

        if is_number(value):
            value = coerce_to_number(value)

            def check(x):
                if isinstance(x, str):
                    # string always compare False unless '!='
                    return op == operator.ne
                else:
                    return op(x, value)
        else:
            value = value.lower()

            def check(x):
                """Compare with a string"""
                if not isinstance(x, str):
                    # non string always compare False unless '!='
                    return op == operator.ne
                else:
                    return op(x.lower(), value)

    else:
        raise ValueError("Couldn't parse criteria: {}".format(criteria))

    return check


def find_corresponding_index(rng, criteria):
    return tuple(find_corresponding_index_generator(rng, criteria))


def find_corresponding_index_generator(rng, criteria):
    # parse criteria, build a criteria check
    check = criteria_parser(criteria)

    assert_list_like(rng)
    return (index for index, item in enumerate(rng) if check(item))


def list_like(data):
    return isinstance(data, (list, tuple, np.ndarray))


def assert_list_like(data):
    if not list_like(data):
        raise TypeError('Must be a list like: {}'.format(data))


def type_cmp_value(value):
    """ Excel compares bools above strings which are above numbers

    https://stackoverflow.com/a/35051992/7311767

    :param value: Operand
    :return: tuple of type precedence and the default to use
    """
    if value in ERROR_CODES:
        return 3, value
    elif isinstance(value, bool):
        return 2, False
    elif isinstance(value, str):
        return 1, ''
    else:
        return 0, 0.0


class ExcelCmp(collections.namedtuple('ExcelCmp', 'cmp_type value empty')):

    def __new__(cls, value, empty=None):
        if isinstance(value, ExcelCmp):
            return value

        # empty as the searched for, becomes 0.0
        if value is None:
            cmp_type = 0 if empty is None else empty.cmp_type
            default_empty = 0.0 if empty is None else empty.empty
            value = default_empty
        else:
            cmp_type, default_empty = type_cmp_value(value)

        if cmp_type == 1:
            value = value.lower()

        return super(ExcelCmp, cls).__new__(cls, cmp_type, value, default_empty)

    def __lt__(self, other):
        other = ExcelCmp(other, empty=self)
        return super().__lt__(other)

    def __le__(self, other):
        other = ExcelCmp(other, empty=self)
        return super().__le__(other)

    def __gt__(self, other):
        other = ExcelCmp(other, empty=self)
        return super().__gt__(other)

    def __ge__(self, other):
        other = ExcelCmp(other, empty=self)
        return super().__ge__(other)

    def __eq__(self, other):
        other = ExcelCmp(other, empty=self)
        return self[0] == other[0] and self[1] == other[1]

    def __ne__(self, other):
        return not self == other


def build_operator_operand_fixup(capture_error_state):

    def fixup(left_op, op, right_op):
        """Fix up python operations to be more excel like in these cases:

            Operand error

            Empty cells
            Case-insensitive string compare
            String to Number coercion
            String / Number multiplication
        """
        if isinstance(left_op, list) or isinstance(right_op, list):
            raise NotImplementedError('Array Formulas not implemented')

        if left_op in ERROR_CODES:
            return left_op

        if right_op in ERROR_CODES:
            return right_op

        if op in COMPARISION_OPS:
            if left_op in (None, EMPTY):
                left_op = type_cmp_value(right_op)[1]

            if right_op in (None, EMPTY):
                right_op = type_cmp_value(left_op)[1]

            left_op = ExcelCmp(left_op)
            right_op = ExcelCmp(right_op)

        elif op == 'BitAnd':
            # use bitwise-and '&' as string concat not '+'
            op = 'Add'

            if left_op in (None, EMPTY):
                left_op = ''
            elif isinstance(left_op, bool):
                left_op = str(left_op).upper()
            else:
                left_op = str(coerce_to_number(left_op))

            if right_op in (None, EMPTY):
                right_op = ''
            elif isinstance(right_op, bool):
                right_op = str(right_op).upper()
            else:
                right_op = str(coerce_to_number(right_op))

        else:
            left_op = coerce_to_number(left_op)
            right_op = coerce_to_number(right_op)

            if left_op in (None, EMPTY) and is_number(right_op):
                left_op = 0

            if right_op in (None, EMPTY) and is_number(left_op):
                right_op = 0

            if not (is_number(left_op) and is_number(right_op)
                    or isinstance(left_op, AddressRange)
                    and isinstance(right_op, AddressRange)):
                if op != 'USub':
                    capture_error_state(
                        True, 'Values: {} {} {}'.format(left_op, op, right_op))
                    return VALUE_ERROR

            if isinstance(left_op, bool):
                left_op = (str(left_op).upper()
                           if not is_number(right_op)
                           else int(left_op))

            if isinstance(right_op, bool):
                right_op = (str(right_op).upper()
                            if not is_number(left_op)
                            else int(right_op))

        try:
            if op == 'USub':
                return PYTHON_AST_OPERATORS[op](right_op)
            else:
                return PYTHON_AST_OPERATORS[op](left_op, right_op)
        except ZeroDivisionError:
            capture_error_state(
                True, 'Values: {} {} {}'.format(left_op, op, right_op))
            return DIV0
        except TypeError:
            capture_error_state(
                True, 'Values: {} {} {}'.format(left_op, op, right_op))
            return VALUE_ERROR

    return fixup
