"""
    ExcelComWrapper : Must be run on Windows as it requires a COM link
                      to an Excel instance.
    ExcelOpxWrapper : Can be run anywhere but only with post 2010 Excel formats
"""

import abc
import collections
import itertools as it
import os

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.cell.read_only import EMPTY_CELL
from pycel.excelutil import AddressCell, AddressRange


class ExcelWrapper:
    __metaclass__ = abc.ABCMeta

    @abc.abstractmethod
    def connect(self):
        """"""

    @abc.abstractmethod
    def get_range(self, address):
        """"""

    @abc.abstractmethod
    def get_used_range(self):
        """"""

    @abc.abstractmethod
    def get_active_sheet_name(self):
        """"""

    def get_formula_from_range(self, address):
        f = self.get_range(address).Formula
        if isinstance(f, (list, tuple)):
            if any(x for x in f if x[0].startswith("=")):
                return [x[0] for x in f]
            else:
                return None
        else:
            return f if f.startswith("=") else None

    def get_formula_or_value(self, name):
        r = self.get_range(name)
        return r.Formula or r.Value


class _OpxRange:
    """ Excel range wrapper that distributes reduced api used by compiler
        (Formula & Value)
    """
    def __init__(self, cells, cells_dataonly):
        self.formulas = tuple(tuple(self.cell_to_formula(cell) for cell in row)
                              for row in cells)
        self.values = tuple(tuple(self.cell_to_value(cell) for cell in row)
                            for row in cells_dataonly)

    @classmethod
    def cell_to_formula(cls, cell):
        return str(cell.value) if cell.value is not None else ''

    @classmethod
    def cell_to_value(cls, cell):
        return None if cell.data_type is Cell.TYPE_FORMULA else cell.value

    @property
    def Formula(self):
        return self.formulas

    @property
    def Value(self):
        return self.values


class _OpxCell(_OpxRange):
    """ Excel cell wrapper that distributes reduced api used by compiler
        (Formula & Value)
    """
    def __init__(self, cell, cell_dataonly):
        self.formulas = self.cell_to_formula(cell)
        self.values = self.cell_to_value(cell_dataonly)


class ExcelOpxWrapper(ExcelWrapper):
    """ OpenPyXl implementation for ExcelWrapper interface """

    def __init__(self, filename, app=None):
        super(ExcelWrapper, self).__init__()

        self.filename = os.path.abspath(filename)
        self._defined_names = None
        self._tables = None
        self.workbook = None
        self.workbook_dataonly = None

    @property
    def defined_names(self):
        if self.workbook is not None and self._defined_names is None:
            self._defined_names = {}

            for defined_name in self.workbook.defined_names.definedName:
                for worksheet, range_alias in defined_name.destinations:
                    if worksheet in self.workbook:
                        self._defined_names[str(defined_name.name)] = (
                            range_alias, worksheet)
        return self._defined_names

    def table(self, table_name):
        """ Return the table and the sheet it was found on

        :param table_name: name of table to retrieve
        :return: table, sheet_name
        """
        # table names are case insensitive
        if self._tables is None:
            TableAndSheet = collections.namedtuple(
                'TableAndSheet', 'table, sheet_name')
            self._tables = {
                t.name.lower(): TableAndSheet(t, ws.title)
                for ws in self.workbook for t in ws._tables}
            self._tables[None] = TableAndSheet(None, None)
        return self._tables.get(table_name.lower(), self._tables[None])

    def connect(self):
        self.workbook = load_workbook(self.filename)
        self.workbook_dataonly = load_workbook(
            self.filename, data_only=True, read_only=True)

        for ws in self.workbook:  # pragma: no cover
            # ::TODO:: this is simple hack so that we won't try to eval
            # array formulas since they are not implemented
            for address, props in ws.formula_attributes.items():
                if props.get('t') == 'array':
                    formula = '{%s}' % ws[address].value
                    addrs = it.chain.from_iterable(
                        AddressRange(props.get('ref')).rows)
                    for addr in addrs:
                        ws[addr.coordinate] = formula

    def set_sheet(self, s):
        self.workbook.active = self.workbook.index(self.workbook[s])
        self.workbook_dataonly.active = self.workbook_dataonly.index(
            self.workbook_dataonly[s])
        return self.workbook.active

    def get_range(self, address):
        if not isinstance(address, (AddressRange, AddressCell)):
            address = AddressRange(address)

        if address.has_sheet:
            sheet = self.workbook[address.sheet]
            sheet_dataonly = self.workbook_dataonly[address.sheet]
        else:
            sheet = self.workbook.active
            sheet_dataonly = self.workbook_dataonly.active

        cells = sheet[address.coordinate]
        if isinstance(cells, Cell):
            cell = cells
            cell_dataonly = sheet_dataonly[address.coordinate]
            return _OpxCell(cell, cell_dataonly)

        else:
            cells_dataonly = sheet_dataonly[address.coordinate]

            if len(cells) != len(cells_dataonly):
                # The read_only version of an openpyxl worksheet has the
                # somewhat annoying property of not giving empty rows at the
                # end.  Which is not the same behavior as the non-readonly
                # version.  So we need to align the data here by adding
                # empty rows.
                empty_row = (EMPTY_CELL, ) * len(cells[0])
                empty_rows = (empty_row, ) * (len(cells) - len(cells_dataonly))
                cells_dataonly += empty_rows

            return _OpxRange(cells, cells_dataonly)

    def get_used_range(self):
        return self.workbook.active.iter_rows()

    def get_active_sheet_name(self):
        return self.workbook.active.title
