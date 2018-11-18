"""
    ExcelComWrapper : Must be run on Windows as it requires a COM link
                      to an Excel instance.
    ExcelOpxWrapper : Can be run anywhere but only with post 2010 Excel formats
"""

import abc
import os

from openpyxl import load_workbook
from openpyxl.cell import Cell
from openpyxl.cell.read_only import EMPTY_CELL
from openpyxl.formula.tokenizer import TokenizerError

from pycel.excelutil import AddressCell, AddressRange


class ExcelWrapper(object):
    __metaclass__ = abc.ABCMeta

    @abc.abstractproperty
    def rangednames(self):
        """"""

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


class ExcelComWrapper(ExcelWrapper):  # pragma: no cover
    """ Excel COM wrapper implementation for ExcelWrapper interface """
    def __init__(self, filename, app=None):

        super(ExcelWrapper, self).__init__()

        self.filename = os.path.abspath(filename)
        self.app = app

    @property
    def rangednames(self):
        return self._rangednames

    def connect(self):
        # http://devnulled.com/content/2004/01/
        #   com-objects-and-threading-in-python/
        # TODO: dont need to uninit?
        # pythoncom.CoInitialize()
        if not self.app:
            from win32com.client import Dispatch
            self.app = Dispatch("Excel.Application")
            self.app.Visible = True
            self.app.DisplayAlerts = 0
            self.app.Workbooks.Open(self.filename)
        # else -> if we are running as an excel addin, this gets passed to us

        # Range Names reading
        # WARNING: by default numpy array require dtype declaration to
        #   specify character length (here 'S200', i.e. 200 characters)
        # WARNING: win32.com cannot get ranges with single column/line, would
        # require way to read Office Open XML
        # TODO: automate detection of max string len to set up numpy array
        # TODO: discriminate between worksheet & workbook ranged names
        import numpy as np
        self._rangednames = np.zeros(
            shape=(int(self.app.ActiveWorkbook.Names.Count), 1),
            dtype=[('id', 'int_'), ('name', 'S200'), ('formula', 'S200')])
        for i in range(0, self.app.ActiveWorkbook.Names.Count):
            self._rangednames[i]['id'] = int(i + 1)
            self._rangednames[i]['name'] = str(
                self.app.ActiveWorkbook.Names.Item(i + 1).Name)
            self._rangednames[i]['formula'] = str(
                self.app.ActiveWorkbook.Names.Item(i + 1).Value)

    def save(self):
        self.app.ActiveWorkbook.Save()

    def save_as(self, filename, delete_existing=False):
        if delete_existing and os.path.exists(filename):
            os.remove(filename)
        self.app.ActiveWorkbook.SaveAs(filename)

    def close(self):
        self.app.ActiveWorkbook.Close(False)

    def quit(self):
        return self.app.Quit()

    def get_range(self, address):
        if address.find('!') > 0:
            sheet, address = address.split('!')
            return self.app.ActiveWorkbook.Worksheets(sheet).Range(address)
        else:
            return self.app.ActiveWorkbook.ActiveSheet.Range(address)

    def get_used_range(self):
        return self.app.ActiveWorkbook.ActiveSheet.UsedRange

    def get_active_sheet_name(self):
        return self.app.ActiveWorkbook.ActiveSheet.Name

    def set_calc_mode(self, automatic=True):
        from win32com.client import constants
        if automatic:
            self.app.Calculation = constants.xlCalculationAutomatic
        else:
            self.app.Calculation = constants.xlCalculationManual

    def set_screen_updating(self, update):
        self.app.ScreenUpdating = update

    def run_macro(self, macro):
        self.app.Run(macro)


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
        self.values = self.cell_to_value(cell)


class ExcelOpxWrapper(ExcelWrapper):
    """ OpenPyXl implementation for ExcelWrapper interface """

    def __init__(self, filename, app=None):
        super(ExcelWrapper, self).__init__()

        self.filename = os.path.abspath(filename)
        self._defined_names = None
        self._rangednames = None
        self.workbook = None
        self.workbook_dataonly = None

    @property
    def defined_names(self):
        if self.workbook is not None and self._defined_names is None:
            self._defined_names = {}

            for defined_name in self.workbook.defined_names.definedName:
                try:
                    for worksheet, range_alias in defined_name.destinations:
                        if worksheet in self.workbook:
                            self._defined_names[str(defined_name.name)] = (
                                range_alias, worksheet)
                except TokenizerError:
                    # ::TODO:: this is a workaround for openpyxl throwing
                    # this exception when given a range of sheet!#REF!
                    pass
        return self._defined_names

    @property
    def rangednames(self):
        if self.workbook is not None and self._rangednames is None:
            rangednames = []

            for named_range in self.workbook.defined_names.definedName:
                try:
                    for worksheet, range_alias in named_range.destinations:
                        tuple_name = (
                            len(rangednames) + 1,
                            str(named_range.name),
                            str(self.workbook[worksheet].title + '!' +
                                range_alias)
                        )
                        rangednames.append([tuple_name])
                except TokenizerError:
                    # ::TODO:: this is a workaround for openpyxl throwing
                    # this exception when given a range of sheet!#REF!
                    pass

            self._rangednames = rangednames
        return self._rangednames

    def connect(self):
        self.workbook = load_workbook(self.filename)
        self.workbook_dataonly = load_workbook(
            self.filename, data_only=True, read_only=True)

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
