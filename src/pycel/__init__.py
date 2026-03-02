def _install_openpyxl_compat():
    # openpyxl 3.1 removed .append on DefinedNameDict
    try:
        from openpyxl.workbook.defined_name import DefinedNameDict
        if not hasattr(DefinedNameDict, 'append'):  # pragma: no branch
            DefinedNameDict.append = lambda self, value: self.add(value)
    except Exception:  # pragma: no cover
        pass

    # openpyxl 3.1 removed Worksheet.formula_attributes
    try:
        from openpyxl.worksheet.worksheet import Worksheet
        if not hasattr(Worksheet, 'formula_attributes'):  # pragma: no branch
            def _formula_attributes(self):
                return self.__dict__.setdefault('_pycel_formula_attributes', {})
            Worksheet.formula_attributes = property(_formula_attributes)
    except Exception:  # pragma: no cover
        pass


_install_openpyxl_compat()

# Imports are intentionally after compatibility patching, so pycel users get
# openpyxl shims installed as soon as the package is imported.
from .excelcompiler import ExcelCompiler  # noqa: E402
from .excelutil import AddressCell, AddressRange, PyCelException  # noqa: E402
from .version import __version__  # noqa: E402
