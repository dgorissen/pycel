1.0b23 (2020-06-20)
===================

Fixes
-----
* Fix #89, NPV function behaviour


1.0b22 (2019-10-17)
===================

Fixes
-----
* Fix #80, incompatible w/ networkx 2.4


1.0b21 (2019-10-13)
===================

Changes
-------

* Speed up compile
* Implement defined names in multicolon ranges
* Tokenize ':' when adjoining functions as infix operator
* Various changes in prep to improve references, including
* Add reference expansion to function helpers
* Add sheet to indirect() and ref_param=0 to offset()
* Implement is_address() helper
* Implement intersection and union for AddressCell

Fixes
-----
* Fix #77, empty arg in IFERROR()
* Fix #78, None compare and cleanup error handling for various IFS() funcs


1.0b20 (2019-09-22)
===================

Changes
-------

* Implement multi colon ranges
* Add support for missing (empty) function parameters

Fixes
-----
* Fix threading issue in iterative evaluator
* Fix range intersection with null result for ROW and COLUMN
* Fix #74 - Count not working for ranges


1.0b19 (2019-09-12)
===================

Changes
-------

* Implement INDIRECT & OFFSET
* Implement SMALL, LARGE & ROUNDDOWN
* Add error message for unhandled missing function parameter

Fixes
-----
* Fix threading issue w/ CSE evaluator


1.0b18 (2019-09-07)
===================

Changes
-------

* Implement CEILING_MATH, CEILING_PRECISION, FLOOR_MATH & FLOOR_PRECISION
* Implement FACT & FACTDOUBLE
* Implement AVERAGEIF, MAXIFS, MINIFS
* Implement ODD, EVEN, ISODD, ISEVEN, SIGN

Fixes
-----
* Fix #67 - Evaluation with unbounded range
* Fix bugs w/ single cells for xIFS functions


1.0b17 (2019-09-02)
===================

Changes
-------
* Add Formula Support for Multi Area Ranges from defined names
* Allow ExcelCompiler init from openpyxl workbook
* Implement LOWER(), REPLACE(), TRIM() & UPPER()
* Implement DATEVALUE(), IFS() and ISERR()

* Reorganized time and time utils and text functions
* Add excelutil.AddressMultiAreaRange.
* Add abs_coordinate() property to AddressRange and AddressCell
* Cleanup import statements

Fixes
-----
* Resolved tox version issue on travis
* Fix defined names with Multi Area Range


1.0b16 (2019-07-07)
===================

Changes
-------
* Add twelve date and time functions
* Serialize workbook filename and use it instead of the serialization filename


1.0b15 (2019-06-30)
===================

Changes
-------
* Implement AVERAGEIFS()
* Take Iterative Calc Parameter defaults from workbook

Bug Fixes
---------
* #60, Binder Notebook Example not Working


1.0b14 (2019-06-16)
===================

Changes
-------
* Added method to evaluate the conditional format (formulas) for a cell or cells
* Added ExcelCompiler(..., cycles=True) to allow Excel iterative calculations


1.0b13 (2019-05-10)
===================

Changes
-------
* Implement VALUE()
* Improve compile performance reversion from CSE work

Bug Fixes
---------
* #54, In normalize_year(), month % 12 can be 0 -> IllegalMonthError


1.0b12 (2019-04-22)
===================

Changes
-------
* Add library plugin support
* Improve evaluate of unbounded row/col (ie: A:B)
* Fix some regressions from 1.0b11


1.0b11 (2019-04-21)
===================

Major Changes
-------------

* Refactored ExcelWrapper, ExcelFormula & ExcelCompiler to allow...
* Added considerable extensions to CSE Array Formula Support
* Refactored function_helpers to add decorators for excelizing library functions
* Added a dozen new functions
* Improved various messages and exceptions in validate_calcs() and trim_graph()

Changes
-------

* Implement LEFT()
* Implement ISERROR()
* Implement FIND()
* Implement ISNUMBER()
* Implement SUMPRODUCT()
* Implement CEILING()
* Implement TRUNC() and FLOOR()
* Add support for LOG()
* Improve ABS(), INT() and ROUND()

* Add quoted_address() method to AddressRange and AddressCell
* Add public interface to get list of formula_cells()
* Improve Some NotImplementedError() messages
* Add NotImplementedError for "linked" sheet names
* Fix function info for paired functions from same line
* Add reference URL to function info
* Only build compiler eval context once
* Address Range Union and Intersection need sheet_name

CSE Array Changes
-----------------

* Add CSE Array handling to excelformula and excelcompiler
* Change Row, Column & Index to rectangular arrays only
* Add in_array_formula_context
* Add cse_array_wrapper() to allow calling functions in array context
* Add error_string_wrapper() to check for excel errors
* Move math_wrap() to function_helpers.
* Handle Direct CSE Array in cell
* Reorganize CSE Array Formula handling in excelwrapper
* For CSE Arrays that are smaller than target fill w/ None
* Trim oversize array results to fit target range
* Improve needed addresses parser from python code
* Improve _coerce_to_number() and _numerics() for CSE arrays
* Remove formulas from excelwrapper._OpxRange()

Bug Fixes
---------

* Fix Range Intersection
* Fix Unary Minus on Empty cell
* Fix ISNA()
* Fix AddressCell create from tuple
* Power(0,-1) now returns DIV0
* Cleanup index()


1.0b8 (2019-03-20)
==================

Changes
-------

* Implement operators for Array Formulas
* Implement concatenate and concat
* Implement subtotal
* Add support for expanding array formulas
* Fix column and row for array formulas
* Add support for table relative references
* Add function information methods
* Improve messages for validate_calcs and not implemented functions


1.0b7 (2019-03-10)
==================

Changes
-------

* Implement Array (CSE) Formulas

Bug Fixes
---------

* Fix #45 - Unbounded Range Addresses (ie: A:B or 1:2) broken


1.0b6 (2019-03-03)
==================

Bug Fixes
---------

* Fix #42 - 'ReadOnlyWorksheet' object has no attribute 'iter_cols'
* Fix #43 - Fix error with leading/trailing whitespace


1.0b5 (2019-02-24)
==================

Changes
-------

* Implement XOR(), NOT(), TRUE(), FALSE()
* Improve error handling for AND(), OR()
* Implement POWER() function


1.0b4 (2019-02-17)
==================

Changes
-------

* Move to openpyxl 2.6+
* Remove support for Python 3.4


1.0b3 (2019-02-02)
==================

Changes
-------

* Work around openpyxl returning datetimes
* Pin to openpyxl 2.5.12 to avoid bug in 2.5.14 (fixed in PR #315)


1.0b2 (2019-01-05)
==================

Changes
-------

* Much work to better match Excel error processing
* Extend validate_calcs() to allow testing entire workbook
* Improvements to match(), including wildcard support
* Finished implementing match(), lookup(), vlookup() and hlookup()
* Implement COLUMN() and ROW()
* Implement % operator
* Implement len()
* Implement binary base number Excel functions (hex2dec, etc.)
* Fix PI()


1.0b0 (2018-12-25)
===================

Major changes
-------------

* Converted to Python 3.4+
* Removed Windows Excel COM driver (openpyxl is used for all xlsx reading)
* Add support for defined names
* Add support for structured references
* Fix support for relative formulas
* set_value() and evaluate() support ranges and lists
* Add several more library functions
* Add AddressRange and AddressCell classes to encapsulate address calcs
* Add validate_calcs() to aid debugging excellib functions
* Add `build` feature which can limit recompile to only when excel file changes
* Improved handling for #DIV0! and #VALUE!


Compatibility
-------------

* Tests run on Python 3.4, 3.5, 3.6, 3.7 (via tox)
* Python 2 no longer supported


Other Changes
-------------

* Heavily refactored ExcelCompiler
* Moved all formula evaluation, parsing, etc, code to ExcelFormula class
* Convert to using openpyxl tokenizer
* Converted prints to logging calls
* Convert to using pytest
* Add support for travis and codecov.io
* 100% unit test coverage (mostly)
* Add debuggable formula evaluation
* Cleanup generated Python code to make easier to read
* Add a text format (yaml or json) serialization format
* flake8 (pep8) checks added
* pip now handles which Python versions can be used
* Release to PyPI
* Docs updated


Bugfixes
--------

* Numerous


0.0.1 (unreleased)
===================

* Original version available from `Dirk Ggorissen's Pycel Github Page`_.
* Supports Python 2

.. _Dirk Ggorissen's Pycel Github Page: https://github.com/dgorissen/pycel/tree/33c1370d499c629476c5506c7da308713b5842dc
