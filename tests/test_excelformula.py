# -*- coding: UTF-8 -*-
#
# Copyright 2011-2019 by Dirk Gorissen, Stephen Rauch and Contributors
# All rights reserved.
# This file is part of the Pycel Library, Licensed under GPLv3 (the 'License')
# You may not use this work except in compliance with the License.
# You may obtain a copy of the Licence at:
#   https://www.gnu.org/licenses/gpl-3.0.en.html

import collections
import logging
import os
import pickle
from unittest import mock

import pytest

from pycel.excelformula import (
    ASTNode,
    ExcelFormula,
    FormulaEvalError,
    FormulaParserError,
    Token,
    UnknownFunction,
)
from pycel.excelutil import (
    AddressCell,
    DIV0,
    NAME_ERROR,
    NULL_ERROR,
    VALUE_ERROR,
)


FormulaTest = collections.namedtuple('FormulaTest', 'formula rpn python_code')


@pytest.fixture(scope='session')
def empty_eval_context():
    return ExcelFormula.build_eval_context(
        None, None, logging.getLogger('pycel_x'))


def stringify_rpn(e):
    return "|".join([str(x) for x in e])


range_inputs = [
    FormulaTest(
        '=$A1',
        '$A1',
        '_C_("A1")'),
    FormulaTest(
        '=$B$2',
        '$B$2',
        '_C_("B2")'),
    FormulaTest(
        '=SUM(B5:B15)',
        'B5:B15|SUM',
        'xsum(_R_("B5:B15"))'),
    FormulaTest(
        '=SUM(B5:B15,D5:D15)',
        'B5:B15|D5:D15|SUM',
        'xsum(_R_("B5:B15"), _R_("D5:D15"))'),
    FormulaTest(
        '=SUM(B5:B15 A7:D7)',
        'B5:B15|A7:D7| |SUM',
        'xsum(_R_(str(_REF_("B5:B15") & _REF_("A7:D7"))))'),
    FormulaTest(
        '=SUM((A:A,1:1))',
        'A:A|1:1|,|SUM',
        'xsum(_R_("A:A"), _R_("1:1"))'),
    FormulaTest(
        '=SUM((A:A A1:B1))',
        'A:A|A1:B1| |SUM',
        'xsum(_R_(str(_REF_("A:A") & _REF_("A1:B1"))))'),
    FormulaTest(
        '=SUM(D9:D11,E9:E11,F9:F11)',
        'D9:D11|E9:E11|F9:F11|SUM',
        'xsum(_R_("D9:D11"), _R_("E9:E11"), _R_("F9:F11"))'),
    FormulaTest(
        '=SUM((D9:D11,(E9:E11,F9:F11)))',
        'D9:D11|E9:E11|F9:F11|,|,|SUM',
        'xsum(_R_("D9:D11"), (_R_("E9:E11"), _R_("F9:F11")))'),
    FormulaTest(
        '={SUM(B2:D2*B3:D3)}',
        'B2:D2|B3:D3|*|SUM|ARRAYROW|ARRAY',
        '((xsum(_R_("B2:D2") * _R_("B3:D3")),),)'),
    FormulaTest(
        '=RIGHT({"A","B"},A2:B2)',
        '"A"|"B"|ARRAYROW|ARRAY|A2:B2|RIGHT',
        'right((("A", "B",),), _R_("A2:B2"))'),
    FormulaTest(
        '=LEFT({"A";"B"},B1:B2)',
        '"A"|ARRAYROW|"B"|ARRAYROW|ARRAY|B1:B2|LEFT',
        'left((("A",), ("B",),), _R_("B1:B2"))'),
    FormulaTest(
        '=MID({"A","B";"C","D"},A1:B2)',
        '"A"|"B"|ARRAYROW|"C"|"D"|ARRAYROW|ARRAY|A1:B2|MID',
        'mid((("A", "B",), ("C", "D",),), _R_("A1:B2"))'),
    FormulaTest(
        '=SUM(123 + SUM(456) + (45<6))+456+789',
        '123|456|SUM|+|45|6|<|+|SUM|456|+|789|+',
        '(xsum((123 + xsum(456)) + (45 < 6)) + 456) + 789'),
    FormulaTest(
        '=AVG(((((123 + 4 + AVG(A1:A2))))))',
        '123|4|+|A1:A2|AVG|+|AVG',
        'avg((123 + 4) + avg(_R_("A1:A2")))'),
]

basic_inputs = [
    FormulaTest(
        '=SUM((A:A 1:1))',
        'A:A|1:1| |SUM',
        'xsum(_R_(str(_REF_("A:A") & _REF_("1:1"))))'),
    FormulaTest(
        '=A1',
        'A1',
        '_C_("A1")'),
    FormulaTest(
        '=50',
        '50',
        '50'),
    FormulaTest(
        '=1+1',
        '1|1|+',
        '1 + 1'),
    FormulaTest(
        '=atan2(A1,B1)',
        'A1|B1|atan2',
        'xatan2(_C_("A1"), _C_("B1"))'),
    FormulaTest(
        '=5*log(sin()+2)',
        '5|sin|2|+|log|*',
        '5 * log(sin() + 2)'),
    FormulaTest(
        '=5*log(sin(3,7,9)+2)',
        '5|3|7|9|sin|2|+|log|*',
        '5 * log(sin(3, 7, 9) + 2)'),
    FormulaTest(
        '="x"="y"',
        '"x"|"y"|=',
        '"x" == "y"'),
    FormulaTest(
        '="x"=1',
        '"x"|1|=',
        '"x" == 1'),
    FormulaTest(
        '=3 +1-5',
        '3|1|+|5|-',
        '(3 + 1) - 5'),
    FormulaTest(
        '=3 + 4 * 5',
        '3|4|5|*|+',
        '3 + (4 * 5)'),
    FormulaTest(
        '=+3',
        '3',
        '3'),
    FormulaTest(
        '=PI()',
        'PI',
        'pi'),
    FormulaTest(
        '=_xlfn.FUNCTION(L45)',
        'L45|_xlfn.FUNCTION',
        'function(_C_("L45"))'),
    FormulaTest(
        '=FLOOR.MATH(L45)',
        'L45|FLOOR.MATH',
        'floor_math(_C_("L45"))'),
    FormulaTest(
        '=100%',
        '100|%',
        '100 / 100'),
    FormulaTest(
        '=100^100%',
        '100|100|%|^',
        '100 ** (100 / 100)'),
    FormulaTest(
        '=SUM(B5:B15,D5:D15)%',
        'B5:B15|D5:D15|SUM|%',
        'xsum(_R_("B5:B15"), _R_("D5:D15")) / 100'),
    FormulaTest(
        '=AND(G3, 1)',
        'G3|1|AND',
        'x_and(_C_("G3"), 1)'),
    FormulaTest(
        '=OR(TRUE, TRUE(), FALSE, FALSE())',
        'TRUE|TRUE|FALSE|FALSE|OR',
        'x_or(True, True, False, False)'),
]

whitespace_inputs = [
    FormulaTest(
        '=3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3',
        '3|4|2|*|1|5|-|2|^|3|^|/|+',
        '3 + ((4 * 2) / (((1 - 5) ** 2) ** 3))'),
    FormulaTest(
        '=1+3+5',
        '1|3|+|5|+',
        '(1 + 3) + 5'),
    FormulaTest(
        '=3 * 4 + 5',
        '3|4|*|5|+',
        '(3 * 4) + 5'),
    FormulaTest(
        '= (1,5 * (1 + B11 *B3 ^ B12) + 5) + 10 ',
        '1|5|,|1|B11|B3|B12|^|*|+|*|5|+|10|+',
        '(((1, 5) * (1 + (_C_("B11") * (_C_("B3") ** _C_("B12"))))) + 5) + 10',
    ),
    FormulaTest(
        '=f(,1)',
        '|1|f',
        'f(None, 1)',
    ),
    FormulaTest(
        '=f(1,,)',
        '1|||f',
        'f(1, None, None)',
    ),
]

if_inputs = [
    FormulaTest(
        '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &'
        '   "  more ""test"" text"',
        '"a"|"a"|"b"|ARRAYROW|"c"|#N/A|ARRAYROW|1|-|TRUE|ARRAYROW|ARRAY|=|'
        '"yes"|"no"|IF|"  more ""test"" text"|&',
        'x_if("a" == (("a", "b",), ("c", "#N/A",), (-1, True,),), "yes", "no")'
        ' & "  more \\"test\\" text"'),
    FormulaTest(
        '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2]'
        ',0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
        'R13C3|2002|1|6|DATE|>|0|R[41]C[2]|ISERROR|0|R13C3|R[41]C[2]|>=|0|'
        'R[23]C[11]|55|>=|R[24]C[11]|20|>=|AND|R53C3|0|IF|IF|IF|IF',
        'x_if(_C_("C13") > date(2002, 1, 6), 0, x_if(iserror(_C_("C42")), 0, '
        'x_if(_C_("C13") >= _C_("C42"), 0, x_if(x_and('
        '_C_("L24") >= 55, _C_("L25") >= 20), _C_("C53"), 0))))'),
    FormulaTest(
        '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND('
        'R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),'
        'R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,'
        'R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],'
        'R[43]C[11])),0))',
        'R[39]C[11]|65|>|R[25]C[42]|R[11]C[11]|R[39]C[11]|55|>=|'
        'R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|'
        'R[44]C[11]|R[43]C[11]|IF|*|R[14]C[11]|R[39]C[11]|55|>=|'
        'R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|'
        'R[45]C[11]|R[43]C[11]|IF|*|+|0|ROUND|IF',
        'x_if(_C_("L40") > 65, _C_("AQ26"), x_round((_C_("L12") * x_if(x_or('
        'x_and(_C_("L40") >= 55, _C_("L41") >= 20), x_and(_C_("L41") >= 20, '
        '_C_("C11") == "YES")), _C_("L45"), _C_("L44"))) + (_C_("L15") * '
        'x_if(x_or(x_and(_C_("L40") >= 55, _C_("L41") >= 20), x_and(_C_("L41") '
        '>= 20, _C_("C11") == "YES")), _C_("L46"), _C_("L44"))), 0))'),
    FormulaTest(
        '=IF(AI119="","",E119)',
        'AI119|""|=|""|E119|IF',
        'x_if(_C_("AI119") == "", "", _C_("E119"))'),
    FormulaTest(
        '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",'
        'IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
        'P5|1.0|=|"NA"|P5|2.0|=|"A"|P5|3.0|=|"B"|P5|4.0|=|"C"|P5|5.0|=|'
        '"D"|P5|6.0|=|"E"|P5|7.0|=|"F"|P5|8.0|=|"G"|IF|IF|IF|IF|IF|IF|IF|IF',
        'x_if(_C_("P5") == 1.0, "NA", x_if(_C_("P5") == 2.0, "A", '
        'x_if(_C_("P5") == 3.0, "B", x_if(_C_("P5") == 4.0, "C", '
        'x_if(_C_("P5") == 5.0, "D", x_if(_C_("P5") == 6.0, "E", '
        'x_if(_C_("P5") == 7.0, "F", x_if(_C_("P5") == 8.0, "G"))))))))'),
]

fancy_reference_inputs = [
    FormulaTest(
        '=A8:index(B2,1)',
        'A8|B2|1|index|:',
        '_R_(str(_REF_("A8") | index(_REF_("B2"), 1)))'),
    FormulaTest(
        '=A8:B8:index(B2,1)',
        'A8:B8|B2|1|index|:',
        '_R_(str(_REF_("A8:B8") | index(_REF_("B2"), 1)))'),
    FormulaTest(
        '=index(B2,1):A8',
        'B2|1|index|A8|:',
        '_R_(str(index(_REF_("B2"), 1) | _REF_("A8")))'),
    FormulaTest(
        '=index(B2,1):A8:B8',
        'B2|1|index|A8:B8|:',
        '_R_(str(index(_REF_("B2"), 1) | _REF_("A8:B8")))'),
    FormulaTest(
        '=A8:index(B2,1):B2',
        'A8|B2|1|index|:|B2|:',
        '_R_(str((_REF_(str(_REF_("A8") | index(_REF_("B2"), 1)))) | '
        '_REF_("B2")))'),
    FormulaTest(
        '=SUM(sheet1!$A$1:$B$2)',
        'sheet1!$A$1:$B$2|SUM',
        'xsum(_R_("sheet1!A1:B2"))'),
    FormulaTest(
        '=[data.xls]sheet1!$A$1',
        '[data.xls]sheet1!$A$1',
        '_C_("[data.xls]sheet1!A1")'),
    FormulaTest(
        '=(propellor_charts!B22*(propellor_charts!E21+propellor_charts!D21*'
        '(engine_data!O16*D70+engine_data!P16)+propellor_charts!C21*'
        '(engine_data!O16*D70+engine_data!P16)^2+propellor_charts!B21*'
        '(engine_data!O16*D70+engine_data!P16)^3)^2)^(1/3)*'
        '(1*D70/5.33E-18)^(2/3)*0.0000000001*28.3495231*9.81/1000',
        'propellor_charts!B22|propellor_charts!E21|propellor_charts!D21|'
        'engine_data!O16|D70|*|engine_data!P16|+|*|+|propellor_charts!C21|'
        'engine_data!O16|D70|*|engine_data!P16|+|2|^|*|+|propellor_charts!B21|'
        'engine_data!O16|D70|*|engine_data!P16|+|3|^|*|+|2|^|*|1|3|/|^|1|D70|*|'
        '5.33E-18|/|2|3|/|^|*|0.0000000001|*|28.3495231|*|9.81|*|1000|/',
        '((((((_C_("propellor_charts!B22") * ((((_C_("propellor_charts!E21")'
        ' + (_C_("propellor_charts!D21") * ((_C_("engine_data!O16") '
        '* _C_("D70")) + _C_("engine_data!P16")))) + ('
        '_C_("propellor_charts!C21") * (((_C_("engine_data!O16") * _C_("D70"))'
        ' + _C_("engine_data!P16")) ** 2))) + (_C_("propellor_charts!B21") '
        '* (((_C_("engine_data!O16") * _C_("D70")) + _C_("engine_data!P16")'
        ') ** 3))) ** 2)) ** (1 / 3)) * (((1 * _C_("D70")) / 5.33E-18) ** '
        '(2 / 3))) * 0.0000000001) * 28.3495231) * 9.81) / 1000'),
    FormulaTest(
        '=IF(configurations!$G$22=3,sizing!$C$303,M14)',
        'configurations!$G$22|3|=|sizing!$C$303|M14|IF',
        'x_if(_C_("configurations!G22") == 3, _C_("sizing!C303"), _C_("M14"))'),
    FormulaTest(
        '=TableX[[#This Row],[COL1]]&"-"&TableX[[#This Row],[COL2]]',
        'TableX[[#This Row],[COL1]]|"-"|&|TableX[[#This Row],[COL2]]|&',
        ''),
]

math_inputs = [
    FormulaTest(
        '=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))',
        '3600|1000|/|E40|*|E8|E39|/|*|E15|E19|/|*|E54|E54|E48|-|/|LN|*',
        '((((3600 / 1000) * _C_("E40")) * (_C_("E8") / _C_("E39"))) '
        '* (_C_("E15") / _C_("E19"))) * ln(_C_("E54") / (_C_("E54") '
        '- _C_("E48")))'),
    FormulaTest(
        '=0.000001042*E226^3-0.00004777*E226^2+0.0007646*E226-0.00075',
        '0.000001042|E226|3|^|*|0.00004777|E226|2|^|*|-|0.0007646|E226|*|'
        '+|0.00075|-',
        '(((0.000001042 * (_C_("E226") ** 3)) - (0.00004777 * '
        '(_C_("E226") ** 2))) + (0.0007646 * _C_("E226"))) - 0.00075'),
]

linest_inputs = [
    FormulaTest(
        '=LINEST(X5:X32,W5:W32^{1,2,3})',
        'X5:X32|W5:W32|1|2|3|ARRAYROW|ARRAY|^|LINEST',
        'linest(_R_("X5:X32"), _R_("W5:W32"), degree=-1)[-2]'),
    FormulaTest(
        '=LINEST(G2:G17,E2:E17,FALSE)',
        'G2:G17|E2:E17|FALSE|LINEST',
        'linest(_R_("G2:G17"), _R_("E2:E17"), False, degree=-1)[-2]'),
    FormulaTest(
        '=LINEST(B32:(INDEX(B32:B119,MATCH(0,B32:B119,-1),1)),(F32:(INDEX('
        'B32:F119,MATCH(0,B32:B119,-1),5)))^{1,2,3,4})',
        'B32|B32:B119|0|B32:B119|1|-|MATCH|1|INDEX||:|F32|B32:F119|0|'
        'B32:B119|1|-|MATCH|5|INDEX||:|1|2|3|4|ARRAYROW|ARRAY|^|LINEST',
        'linest(_R_(str(_REF_("B32") | (index(_REF_("B32:B119"),'
        ' match(0, _REF_("B32:B119"), -1), 1)))), (_R_(str(_REF_("F32") | '
        '(index(_REF_("B32:F119"), match(0, _REF_("B32:B119"), -1), 5))))), '
        'degree=-1)[-2]'),
    FormulaTest(
        '=LINESTMARIO(G2:G17,E2:E17,FALSE)',
        'G2:G17|E2:E17|FALSE|LINESTMARIO',
        'linestmario(_R_("G2:G17"), _R_("E2:E17"), False)[-2]'),
]

reference_inputs = [
    FormulaTest(
        '=ROW(4:7)',
        '4:7|ROW',
        'row(_REF_("4:7"))'),
    FormulaTest(
        '=ROW(D1:E1)',
        'D1:E1|ROW',
        'row(_REF_("D1:E1"))'),
    FormulaTest(
        '=COLUMN(D1:D2)',
        'D1:D2|COLUMN',
        'column(_REF_("D1:D2"))'),
    FormulaTest(
        '=ROW(D1:E2)',
        'D1:E2|ROW',
        'row(_REF_("D1:E2"))'),
    FormulaTest(
        '=ROW(B53:D54 C54:E54)',
        'B53:D54|C54:E54| |ROW',
        'row(_REF_("B53:D54") & _REF_("C54:E54"))'),
    FormulaTest(
        '=COLUMN(L45)',
        'L45|COLUMN',
        'column(_REF_("L45"))'),
]


def dump_test_case(formula, python_code, rpn):
    escaped_python_code = python_code.replace('\\', r'\\')

    print('    FormulaTest(')
    print("        '{}',".format(formula))
    print("        '{}',".format(rpn))
    print("        '{}'),".format(escaped_python_code))


def dump_parse(to_dump, ATestCell):
    cell = ATestCell('A', 1)

    print('[')
    for formula in to_dump:
        excel_formula = ExcelFormula(formula, cell=cell)
        parsed = excel_formula.rpn
        ast_root = excel_formula.ast
        result_rpn = "|".join(str(x) for x in parsed)
        try:
            result_python_code = ast_root.emit
        except:  # noqa: E722
            result_python_code = ''
        dump_test_case(formula, result_python_code, result_rpn)
    print(']')


test_names = (
    'range_inputs', 'basic_inputs', 'whitespace_inputs', 'if_inputs',
    'fancy_reference_inputs', 'math_inputs', 'linest_inputs',
    'reference_inputs',
)

test_data = []
for test_name in test_names:
    for i, test in enumerate(globals()[test_name]):
        test_data.append(
            ('{}_{}'.format(test_name, i + 1), test[0], test[1], test[2]))


def dump_all_test_cases():
    for name in test_names:
        print('{} = '.format(name), end='')
        dump_parse(t.formula for t in globals()[name])
        print()


@pytest.mark.parametrize('test_number, formula, rpn, python_code', test_data)
def test_tokenizer(test_number, formula, rpn, python_code):
    assert rpn == stringify_rpn(ExcelFormula(formula).rpn)


@pytest.mark.parametrize('test_number, formula, rpn, python_code', test_data)
def test_parse(test_number, formula, rpn, python_code, ATestCell):
    cell = ATestCell('A', 1)

    excel_formula = ExcelFormula(formula, cell=cell)
    parsed = excel_formula.rpn
    result_rpn = "|".join(str(x) for x in parsed)
    try:
        result_python_code = excel_formula.python_code
    except AttributeError as exc:
        # we have not mocked the excel table, so this test doesn't work
        if "no attribute 'table'" in str(exc):
            return
        raise

    assert result_python_code == excel_formula.ast.emit

    if (rpn, python_code) != (result_rpn, result_python_code):
        print("***Expected: ")
        dump_test_case(formula, python_code, rpn)

        print("***Result: ")
        dump_test_case(formula, result_python_code, result_rpn)

        print('--------------')

    assert python_code == result_python_code


def test_table_relative_address(ATestCell):
    cell = ATestCell('A', 1, sheet='s')

    excel_formula = ExcelFormula('=junk')
    assert '"#NAME?"' == excel_formula.ast.emit

    excel_formula = ExcelFormula('=junk', cell=cell)
    assert '"#NAME?"' == excel_formula.ast.emit

    excel_formula = ExcelFormula('=[col1]', cell=cell)
    assert '"#NAME?"' == excel_formula.ast.emit

    with mock.patch.object(cell, 'excel') as excel, \
            mock.patch.object(excel, 'table') as get_table, \
            mock.patch.object(excel, 'table_name_containing') as etnc:
        excel.defined_names = {}
        table = mock.Mock()
        table.ref = 'A1:B2'
        table.headerRowCount = 0
        table.totalsRowCount = 0
        table.tableColumns = [mock.Mock()]
        table.tableColumns[0].name = 'col1'

        get_table.return_value = table, 's'
        etnc.return_value = 'Table'

        excel_formula = ExcelFormula('=[col1]', cell=cell)
        assert '_R_("s!A1:A2")' == excel_formula.ast.emit


def test_multi_area_ranges(excel, ATestCell):
    cell = ATestCell('A', 1, excel=excel)

    with mock.patch.object(excel, '_defined_names', {
            'dname': (('$A$1', 's1'), ('$A$3:$A$4', 's2'))}):
        excel_formula = ExcelFormula('=sum(dname)', cell=cell)
        assert excel_formula.ast.emit == 'xsum(_C_("s1!A1"), _R_("s2!A3:A4"))'


def test_str():
    excel_formula = ExcelFormula('=E54-E48')
    assert '=E54-E48' == str(excel_formula)

    assert '_C_("E54") - _C_("E48")' == excel_formula.python_code
    excel_formula.base_formula = None
    assert '_C_("E54") - _C_("E48")' == str(excel_formula)

    excel_formula._ast = None
    excel_formula._rpn = None
    excel_formula._python_code = None
    assert '' == str(excel_formula)


def test_descendants():

    excel_formula = ExcelFormula('=E54-E48')
    descendants = excel_formula.ast.descendants
    assert descendants == excel_formula.ast.descendants

    assert 2 == len(descendants)
    assert 'OPERAND' == descendants[0][0].type
    assert 'OPERAND' == descendants[1][0].type
    assert {'E48', 'E54'} == {
        descendants[0][0].value, descendants[1][0].value
    }


def test_ast_node():
    with pytest.raises(FormulaParserError):
        ASTNode.create(Token('a_value', None, None))

    node = ASTNode(Token('a_value', None, None))
    assert 'ASTNode<a_value>' == repr(node)
    assert 'a_value' == str(node)
    assert 'a_value' == node.emit


def test_if_args_error():
    eval_context = ExcelFormula.build_eval_context(lambda x: 1, lambda x: 1)

    assert 0 == eval_context(ExcelFormula('=if(1,0)'))
    assert VALUE_ERROR == eval_context(ExcelFormula('=if(#VALUE!,1)'))
    assert VALUE_ERROR == eval_context(ExcelFormula('=if(#VALUE!,1,0)'))
    assert VALUE_ERROR == eval_context(ExcelFormula('=if(1,#VALUE!,0)'))
    assert VALUE_ERROR == eval_context(ExcelFormula('=if(0,1,#VALUE!)'))


@pytest.mark.parametrize(
    'formula', (
        '=if(1',
        '=G11;',
        '=G11,',
        '=(G11;',
        '=;',
        '=,',
        '=-',
        '=--4',
    )
)
def test_parser_error(formula):
    with pytest.raises(FormulaParserError):
        ExcelFormula(formula).ast


def test_needed_addresses():

    formula = '=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))'
    needed = sorted(('E40', 'E8', 'E39', 'E15', 'E19', 'E54', 'E48'))

    excel_formula = ExcelFormula(formula)

    assert needed == sorted(x.address for x in excel_formula.needed_addresses)
    assert needed == sorted(x.address for x in excel_formula.needed_addresses)

    assert () == ExcelFormula('').needed_addresses

    excel_formula = ExcelFormula('_REF_(_R_("S!A1"))',
                                 formula_is_python_code=True)
    assert excel_formula.needed_addresses == (AddressCell('S!A1'), )


@pytest.mark.parametrize(
    'result, formula', (
        (42, '=2 * 21'),
        (44, '=2 * 21 + A1 + a1:a2'),
        (1, '=1 + sin(0)'),
        (4.1415926, '=1 + PI()'),
    )
)
def test_build_eval_context(result, formula):
    eval_context = ExcelFormula.build_eval_context(lambda x: 1, lambda x: 1)
    assert eval_context(ExcelFormula(formula)) == pytest.approx(result)


def test_math_wrap():
    eval_context = ExcelFormula.build_eval_context(
        lambda x: None, lambda x: DIV0)

    assert 1 == eval_context(ExcelFormula('=1 + sin(A1)'))
    assert DIV0 == eval_context(ExcelFormula('=1 + sin(A1:B1)'))

    assert 1 == eval_context(ExcelFormula('=1 + abs(A1)'))
    assert DIV0 == eval_context(ExcelFormula('=1 + abs(A1:B1)'))


def test_compiled_python_cache():
    formula = ExcelFormula('=1 + 2')
    # first call does the calc, the second uses cached
    compiled_python = formula.compiled_python
    assert compiled_python == formula.compiled_python

    # rebuild from marshalled
    formula._compiled_python = None
    assert compiled_python == formula.compiled_python

    # invalidate the marshalled code, rebuild from source
    formula._compiled_python = None
    formula._marshalled_python = 'junk'
    assert compiled_python == formula.compiled_python


def test_compiled_python_error():
    formula = ExcelFormula('=1 + 2')
    formula._python_code = 'this will be a syntax error'
    with pytest.raises(FormulaParserError,
                       match='Failed to compile expression'):
        formula.compiled_python


def test_save_to_file(fixture_dir):
    formula = ExcelFormula('=1+2')
    filename = os.path.join(fixture_dir, 'formula_save_test.pickle')
    with open(filename, 'wb') as f:
        pickle.dump(formula, f)

    with open(filename, 'rb') as f:
        loaded_formula = pickle.load(f)

    os.unlink(filename)

    assert formula.python_code == loaded_formula.python_code


def test_get_linest_degree_with_cell(ATestCell):
    with mock.patch('pycel.excelformula.get_linest_degree') as get:
        get.return_value = -1, -1

        cell = ATestCell('A', 1, 'Phony Sheet')
        formula = ExcelFormula('=linest(C1)', cell=cell)

        expected = 'linest(_C_("Phony Sheet!C1"), degree=-1)[-2]'
        assert expected == formula.python_code


def test_init_from_python_code():
    excel_formula1 = ExcelFormula('=B32:B119 + P5')
    assert '_R_("B32:B119") + _C_("P5")' == \
        excel_formula1.python_code

    python_code = '=_R_("B32:B119") + _C_("P5")'
    excel_formula2 = ExcelFormula(python_code, formula_is_python_code=True)
    assert excel_formula1.needed_addresses == excel_formula2.needed_addresses


@pytest.mark.parametrize(
    'formula, result', (
        ('=(1=1.0)+("1"=1)+(1="1")', 1),
        ('=("1"="1") + ("x"=1)', 1),
        ('=if("x"<>"x", "a", "b")', 'b'),
    )
)
def test_string_number_compare(formula, result, empty_eval_context):
    assert empty_eval_context(ExcelFormula(formula)) == result


@pytest.mark.parametrize(
    'formula, result', (
        ('=TRUE%', 0.01),
        ('=FALSE%', 0),

        ('=TRUE+5', 6),
        ('=FALSE+5', 5),
        ('=TRUE*5', 5),
        ('=FALSE*5', 0),

        ('=TRUE&"xyzzy"', 'TRUExyzzy'),
        ('=FALSE&"xyzzy"', 'FALSExyzzy'),
    )
)
def test_bool_ops(formula, result):
    eval_ctx = ExcelFormula.build_eval_context(lambda x: None, None)
    assert eval_ctx(ExcelFormula(formula)) == result


@pytest.mark.parametrize(
    'formula, result', (
        ('=NOT(FALSE)', True),
        ('=NOT(TRUE)', False),
        ('=OR(FALSE, FALSE)', False),
        ('=OR(TRUE, FALSE)', True),
        ('=AND(TRUE, FALSE)', False),
        ('=AND(TRUE, TRUE)', True),
        ('=XOR(FALSE, FALSE)', False),
        ('=XOR(TRUE, FALSE)', True),
        ('=XOR(FALSE, TRUE)', True),
        ('=XOR(TRUE, TRUE)', False),
        ('=FALSE()', False),
        ('=TRUE()', True),
    )
)
def test_bool_funcs(formula, result):
    eval_ctx = ExcelFormula.build_eval_context(lambda x: None, None)
    assert eval_ctx(ExcelFormula(formula)) == result


@pytest.mark.parametrize(
    'formula, result', (
        ('=(A1=0) + (A1=1)', 1),
        ('=(A1<0)+(A1<=0)+(A1=0)+(A1>=0)+(A1>0)', 3),
    )
)
def test_empty_cell_logic_op(formula, result):
    eval_ctx = ExcelFormula.build_eval_context(lambda x: None, None)
    assert eval_ctx(ExcelFormula(formula)) == result


@pytest.mark.parametrize(
    'expected, formula', (
        (-1, '=-1'),
        (1, '=+1'),
        (1, '=-1+2'),
        (3, '=+1+2'),
        (-3, '=-1-2'),
        (-1, '=+1-2'),
        (-3, '=-(1+2)'),
        (3, '=+(1+2)'),
        (1, '=-(1-2)'),
        (-1, '=+(1-2)'),

        (3, '=+sum(+1, 2)'),
        (-3, '=-sum(1, +2)'),
        (1, '=+sum(-1, 2)'),
        (1, '=-sum(1, -2)'),
        (1, '=+sum(+1, "-2")'),
        (-1, '=-sum(1, "+2")'),
    )
)
def test_unary_ops(expected, formula, empty_eval_context):
    assert expected == empty_eval_context(ExcelFormula(formula))


@pytest.mark.parametrize(
    'formula, result', (
        ('=1+2+"4"', 7),
        ('=sum(1, 2, "4")', 3),
        ('=3&"A"', '3A'),
        ('=3.0&"A"', '3A'),
        ('=A1&"A"', '3A'),
    )
)
def test_numerics_type_coercion(formula, result):
    eval_ctx = ExcelFormula.build_eval_context(lambda x: 3.0, None)
    assert eval_ctx(ExcelFormula(formula)) == result


@pytest.mark.parametrize(
    'formula, result', (
        ('=1="a"', False),
        ('=1=2', False),
        ('="a"="b"', False),
        ('=1=1', True),
        ('="A"="a"', True),
        ('="a"="A"', True),
    )
)
def test_string_compare(formula, result, empty_eval_context):
    assert empty_eval_context(ExcelFormula(formula)) == result


@pytest.mark.parametrize(
    'formula, result', (
        ('=2*3&"A"', '6A'),
        ('=1&"a"', '1a'),
        ('="1"&2', '12'),
        ('="a"&"b"', 'ab'),
        ('=1&1', '11'),
        ('="A"&"a"', 'Aa'),
        ('="a"&"A"', 'aA'),
    )
)
def test_string_concat(formula, result, empty_eval_context):
    assert empty_eval_context(ExcelFormula(formula)) == result


@pytest.mark.parametrize(
    'formula, result, cell', (
        ('=COLUMN(L45)', 12, None),
        ('=COLUMN(B:E)', ((2, 3, 4, 5),), None),
        ('=COLUMN(4:7)', range(1, 16385), None),
        ('=COLUMN(D1:E1)', ((4, 5),), None),
        ('=COLUMN(D1:D2)', ((4,),), None),
        ('=COLUMN(D1:E2)', ((4, 5),), None),
        ('=COLUMN()', 2, "ATestCell('B', 3)"),
        ('=COLUMN(B6:D9 C7:E8)', ((3, 4), ), None),
        ('=COLUMN(B6:D9 E7:F8)', NULL_ERROR, None),
    )
)
def test_column(formula, result, cell, empty_eval_context, ATestCell):
    if cell is not None:
        cell = eval(cell)
    assert empty_eval_context(ExcelFormula(formula, cell=cell)) == result


@pytest.mark.parametrize(
    'formula, result, cell', (
        ('=ROW(L45)', 45, None),
        ('=ROW(B:E)', range(1, 1048577), None),
        ('=ROW(4:7)', ((4,), (5,), (6,), (7,)), None),
        ('=ROW(D1:E1)', ((1,), ), None),
        ('=ROW(D1:D2)', ((1,), (2,)), None),
        ('=ROW(D1:E2)', ((1,), (2,)), None),
        ('=ROW()', 3, "ATestCell('B', 3)"),
        ('=ROW(B6:D9 C7:E8)', ((7,), (8,)), None),
        ('=ROW(B6:D9 E7:F8)', NULL_ERROR, None),
    )
)
def test_row(formula, result, cell, empty_eval_context, ATestCell):
    if cell is not None:
        cell = eval(cell)
    assert empty_eval_context(ExcelFormula(formula, cell=cell)) == result


@pytest.mark.parametrize(
    'formula, result', (
        ('=subtotal(01,A1:B3)', 'average(_R_("A1:B3"))'),
        ('=subtotal(02,A1:B3)', 'count(_R_("A1:B3"))'),
        ('=subtotal(03,A1:B3)', 'counta(_R_("A1:B3"))'),
        ('=subtotal(04,A1:B3)', 'xmax(_R_("A1:B3"))'),
        ('=subtotal(05,A1:B3)', 'xmin(_R_("A1:B3"))'),
        ('=subtotal(06,A1:B3)', 'product(_R_("A1:B3"))'),
        ('=subtotal(07,A1:B3)', 'stdev(_R_("A1:B3"))'),
        ('=subtotal(08,A1:B3)', 'stdevp(_R_("A1:B3"))'),
        ('=subtotal(09,A1:B3)', 'xsum(_R_("A1:B3"))'),
        ('=subtotal(10,A1:B3)', 'var(_R_("A1:B3"))'),
        ('=subtotal(11,A1:B3)', 'varp(_R_("A1:B3"))'),
    )
)
def test_subtotal(formula, result):
    assert ExcelFormula(formula).ast.emit == result
    assert ExcelFormula(formula.replace('tal(', 'tal(1')).ast.emit == result


@pytest.mark.parametrize(
    'formula', (
        '=subtotal(0)',
        '=subtotal(12)',
        '=subtotal(100)',
        '=subtotal(112)',
    )
)
def test_subtotal_errors(formula):
    with pytest.raises(ValueError):
        ExcelFormula(formula).ast.emit


def test_plugins():
    with mock.patch('pycel.excelformula.ExcelFormula.default_modules', ()):
        eval_ctx = ExcelFormula.build_eval_context(None, None)
        with pytest.raises(UnknownFunction):
            eval_ctx(ExcelFormula('=sum({1,2,3})'))

    with mock.patch('pycel.excelformula.ExcelFormula.default_modules', ()):
        eval_ctx = ExcelFormula.build_eval_context(
            None, None, plugins=('pycel.excellib', ))
        assert eval_ctx(ExcelFormula('=sum({1,2,3})')) == 6

    with mock.patch('pycel.excelformula.ExcelFormula.default_modules', ()):
        eval_ctx = ExcelFormula.build_eval_context(
            None, None, plugins='pycel.excellib')
        assert eval_ctx(ExcelFormula('=sum({1,2,3})')) == 6

    with mock.patch('pycel.excelformula.ExcelFormula.default_modules',
                    ('pycel.excellib', )):
        eval_ctx = ExcelFormula.build_eval_context(None, None)
        assert eval_ctx(ExcelFormula('=sum({1,2,3})')) == 6


def test_unknown_name(empty_eval_context):
    assert NAME_ERROR == empty_eval_context(ExcelFormula('=CE'))


@pytest.mark.parametrize(
    'formula', (
        '=sum(A1)',
        '=sum(A1:B2)',
        '=a1=1',
        '=a1+"l"',
        '=1 - (1 / 0)',
    )
)
def test_div_zero(formula):
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: DIV0, lambda x: [[1, 1], [1, DIV0]],
        logging.getLogger('pycel_x'))
    assert eval_ctx(ExcelFormula(formula)) == DIV0


def test_error_logging(caplog):
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: DIV0, lambda x: [[1, 1], [1, DIV0]],
        logging.getLogger('pycel_x'))

    caplog.set_level(logging.INFO)
    assert 3 == eval_ctx(ExcelFormula('=iferror(1/0,3)'))
    assert 1 == len(caplog.records)
    assert "INFO" == caplog.records[0].levelname
    assert "1 Div 0" in caplog.records[0].message

    assert DIV0 == eval_ctx(ExcelFormula('=1/0'))
    assert 2 == len(caplog.records)
    assert "WARNING" == caplog.records[1].levelname

    message = """return PYTHON_AST_OPERATORS[op](left_op, right_op)
ZeroDivisionError: division by zero
Eval: 1 / 0
Values: 1 Div 0"""
    assert message in caplog.records[1].message


@pytest.mark.parametrize(
    'formula, result', (
        ('=sum(A1)', VALUE_ERROR),
        ('=sum(A1:B2)', VALUE_ERROR),
        ('=a1=1', VALUE_ERROR),
        ('=a1+"l"', VALUE_ERROR),
        ('=iferror(1+"A",3)', 3),
        ('=iferror(1+"A",)', 0),
        ('=1+"A"', VALUE_ERROR),
    )
)
def test_value_error(formula, result):
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: VALUE_ERROR, lambda x: [[1, 1], [1, VALUE_ERROR]],
        logging.getLogger('pycel_x'))

    assert eval_ctx(ExcelFormula(formula)) == result


def test_eval_exception():
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: 1 + 'a', lambda x: [[1, 1], [1, DIV0]],
        logging.getLogger('pycel'))

    with pytest.raises(FormulaEvalError):
        eval_ctx(ExcelFormula('=a1'))


def test_lineno_on_error_reporting(empty_eval_context):
    excel_formula = ExcelFormula('')
    excel_formula._python_code = 'X'
    excel_formula.lineno = 6
    excel_formula.filename = 'a_file'

    msg = 'File "a_file", line 6,'
    with pytest.raises(UnknownFunction, match=msg):
        empty_eval_context(excel_formula)

    excel_formula._python_code = '(x)'
    excel_formula._compiled_python = None
    excel_formula._marshalled_python = None
    excel_formula.compiled_lambda = None
    excel_formula.lineno = 60

    with pytest.raises(UnknownFunction, match='File "a_file", line 60,'):
        empty_eval_context(excel_formula)


@pytest.mark.parametrize(
    'msg, formula', (
        ("Function XYZZY is not implemented. "
         "XYZZY is not a known Excel function", '=xyzzy()'),
        ("Function PLUGH is not implemented. "
         "PLUGH is not a known Excel function\n"
         "Function XYZZY is not implemented. "
         "XYZZY is not a known Excel function", '=xyzzy() + plugh()'),
        ('Function ARABIC is not implemented. '
         'ARABIC is in the "Math and trigonometry" group, '
         'and was introduced in Excel 2013',
         '=ARABIC()'),
    )
)
def test_unknown_function(msg, formula, empty_eval_context):
    compiled = ExcelFormula(formula)

    with pytest.raises(UnknownFunction, match=msg):
        empty_eval_context(compiled)

    # second time is needed to test cached msg
    with pytest.raises(UnknownFunction, match=msg):
        empty_eval_context(compiled)


if __name__ == '__main__':
    dump_parse()
