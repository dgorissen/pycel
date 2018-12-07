import logging
import os
import pickle
from unittest import mock

import pytest

from pycel.excelformula import (
    ASTNode,
    FormulaEvalError,
    ExcelFormula,
    FormulaParserError,
    Token,
)
from pycel.excelutil import DIV0, VALUE_ERROR

from test_excelutil import ATestCell


def stringify(e):
    return "|".join([str(x) for x in e])


range_inputs = [
    ('=$A1',
     '$A1'),

    ('=$B$2',
     '$B$2'),

    ('=SUM(B5:B15)',
     'B5:B15|SUM'),

    ('=SUM(B5:B15,D5:D15)',
     'B5:B15|D5:D15|SUM'),

    ('=SUM(B5:B15 A7:D7)',
     'B5:B15|A7:D7| |SUM'),

    ('=SUM((A:A,1:1))',
     'A:A|1:1|,|SUM'),

    ('=SUM((A:A A1:B1))',
     'A:A|A1:B1| |SUM'),

    ('=SUM(D9:D11,E9:E11,F9:F11)',
     'D9:D11|E9:E11|F9:F11|SUM'),

    ('=SUM((D9:D11,(E9:E11,F9:F11)))',
     'D9:D11|E9:E11|F9:F11|,|,|SUM'),

    ('={SUM(B2:D2*B3:D3)}',
     'B2:D2|B3:D3|*|SUM|ARRAYROW|ARRAY'),

    ('=SUM(123 + SUM(456) + (45<6))+456+789',
     '123|456|SUM|+|45|6|<|+|SUM|456|+|789|+'),

    ('=AVG(((((123 + 4 + AVG(A1:A2))))))',
     '123|4|+|A1:A2|AVG|+|AVG'),
]

basic_inputs = [
    ('=SUM((A:A 1:1))',
     'A:A|1:1| |SUM'),

    ('=A1',
     'A1'),

    ('=50',
     '50'),

    ('=1+1',
     '1|1|+'),

    ('=atan2(A1,B1)',
     'A1|B1|atan2'),

    ('=5*log(sin()+2)',
     '5|sin|2|+|log|*'),

    ('=5*log(sin(3,7,9)+2)',
     '5|3|7|9|sin|2|+|log|*'),
]

whitespace_inputs = [
    ('=3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3',
     '3|4|2|*|1|5|-|2|^|3|^|/|+'),

    ('=1+3+5',
     '1|3|+|5|+'),

    ('=3 * 4 + 5',
     '3|4|*|5|+'),
]

if_inputs = [
    (
    '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
    '"a"|"a"|"b"|ARRAYROW|"c"|#N/A|ARRAYROW|1|-|TRUE|ARRAYROW|ARRAY|=|"yes"|"no"|IF|"  more ""test"" text"|&'),

    (
    '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
    'R13C3|2002|1|6|DATE|>|0|R[41]C[2]|ISERROR|0|R13C3|R[41]C[2]|>=|0|R[23]C[11]|55|>=|R[24]C[11]|20|>=|AND|R53C3|0|IF|IF|IF|IF'),

    (
    '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))',
    'R[39]C[11]|65|>|R[25]C[42]|R[11]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|R[44]C[11]|R[43]C[11]|IF|*|R[14]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|R[45]C[11]|R[43]C[11]|IF|*|+|0|ROUND|IF'),

    ('=IF(AI119="","",E119)',
     'AI119|""|=|""|E119|IF'),

    (
    '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
    'P5|1.0|=|"NA"|P5|2.0|=|"A"|P5|3.0|=|"B"|P5|4.0|=|"C"|P5|5.0|=|"D"|P5|6.0|=|"E"|P5|7.0|=|"F"|P5|8.0|=|"G"|IF|IF|IF|IF|IF|IF|IF|IF'),

]

fancy_reference_inputs = [
    ('=SUM(sheet1!$A$1:$B$2)',
     'sheet1!$A$1:$B$2|SUM'),

    ('=[data.xls]sheet1!$A$1',
     '[data.xls]sheet1!$A$1'),

    (
    '=(propellor_charts!B22*(propellor_charts!E21+propellor_charts!D21*(engine_data!O16*D70+engine_data!P16)+propellor_charts!C21*(engine_data!O16*D70+engine_data!P16)^2+propellor_charts!B21*(engine_data!O16*D70+engine_data!P16)^3)^2)^(1/3)*(1*D70/5.33E-18)^(2/3)*0.0000000001*28.3495231*9.81/1000',
    'propellor_charts!B22|propellor_charts!E21|propellor_charts!D21|engine_data!O16|D70|*|engine_data!P16|+|*|+|propellor_charts!C21|engine_data!O16|D70|*|engine_data!P16|+|2|^|*|+|propellor_charts!B21|engine_data!O16|D70|*|engine_data!P16|+|3|^|*|+|2|^|*|1|3|/|^|1|D70|*|5.33E-18|/|2|3|/|^|*|0.0000000001|*|28.3495231|*|9.81|*|1000|/'),

    ('=IF(configurations!$G$22=3,sizing!$C$303,M14)',
     'configurations!$G$22|3|=|sizing!$C$303|M14|IF'),

    ('=TableX[[#This Row],[COL1]]&"-"&TableX[[#This Row],[COL2]]',
     'TableX[[#This Row],[COL1]]|"-"|&|TableX[[#This Row],[COL2]]|&'),
]

math_inputs = [
    ('=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))',
     '3600|1000|/|E40|*|E8|E39|/|*|E15|E19|/|*|E54|E54|E48|-|/|LN|*'),

    ('=0.000001042*E226^3-0.00004777*E226^2+0.0007646*E226-0.00075',
     '0.000001042|E226|3|^|*|0.00004777|E226|2|^|*|-|0.0007646|E226|*|+|0.00075|-'),
]

linest_inputs = [
    ('=LINEST(X5:X32,W5:W32^{1,2,3})',
     'X5:X32|W5:W32|1|2|3|ARRAYROW|ARRAY|^|LINEST'),

    ('=LINEST(G2:G17,E2:E17,FALSE)',
     'G2:G17|E2:E17|FALSE|LINEST'),
    (
    '=LINEST(B32:(INDEX(B32:B119,MATCH(0,B32:B119,-1),1)),(F32:(INDEX(B32:F119,MATCH(0,B32:B119,-1),5)))^{1,2,3,4})',
    'B32:B119|0|B32:B119|1|-|MATCH|1|INDEX|B32:|B32:F119|0|B32:B119|1|-|MATCH|5|INDEX|F32:|1|2|3|4|ARRAYROW|ARRAY|^|LINEST'),
]


test_names = (
    'range_inputs', 'if_inputs', 'whitespace_inputs', 'basic_inputs',
    'math_inputs', 'linest_inputs', 'fancy_reference_inputs')
test_data = []
for test_name in test_names:
    for i, test in enumerate(globals()[test_name]):
        test_data.append(('{}_{}'.format(test_name, i + 1), *test))


@pytest.mark.parametrize('test_number, formula, rpn', test_data)
def test_tokenizer(test_number, formula, rpn):
    assert rpn == stringify(ExcelFormula(formula).rpn)


test_data = [
    dict(
        formula='=SUM((A:A 1:1))',
        rpn='A:A|1:1| |SUM',
        python_code='xsum(_R_("A:A") + _R_("1:1"))',
    ),
    dict(
        formula='=A1',
        rpn='A1',
        python_code='_C_("A1")',
    ),
    dict(
        formula='="x"="y"',
        rpn='"x"|"y"|=',
        python_code='"x" == "y"',
    ),
    dict(
        formula='="x"=1',
        rpn='"x"|1|=',
        python_code='"x" == 1',
    ),
    dict(
        formula='=atan2(A1,B1)',
        rpn='A1|B1|atan2',
        python_code='atan2(_C_("B1"), _C_("A1"))',
    ),
    dict(
        formula='=5*log(sin()+2)',
        rpn='5|sin|2|+|log|*',
        python_code='5 * log(sin() + 2)',
    ),
    dict(
        formula='=5*log(sin(3,7,9)+2)',
        rpn='5|3|7|9|sin|2|+|log|*',
        python_code='5 * log(sin(3, 7, 9) + 2)',
    ),
    dict(
        formula='=3 +1-5',
        rpn='3|1|+|5|-',
        python_code='(3 + 1) - 5',
    ),
    dict(
        formula='=3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3',
        rpn='3|4|2|*|1|5|-|2|^|3|^|/|+',
        python_code='3 + ((4 * 2) / (((1 - 5) ** 2) ** 3))',
    ),
    dict(
        formula='=1+3+5',
        rpn='1|3|+|5|+',
        python_code='(1 + 3) + 5',
    ),
    dict(
        formula='=3 + 4 * 5',
        rpn='3|4|5|*|+',
        python_code='3 + (4 * 5)',
    ),
    dict(
        formula='=3 * 4 + 5',
        rpn='3|4|*|5|+',
        python_code='(3 * 4) + 5',
    ),
    dict(
        formula='=50',
        rpn='50',
        python_code='50',
    ),
    dict(
        formula='=1+1',
        rpn='1|1|+',
        python_code='1 + 1',
    ),
    dict(
        formula='=$A1',
        rpn='$A1',
        python_code='_C_("A1")',
    ),
    dict(
        formula='=$B$2',
        rpn='$B$2',
        python_code='_C_("B2")',
    ),
    dict(
        formula='=PI()',
        rpn='PI',
        python_code='pi',
    ),
    dict(
        formula='=SUM(B5:B15)',
        rpn='B5:B15|SUM',
        python_code='xsum(_R_("B5:B15"))',
    ),
    dict(
        formula='=SUM(B5:B15,D5:D15)',
        rpn='B5:B15|D5:D15|SUM',
        python_code='xsum(_R_("B5:B15"), _R_("D5:D15"))',
    ),
    dict(
        formula='=SUM(B5:B15 A7:D7)',
        rpn='B5:B15|A7:D7| |SUM',
        python_code='xsum(_R_("B5:B15") + _R_("A7:D7"))',
    ),
    dict(
        formula='=SUM(sheet1!$A$1:$B$2)',
        rpn='sheet1!$A$1:$B$2|SUM',
        python_code='xsum(_R_("sheet1!A1:B2"))',
    ),
    dict(
        formula='=[data.xls]sheet1!$A$1',
        rpn='[data.xls]sheet1!$A$1',
        python_code='_C_("[data.xls]sheet1!A1")',
    ),
    dict(
        formula='=SUM((A:A,1:1))',
        rpn='A:A|1:1|,|SUM',
        python_code='xsum(_R_("A:A"), _R_("1:1"))',
    ),
    dict(
        formula='=SUM((A:A A1:B1))',
        rpn='A:A|A1:B1| |SUM',
        python_code='xsum(_R_("A:A") + _R_("A1:B1"))',
    ),
    dict(
        formula='=SUM(D9:D11,E9:E11,F9:F11)',
        rpn='D9:D11|E9:E11|F9:F11|SUM',
        python_code='xsum(_R_("D9:D11"), _R_("E9:E11"), _R_("F9:F11"))',
    ),
    dict(
        formula='=SUM((D9:D11,(E9:E11,F9:F11)))',
        rpn='D9:D11|E9:E11|F9:F11|,|,|SUM',
        python_code='xsum(_R_("D9:D11"), (_R_("E9:E11"), _R_("F9:F11")))',
    ),
    dict(
        formula='=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
        rpn='P5|1.0|=|"NA"|P5|2.0|=|"A"|P5|3.0|=|"B"|P5|4.0|=|"C"|P5|5.0|=|"D"|P5|6.0|=|"E"|P5|7.0|=|"F"|P5|8.0|=|"G"|IF|IF|IF|IF|IF|IF|IF|IF',
        python_code='("NA" if _C_("P5") == 1.0 else ("A" if _C_("P5") == 2.0 else ("B" if _C_("P5") == 3.0 else ("C" if _C_("P5") == 4.0 else ("D" if _C_("P5") == 5.0 else ("E" if _C_("P5") == 6.0 else ("F" if _C_("P5") == 7.0 else ("G" if _C_("P5") == 8.0 else 0))))))))',
    ),
    dict(
        formula='={SUM(B2:D2*B3:D3)}',
        rpn='B2:D2|B3:D3|*|SUM|ARRAYROW|ARRAY',
        python_code='[xsum(_R_("B2:D2") * _R_("B3:D3"))]',
    ),
    dict(
        formula='=SUM(123 + SUM(456) + (45<6))+456+789',
        rpn='123|456|SUM|+|45|6|<|+|SUM|456|+|789|+',
        python_code='(xsum((123 + xsum(456)) + (45 < 6)) + 456) + 789',
    ),
    dict(
        formula='=AVG(((((123 + 4 + AVG(A1:A2))))))',
        rpn='123|4|+|A1:A2|AVG|+|AVG',
        python_code='avg((123 + 4) + avg(_R_("A1:A2")))',
    ),

    # E. W. Bachtal's test formulae
    dict(
        formula='=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
        rpn='"a"|"a"|"b"|ARRAYROW|"c"|#N/A|ARRAYROW|1|-|TRUE|ARRAYROW|ARRAY|=|"yes"|"no"|IF|"  more ""test"" text"|&',
        python_code='("yes" if "a" == [["a", "b"], ["c", "#N/A"], [-1, True]] else "no") & "  more \\"test\\" text"',
    ),
    dict(
        formula='=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
        rpn='R13C3|2002|1|6|DATE|>|0|R[41]C[2]|ISERROR|0|R13C3|R[41]C[2]|>=|0|R[23]C[11]|55|>=|R[24]C[11]|20|>=|AND|R53C3|0|IF|IF|IF|IF',
        python_code='(0 if _C_("C13") > date(2002, 1, 6) else (0 if iserror(_C_("C42")) else (0 if _C_("C13") >= _C_("C42") else (_C_("C53") if all((_C_("L24") >= 55, _C_("L25") >= 20,)) else 0))))',    ),
    dict(
        formula='=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))',
        rpn='R[39]C[11]|65|>|R[25]C[42]|R[11]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|R[44]C[11]|R[43]C[11]|IF|*|R[14]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|R[45]C[11]|R[43]C[11]|IF|*|+|0|ROUND|IF',
        python_code='(_C_("AQ26") if _C_("L40") > 65 else xround((_C_("L12") * (_C_("L45") if any((all((_C_("L40") >= 55, _C_("L41") >= 20,)), all((_C_("L41") >= 20, _C_("C11") == "YES",)),)) else _C_("L44"))) + (_C_("L15") * (_C_("L46") if any((all((_C_("L40") >= 55, _C_("L41") >= 20,)), all((_C_("L41") >= 20, _C_("C11") == "YES",)),)) else _C_("L44"))), 0))',
    ),
    dict(
        formula='=(propellor_charts!B22*(propellor_charts!E21+propellor_charts!D21*(engine_data!O16*D70+engine_data!P16)+propellor_charts!C21*(engine_data!O16*D70+engine_data!P16)^2+propellor_charts!B21*(engine_data!O16*D70+engine_data!P16)^3)^2)^(1/3)*(1*D70/5.33E-18)^(2/3)*0.0000000001*28.3495231*9.81/1000',
        rpn='propellor_charts!B22|propellor_charts!E21|propellor_charts!D21|engine_data!O16|D70|*|engine_data!P16|+|*|+|propellor_charts!C21|engine_data!O16|D70|*|engine_data!P16|+|2|^|*|+|propellor_charts!B21|engine_data!O16|D70|*|engine_data!P16|+|3|^|*|+|2|^|*|1|3|/|^|1|D70|*|5.33E-18|/|2|3|/|^|*|0.0000000001|*|28.3495231|*|9.81|*|1000|/',
        python_code='((((((_C_("propellor_charts!B22") * ((((_C_("propellor_charts!E21") + (_C_("propellor_charts!D21") * ((_C_("engine_data!O16") * _C_("D70")) + _C_("engine_data!P16")))) + (_C_("propellor_charts!C21") * (((_C_("engine_data!O16") * _C_("D70")) + _C_("engine_data!P16")) ** 2))) + (_C_("propellor_charts!B21") * (((_C_("engine_data!O16") * _C_("D70")) + _C_("engine_data!P16")) ** 3))) ** 2)) ** (1 / 3)) * (((1 * _C_("D70")) / 5.33E-18) ** (2 / 3))) * 0.0000000001) * 28.3495231) * 9.81) / 1000',
    ),
    dict(
        formula='=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))',
        rpn='3600|1000|/|E40|*|E8|E39|/|*|E15|E19|/|*|E54|E54|E48|-|/|LN|*',
        python_code='((((3600 / 1000) * _C_("E40")) * (_C_("E8") / _C_("E39"))) * (_C_("E15") / _C_("E19"))) * xlog(_C_("E54") / (_C_("E54") - _C_("E48")))',
    ),
    dict(
        formula='=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
        rpn='P5|1.0|=|"NA"|P5|2.0|=|"A"|P5|3.0|=|"B"|P5|4.0|=|"C"|P5|5.0|=|"D"|P5|6.0|=|"E"|P5|7.0|=|"F"|P5|8.0|=|"G"|IF|IF|IF|IF|IF|IF|IF|IF',
        python_code='("NA" if _C_("P5") == 1.0 else ("A" if _C_("P5") == 2.0 else ("B" if _C_("P5") == 3.0 else ("C" if _C_("P5") == 4.0 else ("D" if _C_("P5") == 5.0 else ("E" if _C_("P5") == 6.0 else ("F" if _C_("P5") == 7.0 else ("G" if _C_("P5") == 8.0 else 0))))))))',
    ),
    dict(
        formula='=LINEST(X5:X32,W5:W32^{1,2,3})',
        rpn='X5:X32|W5:W32|1|2|3|ARRAYROW|ARRAY|^|LINEST',
        python_code='linest(_R_("X5:X32"), _R_("W5:W32"), degree=-1)[-2]',
    ),
    dict(
        formula='=IF(configurations!$G$22=3,sizing!$C$303,M14)',
        rpn='configurations!$G$22|3|=|sizing!$C$303|M14|IF',
        python_code='(_C_("sizing!C303") if _C_("configurations!G22") == 3 else _C_("M14"))',
    ),
    dict(
        formula='=0.000001042*E226^3-0.00004777*E226^2+0.0007646*E226-0.00075',
        rpn='0.000001042|E226|3|^|*|0.00004777|E226|2|^|*|-|0.0007646|E226|*|+|0.00075|-',
        python_code='(((0.000001042 * (_C_("E226") ** 3)) - (0.00004777 * (_C_("E226") ** 2))) + (0.0007646 * _C_("E226"))) - 0.00075',
    ),
    dict(
        formula='=LINEST(G2:G17,E2:E17,FALSE)',
        rpn='G2:G17|E2:E17|FALSE|LINEST',
        python_code='linest(_R_("G2:G17"), _R_("E2:E17"), False, degree=-1)[-2]',
    ),
    dict(
        formula='=LINESTMARIO(G2:G17,E2:E17,FALSE)',
        rpn='G2:G17|E2:E17|FALSE|LINESTMARIO',
        python_code='linestmario(_R_("G2:G17"), _R_("E2:E17"), False)[-2]',
    ),
    dict(
        formula='=IF(AI119="","",E119)',
        rpn='AI119|""|=|""|E119|IF',
        python_code='("" if _C_("AI119") == "" else _C_("E119"))',
    ),
    dict(
        formula='=LINEST(B32:(INDEX(B32:B119,MATCH(0,B32:B119<6,-1),1)),(F32:(INDEX(B32:F119,MATCH(0,B32:B119,-1),5)))^{1,2,3,4})',
        rpn='B32:B119|0|B32:B119|6|<|1|-|MATCH|1|INDEX|B32:|B32:F119|0|B32:B119|1|-|MATCH|5|INDEX|F32:|1|2|3|4|ARRAYROW|ARRAY|^|LINEST',
        python_code='linest(b32:(index(_R_("B32:B119"), match(0, _R_("B32:B119") < 6, -1), 1)), f32:(index(_R_("B32:F119"), match(0, _R_("B32:B119"), -1), 5)), degree=-1)[-2]',
    ),
]


def dump_test_case(formula, python_code, rpn):
    escaped_python_code = python_code.replace('\\', r'\\')

    print('    dict(')
    print("        formula='{}',".format(formula))
    print("        rpn='{}',".format(rpn))
    print("        python_code='{}',".format(escaped_python_code))
    print('    ),')


def dump_parse():
    print('[')
    for test_case in test_data:
        excel_formula = ExcelFormula(test_case['formula'])
        parsed = excel_formula.rpn
        ast_root = excel_formula.ast
        result_rpn = "|".join(str(x) for x in parsed)
        result_python_code = ast_root.emit()
        dump_test_case(test_case['formula'], result_python_code, result_rpn)
    print(']')


sorted_keys = tuple(map(str, sorted(test_data[0])))


@pytest.mark.parametrize(
    ', '.join(sorted_keys),
    [tuple(test_case[k] for k in sorted_keys) for test_case in test_data]
)
def test_parse(formula, python_code, rpn):
    cell = ATestCell('A', 1)

    excel_formula = ExcelFormula(formula, cell=cell)
    parsed = excel_formula.rpn
    result_rpn = "|".join(str(x) for x in parsed)
    result_python_code = excel_formula.python_code
    assert result_python_code == excel_formula.ast.emit()

    if (rpn, python_code) != (result_rpn, result_python_code):
        print("***Expected: ")
        dump_test_case(formula, python_code, rpn)

        print("***Result: ")
        dump_test_case(formula, result_python_code, result_rpn)

        print('--------------')

    assert rpn == result_rpn
    assert python_code == result_python_code


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
    assert 'a_value' == node.emit()


def test_if_args_error():
    with pytest.raises(FormulaParserError):
        ExcelFormula('=if(1)').python_code


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


def test_build_eval_context():
    eval_context = ExcelFormula.build_eval_context(lambda x: 1, lambda x: 1)

    assert 42 == eval_context(ExcelFormula('=2 * 21'))
    assert 44 == eval_context(ExcelFormula('=2 * 21 + A1 + a1:a2'))
    assert 1 == eval_context(ExcelFormula('=1 + sin(0)'))

    with pytest.raises(FormulaEvalError,
                       match="name 'unknown_function' is not defined"):
        eval_context(ExcelFormula('=unknown_function(0)'))


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


def test_get_linest_degree_with_cell():
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
    

def test_string_number_compare():
    eval_ctx = ExcelFormula.build_eval_context(None, None)

    assert 1 == eval_ctx(ExcelFormula('=sum(1=1.0, "1"=1, 1="1")'))
    assert 1 == eval_ctx(ExcelFormula('=sum("1"="1", "x"=1)'))
    assert 'b' == eval_ctx(ExcelFormula('=if("x"<>"x", "a", "b")'))


def test_empty_cell_logic_op():
    eval_ctx = ExcelFormula.build_eval_context(lambda x: None, None)
    assert 1 == eval_ctx(ExcelFormula('=sum(A1=0, A1=1)'))
    assert 3 == eval_ctx(ExcelFormula('=sum(A1<0, A1<=0, A1=0, A1>=0, A1>0)'))


def test_numerics_type_coercion():
    eval_ctx = ExcelFormula.build_eval_context(lambda x: 3.0, None)
    assert 7 == eval_ctx(ExcelFormula('=1+2+"4"'))
    assert 7 == eval_ctx(ExcelFormula('=sum(1, 2, "4")'))

    assert '3A' == eval_ctx(ExcelFormula('=3&"A"'))
    assert '3A' == eval_ctx(ExcelFormula('=3.0&"A"'))
    assert '3A' == eval_ctx(ExcelFormula('=A1&"A"'))


def test_string_compare():
    eval_ctx = ExcelFormula.build_eval_context(None, None)

    assert not eval_ctx(ExcelFormula('=1="a"'))
    assert not eval_ctx(ExcelFormula('=1=2'))
    assert not eval_ctx(ExcelFormula('="a"="b"'))

    assert eval_ctx(ExcelFormula('=1=1'))
    assert eval_ctx(ExcelFormula('="A"="a"'))
    assert eval_ctx(ExcelFormula('="a"="A"'))


def test_string_concat():
    eval_ctx = ExcelFormula.build_eval_context(None, None)

    assert '6A' == eval_ctx(ExcelFormula('=2*3&"A"'))

    assert '1a' == eval_ctx(ExcelFormula('=1&"a"'))
    assert '12' == eval_ctx(ExcelFormula('="1"&2'))
    assert 'ab' == eval_ctx(ExcelFormula('="a"&"b"'))
    assert '11' == eval_ctx(ExcelFormula('=1&1'))
    assert 'Aa' == eval_ctx(ExcelFormula('="A"&"a"'))
    assert 'aA' == eval_ctx(ExcelFormula('="a"&"A"'))


def test_div_zero(caplog):
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: DIV0, lambda x: [[1, 1], [1, DIV0]],
        logging.getLogger('pycel_x'))

    assert DIV0 == eval_ctx(ExcelFormula('=sum(A1)'))
    assert DIV0 == eval_ctx(ExcelFormula('=sum(A1:B2)'))
    assert DIV0 == eval_ctx(ExcelFormula('=a1=1'))
    assert DIV0 == eval_ctx(ExcelFormula('=a1+"l"'))

    assert 3 == eval_ctx(ExcelFormula('=iferror(1/0,3)'))
    assert 1 == len(caplog.records)
    assert "WARNING" == caplog.records[0].levelname
    assert "1 Div 0" in caplog.records[0].message

    assert DIV0 == eval_ctx(ExcelFormula('=1/0'))
    assert 2 == len(caplog.records)
    assert "WARNING" == caplog.records[1].levelname

    message = """return PYTHON_AST_OPERATORS[op](left_op, right_op)
ZeroDivisionError: division by zero
Eval: 1 / 0
Values: 1 Div 0"""
    assert message in caplog.records[1].message


def test_value_error(caplog):
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: VALUE_ERROR, lambda x: [[1, 1], [1, VALUE_ERROR]],
        logging.getLogger('pycel_x'))

    assert VALUE_ERROR == eval_ctx(ExcelFormula('=sum(A1)'))
    assert VALUE_ERROR == eval_ctx(ExcelFormula('=sum(A1:B2)'))
    assert VALUE_ERROR == eval_ctx(ExcelFormula('=a1=1'))
    assert VALUE_ERROR == eval_ctx(ExcelFormula('=a1+"l"'))

    assert 3 == eval_ctx(ExcelFormula('=iferror(1+"A",3)'))
    assert 1 == len(caplog.records)
    assert "WARNING" == caplog.records[0].levelname
    assert "unsupported operand type(s)" in caplog.records[0].message

    assert VALUE_ERROR == eval_ctx(ExcelFormula('=1+"A"'))
    assert 2 == len(caplog.records)
    assert "WARNING" == caplog.records[1].levelname

    message = """return PYTHON_AST_OPERATORS[op](left_op, right_op)
TypeError: unsupported operand type(s) for +: 'int' and 'str'
Eval: 1 + "A"
Values: 1 Add A"""
    assert message in caplog.records[1].message


def test_string_number_mult(caplog):
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: VALUE_ERROR, lambda x: [[1, 1], [1, VALUE_ERROR]],
        logging.getLogger('pycel_x'))

    assert 3 == eval_ctx(ExcelFormula('=iferror(2*"A",3)'))
    assert 1 == len(caplog.records)
    assert "WARNING" == caplog.records[0].levelname
    assert "Cannot multiple type:" in caplog.records[0].message

    assert VALUE_ERROR == eval_ctx(ExcelFormula('="a"*2'))
    assert 2 == len(caplog.records)
    assert "WARNING" == caplog.records[1].levelname

    message = """Eval: "a" * 2
Cannot multiple type: str(a) * int(2)"""
    assert message in caplog.records[1].message


def test_eval_exception():
    eval_ctx = ExcelFormula.build_eval_context(
        lambda x: 1 + 'a', lambda x: [[1, 1], [1, DIV0]],
        logging.getLogger('pycel'))

    with pytest.raises(FormulaEvalError):
        eval_ctx(ExcelFormula('=a1'))


def test_lineno_on_error_reporting():
    eval_ctx = ExcelFormula.build_eval_context(None, None)

    excel_formula = ExcelFormula('')
    excel_formula._python_code = 'X'
    excel_formula.lineno = 6
    excel_formula.filename = 'a_file'

    with pytest.raises(FormulaEvalError, match='File "a_file", line 6'):
        eval_ctx(excel_formula)

    excel_formula._python_code = '(x)'
    excel_formula._compiled_python = None
    excel_formula.compiled_lambda = None
    excel_formula.lineno = 60

    with pytest.raises(FormulaEvalError, match=', line 60,'):
        eval_ctx(excel_formula)


if __name__ == '__main__':
    dump_parse()
