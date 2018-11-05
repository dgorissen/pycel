import os
import pickle
from unittest import mock

import pytest
from pycel.excelformula import (
    ASTNode,
    ExcelFormula,
    FormulaParserError,
    Token,
)

from tests.test_excelutil import TestCell


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
        python_code='xsum(eval_range("A:A") + eval_range("1:1"))',
    ),
    dict(
        formula='=A1',
        rpn='A1',
        python_code='eval_cell("A1")',
    ),
    dict(
        formula='=atan2(A1,B1)',
        rpn='A1|B1|atan2',
        python_code='atan2(eval_cell("B1"), eval_cell("A1"))',
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
        python_code='eval_cell("A1")',
    ),
    dict(
        formula='=$B$2',
        rpn='$B$2',
        python_code='eval_cell("B2")',
    ),
    dict(
        formula='=PI()',
        rpn='PI',
        python_code='pi',
    ),
    dict(
        formula='=SUM(B5:B15)',
        rpn='B5:B15|SUM',
        python_code='xsum(eval_range("B5:B15"))',
    ),
    dict(
        formula='=SUM(B5:B15,D5:D15)',
        rpn='B5:B15|D5:D15|SUM',
        python_code='xsum(eval_range("B5:B15"), eval_range("D5:D15"))',
    ),
    dict(
        formula='=SUM(B5:B15 A7:D7)',
        rpn='B5:B15|A7:D7| |SUM',
        python_code='xsum(eval_range("B5:B15") + eval_range("A7:D7"))',
    ),
    dict(
        formula='=SUM(sheet1!$A$1:$B$2)',
        rpn='sheet1!$A$1:$B$2|SUM',
        python_code='xsum(eval_range("sheet1!A1:B2"))',
    ),
    dict(
        formula='=[data.xls]sheet1!$A$1',
        rpn='[data.xls]sheet1!$A$1',
        python_code='eval_cell("[data.xls]sheet1!A1")',
    ),
    dict(
        formula='=SUM((A:A,1:1))',
        rpn='A:A|1:1|,|SUM',
        python_code='xsum(eval_range("A:A"), eval_range("1:1"))',
    ),
    dict(
        formula='=SUM((A:A A1:B1))',
        rpn='A:A|A1:B1| |SUM',
        python_code='xsum(eval_range("A:A") + eval_range("A1:B1"))',
    ),
    dict(
        formula='=SUM(D9:D11,E9:E11,F9:F11)',
        rpn='D9:D11|E9:E11|F9:F11|SUM',
        python_code='xsum(eval_range("D9:D11"), eval_range("E9:E11"), eval_range("F9:F11"))',
    ),
    dict(
        formula='=SUM((D9:D11,(E9:E11,F9:F11)))',
        rpn='D9:D11|E9:E11|F9:F11|,|,|SUM',
        python_code='xsum(eval_range("D9:D11"), (eval_range("E9:E11"), eval_range("F9:F11")))',
    ),
    dict(
        formula='=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
        rpn='P5|1.0|=|"NA"|P5|2.0|=|"A"|P5|3.0|=|"B"|P5|4.0|=|"C"|P5|5.0|=|"D"|P5|6.0|=|"E"|P5|7.0|=|"F"|P5|8.0|=|"G"|IF|IF|IF|IF|IF|IF|IF|IF',
        python_code='("NA" if eval_cell("P5") == 1.0 else ("A" if eval_cell("P5") == 2.0 else ("B" if eval_cell("P5") == 3.0 else ("C" if eval_cell("P5") == 4.0 else ("D" if eval_cell("P5") == 5.0 else ("E" if eval_cell("P5") == 6.0 else ("F" if eval_cell("P5") == 7.0 else ("G" if eval_cell("P5") == 8.0 else 0))))))))',
    ),
    dict(
        formula='={SUM(B2:D2*B3:D3)}',
        rpn='B2:D2|B3:D3|*|SUM|ARRAYROW|ARRAY',
        python_code='[xsum(eval_range("B2:D2") * eval_range("B3:D3"))]',
    ),
    dict(
        formula='=SUM(123 + SUM(456) + (45<6))+456+789',
        rpn='123|456|SUM|+|45|6|<|+|SUM|456|+|789|+',
        python_code='(xsum((123 + xsum(456)) + (45 < 6)) + 456) + 789',
    ),
    dict(
        formula='=AVG(((((123 + 4 + AVG(A1:A2))))))',
        rpn='123|4|+|A1:A2|AVG|+|AVG',
        python_code='avg((123 + 4) + avg(eval_range("A1:A2")))',
    ),

    # E. W. Bachtal's test formulae
    dict(
        formula='=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
        rpn='"a"|"a"|"b"|ARRAYROW|"c"|#N/A|ARRAYROW|1|-|TRUE|ARRAYROW|ARRAY|=|"yes"|"no"|IF|"  more ""test"" text"|&',
        python_code='("yes" if "a" == [["a", "b"], ["c", "#N/A"], [-1, True]] else "no") + "  more \\"test\\" text"',
    ),
    dict(
        formula='=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
        rpn='R13C3|2002|1|6|DATE|>|0|R[41]C[2]|ISERROR|0|R13C3|R[41]C[2]|>=|0|R[23]C[11]|55|>=|R[24]C[11]|20|>=|AND|R53C3|0|IF|IF|IF|IF',
        python_code='(0 if eval_cell("C13") > (date(2002, 1, 6) if date(2002, 1, 6) is not None else 0) else (0 if iserror(eval_cell("C42")) else (0 if eval_cell("C13") >= (eval_cell("C42") if eval_cell("C42") is not None else 0) else (eval_cell("C53") if all([eval_cell("L24") >= 55, eval_cell("L25") >= 20]) else 0))))',    ),
    dict(
        formula='=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), R[45]C[11],R[43]C[11])),0))',
        rpn='R[39]C[11]|65|>|R[25]C[42]|R[11]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|R[44]C[11]|R[43]C[11]|IF|*|R[14]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|"YES"|=|AND|OR|R[45]C[11]|R[43]C[11]|IF|*|+|0|ROUND|IF',
        python_code='(eval_cell("AQ26") if eval_cell("L40") > 65 else xround((eval_cell("L12") * (eval_cell("L45") if any([all([eval_cell("L40") >= 55, eval_cell("L41") >= 20]), all([eval_cell("L41") >= 20, eval_cell("C11") == "YES"])]) else eval_cell("L44"))) + (eval_cell("L15") * (eval_cell("L46") if any([all([eval_cell("L40") >= 55, eval_cell("L41") >= 20]), all([eval_cell("L41") >= 20, eval_cell("C11") == "YES"])]) else eval_cell("L44"))), 0))',
    ),
    dict(
        formula='=(propellor_charts!B22*(propellor_charts!E21+propellor_charts!D21*(engine_data!O16*D70+engine_data!P16)+propellor_charts!C21*(engine_data!O16*D70+engine_data!P16)^2+propellor_charts!B21*(engine_data!O16*D70+engine_data!P16)^3)^2)^(1/3)*(1*D70/5.33E-18)^(2/3)*0.0000000001*28.3495231*9.81/1000',
        rpn='propellor_charts!B22|propellor_charts!E21|propellor_charts!D21|engine_data!O16|D70|*|engine_data!P16|+|*|+|propellor_charts!C21|engine_data!O16|D70|*|engine_data!P16|+|2|^|*|+|propellor_charts!B21|engine_data!O16|D70|*|engine_data!P16|+|3|^|*|+|2|^|*|1|3|/|^|1|D70|*|5.33E-18|/|2|3|/|^|*|0.0000000001|*|28.3495231|*|9.81|*|1000|/',
        python_code='((((((eval_cell("propellor_charts!B22") * ((((eval_cell("propellor_charts!E21") + (eval_cell("propellor_charts!D21") * ((eval_cell("engine_data!O16") * eval_cell("D70")) + eval_cell("engine_data!P16")))) + (eval_cell("propellor_charts!C21") * (((eval_cell("engine_data!O16") * eval_cell("D70")) + eval_cell("engine_data!P16")) ** 2))) + (eval_cell("propellor_charts!B21") * (((eval_cell("engine_data!O16") * eval_cell("D70")) + eval_cell("engine_data!P16")) ** 3))) ** 2)) ** (1 / 3)) * (((1 * eval_cell("D70")) / 5.33E-18) ** (2 / 3))) * 0.0000000001) * 28.3495231) * 9.81) / 1000',
    ),
    dict(
        formula='=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))',
        rpn='3600|1000|/|E40|*|E8|E39|/|*|E15|E19|/|*|E54|E54|E48|-|/|LN|*',
        python_code='((((3600 / 1000) * eval_cell("E40")) * (eval_cell("E8") / eval_cell("E39"))) * (eval_cell("E15") / eval_cell("E19"))) * xlog(eval_cell("E54") / (eval_cell("E54") - eval_cell("E48")))',
    ),
    dict(
        formula='=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
        rpn='P5|1.0|=|"NA"|P5|2.0|=|"A"|P5|3.0|=|"B"|P5|4.0|=|"C"|P5|5.0|=|"D"|P5|6.0|=|"E"|P5|7.0|=|"F"|P5|8.0|=|"G"|IF|IF|IF|IF|IF|IF|IF|IF',
        python_code='("NA" if eval_cell("P5") == 1.0 else ("A" if eval_cell("P5") == 2.0 else ("B" if eval_cell("P5") == 3.0 else ("C" if eval_cell("P5") == 4.0 else ("D" if eval_cell("P5") == 5.0 else ("E" if eval_cell("P5") == 6.0 else ("F" if eval_cell("P5") == 7.0 else ("G" if eval_cell("P5") == 8.0 else 0))))))))',
    ),
    dict(
        formula='=LINEST(X5:X32,W5:W32^{1,2,3})',
        rpn='X5:X32|W5:W32|1|2|3|ARRAYROW|ARRAY|^|LINEST',
        python_code='linest(eval_range("X5:X32"), eval_range("W5:W32"), degree=-1)[-2]',
    ),
    dict(
        formula='=IF(configurations!$G$22=3,sizing!$C$303,M14)',
        rpn='configurations!$G$22|3|=|sizing!$C$303|M14|IF',
        python_code='(eval_cell("sizing!C303") if eval_cell("configurations!G22") == 3 else eval_cell("M14"))',
    ),
    dict(
        formula='=0.000001042*E226^3-0.00004777*E226^2+0.0007646*E226-0.00075',
        rpn='0.000001042|E226|3|^|*|0.00004777|E226|2|^|*|-|0.0007646|E226|*|+|0.00075|-',
        python_code='(((0.000001042 * (eval_cell("E226") ** 3)) - (0.00004777 * (eval_cell("E226") ** 2))) + (0.0007646 * eval_cell("E226"))) - 0.00075',
    ),
    dict(
        formula='=LINEST(G2:G17,E2:E17,FALSE)',
        rpn='G2:G17|E2:E17|FALSE|LINEST',
        python_code='linest(eval_range("G2:G17"), eval_range("E2:E17"), False, degree=-1)[-2]',
    ),
    dict(
        formula='=LINESTMARIO(G2:G17,E2:E17,FALSE)',
        rpn='G2:G17|E2:E17|FALSE|LINESTMARIO',
        python_code='linestmario(eval_range("G2:G17"), eval_range("E2:E17"), False)[-2]',
    ),
    dict(
        formula='=IF(AI119="","",E119)',
        rpn='AI119|""|=|""|E119|IF',
        python_code='("" if eval_cell("AI119") == "" else eval_cell("E119"))',
    ),
    dict(
        formula='=LINEST(B32:(INDEX(B32:B119,MATCH(0,B32:B119<6,-1),1)),(F32:(INDEX(B32:F119,MATCH(0,B32:B119,-1),5)))^{1,2,3,4})',
        rpn='B32:B119|0|B32:B119|6|<|1|-|MATCH|1|INDEX|B32:|B32:F119|0|B32:B119|1|-|MATCH|5|INDEX|F32:|1|2|3|4|ARRAYROW|ARRAY|^|LINEST',
        python_code='linest(b32:(index(eval_range("B32:B119"), match(0, (eval_range("B32:B119") if eval_range("B32:B119") is not None else 0) < 6, -1), 1)), f32:(index(eval_range("B32:F119"), match(0, eval_range("B32:B119"), -1), 5)), degree=-1)[-2]',
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
    cell = TestCell('A', 1)

    excel_formula = ExcelFormula(formula, cell=cell)
    parsed = excel_formula.rpn
    result_rpn = "|".join(str(x) for x in parsed)
    result_python_code = excel_formula.python_code

    if (rpn, python_code) != (result_rpn, result_python_code):
        print("***Expected: ")
        dump_test_case(formula, python_code, rpn)

        print("***Result: ")
        dump_test_case(formula, result_python_code, result_rpn)

        print('--------------')

    assert rpn == result_rpn
    assert python_code == result_python_code


def test_descendants():

    excel_formula = ExcelFormula('=E54-E48')
    descendants = excel_formula.ast.descendants

    assert 2 == len(descendants)
    assert 'OPERAND' == descendants[0][0].type
    assert 'E54' == descendants[0][0].value
    assert 'OPERAND' == descendants[1][0].type
    assert 'E48' == descendants[1][0].value


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

    assert needed == sorted(x.address for x in
                            ExcelFormula(formula).needed_addresses)


def test_build_eval_context():
    null_func = lambda x: 1

    eval_context = ExcelFormula.build_eval_context(null_func, null_func)

    assert 42 == eval_context(ExcelFormula('=2 * 21'))
    assert 44 == eval_context(ExcelFormula('=2 * 21 + A1 + a1:a2'))
    assert 1 == eval_context(ExcelFormula('=1 + sin(0)'))

    with pytest.raises(NameError):
        eval_context(ExcelFormula('=unknown_function(0)'))


def test_compiled_python_error():
    formula = ExcelFormula('=1 + 2')
    formula._python_code = 'this will be a syntax error'
    with pytest.raises(FormulaParserError, match='Failed to compile expression'):
        x = formula.compiled_python


def test_save_to_file(fixture_dir):
    formula = ExcelFormula('=1+2')
    filename = os.path.join(fixture_dir, 'formula_save_test.pickle')
    with open(filename, 'wb') as f:
        pickle.dump(formula, f, protocol=2)

    with open(filename, 'rb') as f:
        loaded_formula = pickle.load(f)

    os.unlink(filename)

    assert formula.python_code == loaded_formula.python_code


def test_get_linest_degree_with_cell():
    with mock.patch('pycel.excelformula.get_linest_degree') as get:
        get.return_value = -1, -1

        cell = TestCell('A', 1, 'Phony Sheet')
        formula = ExcelFormula('=linest(C1)', cell=cell)

        expected = 'linest(eval_cell("Phony Sheet!C1"), degree=-1)[-2]'
        assert expected == formula.python_code


if __name__ == '__main__':
    dump_parse()
