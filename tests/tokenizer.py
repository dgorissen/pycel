import pytest
from pycel.excelcompiler import ExcelCompiler


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
     'B5:B15|A7:D7||SUM'),

    ('=SUM((A:A,1:1))',
     'A:A|1:1|,|SUM'),

    ('=SUM((A:A A1:B1))',
     'A:A|A1:B1||SUM'),

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
     'A:A|1:1||SUM'),

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
    'R[39]C[11]|65|>|R[25]C[42]|R[11]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|YES|=|AND|OR|R[44]C[11]|R[43]C[11]|IF|*|R[14]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|YES|=|AND|OR|R[45]C[11]|R[43]C[11]|IF|*|+|0|ROUND|IF'),

    ('=IF(AI119="","",E119)',
     'AI119||=||E119|IF'),

    (
    '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
    'P5|1.0|=|NA|P5|2.0|=|A|P5|3.0|=|B|P5|4.0|=|C|P5|5.0|=|D|P5|6.0|=|E|P5|7.0|=|F|P5|8.0|=|G|IF|IF|IF|IF|IF|IF|IF|IF'),

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
def tests(test_number, formula, rpn):
    for formula, rpn in fancy_reference_inputs:
        assert rpn == stringify(ExcelCompiler.parse_to_rpn(formula))
