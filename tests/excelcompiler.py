import json
import pytest

from pycel.excelcompiler import ExcelCompiler

test_data = [
    {
        "formula": "=SUM((A:A 1:1))",
        "rpn": "A:A|1:1| |SUM",
        "python_code": "xsum(eval_range(\"A:A\") + eval_range(\"1:1\"))"
    },
    {
        "formula": "=A1",
        "rpn": "A1",
        "python_code": "eval_cell(\"A1\")"
    },
    {
        "formula": "=atan2(A1,B1)",
        "rpn": "A1|B1|atan2",
        "python_code": "atan2(eval_cell(\"B1\"), eval_cell(\"A1\"))"
    },
    {
        "formula": "=5*log(sin()+2)",
        "rpn": "5|sin|2|+|log|*",
        "python_code": "5 * log(sin() + 2)"
    },
    {
        "formula": "=5*log(sin(3,7,9)+2)",
        "rpn": "5|3|7|9|sin|2|+|log|*",
        "python_code": "5 * log(sin(3, 7, 9) + 2)"
    },
    {
        "formula": "=3 +1-5",
        "rpn": "3|1|+|5|-",
        "python_code": "(3 + 1) - 5"
    },
    {
        "formula": "=3 + 4 * 2 / ( 1 - 5 ) ^ 2 ^ 3",
        "rpn": "3|4|2|*|1|5|-|2|^|3|^|/|+",
        "python_code": "3 + ((4 * 2) / (((1 - 5) ** 2) ** 3))"
    },
    {
        "formula": "=1+3+5",
        "rpn": "1|3|+|5|+",
        "python_code": "(1 + 3) + 5"
    },
    {
        "formula": "=3 + 4 * 5",
        "rpn": "3|4|5|*|+",
        "python_code": "3 + (4 * 5)"
    },
    {
        "formula": "=3 * 4 + 5",
        "rpn": "3|4|*|5|+",
        "python_code": "(3 * 4) + 5"
    },
    {
        "formula": "=50",
        "rpn": "50",
        "python_code": "50"
    },
    {
        "formula": "=1+1",
        "rpn": "1|1|+",
        "python_code": "1 + 1"
    },
    {
        "formula": "=$A1",
        "rpn": "$A1",
        "python_code": "eval_cell(\"A1\")"
    },
    {
        "formula": "=$B$2",
        "rpn": "$B$2",
        "python_code": "eval_cell(\"B2\")"
    },
    {
        "formula": "=SUM(B5:B15)",
        "rpn": "B5:B15|SUM",
        "python_code": "xsum(eval_range(\"B5:B15\"))"
    },
    {
        "formula": "=SUM(B5:B15,D5:D15)",
        "rpn": "B5:B15|D5:D15|SUM",
        "python_code": "xsum(eval_range(\"B5:B15\"), eval_range(\"D5:D15\"))"
    },
    {
        "formula": "=SUM(B5:B15 A7:D7)",
        "rpn": "B5:B15|A7:D7| |SUM",
        "python_code": "xsum(eval_range(\"B5:B15\") + eval_range(\"A7:D7\"))"
    },
    {
        "formula": "=SUM(sheet1!$A$1:$B$2)",
        "rpn": "sheet1!$A$1:$B$2|SUM",
        "python_code": "xsum(eval_range(\"sheet1!A1:B2\"))"
    },
    {
        "formula": "=[data.xls]sheet1!$A$1",
        "rpn": "[data.xls]sheet1!$A$1",
        "python_code": "eval_cell(\"[data.xls]sheet1!A1\")"
    },
    {
        "formula": "=SUM((A:A,1:1))",
        "rpn": "A:A|1:1|,|SUM",
        "python_code": "xsum(eval_range(\"A:A\"), eval_range(\"1:1\"))"
    },
    {
        "formula": "=SUM((A:A A1:B1))",
        "rpn": "A:A|A1:B1| |SUM",
        "python_code": "xsum(eval_range(\"A:A\") + eval_range(\"A1:B1\"))"
    },
    {
        "formula": "=SUM(D9:D11,E9:E11,F9:F11)",
        "rpn": "D9:D11|E9:E11|F9:F11|SUM",
        "python_code": "xsum(eval_range(\"D9:D11\"), eval_range(\"E9:E11\"), eval_range(\"F9:F11\"))"
    },
    {
        "formula": "=SUM((D9:D11,(E9:E11,F9:F11)))",
        "rpn": "D9:D11|E9:E11|F9:F11|,|,|SUM",
        "python_code": "xsum(eval_range(\"D9:D11\"), (eval_range(\"E9:E11\"), eval_range(\"F9:F11\")))"
    },
    {
        "formula": "=IF(P5=1.0,\"NA\",IF(P5=2.0,\"A\",IF(P5=3.0,\"B\",IF(P5=4.0,\"C\",IF(P5=5.0,\"D\",IF(P5=6.0,\"E\",IF(P5=7.0,\"F\",IF(P5=8.0,\"G\"))))))))",
        "rpn": "P5|1.0|=|\"NA\"|P5|2.0|=|\"A\"|P5|3.0|=|\"B\"|P5|4.0|=|\"C\"|P5|5.0|=|\"D\"|P5|6.0|=|\"E\"|P5|7.0|=|\"F\"|P5|8.0|=|\"G\"|IF|IF|IF|IF|IF|IF|IF|IF",
        "python_code": "(\"NA\" if eval_cell(\"P5\") == 1.0 else (\"A\" if eval_cell(\"P5\") == 2.0 else (\"B\" if eval_cell(\"P5\") == 3.0 else (\"C\" if eval_cell(\"P5\") == 4.0 else (\"D\" if eval_cell(\"P5\") == 5.0 else (\"E\" if eval_cell(\"P5\") == 6.0 else (\"F\" if eval_cell(\"P5\") == 7.0 else (\"G\" if eval_cell(\"P5\") == 8.0 else 0))))))))"
    },
    {
        "formula": "={SUM(B2:D2*B3:D3)}",
        "rpn": "B2:D2|B3:D3|*|SUM|ARRAYROW|ARRAY",
        "python_code": "[xsum(eval_range(\"B2:D2\") * eval_range(\"B3:D3\"))]"
    },
    {
        "formula": "=SUM(123 + SUM(456) + (45<6))+456+789",
        "rpn": "123|456|SUM|+|45|6|<|+|SUM|456|+|789|+",
        "python_code": "(xsum((123 + xsum(456)) + (45 < 6)) + 456) + 789"
    },
    {
        "formula": "=AVG(((((123 + 4 + AVG(A1:A2))))))",
        "rpn": "123|4|+|A1:A2|AVG|+|AVG",
        "python_code": "avg((123 + 4) + avg(eval_range(\"A1:A2\")))"
    },

    # E. W. Bachtal's test formulae
    {
        "formula": "=IF(\"a\"={\"a\",\"b\";\"c\",#N/A;-1,TRUE}, \"yes\", \"no\") &   \"  more \"\"test\"\" text\"",
        "rpn": "\"a\"|\"a\"|\"b\"|ARRAYROW|\"c\"|#N/A|ARRAYROW|1|-|TRUE|ARRAYROW|ARRAY|=|\"yes\"|\"no\"|IF|\"  more \"\"test\"\" text\"|&",
        "python_code": "(\"yes\" if \"a\" == [[\"a\", \"b\"], [\"c\", \"#N/A\"], [-1, True]] else \"no\") + \"  more \\\"test\\\" text\""
    },
    {
        "formula": "=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))",
        "rpn": "R13C3|2002|1|6|DATE|>|0|R[41]C[2]|ISERROR|0|R13C3|R[41]C[2]|>=|0|R[23]C[11]|55|>=|R[24]C[11]|20|>=|AND|R53C3|0|IF|IF|IF|IF",
        "python_code": "(0 if eval_cell(\"R13C3\") > (date(2002, 1, 6) if date(2002, 1, 6) is not None else 0) else (0 if iserror(eval_cell(\"R[41]C[2]\")) else (0 if eval_cell(\"R13C3\") >= (eval_cell(\"R[41]C[2]\") if eval_cell(\"R[41]C[2]\") is not None else 0) else (eval_cell(\"R53C3\") if all([eval_cell(\"R[23]C[11]\") >= 55, eval_cell(\"R[24]C[11]\") >= 20]) else 0))))"
    },
    {
        "formula": "=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3=\"YES\")),R[44]C[11],R[43]C[11]))+(R[14]C[11] *IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3=\"YES\")), R[45]C[11],R[43]C[11])),0))",
        "rpn": "R[39]C[11]|65|>|R[25]C[42]|R[11]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|\"YES\"|=|AND|OR|R[44]C[11]|R[43]C[11]|IF|*|R[14]C[11]|R[39]C[11]|55|>=|R[40]C[11]|20|>=|AND|R[40]C[11]|20|>=|R11C3|\"YES\"|=|AND|OR|R[45]C[11]|R[43]C[11]|IF|*|+|0|ROUND|IF",
        "python_code": "(eval_cell(\"R[25]C[42]\") if eval_cell(\"R[39]C[11]\") > 65 else xround((eval_cell(\"R[11]C[11]\") * (eval_cell(\"R[44]C[11]\") if any([all([eval_cell(\"R[39]C[11]\") >= 55, eval_cell(\"R[40]C[11]\") >= 20]), all([eval_cell(\"R[40]C[11]\") >= 20, eval_cell(\"R11C3\") == \"YES\"])]) else eval_cell(\"R[43]C[11]\"))) + (eval_cell(\"R[14]C[11]\") * (eval_cell(\"R[45]C[11]\") if any([all([eval_cell(\"R[39]C[11]\") >= 55, eval_cell(\"R[40]C[11]\") >= 20]), all([eval_cell(\"R[40]C[11]\") >= 20, eval_cell(\"R11C3\") == \"YES\"])]) else eval_cell(\"R[43]C[11]\"))), 0))"
    },
    {
        "formula": "=(propellor_charts!B22*(propellor_charts!E21+propellor_charts!D21*(engine_data!O16*D70+engine_data!P16)+propellor_charts!C21*(engine_data!O16*D70+engine_data!P16)^2+propellor_charts!B21*(engine_data!O16*D70+engine_data!P16)^3)^2)^(1/3)*(1*D70/5.33E-18)^(2/3)*0.0000000001*28.3495231*9.81/1000",
        "rpn": "propellor_charts!B22|propellor_charts!E21|propellor_charts!D21|engine_data!O16|D70|*|engine_data!P16|+|*|+|propellor_charts!C21|engine_data!O16|D70|*|engine_data!P16|+|2|^|*|+|propellor_charts!B21|engine_data!O16|D70|*|engine_data!P16|+|3|^|*|+|2|^|*|1|3|/|^|1|D70|*|5.33E-18|/|2|3|/|^|*|0.0000000001|*|28.3495231|*|9.81|*|1000|/",
        "python_code": "((((((eval_cell(\"propellor_charts!B22\") * ((((eval_cell(\"propellor_charts!E21\") + (eval_cell(\"propellor_charts!D21\") * ((eval_cell(\"engine_data!O16\") * eval_cell(\"D70\")) + eval_cell(\"engine_data!P16\")))) + (eval_cell(\"propellor_charts!C21\") * (((eval_cell(\"engine_data!O16\") * eval_cell(\"D70\")) + eval_cell(\"engine_data!P16\")) ** 2))) + (eval_cell(\"propellor_charts!B21\") * (((eval_cell(\"engine_data!O16\") * eval_cell(\"D70\")) + eval_cell(\"engine_data!P16\")) ** 3))) ** 2)) ** (1 / 3)) * (((1 * eval_cell(\"D70\")) / 5.33E-18) ** (2 / 3))) * 0.0000000001) * 28.3495231) * 9.81) / 1000"
    },
    {
        "formula": "=(3600/1000)*E40*(E8/E39)*(E15/E19)*LN(E54/(E54-E48))",
        "rpn": "3600|1000|/|E40|*|E8|E39|/|*|E15|E19|/|*|E54|E54|E48|-|/|LN|*",
        "python_code": "((((3600 / 1000) * eval_cell(\"E40\")) * (eval_cell(\"E8\") / eval_cell(\"E39\"))) * (eval_cell(\"E15\") / eval_cell(\"E19\"))) * xlog(eval_cell(\"E54\") / (eval_cell(\"E54\") - eval_cell(\"E48\")))"
    },
    {
        "formula": "=IF(P5=1.0,\"NA\",IF(P5=2.0,\"A\",IF(P5=3.0,\"B\",IF(P5=4.0,\"C\",IF(P5=5.0,\"D\",IF(P5=6.0,\"E\",IF(P5=7.0,\"F\",IF(P5=8.0,\"G\"))))))))",
        "rpn": "P5|1.0|=|\"NA\"|P5|2.0|=|\"A\"|P5|3.0|=|\"B\"|P5|4.0|=|\"C\"|P5|5.0|=|\"D\"|P5|6.0|=|\"E\"|P5|7.0|=|\"F\"|P5|8.0|=|\"G\"|IF|IF|IF|IF|IF|IF|IF|IF",
        "python_code": "(\"NA\" if eval_cell(\"P5\") == 1.0 else (\"A\" if eval_cell(\"P5\") == 2.0 else (\"B\" if eval_cell(\"P5\") == 3.0 else (\"C\" if eval_cell(\"P5\") == 4.0 else (\"D\" if eval_cell(\"P5\") == 5.0 else (\"E\" if eval_cell(\"P5\") == 6.0 else (\"F\" if eval_cell(\"P5\") == 7.0 else (\"G\" if eval_cell(\"P5\") == 8.0 else 0))))))))"
    },
    {
        "formula": "=LINEST(X5:X32,W5:W32^{1,2,3})",
        "rpn": "X5:X32|W5:W32|1|2|3|ARRAYROW|ARRAY|^|LINEST",
        "python_code": "linest(eval_range(\"X5:X32\"), eval_range(\"W5:W32\"), degree=-1)[-2]"
    },
    {
        "formula": "=IF(configurations!$G$22=3,sizing!$C$303,M14)",
        "rpn": "configurations!$G$22|3|=|sizing!$C$303|M14|IF",
        "python_code": "(eval_cell(\"sizing!C303\") if eval_cell(\"configurations!G22\") == 3 else eval_cell(\"M14\"))"
    },
    {
        "formula": "=0.000001042*E226^3-0.00004777*E226^2+0.0007646*E226-0.00075",
        "rpn": "0.000001042|E226|3|^|*|0.00004777|E226|2|^|*|-|0.0007646|E226|*|+|0.00075|-",
        "python_code": "(((0.000001042 * (eval_cell(\"E226\") ** 3)) - (0.00004777 * (eval_cell(\"E226\") ** 2))) + (0.0007646 * eval_cell(\"E226\"))) - 0.00075"
    },
    {
        "formula": "=LINEST(G2:G17,E2:E17,FALSE)",
        "rpn": "G2:G17|E2:E17|FALSE|LINEST",
        "python_code": "linest(eval_range(\"G2:G17\"), eval_range(\"E2:E17\"), False, degree=-1)[-2]"
    },
    {
        "formula": "=IF(AI119=\"\",\"\",E119)",
        "rpn": "AI119|\"\"|=|\"\"|E119|IF",
        "python_code": "(\"\" if eval_cell(\"AI119\") == \"\" else eval_cell(\"E119\"))"
    },
    {
        "formula": "=LINEST(B32:(INDEX(B32:B119,MATCH(0,B32:B<6119,-1),1)),(F32:(INDEX(B32:F119,MATCH(0,B32:B119,-1),5)))^{1,2,3,4})",
        "rpn": "B32:B119|0|B32:B|6119|<|1|-|MATCH|1|INDEX|B32:|B32:F119|0|B32:B119|1|-|MATCH|5|INDEX|F32:|1|2|3|4|ARRAYROW|ARRAY|^|LINEST",
        "python_code": "linest(b32:(index(eval_range(\"B32:B119\"), match(0, (eval_range(\"B32:B\") if eval_range(\"B32:B\") is not None else 0) < 6119, -1), 1)), f32:(index(eval_range(\"B32:F119\"), match(0, eval_range(\"B32:B119\"), -1), 5)), degree=-1)[-2]"
    }
]


def dump_parse():
    for test_case in test_data:

        parsed = ExcelCompiler.parse_to_rpn(test_case['formula'])
        graph, root = ExcelCompiler.build_ast(parsed)
        result_rpn = "|".join(str(x) for x in parsed)
        result_python_code = root.emit(graph)

        print(json.dumps(
            dict(formula=test_case['formula'], rpn=result_rpn,
                 python_code=result_python_code),
            indent=4))


sorted_keys = tuple(map(str, sorted(test_data[0])))


@pytest.mark.parametrize(
    ', '.join(sorted_keys),
    [tuple(test_case[k] for k in sorted_keys) for test_case in test_data]
)
def test_parse(formula, python_code, rpn):

        parsed = ExcelCompiler.parse_to_rpn(formula)
        ast_root = ExcelCompiler.build_ast(parsed)
        result_rpn = "|".join(str(x) for x in parsed)
        result_python_code = ast_root.emit()

        if (rpn, python_code) != (result_rpn, result_python_code):

            print("Formula: ", formula)

            if rpn != result_rpn:
                print("***RPN_e: ", rpn)
                print("***RPN_r: ", result_rpn)
                print('***JSON:', json.dumps(result_rpn))
                print('--------------')

            if python_code != result_python_code:
                print("***Python Code_e: ", python_code)
                print("***Python Code_r: ", result_python_code)
                print('***JSON:', json.dumps(result_python_code))
                print('--------------')

        assert rpn == result_rpn
        assert python_code == result_python_code


def test_end_2_end(excel, example_xls_path):
    # load & compile the file to a graph, starting from D1
    excel = ExcelCompiler(excel=excel)

    # test evaluation
    assert -0.02286 == round(excel.evaluate('Sheet1!D1'), 5)

    excel.set_value('Sheet1!A1', 200)
    assert -0.00331 == round(excel.evaluate('Sheet1!D1'), 5)

    # show the graph usisng matplotlib
    # sp.plot_graph()

    # export the graph, can be loaded by a viewer like gephi
    # sp.export_to_gexf(fname + ".gexf")

    # Serializing to disk...
    excel.save_to_file(example_xls_path + ".pickle")



if __name__ == '__main__':
    dump_parse()
