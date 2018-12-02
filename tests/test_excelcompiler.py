import os
import pytest

from unittest import mock

from pycel.excelcompiler import Cell, CellRange, ExcelCompiler
from pycel.excelformula import FormulaEvalError
from pycel.excelutil import AddressRange


# ::TODO:: need some rectangular ranges for testing


def test_end_2_end(excel, example_xls_path):
    # load & compile the file to a graph, starting from D1
    for excel_compiler in (ExcelCompiler(excel=excel),
                           ExcelCompiler(example_xls_path)):

        # test evaluation
        assert -0.02286 == round(excel_compiler.evaluate('Sheet1!D1'), 5)

        excel_compiler.set_value('Sheet1!A1', 200)
        assert -0.00331 == round(excel_compiler.evaluate('Sheet1!D1'), 5)

    # show the graph usisng matplotlib
    # sp.plot_graph()

    # export the graph, can be loaded by a viewer like gephi
    # sp.export_to_gexf(fname + ".gexf")

    # Serializing to disk...
    # excel_compiler.save_to_file(example_xls_path + ".pickle")


def test_round_trip_through_json_and_pickle(excel, example_xls_path):
    excel_compiler = ExcelCompiler(excel=excel)
    excel_compiler.evaluate('Sheet1!D1')
    excel_compiler.extra_data = {1: 3}
    excel_compiler.to_file()
    excel_compiler.to_yaml()
    excel_compiler.to_json()

    # read the spreadsheet from json
    excel_compiler_json = ExcelCompiler.from_json(excel.filename)
    excel_compiler_yaml = ExcelCompiler.from_yaml(excel.filename)
    excel_compiler = ExcelCompiler.from_file(excel.filename)

    # test evaluation
    assert -0.02286 == round(excel_compiler_json.evaluate('Sheet1!D1'), 5)
    assert -0.02286 == round(excel_compiler_yaml.evaluate('Sheet1!D1'), 5)
    assert -0.02286 == round(excel_compiler.evaluate('Sheet1!D1'), 5)

    excel_compiler_json.set_value('Sheet1!A1', 200)
    assert -0.00331 == round(excel_compiler_json.evaluate('Sheet1!D1'), 5)

    excel_compiler_yaml.set_value('Sheet1!A1', 200)
    assert -0.00331 == round(excel_compiler_yaml.evaluate('Sheet1!D1'), 5)

    excel_compiler.set_value('Sheet1!A1', 200)
    assert -0.00331 == round(excel_compiler.evaluate('Sheet1!D1'), 5)


def test_filename_ext(excel, example_xls_path):
    excel_compiler = ExcelCompiler(excel=excel)
    excel_compiler.evaluate('Sheet1!D1')
    excel_compiler.extra_data = {1: 3}
    pickle_name = excel_compiler.filename + '.pkl'
    yaml_name = excel_compiler.filename + '.yaml'
    json_name = excel_compiler.filename + '.json'

    for name in (pickle_name, yaml_name, json_name):
        try:
            os.unlink(name)
        except FileNotFoundError:
            pass

    excel_compiler.to_file(pickle_name)
    excel_compiler.to_yaml(yaml_name)
    excel_compiler.to_json(json_name)

    assert os.path.exists(pickle_name)
    assert os.path.exists(yaml_name)
    assert os.path.exists(json_name)


def test_hash_matches(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    assert excel_compiler.hash_matches

    excel_compiler._excel_file_md5_digest = 0
    assert not excel_compiler.hash_matches


def test_reset(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    in_address = AddressRange('Sheet1!A1')
    out_address = AddressRange('Sheet1!D1')

    assert -0.02286 == round(excel_compiler.evaluate(out_address), 5)

    in_value = excel_compiler.cell_map[in_address].value

    excel_compiler.reset(excel_compiler.cell_map[in_address])
    assert excel_compiler.cell_map[out_address].value is None

    excel_compiler.reset(excel_compiler.cell_map[in_address])
    assert excel_compiler.cell_map[out_address].value is None

    excel_compiler.cell_map[in_address].value = in_value
    assert -0.02286 == round(excel_compiler.evaluate(out_address), 5)
    assert -0.02286 == round(excel_compiler.cell_map[out_address].value, 5)


def test_recalculate(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    out_address = AddressRange('Sheet1!D1')

    assert -0.02286 == round(excel_compiler.evaluate(out_address), 5)
    excel_compiler.cell_map[out_address].value = None

    excel_compiler.recalculate()
    assert -0.02286 == round(excel_compiler.cell_map[out_address].value, 5)


def test_evaluate_range(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    result = excel_compiler.evaluate_range('trim-range!B2')
    assert 136 == result


def test_evaluate_empty(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    assert 0 == excel_compiler.evaluate('Empty!B1')
    excel_compiler.recalculate()
    assert '#EMPTY!' == excel_compiler.evaluate('Empty!B1')


def test_gen_graph(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    excel.set_sheet('trim-range')
    excel_compiler.gen_graph('B2')

    with pytest.raises(ValueError, match='Unknown seed'):
        excel_compiler.gen_graph(None)


def test_value_tree_str(excel):
    out_address = AddressRange('trim-range!B2')
    excel_compiler = ExcelCompiler(excel=excel)
    excel_compiler.evaluate(out_address)

    expected = [
        'trim-range!B2 = 136',
        ' trim-range!B1 = 24',
        '  trim-range!D1:E3 = [[1, 5], [2, 6], [3, 7]]',
        '   trim-range!D1 = 1',
        '   trim-range!D2 = 2',
        '   trim-range!D3 = 3',
        '   trim-range!E1 = 5',
        '   trim-range!E2 = 6',
        '   trim-range!E3 = 7',
        ' trim-range!D4:E4 = [4, 8]',
        '  trim-range!D4 = 4',
        '  trim-range!E4 = 8',
        ' trim-range!D5 = 100'
    ]
    assert expected == list(excel_compiler.value_tree_str(out_address))


def test_trim_cells(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = ['trim-range!D5']
    output_addrs = ['trim-range!B2']

    old_value = excel_compiler.evaluate(output_addrs[0])

    excel_compiler.trim_graph(input_addrs, output_addrs)
    excel_compiler.to_json()

    new_value = ExcelCompiler.from_json(
        excel_compiler.filename).evaluate(output_addrs[0])

    assert old_value == new_value


def test_trim_cells_warn_address_not_found(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = ['trim-range!D5', 'trim-range!H1']
    output_addrs = ['trim-range!B2']

    excel_compiler.evaluate(output_addrs[0])
    excel_compiler.log.warning = mock.Mock()
    excel_compiler.trim_graph(input_addrs, output_addrs)
    assert 1 == excel_compiler.log.warning.call_count


def test_trim_cells_exception_input_unused(excel):

    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = ['trim-range!G1']
    output_addrs = ['trim-range!B2']
    excel_compiler.evaluate(output_addrs[0])
    excel_compiler.evaluate(input_addrs[0])

    with pytest.raises(
            ValueError,
            match=' which usually means no outputs are dependant on it'):
        excel_compiler.trim_graph(input_addrs, output_addrs)


def test_compile_error_message_line_number(excel):
    excel_compiler = ExcelCompiler(excel=excel)

    input_addrs = [AddressRange('trim-range!D5')]
    output_addrs = [AddressRange('trim-range!B2')]

    excel_compiler.trim_graph(input_addrs, output_addrs)
    excel_compiler.to_file()

    excel_compiler = ExcelCompiler.from_file(excel_compiler.filename)
    formula = excel_compiler.cell_map[output_addrs[0]].formula
    formula._python_code = '(x)'
    formula.lineno = 3000
    formula.filename = 'a_file'
    with pytest.raises(
            FormulaEvalError, match='File "a_file", line 3000'):
        excel_compiler.evaluate(output_addrs[0])


def test_init_cell_address_error(excel):
    with pytest.raises(ValueError):
        CellRange('A1', excel)


def test_cell_range_repr(excel):
    cell_range = CellRange('sheet!A1', excel)
    assert 'sheet!A1' == repr(cell_range)


def test_cell_repr(excel):
    cell_range = Cell('sheet!A1', value=0)
    assert 'sheet!A1 -> 0' == repr(cell_range)


def test_gen_gexf(excel, tmpdir):
    excel_compiler = ExcelCompiler(excel=excel)
    filename = os.path.join(tmpdir, 'test.gexf')
    assert not os.path.exists(filename)
    excel_compiler.export_to_gexf(filename)
    assert os.path.exists(filename)


def test_gen_dot(excel, tmpdir):
    excel_compiler = ExcelCompiler(excel=excel)
    filename = os.path.join(tmpdir, 'test.dot')
    assert not os.path.exists(filename)
    excel_compiler.export_to_dot(filename)
    assert os.path.exists(filename)


def test_plot_graph(excel, tmpdir):
    from unittest import mock
    import sys
    mock_imports = (
        'matplotlib',
        'matplotlib.pyplot',
        'matplotlib.cbook',
        'matplotlib.colors',
        'matplotlib.collections',
        'matplotlib.patches',
    )
    for mock_import in mock_imports:
        sys.modules[mock_import] = mock.MagicMock()
    out_address = AddressRange('trim-range!B2')
    excel_compiler = ExcelCompiler(excel=excel)
    excel_compiler.evaluate(out_address)

    with mock.patch('pycel.excelcompiler.nx'):
        excel_compiler.plot_graph()
