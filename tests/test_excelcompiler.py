import os
import shutil
from unittest import mock

import pytest
from pycel.excelcompiler import _Cell, _CellRange, ExcelCompiler
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


def test_round_trip_through_json_yaml_and_pickle(excel, example_xls_path):
    excel_compiler = ExcelCompiler(excel=excel)
    excel_compiler.evaluate('Sheet1!D1')
    excel_compiler.extra_data = {1: 3}
    excel_compiler.to_file(file_types=('pickle', ))
    excel_compiler.to_file(file_types=('yml', ))
    excel_compiler.to_file(file_types=('json', ))

    # read the spreadsheet from json, yaml and pickle
    excel_compiler_json = ExcelCompiler.from_file(excel.filename + '.json')
    excel_compiler_yaml = ExcelCompiler.from_file(excel.filename + '.yml')
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
    yaml_name = excel_compiler.filename + '.yml'
    json_name = excel_compiler.filename + '.json'

    for name in (pickle_name, yaml_name, json_name):
        try:
            os.unlink(name)
        except FileNotFoundError:
            pass

    excel_compiler.to_file(excel_compiler.filename)
    excel_compiler.to_file(json_name, file_types=('json', ))

    assert os.path.exists(pickle_name)
    assert os.path.exists(yaml_name)
    assert os.path.exists(json_name)


def test_filename_extension_errors(excel, example_xls_path):
    with pytest.raises(ValueError, match='Unrecognized file type'):
        ExcelCompiler.from_file(excel.filename + '.xyzzy')

    excel_compiler = ExcelCompiler(excel=excel)

    with pytest.raises(ValueError, match='Only allowed one extension'):
        excel_compiler.to_file(file_types=('pkl', 'pickle'))

    with pytest.raises(ValueError, match='Only allowed one '):
        excel_compiler.to_file(file_types=('pkl', 'yml', 'json'))

    with pytest.raises(ValueError, match='Unknown file types: pkly'):
        excel_compiler.to_file(file_types=('pkly',))


def test_hash_matches(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    assert excel_compiler.hash_matches

    excel_compiler._excel_file_md5_digest = 0
    assert not excel_compiler.hash_matches


def test_pickle_file_rebuilding(excel):

    input_addrs = ['Sheet1!A11']
    output_addrs = ['Sheet1!D1']

    excel_compiler = ExcelCompiler(excel=excel)
    excel_compiler.trim_graph(input_addrs, output_addrs)
    excel_compiler.to_file()

    pickle_name = excel_compiler.filename + '.pkl'
    yaml_name = excel_compiler.filename + '.yml'

    assert os.path.exists(pickle_name)
    old_hash = excel_compiler._compute_file_md5_digest(pickle_name)

    excel_compiler.to_file()
    assert old_hash == excel_compiler._compute_file_md5_digest(pickle_name)

    os.unlink(yaml_name)
    excel_compiler.to_file()
    new_hash = excel_compiler._compute_file_md5_digest(pickle_name)
    assert old_hash != new_hash

    shutil.copyfile(pickle_name, yaml_name)
    excel_compiler.to_file()
    assert new_hash != excel_compiler._compute_file_md5_digest(pickle_name)


def test_reset(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    in_address = 'Sheet1!A1'
    out_address = 'Sheet1!D1'

    assert -0.02286 == round(excel_compiler.evaluate(out_address), 5)

    in_value = excel_compiler.cell_map[in_address].value

    excel_compiler._reset(excel_compiler.cell_map[in_address])
    assert excel_compiler.cell_map[out_address].value is None

    excel_compiler._reset(excel_compiler.cell_map[in_address])
    assert excel_compiler.cell_map[out_address].value is None

    excel_compiler.cell_map[in_address].value = in_value
    assert -0.02286 == round(excel_compiler.evaluate(out_address), 5)
    assert -0.02286 == round(excel_compiler.cell_map[out_address].value, 5)


def test_recalculate(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    out_address = 'Sheet1!D1'

    assert -0.02286 == round(excel_compiler.evaluate(out_address), 5)
    excel_compiler.cell_map[out_address].value = None

    excel_compiler.recalculate()
    assert -0.02286 == round(excel_compiler.cell_map[out_address].value, 5)


def test_evaluate_from_generator(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    result = excel_compiler.evaluate(
        a for a in ('trim-range!B1', 'trim-range!B2'))
    assert (24, 136) == result


def test_evaluate_empty(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    assert 0 == excel_compiler.evaluate('Empty!B1')

    excel_compiler.recalculate()
    assert 0 == excel_compiler.evaluate('Empty!B1')

    input_addrs = ['Empty!C1', 'Empty!B2']
    output_addrs = ['Empty!B1', 'Empty!B2']

    excel_compiler.trim_graph(input_addrs, output_addrs)
    excel_compiler._to_text(is_json=True)
    text_excel_compiler = ExcelCompiler._from_text(
        excel_compiler.filename, is_json=True)

    assert [0, None] == text_excel_compiler.evaluate(output_addrs)
    text_excel_compiler.set_value(input_addrs[0], 10)
    assert [10, None] == text_excel_compiler.evaluate(output_addrs)

    text_excel_compiler.set_value(input_addrs[1], 20)
    assert [10, 20] == text_excel_compiler.evaluate(output_addrs)


def test_gen_graph(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    excel.set_sheet('trim-range')
    excel_compiler._gen_graph('B2')

    with pytest.raises(ValueError, match='Unknown seed'):
        excel_compiler._gen_graph(None)


def test_value_tree_str(excel):
    out_address = 'trim-range!B2'
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
    excel_compiler._to_text(is_json=True)

    new_value = ExcelCompiler._from_text(
        excel_compiler.filename, is_json=True).evaluate(output_addrs[0])

    assert old_value == new_value


def test_trim_cells_range(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = [AddressRange('trim-range!D4:E4')]
    output_addrs = ['trim-range!B2']

    old_value = excel_compiler.evaluate(output_addrs[0])

    excel_compiler.trim_graph(input_addrs, output_addrs)

    excel_compiler._to_text()
    excel_compiler = ExcelCompiler._from_text(excel_compiler.filename)
    assert old_value == excel_compiler.evaluate(output_addrs[0])

    excel_compiler.set_value(input_addrs[0], [5, 6])
    assert old_value - 1 == excel_compiler.evaluate(output_addrs[0])

    excel_compiler.set_value(input_addrs[0], [4, 6])
    assert old_value - 2 == excel_compiler.evaluate(output_addrs[0])

    excel_compiler.set_value(tuple(next(input_addrs[0].rows)), [5, 6])
    assert old_value - 1 == excel_compiler.evaluate(output_addrs[0])


def test_evaluate_from_non_cells(excel):
    excel_compiler = ExcelCompiler(excel=excel)

    input_addrs = ['Sheet1!A11']
    output_addrs = ['Sheet1!A11:A13', 'Sheet1!D1', 'Sheet1!B11', ]

    old_values = excel_compiler.evaluate(output_addrs)

    excel_compiler.trim_graph(input_addrs, output_addrs)

    excel_compiler.to_file(file_types='yml')
    excel_compiler = ExcelCompiler.from_file(excel_compiler.filename)
    for expected, result in zip(
            old_values, excel_compiler.evaluate(output_addrs)):
        assert expected == pytest.approx(result)

    range_cell = excel_compiler.cell_map[output_addrs[0]]
    excel_compiler._reset(range_cell)
    range_value = excel_compiler.evaluate(range_cell.address)
    assert old_values[0] == range_value


def test_validate_calcs(excel, capsys):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = ['trim-range!D5']
    output_addrs = ['trim-range!B2']

    excel_compiler.trim_graph(input_addrs, output_addrs)
    excel_compiler.cell_map[output_addrs[0]].value = 'JUNK'
    failed_cells = excel_compiler.validate_calcs(output_addrs)

    assert {'trim-range!B2': ('JUNK', 136)} == failed_cells

    out, err = capsys.readouterr()
    assert '' == err
    assert 'JUNK' in out


def test_trim_cells_warn_address_not_found(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = ['trim-range!D5', 'trim-range!H1']
    output_addrs = ['trim-range!B2']

    excel_compiler.evaluate(output_addrs[0])
    excel_compiler.log.warning = mock.Mock()
    excel_compiler.trim_graph(input_addrs, output_addrs)
    assert 1 == excel_compiler.log.warning.call_count


def test_trim_cells_info_buried_input(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = ['trim-range!B1', 'trim-range!D1']
    output_addrs = ['trim-range!B2']

    excel_compiler.evaluate(output_addrs[0])
    excel_compiler.log.info = mock.Mock()
    excel_compiler.trim_graph(input_addrs, output_addrs)
    assert 2 == excel_compiler.log.info.call_count
    assert 'not a leaf node' in excel_compiler.log.info.mock_calls[1][1][0]


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

    input_addrs = ['trim-range!D5']
    output_addrs = ['trim-range!B2']

    excel_compiler.trim_graph(input_addrs, output_addrs)

    filename = excel_compiler.filename + '.pickle'
    excel_compiler.to_file(filename)

    excel_compiler = ExcelCompiler.from_file(filename)
    formula = excel_compiler.cell_map[output_addrs[0]].formula
    formula._python_code = '(x)'
    formula.lineno = 3000
    formula.filename = 'a_file'
    with pytest.raises(
            FormulaEvalError, match='File "a_file", line 3000'):
        excel_compiler.evaluate(output_addrs[0])


def test_init_cell_address_error(excel):
    with pytest.raises(ValueError):
        _CellRange('A1', excel)


def test_cell_range_repr(excel):
    cell_range = _CellRange('sheet!A1', excel)
    assert 'sheet!A1' == repr(cell_range)


def test_cell_repr(excel):
    cell_range = _Cell('sheet!A1', value=0)
    assert 'sheet!A1 -> 0' == repr(cell_range)


def test_gen_gexf(excel, tmpdir):
    excel_compiler = ExcelCompiler(excel=excel)
    filename = os.path.join(str(tmpdir), 'test.gexf')
    assert not os.path.exists(filename)
    excel_compiler.export_to_gexf(filename)

    # ::TODO: it would good to test this by comparing to an fixture/artifact
    assert os.path.exists(filename)


def test_gen_dot(excel, tmpdir):
    from unittest import mock

    excel_compiler = ExcelCompiler(excel=excel)
    with pytest.raises(ImportError, match="Package 'pydot' is not installed"):
        excel_compiler.export_to_dot('test.dot')

    import sys
    mock_imports = (
        'pydot',
    )
    for mock_import in mock_imports:
        sys.modules[mock_import] = mock.MagicMock()

    with mock.patch('networkx.drawing.nx_pydot.write_dot'):
        excel_compiler.export_to_dot('test.dot')


def test_plot_graph(excel, tmpdir):
    from unittest import mock

    excel_compiler = ExcelCompiler(excel=excel)
    with pytest.raises(ImportError,
                       match="Package 'matplotlib' is not installed"):
        excel_compiler.plot_graph()

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
    out_address = 'trim-range!B2'
    excel_compiler.evaluate(out_address)

    with mock.patch('pycel.excelcompiler.nx'):
        excel_compiler.plot_graph()


def test_structured_ref(excel):
    excel_compiler = ExcelCompiler(excel=excel)
    input_addrs = ['sref!F3']
    output_addrs = ['sref!B3']

    assert 15 == excel_compiler.evaluate(output_addrs[0])
    excel_compiler.trim_graph(input_addrs, output_addrs)

    assert 15 == excel_compiler.evaluate(output_addrs[0])

    excel_compiler.set_value(input_addrs[0], 11)
    assert 20 == excel_compiler.evaluate(output_addrs[0])
