import shutil
import os
from pathlib import Path

import pytest

import excels2vensim as e2v

cdir = Path(__file__).parent


def test_col_to_num():
    """
    ExternalVariable._num_to_col and ExternalVariable._col_to_num test
    """

    col_to_num = e2v.Constants._col_to_num
    num_to_col = e2v.Constants._num_to_col

    # Check col_to_num
    assert col_to_num("A") == 0
    assert col_to_num("Z") == 25
    assert col_to_num("a") == col_to_num("B")-1
    assert col_to_num("Z") == col_to_num("aa")-1
    assert col_to_num("Zz") == col_to_num("AaA")-1

    # Check num_to_col
    assert num_to_col(col_to_num("A")) == "A"
    assert num_to_col(col_to_num("Z")) == "Z"
    assert num_to_col(col_to_num("QS")) == "QS"
    assert num_to_col(col_to_num("AZb")) == "AZB"


def test_split_excel_cell():
    """
    ExternalVariable._split_excel_cell test
    """

    split_excel_cell = e2v.Constants._split_excel_cell

    # No cells, function must return nothing
    nocells = ["A2A", "H0", "0", "5A", "A_1", "ZZZZ1", "A"]

    for nocell in nocells:
        assert not split_excel_cell(nocell)

    # Cells
    cells = [(1, 0, "A2"), (573, 7, "h574"),
             (1, 572, "Va2"), (1, 728, "ABA2")]

    for row, col, cell in cells:
        assert (row, col) == split_excel_cell(cell)


def test_clean_identifier():
    """
    Test for clean identifier
    """

    clean_identifier = e2v.Constants._clean_identifier

    assert clean_identifier('huyauJHGYG87272j') == 'huyauJHGYG87272j'
    assert clean_identifier('   huyauJHGYG87272j_*') == 'huyauJHGYG87272j'
    assert clean_identifier(' L  huyauJHGYG87272j_*') == 'L_huyauJHGYG87272j'
    assert clean_identifier('   huyauJ_h  ?  HGYG8-727+2j_*')\
        == 'huyauJ_h_HGYG8_727_2j'


def test_subscripts_errors(tmp_path):
    """
    Test for errors when adding subscripts
    """
    # Add two subscripts with step 1
    e2v.Subscripts.set({
        'source': ['Gas', 'Oil', 'Coal'],
        'sector': ['A', 'B', 'C']
        })

    obj1 = e2v.Constants(
        'q_row',
        ['source', 'sector'],
        'A24',
        'This is my variable q_row',
        'Twh',
        'tmp_dir/inputs.xlsx',
        'Region1')

    obj1.add_dimension('sector', 'col')
    obj1.add_dimension('source', 'col')

    expected = "\nTwo or more dimensions are defined along col with step 1."

    with pytest.raises(ValueError, match=expected):
        obj1.get_vensim()

    # Add two subscripts across sheet
    obj2 = e2v.Constants(
        'q_row',
        ['source', 'sector'],
        'A24',
        'This is my variable q_row',
        'Twh',
        'tmp_dir/inputs.xlsx',
        'Region1')

    obj2.add_dimension('source', 'sheet', ['A', 'B', 'C'])
    obj2.add_dimension('sector', 'sheet', ['A', 'B', 'C'])

    expected = "\nTwo or more dimensions are defined along sheet."

    with pytest.raises(ValueError, match=expected):
        obj2.get_vensim()

    # Add dimension not added to Subscripts
    obj3 = e2v.Constants(
        'q_row',
        ['region'],
        'A24',
        'This is my variable q_row',
        'Twh',
        'tmp_dir/inputs.xlsx',
        'Region1')

    expected = r"\n'region' is not in the list of subscript ranges:\n\t.*"

    with pytest.raises(ValueError, match=expected):
        obj3.add_dimension('region', 'col')

    # Add dimension with non-valid read_along
    obj4 = e2v.Constants(
        'q_row',
        ['source', 'sector'],
        'A24',
        'This is my variable q_row',
        'Twh',
        'tmp_dir/inputs.xlsx',
        'Region1')

    expected = r"\nread_along must be 'row', 'col', 'sheet' or 'file'\."

    with pytest.raises(ValueError, match=expected):
        obj4.add_dimension('source', 'error')


def test_add_series(tmp_path):
    """
    Test for add_series (add_x and add_time) method
    """
    # add_series is the method called when calling Lookups.add_x
    # and Data.add_time

    obj = e2v.Lookups(
        'q_row',
        ['source', 'sector'],
        'A24',
        'This is my variable q_row',
        'Twh',
        'tmp_dir/inputs.xlsx',
        'Region1')

    # across col
    obj.add_series('time', 'B5', 'col', 10)

    assert obj.series['cellrange'] == '$B$5:$K$5'

    # across row
    obj.add_series('time', 'B5', 'row', 10)

    assert obj.series['cellrange'] == '$B$5:$B$14'

    # across error
    expected = "\nread_along must be 'row' or 'col'."

    with pytest.raises(ValueError, match=expected):
        obj.add_series('time', 'B5', 'file', 10)


def test_write_cell_range(tmp_path):
    """
    Test for write_cell_range and Excels class
    """
    os.chdir(cdir.joinpath("tmp_dir"))
    # copy original file without cellranges
    shutil.copy2(cdir.joinpath('original_files/white.xlsx'),
                 'white.xlsx')

    name = 'my_cellrange'
    sheet = 'Sheet1'
    file = 'white.xlsx'

    write_cellrange = e2v.Constants._write_cellrange

    # write first cellrange
    write_cellrange(name, file, sheet,
                    sheet+'!$A$1:$B$2', False)

    assert file in e2v.Excels._Excels
    wb = e2v.Excels._Excels[file]

    sheetId = [sheetname_wb.lower() for sheetname_wb
               in wb.sheetnames].index(sheet.lower())

    assert name in wb.defined_names.localnames(sheetId)
    assert wb.defined_names.get(name, sheetId).attr_text == 'Sheet1!$A$1:$B$2'

    # write same cellrange (pass)
    write_cellrange(name, file, sheet,
                    sheet+'!$A$1:$B$2', False)

    assert name in wb.defined_names.localnames(sheetId)
    assert wb.defined_names.get(name, sheetId).attr_text == 'Sheet1!$A$1:$B$2'

    # write cellrange with different values (error)
    expected = f"\nTrying to write a cellrange with name '{name}'"\
               + f" at '{sheet}"\
               + r"!\$A\$3:\$B\$4'\. However, "\
               + f"'{name}' already exist in '{sheet}"\
               + r"!\$A\$1:\$B\$2'\n"\
               + r"Use force=True to overwrite it\."

    with pytest.raises(ValueError, match=expected):
        write_cellrange(name, file, sheet,
                        sheet+'!$A$3:$B$4', False)

    assert name in wb.defined_names.localnames(sheetId)
    assert wb.defined_names.get(name, sheetId).attr_text == 'Sheet1!$A$1:$B$2'

    # write cellrange with different values (force to remove old)
    write_cellrange(name, file, sheet,
                    sheet+'!$A$3:$B$4', True)

    assert name in wb.defined_names.localnames(sheetId)
    assert wb.defined_names.get(name, sheetId).attr_text == 'Sheet1!$A$3:$B$4'

    # close file
    e2v.Excels.save_and_close()

    assert file not in e2v.Excels._Excels


def test_load_from_json():
    """
    Test for load_from_json for non-navild type. Valid ones are tested in
    tests_excels2vensim.py
    """

    expected = "\n Invalid type of variable 'NonValid' for 'share_energy'."\
               + " It must be 'constants', 'lookups' or 'data'."
    # invalid var name
    with pytest.raises(ValueError, match=expected):
        e2v.load_from_json(cdir.joinpath('jsons/non_valid.json'))
