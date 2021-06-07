"""
Tests for the general functioning of the library
"""

import shutil
from pathlib import Path

import excels2vensim


def test_constants():

    # copy original file without data
    Path("new_files").mkdir(parents=True, exist_ok=True)
    shutil.copy2('original_files/inputs.xlsx', 'new_files/inputs.xlsx')

    obj1 = excels2vensim.Constants(
        'q_row',
        ['source'],
        'A24',
        'This is my variable q_row',
        'Twh',
        'new_files/inputs.xlsx',
        'Region1')

    obj1.add_dimension('source', 'col')

    obj1.get_vensim()

    obj2 = excels2vensim.Constants(
        'q_col',
        ['source'],
        'B18',
        'This is my variable q_col',
        'Twh',
        'new_files/inputs.xlsx',
        'Region1')

    obj2.add_dimension('source', 'row')

    obj2.get_vensim()

    obj4 = excels2vensim.Constants(
        'share_energy',
        ['source', 'sector', 'region', 'out'],
        'C4',
        'This is my variable',
        'dmnl',
        'new_files/inputs.xlsx',
        None)

    obj4.add_dimension('source', 'row')
    obj4.add_dimension('sector', 'col')
    obj4.add_dimension('out', 'row', 3)
    obj4.add_dimension('region', 'sheet',
                       ['Region1', 'Region2', 'Region3', 'Region4'])

    obj4.get_vensim()

def test_data():

    # copy original file without cellranges
    Path("new_files").mkdir(parents=True, exist_ok=True)
    shutil.copy2('original_files/inputs_data.xlsx',
                 'new_files/inputs_data.xlsx')

    obj1 = excels2vensim.Data(
        'var1',
        ['age', 'regions9', 'gender'],
        'E5',
        'Total population',
        'people',
        'new_files/inputs_data.xlsx',
        'GPH')

    obj1.add_dimension('regions9', 'row', 34)
    obj1.add_dimension('gender', 'row', 17)
    obj1.add_dimension('age', 'row')
    obj1.add_time('year', 'E4', 'col', 16)

    obj1.get_vensim()

    # copy original file without cellranges
    shutil.copy2('original_files/inputs_data2.xlsx',
                 'new_files/inputs_data2.xlsx')

    obj2 = excels2vensim.Data(
        'var2',
        ['age', 'regions9', 'gender'],
        'D5',
        'Total population',
        'people',
        'new_files/inputs_data2.xlsx',
        'GPH',
        'hold backward')

    obj2.add_dimension('regions9', 'sheet', [
        'EU27', 'UK', 'CNHK', 'EASTOC', 'IND',
        'LATAM', 'RUS', 'USMCA', 'LROW'
        ])
    obj2.add_dimension('gender', 'row', 17)
    obj2.add_dimension('age', 'row')
    obj2.add_time('year', 'd4', 'col', 16)

    obj2.get_vensim()

def test_lookup():

    # copy original file without cellranges
    Path("new_files").mkdir(parents=True, exist_ok=True)
    shutil.copy2('original_files/inputs_data.xlsx',
                 'new_files/inputs_data_l.xlsx')

    obj1 = excels2vensim.Lookups(
        'var1',
        ['age', 'gender', 'regions9'],
        'E5',
        'Total population',
        'people',
        'new_files/inputs_data_l.xlsx',
        'GPH')

    obj1.add_dimension('regions9', 'row', 34)
    obj1.add_dimension('gender', 'row', 17)
    obj1.add_dimension('age', 'row')
    obj1.add_x('year', 'E4', 'col', 16)

    obj1.get_vensim()

    # copy original file without cellranges
    shutil.copy2('original_files/inputs_data2.xlsx',
                 'new_files/inputs_data2_l.xlsx')

    obj2 = excels2vensim.Lookups(
        'var2',
        ['regions9', 'gender', 'age'],
        'D5',
        'Total population',
        'people',
        'new_files/inputs_data2_l.xlsx',
        'GPH')

    obj2.add_dimension('regions9', 'sheet', [
        'EU27', 'UK', 'CNHK', 'EASTOC', 'IND',
        'LATAM', 'RUS', 'USMCA', 'LROW'
        ])
    obj2.add_dimension('gender', 'row', 17)
    obj2.add_dimension('age', 'row')
    obj2.add_x('year', 'd4', 'col', 16)

    obj2.get_vensim()