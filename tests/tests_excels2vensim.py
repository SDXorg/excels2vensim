"""
Tests for the general functioning of the library
"""

import subprocess
import shutil
from pathlib import Path

import excels2vensim as e2v


def test_constants():

    # copy original file without data
    Path("new_files").mkdir(parents=True, exist_ok=True)
    shutil.copy2('original_files/inputs.xlsx', 'new_files/inputs.xlsx')

    e2v.Subscripts.set({'source': ['Gas', 'Oil', 'Coal']})

    obj1 = e2v.Constants(
        'q_row',
        ['source'],
        'A24',
        'This is my variable q_row',
        'Twh',
        'new_files/inputs.xlsx',
        'Region1')

    obj1.add_dimension('source', 'col')

    obj1.get_vensim()

    obj2 = e2v.Constants(
        'q_col',
        ['source'],
        'B18',
        'This is my variable q_col',
        'Twh',
        'new_files/inputs.xlsx',
        'Region1')

    obj2.add_dimension('source', 'row')

    obj2.get_vensim()

    # add more subscripts (check update method)
    e2v.Subscripts.update({
        'sector': ['A', 'B', 'C', 'D'],
        'region': ['Region1', 'Region2', 'Region3', 'Region4'],
        'out': ['Elec', 'Heat', 'Solid', 'Liquid']})

    e2v.load_from_json('jsons/constants.json')


def test_data():

    # copy original file without cellranges
    Path("new_files").mkdir(parents=True, exist_ok=True)
    shutil.copy2('original_files/inputs_data.xlsx',
                 'new_files/inputs_data.xlsx')

    # read subscripts from a mdl file
    e2v.Subscripts.read('subscripts/data.mdl')

    # test load_from_json
    e2v.load_from_json('jsons/data.json')


def test_lookup():

    # copy original file without cellranges
    Path("new_files").mkdir(parents=True, exist_ok=True)
    shutil.copy2('original_files/inputs_data2.xlsx',
                 'new_files/inputs_data2.xlsx')

    # test command line
    result = subprocess.run([
        "python3", "-m", "excels2vensim", "subscripts/data_subscripts.json",
        "jsons/lookups.json"], capture_output=True)


