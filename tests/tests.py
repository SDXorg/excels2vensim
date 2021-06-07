"""
Tests for the general functioning of the library
"""
import os
import pytest
import sdqc


def test_invalid_value():
    """
    test_invalid_value for loadinf external objects
    """

    with pytest.raises(ValueError,
                       match="Source must be a file_name or a dictionary."):
        sdqc.check(['list_not_supported'])


def test_reading_error():
    """
    test_invalid_value for loadinf external objects
    """
    from sdqc.loading_examples import read_data_william_artifacts

    path = 'jsons/reading_error'

    ext_dict = read_data_william_artifacts(
        os.path.join(path, 'reading_error.json'))

    # Load from json file
    with pytest.warns(UserWarning) as ws:
        sdqc.check(ext_dict,
                   verbose=True)

    assert len(ws) == 2
    assert all(['Not able to initialize the following object'
                in w.message.args[0] for w in ws])


def test_simple_json():
    """
    test_invalid_value for several passing test
    """
    from sdqc.loading_examples import read_data_william_artifacts

    path = 'jsons/simple_json'

    ext_dict = read_data_william_artifacts(
        os.path.join(path, 'simple_json.json'))

    # Load from dict file
    objs = sdqc.check(ext_dict,
                      os.path.join(path, 'simple_json.ini'))

    assert objs.empty

    # Load from dict file using verbose=True
    objs = sdqc.check(ext_dict,
                      os.path.join(path, 'simple_json.ini'),
                      verbose=True)

    assert set(objs['check_name'].values)\
        == set(['missing_values_series', 'missing_values_data',
                'outlier_values', 'series_monotony',
                'series_range', 'series_increment_type'])

    assert all(objs['check_pass'].values)


def test_missing_values():
    """
    test_missing_values using an mdl an py
    """

    path = 'models/missing_values'

    # Load from mdl file
    objs = sdqc.check(os.path.join(path, 'missing_values.mdl'),
                      os.path.join(path, 'missing_values.ini'))

    assert len(objs) == 14
    assert not objs[(objs['check_target'] == 'dataseries')\
                    & (objs['py_name'] == '_ext_data_data_interp1')]\
        ['check_pass'].values[0]
    assert objs[(objs['check_target'] == 'dataseries')\
                & (objs['py_short_name'] == 'data_interp1')]\
        ['check_out'].values[0] == [[4, 9], [3., 6., 10., 11.]]
    assert '_ext_constant_constant' not in objs['py_name'].values
    assert not objs[(objs['check_target'] == 'series')\
                    & (objs['py_short_name'] == 'data_hb3')]\
        ['check_pass'].values[0]

    # Load from py file using verbose=True
    objs = sdqc.check(os.path.join(path, 'missing_values.py'),
                      os.path.join(path, 'missing_values.ini'),
                      verbose=True)

    assert len(objs.index) == 29
    assert not objs[(objs['check_target'] == 'dataseries')\
                    & (objs['py_name'] == '_ext_data_data_interp1')]\
        ['check_pass'].values[0]
    assert objs[(objs['check_target'] == 'dataseries')\
                & (objs['py_short_name'] == 'data_interp1')]\
        ['check_out'].values[0] == [[4, 9], [3., 6., 10., 11.]]
    assert '_ext_constant_constant' in objs['py_name'].values
    assert objs[objs['original_name'] == 'constant']['check_pass'].values[0]
    assert objs[(objs['check_target'] == 'series')\
                & (objs['original_name'] == 'data hb1')]\
        ['check_pass'].values[0]
    assert not objs[(objs['check_target'] == 'series')\
                    & (objs['py_short_name'] == 'data_hb3')]\
        ['check_pass'].values[0]

    # check that output is a list
    objs = sdqc.check(os.path.join(path, 'missing_values.py'),
                      os.path.join(path, 'missing_values.ini'),
                      output='list')

    assert isinstance(objs, list)


def test_non_monotonous_series():
    """
    test_non_monotonous_series using a json
    """
    from sdqc.loading_examples import read_data_william_artifacts

    path = 'jsons/non_monotonous'

    ext_dict = read_data_william_artifacts(
        os.path.join(path, 'non_monotonous.json'))

    # Load from py file
    objs = sdqc.check(ext_dict,
                      os.path.join(path, 'non_monotonous.ini'))

    assert len(objs.index) == 2
    assert not any(objs['check_pass'].values)
    assert all([name == 'series_monotony'
                for name in objs['check_name'].values])
    assert set(['Excel DaTa2', 'Excel DaTa3'])\
        == set(objs['original_name'].values)
