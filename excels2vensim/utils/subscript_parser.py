"""
Functions for parsing the subscript from a .mdl file using PySD.
"""
import os
import warnings
import json

import pysd.py_backend.vensim.vensim2py as pysd_v2py

def get_subscripts(mdl_file, output=None):
    """
    Gets the subscripts from a Vensim .mdl model file.

    Parameters
    ----------
    mdl_file: str
        File path of a vensim model file to translate the subscripts.

    output: str or None (optional)
        If given the translated dictionary from a model file will be saved
        in a JSON file with the given value.

    Returns
    -------
    subscript_dict: dict
        Dictionary of the subscripts.

    """
    subscript_dict = _translate_vensim(mdl_file)

    if output:
        with open(output, 'w') as outfile:
                json.dump(subscript_dict, outfile)

    return subscript_dict

def _translate_vensim(mdl_file):
    """
    Translate Vensim's model file subscripts to a python dictionary.
    Adapted from pysd.py_backend.vensim.vensim2py.translate_vensim.

    Parameters
    ----------
    mdl_file: str
        File path of a vensim model file to translate the subscripts.

    Returns
    -------
    all_subscripts: dict
        Dictionary of the subscripts.

    Examples
    --------
    >>> translate_vensim('my_model.mdl')

    """
    root_path = os.path.split(mdl_file)[0]
    with open(mdl_file, 'r', encoding='UTF-8') as in_file:
        text = in_file.read()

    # extract model elements
    file_sections = pysd_v2py.get_file_sections(text.replace('\n', ''))

    # get subscripts
    all_subscripts = {}
    for section in file_sections:
        all_subscripts.update(_translate_section(section, root_path))

    return all_subscripts


def _translate_section(section, root_path):
    """
    Translate section subscripts to a python dictionary.
    Adapted from pysd.py_backend.vensim.vensim2py.translate_section.

    Parameters
    ----------
    section: str
        Output from pysd.py_backend.vensim.vensim2py.get_file_sections().

    root_path: str
        The root path to the model file, neede for subscripts defined with
        GET XLS/DIRECT SUBSCRIPTS.

    Returns
    -------
    subscript_dict: dict
        Dictionary of the subscripts.

    """
    model_elements = pysd_v2py.get_model_elements(section['string'])

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        # extract equation components with ':' or '<->'
        for entry in model_elements:
            if entry['kind'] == 'entry'\
                and ('<->' in entry['eqn']
                    or (':' in entry['eqn'] and not ':=' in entry['eqn'])):
                entry.update(
                    pysd_v2py.get_equation_components(entry['eqn'], root_path)
                    )


    # Create a namespace for the subscripts
    subscript_dict = {}
    for e in model_elements:
        if e['kind'] == 'subdef':
            subscript_dict[e['real_name']] = e['subs']
            for compatible in e['subs_compatibility']:
                # check if copy
                if not subscript_dict[compatible]:
                    # copy subscript to subscript_dict
                    subscript_dict[compatible] =\
                        subscript_dict[e['subs_compatibility'][compatible][0]]

    return subscript_dict