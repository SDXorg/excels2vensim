"""
Functions for parsing the subscript from a .mdl file using PySD.
"""
import json

from pysd.translators.vensim.vensim_file import VensimFile
from pysd.translators.vensim.vensim_element import SubscriptRange
from pysd.builders.python.subscripts import SubscriptManager


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
    model = VensimFile(mdl_file)
    # parse model file without parsing the sections
    model.parse(parse_all=False)

    subscripts = []
    # parse only first section (main)
    model.sections[0].parse(parse_all=False)
    for element in model.sections[0].elements:
        if (":" in element.equation and ":=" not in element.equation)\
           or "<->" in element.equation:
            # parse elements with ":" or "<->" in the equation
            new_element = element.parse()
            if isinstance(new_element, SubscriptRange):
                # if parsed element is a SubscriptRange get the
                # AbstractSubscriptRange
                subscripts.append(new_element.get_abstract_subscript_range())

    return SubscriptManager(subscripts, model.mdl_path).subscripts
