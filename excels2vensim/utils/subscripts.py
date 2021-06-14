"""
Subscript manager class.
"""
import os
import json

from .subscript_parser import get_subscripts

class Subscripts():
    """
    Class to save the subscript dictionary.
    """
    _subscript_dict = {}

    @classmethod
    def read(cls, file_name):
        """
        Read the subscripts form a .mdl or .json file.

        Parameters
        ----------
        file_name: str
            Full name of the .json or .mdl file.

        """
        file_ini, file_extension = os.path.splitext(file_name)
        if file_extension.lower() == ".mdl":
            cls.set(get_subscripts(
                file_name, file_ini + '_subscripts.json'))
        elif file_extension.lower() == ".json":
            with open(file_name) as file:
                cls.set(json.load(file))
        else:
            pass

    @classmethod
    def get(cls, key):
        """
        Get the value of key of _subscript_dict.

        Parameters
        ----------
        key: str
            Subscript range name to extract the subscripts.

        """
        return cls._subscript_dict[key]

    @classmethod
    def get_ranges(cls):
        """
        Get the list of keys of _subscript_dict.
        """
        return list(cls._subscript_dict)

    @classmethod
    def set(cls, dict):
        """
        Set _subscript_dict to input value.

        Parameters
        ----------
        dict: dict
            The subscripts dictionary.

        """
        cls.clean()
        cls.update(dict)

    @classmethod
    def update(cls, dict):
        """
        Update _subscript_dict removing trailing whitespaces from keys
        and values.

        Parameters
        ----------
        dict: dict
            The subscripts dictionary.

        """
        cls._subscript_dict.update({
            key.strip(): [value.strip() for value in values]
            for key, values in dict.items()})

    @classmethod
    def clean(cls):
        """
        Cleans the subscript dict.
        """
        cls._subscript_dict = {}

