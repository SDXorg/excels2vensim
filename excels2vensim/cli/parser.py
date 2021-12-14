"""
cmdline parser
"""
from pathlib import Path
from argparse import ArgumentParser

from excels2vensim import __version__


parser = ArgumentParser(
    description="Easy generate Vensim GET XLS/DIRECT equations with "
                "cellrange names.",
    prog="excels2vensim")


#########################
# functions and actions #
#########################


def check_file(string):
    """
    Checks that subscripts file ends with .mdl or .json and that exists.

    """
    file_path = Path(string)
    if not file_path.suffix.lower() in [".json", ".mdl"]:
        parser.error(
            f"when parsing '{string}'"
            "\nThe subscript file name must be Vensim model (.mdl)"
            " or JSON (.json) file...")

    if not file_path.is_file():
        parser.error(
            f"when parsing '{string}'"
            "\nThe model/subscripts file does not exist...")

    return file_path


def check_config(string):
    """
    Checks that config file ends with .json and that exists.

    """
    file_path = Path(string)
    if not file_path.suffix.lower() == ".json":
        parser.error(
            f"when parsing '{string}'"
            "\nThe config file name must be a JSON (.json) file...")

    if not file_path.is_file():
        parser.error(
            f"when parsing '{string}'"
            "\nThe config file does not exist...")

    return file_path


###########
# options #
###########

parser.add_argument(
    "-v", "--version",
    action="version", version=f"excels2vensim {__version__}")

parser.add_argument(
    "-o", "--output-file", dest="output_file",
    type=str, metavar="FILE", default=None,
    help="output file to save the vensim equations (.txt recommended), "
         " if not given the output will be printed in the command line")

parser.add_argument(
    "-g", "--gui", dest="gui",
    action="store_true", default=False,
    help="start the GUI")


########################
# Positional arguments #
########################

parser.add_argument("subscript_file", metavar="subscript_file",
                    type=check_file, default=None, nargs="?",
                    help="Vensim, Xmile or PySD model file")

parser.add_argument("config_file", metavar="FILE",
                    type=check_config, nargs="*", default=None,
                    help="configuration json file")


#########
# Usage #
#########

parser.usage = parser.format_usage().replace(
    "usage: excels2vensim", "python -m excels2vensim")
