import sys
import os
from pathlib import Path

from .parser import parser

from excels2vensim import Subscripts, load_from_json
from excels2vensim.gui import start_gui


def main(args):
    """
    Main function. Reads user arguments, loads the models,
    runs it and saves the output

    Parameters
    ----------
    args: list
        User arguments.

    Returns
    -------
    None

    """
    options = parser.parse_args(args)

    if options.gui:  # pragma: no cover
        start_gui(options.subscript_file, options.output_file)
        sys.exit()

    # read the subscripts
    original_wd = Path.cwd()
    Subscripts.read(options.subscript_file)
    model_dir = options.subscript_file.parent
    print(f"Setting current working directory to: {model_dir}")
    os.chdir(model_dir)

    eqs = ""
    # execute json files
    for json_file in options.config_file:
        eqs += load_from_json(original_wd.joinpath(json_file))

    if options.output_file:
        with open(original_wd.joinpath(options.output_file), 'w')\
             as file:
            file.write(eqs)
    else:
        print(eqs)

    sys.exit()
