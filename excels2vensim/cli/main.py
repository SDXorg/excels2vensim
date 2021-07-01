import sys

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

    if options.gui:
        start_gui(options.subscript_file, options.output_file)
        sys.exit()

    # read the subscripts
    Subscripts.read(options.subscript_file)

    eqs = ""
    # execute json files
    for json_file in options.config_file:
        eqs += load_from_json(json_file)

    if options.output_file:
        with open(options.output_file, 'w') as file:
            file.write(eqs)
    else:
        print(eqs)
