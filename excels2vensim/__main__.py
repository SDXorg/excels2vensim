import sys
import getopt
from .excels2vensim import load_from_json, Subscripts

def get_user_input(args):
    """
    Get user input.

    Parameters
    ----------
    args: list
        User arguments list.

    Returns
    -------
    subscript_file, json_files, output: str, list, str

    """
    output = None

    try:
        opts, args = getopt.getopt(
            args, "ho:", ['help', 'output-file=', 'gui'])

        for opt, arg in opts:
            if opt in ('-h', '--help'):
                print(
                    'usage: python -m excels2vensim [options]'
                    ' [subscript_file] [config_file] [config_file] [...]\n'
                    '\npositional arguments:\n'
                    '\tsubscript_file\t\tjson or .mdl file with the '
                    'subscripts\n'
                    '\tconfig_file\t\tconfiguration json file.\n'
                    '\noptional arguments:\n'
                    '-h, --help\t\t help menu \n'
                    '-o OUT_FILE, --output_file=OUT_FILE\t\t the output '
                    'file to save Vensim equations. If not given the '
                    'vensim equation will be printed in the standard output.\n'
                    '--gui\t\t start the GUI.\n\n'
                      )

                sys.exit()

            elif opt in ('-o', '--output-file'):
                output = arg
            elif opt == '--gui':
                from .gui import start_gui
                start_gui()
                sys.exit()
            else:
                pass

        return args[0], args[1:], output

    except getopt.GetoptError:
        raise ValueError(
            "Wrong parameter definition (run 'python -m excels2vensim -h'"
            + "to see the description of available parameters)")


def main(subscript_file, json_files, output=None):
    """
    Run the user input

    Parameters
    ----------
    subscript_file: str
        Subscripts json file.

    json_files: list
        List of the json configuration files for the elements to create.

    output: str or None (optional)
        Name of the file to write Vensim equations. If None they will be
        printed using the standard output. Default is None.

    """
    # read the subscripts
    Subscripts.read(subscript_file)

    eqs = ""
    # execute json files
    for json_file in json_files:
        eqs += load_from_json(json_file)

    if output:
        with open(output, 'w') as file:
            file.write(eqs)
    else:
        print(eqs)


if __name__ == "__main__":
    if len(sys.argv) == 1:
        from .gui import start_gui
        start_gui()
    else:
        subscript_file, json_files, output =\
            get_user_input(sys.argv[1:])
        main(subscript_file, json_files, output)
