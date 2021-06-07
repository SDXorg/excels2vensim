import sys
from .excels2vensim import load_from_json

if __name__ == "__main__":
    if len(sys.argv) == 1:
        raise ValueError(
            "\nThe module must be called with the name or names of JSON"
            + "files to process, e.g.:\n\tpython -m excels2vensim "
            + "example1.json\nor\n\tpython -m excels2vensim "
            + "example1.json example2.json .. examplen.json")
    for json_file in sys.argv[1:]:
        load_from_json(json_file)