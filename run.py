import sys
import json

from excels2vensim import Constants

def constants(json_file):
    # open json
    vars_dict = json.load(open(json_file))

    for var, info in vars_dict.items():
        obj = Constants(
            var,
            **info)
        for dimension, along in info['dimensions'].items():
            obj.add_dimension(dimension, *along)
        obj.get_vensim()

if __name__ == "__main__":
    constants(sys.argv[1])
