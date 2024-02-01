"""
Tests for the argument parser
"""
import subprocess

import pytest

from excels2vensim.cli import parser


def test_help():
    subprocess.run(["python3", "-m", "excels2vensim"])


@pytest.mark.parametrize(
    "string,function,error_message",
    [
        (
            "my_file.mwl",
            parser.check_file,
            "when parsing 'my_file.mwl'"
            "\nThe subscript file name must be Vensim model (.mdl)"
            " or JSON (.json) file..."
        ),
        (
            "my_file.json",
            parser.check_file,
            "when parsing 'my_file.json'"
            "\nThe model/subscripts file does not exist..."
        ),
        (
            "my_file.mwl",
            parser.check_config,
            "when parsing 'my_file.mwl'"
            "\nThe config file name must be a JSON (.json) file..."
        ),
        (
            "my_file.json",
            parser.check_config,
            "when parsing 'my_file.json'"
            "\nThe config file does not exist..."
        ),
    ],
    ids=["check_file_invalid_suffix", "check_file_inexistent",
         "check_config_invalid_suffix", "check_config_inexistent"]
)
def test_parser_errors(capsys, string, function, error_message):
    with pytest.raises(SystemExit) as pytest_wrapped_e:
        function(string)

    assert pytest_wrapped_e.value.code == 2
    captured = capsys.readouterr()
    assert error_message in captured.err
