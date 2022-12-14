Usage
=====

The module can be used from command line. For help about the possible options run::

    python -m excels2vensim --help

In order to define a configuration of a variable it is needed to give basic information,
such as name, description, units, input file, input sheet, list of subscripts ranges. Moreover,
information about each subscript range should be given, which specifies how the subscript is
spread along the file or files. For DATA and LOOKUPS variables, in addition, information
for the interpolation series (time and x, respectibvely) has to be added.

Some examples are given in `examples folder of the GitHub repository <https://github.com/SDXorg/excels2vensim/tree/master/examples>`_.
Please, check them to understand the usage.

.. note::
    The parameter *along* means how the data is spreaded. A row vector is spreaded along columns, while a clumn vector is spreaded along rows.
    Other possible values are sheets or files.

    .. image:: figures/along.png
        :width: 350 px
        :align: center

    The *separation* between a dimension depends on the *along* parameter. For columns and rows, the separator is the number of columns (or rows) until a change of a subscript in that dimension.
    Note that the subscripts in the Excel file should be in the same order that are in the model's subscript range.
    For sheets and files, the separation is the list of sheets (or files), ordered in the same order the subscripts are in the model's subcript range.

    See the examples for more information.


Using the GUI
-------------
A basic GUI can be launch using the following command in the terminal/PowerShell::

    python -m excels2vensim --gui

First, the subscripts file will be asked, this can be provided using the vensim model .mdl file.

.. note::
    Each time the subscript are read from a model *model_name.mdl* file, for both the GUI and the
    json files, a *model_name_subscripts.json* file will be created. For future executions,
    *model_name_subscripts.json* can be used instead of *model_name.json* for reading the subscripts faster.

Second, general information about the variable will be asked.

Third, information about each subscript will be asked and if necessary,
information about the interpolation dimension (for DATA and LOOKUPS)

Last, the user can choose if he wants to create a new variable, save the
configuration to a json file (recommended for recovery) or execute the introduced information.

Using the json files
--------------------
For using json files for configuration the folowing command can be used::

    python -m excels2vensim --output-file=my_var.txt my_model.mdl my_var_conf.json

As output-file was given, the vensim equations will be saved in *my_var.txt*. If not provided they
will be printed in the command line.

Using Python interpreter
------------------------
For using the Python interpreter the examples given above can be checked.
It is also possible to use the classes defined for each variable type.
Check :doc:`API <api>` for more information.
