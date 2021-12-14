.. excels2vensim documentation master file, created by
   sphinx-quickstart on Tue Feb 16 13:30:41 2021.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

Welcome excels2vensim library documentation!
============================================

Easy generate Vensim GET XLS/DIRECT equations with cellrange names.

|made-with-sphinx-doc|
|docs|
|PyPI license|
|Anaconda package|
|PyPI package|
|PyPI status|
|PyPI pyversions|
|pipeline|
|coverage|

.. |made-with-sphinx-doc| image:: https://img.shields.io/badge/Made%20with-Sphinx-1f425f.svg
   :target: https://www.sphinx-doc.org/

.. |docs| image:: https://readthedocs.org/projects/excels2vensim/badge/?version=latest
   :target: https://excels2vensim.readthedocs.io/en/latest/?badge=latest

.. |PyPI license| image:: https://img.shields.io/pypi/l/excels2vensim.svg
   :target: https://gitlab.com/eneko.martin.martinez/excels2vensim/-/blob/master/LICENSE

.. |Anaconda package| image:: https://anaconda.org/conda-forge/excels2vensim/badges/version.svg
   :target: https://anaconda.org/conda-forge/excels2vensim

.. |PyPI package| image:: https://badge.fury.io/py/excels2vensim.svg
    :target: https://badge.fury.io/py/excels2vensim

.. |PyPI pyversions| image:: https://img.shields.io/pypi/pyversions/excels2vensim.svg
   :target: https://pypi.python.org/pypi/excels2vensim/

.. |PyPI status| image:: https://img.shields.io/pypi/status/excels2vensim.svg
   :target: https://pypi.python.org/pypi/excels2vensim/

.. |pipeline| image:: https://gitlab.com/eneko.martin.martinez/excels2vensim/badges/master/pipeline.svg
   :target: https://gitlab.com/eneko.martin.martinez/excels2vensim/

.. |coverage| image:: https://gitlab.com/eneko.martin.martinez/excels2vensim/badges/master/coverage.svg
   :target: https://gitlab.com/eneko.martin.martinez/excels2vensim/


Systems Dynamics models often use large amounts of input data. Reading this data is complicated when it corresponds to multidimensional matrices. In the case of Vensim, if the data is read from '.xlsx' files, the data can only be read in two-dimensional tables at most, which implies the introduction of multiple equations in the model file.

The excels2vensim library aims to simplify the incorporation of equations from external data into Vensim. Given the name of the variable and information on how its dimensions are distributed, the library returns the equations for copying and pasting into the Vensim model. In addition, the library uses cellrange names to write the equations, and automatically creates the cellranges in the '.xlsx' file.

The automation of this process can save a lot of time for the modeller, as well as reduce possible human error by avoiding entering a large number of equations by hand. The library is also flexible and works for any number of dimensions. In addition, it allows reading data of which one dimension is spread over different files or sheets.

This library is able to automate the generation of the following type of equations:

#. GET XLS CONSTANTS
#. GET DIRECT CONSTANTS
#. GET XLS DATA (with or without keywords)
#. GET DIRECT DATA (with or without keywords)
#. GET XLS LOOKUPS
#. GET DIRECT LOOKUPS

The `original code for excels2vensim is available at GitLab <https://gitlab.com/eneko.martin.martinez/excels2vensim>`_. If you find a bug, or are interested in a particular feature, please use the `issue tracker from GitLab <https://gitlab.com/eneko.martin.martinez/excels2vensim/-/issues>`_. For contributions see :doc:`development <development>`.

Requirements
------------
* Python 3.7+
* PySD 2.1.1+

Contents
--------

.. toctree::
   :maxdepth: 2

   installation
   usage
   api
   development

Additional Resources
--------------------
Examples
^^^^^^^^
Some jupyter notebook examples are available in the `examples folder of the GitLab repository <https://gitlab.com/eneko.martin.martinez/excels2vensim/-/tree/master/examples>`_.

Acknowledgmentes
----------------
This library was originally developed for `H2020 LOCOMOTION <https://www.locomotion-h2020.eu>`_ project by `@eneko.martin.martinez <https://gitlab.com/eneko.martin.martinez>`_ at `Centre de Recerca Ecològica i Aplicacions Forestals (CREAF) <http://www.creaf.cat>`_.

This project has received funding from the European Union’s Horizon 2020
research and innovation programme under grant agreement No. 821105.

