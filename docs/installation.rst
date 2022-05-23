Installation
============

Installing using pip
--------------------

To install the *excels2vensim* package from the Python package index use the following command:

.. code-block:: bash

   pip install excels2vensim

Installing with conda
---------------------
To install *excels2vensim* with conda, using the conda-forge channel, into a conda environment, use the following command:

.. code-block:: bash

   conda install -c conda-forge excels2vensim

Installing from source
----------------------
To install from source, clone the project with git:

.. code-block:: bash

   git clone https://gitlab.com/eneko.martin.martinez/excels2vensim

or download the `latest version from GitLab <https://gitlab.com/eneko.martin.martinez/excels2vensim>`_.

From the project's main directory, use the following command to install it:

.. code-block:: bash

   python setup.py install

Required Dependencies
---------------------
*excels2vensim* uses `PySD <https://pysd.readthedocs.io>`_ library for reading subscripts from Vensims model file. It requires at least **Python 3.7** and **PySD 3.0**.

If not installed, *PySD* should be built automatically if you are installing via `pip`, using `conda`, or from source.
