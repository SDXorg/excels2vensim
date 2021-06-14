API
====


.. automodule:: excels2vensim
    :members:

Setting Subscripts
------------------
Setting Subscripts from a file
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
.. automethod:: Subscripts.read

Setting Subscripts from a python dictionary
^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
.. automethod:: Subscripts.set

Reading variable configuration from a json
------------------------------------------
.. autofunction:: load_from_json

Variable classes
----------------

Constants class
^^^^^^^^^^^^^^^
.. autoclass:: Constants

.. automethod:: Constants.add_dimension

.. automethod:: Constants.execute

Lookups class
^^^^^^^^^^^^^
.. autoclass:: Lookups

.. automethod:: Lookups.add_x

.. automethod:: Lookups.add_dimension

.. automethod:: Lookups.execute

Data class
^^^^^^^^^^
.. autoclass:: Data

.. automethod:: Data.add_time

.. automethod:: Data.add_dimension

.. automethod:: Data.execute