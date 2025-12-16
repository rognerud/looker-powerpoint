Text Cell Templating
====================

The Looker PowerPoint tool enables dynamic text replacement in PowerPoint text boxes using Jinja2 templating. By connecting a shape to a Looker query, you can inject data values directly into the text.

Setup
-----

To enable templating, you must first associate the PowerPoint shape with a Looker query. This is done by setting the shape's **Alt-Text** to a JSON object that matches the :class:`~looker_powerpoint.models.LookerReference` model.

For details on the Alt-Text configuration structure, please refer to the :doc:`models` documentation. Crucially, the ``id`` field must be set to the Look ID you wish to reference.

Jinja Context
-------------

When the tool processes a text box, it fetches the data from the associated Look and makes it available to the Jinja template. The following variables are available in the context:

*   **header_rows**: A list of rows where each row is a dictionary mapping column names to values. This allows for accessing column data by name.

    *   Example: ``{{ header_rows[0].my_column_name }}``

*   **indexed_rows**: A list of rows where each row is a list of values. This allows for accessing data by row and column index.

    *   Example: ``{{ indexed_rows[0][1] }}``

*   **headers**: A list of the column headers from the Look.

Syntax Examples
---------------

Header-based Syntax
~~~~~~~~~~~~~~~~~~~

You can reference data using the column name. This is often more readable and robust against changes in column order.

.. code-block:: jinja

    Total Revenue: {{ header_rows[0].total_revenue }}
    User Count: {{ header_rows[0].users }}

Index-based Syntax
~~~~~~~~~~~~~~~~~~

You can also reference data by its position (row index and column index). This is useful if column names are dynamic or uncertain.

.. code-block:: jinja

   First Metric: {{ indexed_rows[0][0] }}
   Second Metric: {{ indexed_rows[0][1] }}


Custom Filters
--------------

colorize_positive
~~~~~~~~~~~~~~~~~

The ``colorize_positive`` filter serves to format numbers with colors based on their sign. It automatically applies marker encoding that the tool uses to colorize the text in the final PowerPoint presentation.

**Usage:**

.. code-block:: jinja

   {{ header_rows[0].growth_rate | colorize_positive }}

**Arguments:**

The filter accepts three optional arguments to customize the colors (provided as hex codes):

1.  ``positive_hex``: Color for positive numbers (default: ``#008000`` - Green)
2.  ``negative_hex``: Color for negative numbers (default: ``#C00000`` - Red)
3.  ``zero_hex``: Color for zero or non-numeric values (default: ``#000000`` - Black)

**Example with custom colors:**

.. code-block:: jinja

   {{ header_rows[0].change | colorize_positive(positive_hex="#0000FF", negative_hex="#FF0000") }}

Note: The input value is robustly parsed. It handles standard numbers, strings with currency symbols, and percentage signs.
