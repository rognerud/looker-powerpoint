Getting Started: Creating Your Presentation
============================================

This guide walks business users and analysts through creating a PowerPoint presentation
that automatically pulls live data from Looker. No coding required — just a few lines
of YAML in a shape's Alt Text field, and the ``lppt`` tool does the rest.

.. contents:: On this page
   :local:
   :depth: 2


Overview
--------

The **Looker PowerPoint CLI** (``lppt``) works by reading a standard ``.pptx`` file,
locating any shapes that have a special YAML snippet in their **Alt Text** description,
fetching the corresponding data from Looker, and writing the results back into a new
copy of the presentation.

You design your slides in PowerPoint as you normally would — adding tables, images,
and text boxes wherever you want data to appear. The only extra step is setting a small
block of YAML in the **Alt Text** panel for each shape that should be populated by Looker.

.. note::

   Before following the steps below, make sure ``lppt`` is installed and your Looker
   credentials are configured.  See :doc:`cli` for environment-variable details and
   the :ref:`quick-start` section on the home page for installation instructions.


Prerequisites
-------------

* A PowerPoint file (``.pptx``) — new or existing.
* Your Looker **Look ID** (the numeric ID shown in the Looker URL when you open a Look,
  e.g. ``https://your-company.looker.com/looks/42`` → ID is ``42``).
* ``lppt`` installed and Looker credentials set (see :doc:`cli`).


.. _adding-alt-text:

Step 1 — Open the Alt Text panel
---------------------------------

In PowerPoint, every shape has an **Alt Text** field. This is normally used to describe
images for screen readers, but ``lppt`` uses the *Description* box to read YAML
configuration.

To open the Alt Text panel:

1. **Right-click** on the shape (table, image, text box, or chart).
2. Choose **"Edit Alt Text…"** from the context menu.

.. figure:: _static/images/alt_text_panel.svg
   :alt: Screenshot showing how to right-click a shape and choose "Edit Alt Text…"
   :align: center
   :width: 80%

   *Replace this placeholder with a screenshot: right-clicking a shape in PowerPoint
   to open the context menu, with "Edit Alt Text…" highlighted.*

The **Alt Text** panel opens on the right side of the screen. You will see two fields:
**Title** and **Description**. Leave the *Title* field empty — enter all your YAML
configuration in the **Description** field only.

.. figure:: _static/images/alt_text_yaml_entry.svg
   :alt: Screenshot of the Alt Text panel with YAML content in the Description field
   :align: center
   :width: 80%

   *Replace this placeholder with a screenshot: the Alt Text panel open in PowerPoint
   with example YAML entered in the Description box.*


Step 2 — Write your YAML configuration
---------------------------------------

The YAML you enter tells ``lppt`` which Looker Look to fetch and how to display the
results. The only **required** field is ``id`` — the Look ID.

The simplest possible configuration is:

.. code-block:: yaml

   id: 42

Paste this into the **Description** field of the Alt Text panel and save. That is all
you need to get started. See :ref:`configuration-patterns` below for more complex
examples.


Step 3 — Run the tool
----------------------

Once you have saved your PowerPoint file with YAML in the Alt Text fields, open a
terminal in the folder containing the file and run:

.. code-block:: bash

   uv run lppt -f my_presentation.pptx

``lppt`` will:

1. Open ``my_presentation.pptx``.
2. Find all shapes with valid YAML in their Alt Text.
3. Fetch the corresponding data from Looker.
4. Write the data into the shapes.
5. Save a new file (e.g. ``my_presentation_2025-01-15.pptx``) in the same folder.

.. figure:: _static/images/run_lppt_terminal.svg
   :alt: Screenshot of a terminal running the lppt command
   :align: center
   :width: 80%

   *Replace this placeholder with a screenshot of a terminal session running
   ``uv run lppt -f my_presentation.pptx`` and showing success output.*

The original file is **never modified** — ``lppt`` always writes to a new output file.
Use ``--self`` / ``-s`` if you want to overwrite the original instead.

For full CLI options, see :doc:`cli`.


.. _configuration-patterns:

Configuration Patterns
----------------------

The sections below show the most common ways to configure a shape using YAML.
For a complete reference of every available field, see :doc:`models`.


Pattern 1 — Populate a table with all results
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You have a PowerPoint **table** shape and want it filled with all
rows and columns from a Look.

.. code-block:: yaml

   id: 42

Add this YAML to the Alt Text of a table shape. ``lppt`` will fill the table with the
results, using the first row as the header and subsequent rows as data rows.

.. tip::

   Make your table large enough to hold the expected number of rows and columns.
   Extra rows are left blank; if the data exceeds the table size, rows are truncated.

.. figure:: _static/images/table_result_example.svg
   :alt: Screenshot of a slide with a table filled with Looker data
   :align: center
   :width: 80%

   *Replace this placeholder with a screenshot: a PowerPoint slide with a table
   shape populated with live Looker data after running lppt.*


Pattern 2 — Extract a single value (row and column selection)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You want to show a single metric value inside a text box, title, or
a small single-cell table.

``row`` and ``column`` are both **0-indexed** (first row = 0, first column = 0).

.. code-block:: yaml

   id: 42
   row: 0
   column: 1

This fetches the value from the first data row (``row: 0``) and the second column
(``column: 1``) of Look 42.


Pattern 3 — Select a value by column label
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You want to pick a value by the **column name** rather than its
position. This is more robust if the column order in the Look might change.

.. code-block:: yaml

   id: 42
   row: 0
   label: "Total Revenue"

The ``label`` value must exactly match the column header label defined in Looker
(including capitalization and any special characters).


Pattern 4 — Embed a Looker chart as an image
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You have an **image placeholder** in your slide and want to replace
it with a Looker visualization rendered as a picture.

Insert any image into your slide (even a blank placeholder image), set it to the size
you want, then add the following YAML to its Alt Text:

.. code-block:: yaml

   id: 42
   result_format: png

``lppt`` will render the Look as a PNG and resize it to fit the shape's dimensions.
You can also specify explicit pixel dimensions:

.. code-block:: yaml

   id: 42
   result_format: png
   image_width: 1200
   image_height: 675

.. figure:: _static/images/image_result_example.svg
   :alt: Screenshot of a slide with an image placeholder replaced by a Looker chart
   :align: center
   :width: 80%

   *Replace this placeholder with a screenshot: a PowerPoint slide with an image
   shape replaced by a Looker visualization after running lppt.*


Pattern 5 — Populate a text box with templated values
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You have a **text box** (or slide title) containing Jinja2 template
syntax that you want to fill with Looker data.

First, write your text box content using Jinja2 double-brace syntax, for example::

   Monthly Revenue: {{ header_rows[0].total_revenue }}
   vs Last Month: {{ header_rows[0].revenue_change }}

Then set the Alt Text of that text box to:

.. code-block:: yaml

   id: 42

``lppt`` will render the Jinja2 template using the Look's results, replacing the
``{{ ... }}`` placeholders with actual values.

.. figure:: _static/images/text_template_result.svg
   :alt: Screenshot of a slide showing a text box with live data values
   :align: center
   :width: 80%

   *Replace this placeholder with a screenshot: a PowerPoint text box before and
   after running lppt, with template placeholders replaced by real Looker values.*

.. note::

   Column names in templates are derived from the Look's column labels: lowercased
   and spaces replaced with underscores. For example, ``"Total Revenue"`` becomes
   ``total_revenue``. See :doc:`templating` for the full Jinja context reference.


Pattern 6 — Color-code numbers with Jinja2 filters
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You want numbers to automatically appear **green** (positive) or
**red** (negative) in your text box.

In your text box, use the ``colorize_positive`` filter:

.. code-block:: jinja

   Growth Rate: {{ header_rows[0].growth_rate | colorize_positive }}

You can also customize the colors:

.. code-block:: jinja

   Change: {{ header_rows[0].change | colorize_positive(positive_hex="#0070C0", negative_hex="#C00000") }}

See :doc:`templating` for all available Jinja2 variables and filters.


Pattern 7 — Filter results dynamically
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You want to run the same presentation for different regions,
products, or time periods without editing it each time.

Add a ``filter`` field pointing to a Looker dimension:

.. code-block:: yaml

   id: 42
   filter: "orders.region"

Then pass the filter value at run time using the ``--filter`` CLI argument:

.. code-block:: bash

   uv run lppt -f my_presentation.pptx --filter "Europe"

This runs Look 42 filtered to rows where ``orders.region = "Europe"``.

You can also hard-code static filter overrides that always apply:

.. code-block:: yaml

   id: 42
   filter_overwrites:
     orders.status: "complete"
     orders.region: "EMEA"

``filter_overwrites`` supports complex filter values and multi-value filters:

**Multi-value filters (list syntax)**

Use a YAML list to filter on multiple values at once.  The list is automatically
joined as a comma-separated string, which Looker interprets as an *is any of*
expression:

.. code-block:: yaml

   id: 42
   filter_overwrites:
     orders.status:
       - complete
       - pending
       - shipped

This is equivalent to writing ``orders.status: "complete,pending,shipped"`` and
produces the Looker filter expression *status is complete, pending, or shipped*.

**Complex Looker filter expressions**

Any valid `Looker filter expression
<https://cloud.google.com/looker/docs/reference/filter-expressions>`_ can be
used as a value:

.. code-block:: yaml

   id: 42
   filter_overwrites:
     orders.created_date: "2024/01/01 to 2024/12/31"
     orders.revenue: ">10000"
     orders.customer_name: "%EMEA%"

**Adding filters not in the original Look**

``filter_overwrites`` can add entirely new filter dimensions — not just override
ones already defined on the underlying Look:

.. code-block:: yaml

   id: 42
   filter_overwrites:
     orders.region: "EMEA"   # added even if the Look has no region filter


Pattern 8 — Apply Looker formatting
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You want Looker to format values (e.g. currency symbols, percentage
signs, thousands separators) exactly as configured in your Looker model.

.. code-block:: yaml

   id: 42
   apply_formatting: true

When ``apply_formatting`` is ``true``, ``lppt`` asks Looker to return pre-formatted
strings (e.g. ``"$1,234,567"`` instead of ``1234567``). This is especially useful for
table shapes.


Pattern 9 — Retry on failure
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You are working with large or slow Looker queries that occasionally
time out.

.. code-block:: yaml

   id: 42
   retries: 3

``lppt`` will retry the Looker API request up to 3 times before marking the shape as
failed.

----

Troubleshooting
---------------

**Shape outlined in red after running lppt**
   ``lppt`` draws a red circle around any shape it could not populate. This usually
   means the Look ID is wrong, the query returned no data, or a column/row index is
   out of range.  Run with ``--verbose`` (or ``-vvv``) to see detailed error messages.
   Use ``--hide-errors`` to suppress the red outlines in the output file.

**Nothing happened to my shape**
   Make sure your YAML is in the **Description** field of Alt Text (not the *Title*
   field), and that it is valid YAML.  You can validate YAML at
   `yaml-online-parser.appspot.com <https://yaml-online-parser.appspot.com/>`_.

**Wrong column values**
   Column names in templates are lowercased with spaces replaced by underscores.
   Open the Look in Looker and check the exact column label. If in doubt, use
   index-based access: ``{{ indexed_rows[0][0] }}``.

----

Next Steps
----------

* :doc:`templating` — Full reference for Jinja2 variables and the ``colorize_positive``
  filter.
* :doc:`models` — Complete field reference for the ``LookerReference`` YAML schema.
* :doc:`cli` — All CLI flags, environment variables, and advanced options.
* :doc:`api` — Auto-generated API reference for developers.
