Looker-PowerPoint Documentation
======================================

A command line interface for Looker PowerPoint integration that embeds YAML metadata in PowerPoint shape alternative text and replaces shapes with live Looker data.

.. toctree::
   :maxdepth: 2
   :caption: Contents:

   cli
   models
   api

Quick Start
===========

1. Install dependencies:

   .. code-block:: bash

      uv sync

2. Set up environment variables:

   .. code-block:: bash

      export LOOKERSDK_BASE_URL=https://your-looker.com
      export LOOKERSDK_CLIENT_ID=your_client_id
      export LOOKERSDK_CLIENT_SECRET=your_secret

3. Process a PowerPoint file:

   .. code-block:: bash

      uv run lppt -f your_presentation.pptx

Entry Point
===========

.. autofunction:: looker_powerpoint.cli.main

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`