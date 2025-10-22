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

      uv add looker_powerpoint

      or

      pip install looker_powerpoint

2. Set up Looker SDK credentials:

   To authenticate with your Looker instance, you need to provide your Looker base URL, client ID, and client secret.
   You can either create a looker.ini file in the project root with the following content:
   .. code-block:: ini

      [looker]
      base_url=https://your-looker.com
      client_id=your_client_id
      client_secret=your_secret

   Or add the environment variables in your shell before running the tool
   .. code-block:: bash

      export LOOKERSDK_BASE_URL=https://your-looker.com
      export LOOKERSDK_CLIENT_ID=your_client_id
      export LOOKERSDK_CLIENT_SECRET=your_secret

   or put the environment variables in a .env file in the project root:
   .. code-block:: bash

      LOOKERSDK_BASE_URL=https://your-looker.com
      LOOKERSDK_CLIENT_ID=your_client_id
      LOOKERSDK_CLIENT_SECRET=your_secret

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
