Command Line Interface
======================

The Looker PowerPoint CLI (``lppt``) processes PowerPoint presentations by replacing shapes with live Looker data.

Main CLI Class
--------------

.. autoclass:: looker_powerpoint.cli.Cli
   :members:
   :undoc-members:
   :show-inheritance:
   :exclude-members: HEADER

Entry Point
-----------

.. autofunction:: looker_powerpoint.cli.main

Environment Variables
---------------------

The CLI requires these environment variables for Looker API access:

.. envvar:: LOOKERSDK_BASE_URL

   Your Looker instance URL

.. envvar:: LOOKERSDK_CLIENT_ID

   Looker API client ID

.. envvar:: LOOKERSDK_CLIENT_SECRET

   Looker API client secret

Additional optional environment variables:
------------------------------
.. envvar:: LOOKERSDK_TIMEOUT

   If you experience timeouts with large queries, you can set this variable to increase the timeout duration (in seconds). The default is 120 seconds for most Looker SDK clients.
   Example: ``LOOKERSDK_TIMEOUT=300``
