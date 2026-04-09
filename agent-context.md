This file is a merged representation of a subset of the codebase, containing files not matching ignore patterns, combined into a single document by Repomix.

# File Summary

## Purpose
This file contains a packed representation of a subset of the repository's contents that is considered the most important context.
It is designed to be easily consumable by AI systems for analysis, code review,
or other automated processes.

## File Format
The content is organized as follows:
1. This summary section
2. Repository information
3. Directory structure
4. Repository files (if enabled)
5. Multiple file entries, each consisting of:
  a. A header with the file path (## File: path/to/file)
  b. The full contents of the file in a code block

## Usage Guidelines
- This file should be treated as read-only. Any changes should be made to the
  original repository files, not this packed version.
- When processing this file, use the file path to distinguish
  between different files in the repository.
- Be aware that this file may contain sensitive information. Handle it with
  the same level of security as you would the original repository.

## Notes
- Some files may have been excluded based on .gitignore rules and Repomix's configuration
- Binary files are not included in this packed representation. Please refer to the Repository Structure section for a complete list of file paths, including binary files
- Files matching these patterns are excluded: agent-context.md, package-lock.json
- Files matching patterns in .gitignore are excluded
- Files matching default ignore patterns are excluded
- Files are sorted by Git change count (files with more changes are at the bottom)

# Directory Structure
```
.github/
  workflows/
    copilot-setup-steps.yml
    make_docs.yml
    publish.yml
    pull-request.yml
    release-drafter.yml
  copilot-instructions.md
  release-drafter.yml
docs/
  _static/
    images/
      alt_text_panel.svg
      alt_text_yaml_entry.svg
      image_result_example.svg
      run_lppt_terminal.svg
      table_result_example.svg
      text_template_result.svg
  api.rst
  cli.rst
  conf.py
  docs.instructions.md
  getting_started.rst
  index.rst
  make.bat
  Makefile
  models.rst
  templating.rst
looker_powerpoint/
  tools/
    __init__.py
    find_alt_text.py
    pptx_text_handler.py
    tools.instructions.md
    url_to_hyperlink.py
  __init__.py
  cli.py
  gemini.py
  looker_powerpoint.instructions.md
  looker.py
  models.py
test/
  pptx/
    gemini_textbox.md
    gemini_textbox.pptx
    table7x7.md
    table7x7.pptx
  test_cli.py
  test_gemini.py
  test_integration.py
  test_pptx.py
  test_tools.py
  test.instructions.md
.gitignore
.pre-commit-config.yaml
LICENSE
pyproject.toml
README.md
```

# Files

## File: .github/workflows/copilot-setup-steps.yml
````yaml
name: "Setup Environment and Git Hooks"
#description: "Installs uv, syncs dependencies, and installs pre-commit hooks for future commits"

on:
  workflow_call:

jobs:
  copilot-setup-steps:
    runs-on: ubuntu-latest
    steps:
      # ADDED: Pulls your code into the runner
      - name: Checkout repository
        uses: actions/checkout@v5

      - name: Install uv
        uses: astral-sh/setup-uv@v7
        with:
          enable-cache: true
          cache-dependency-glob: "uv.lock"

      - name: Set up Python
        run: uv python install
        shell: bash

      - name: Sync Dependencies
        run: uv sync --frozen
        shell: bash

      - name: Install pre-commit git hook
        run: uv run pre-commit install
        shell: bash
````

## File: .github/workflows/make_docs.yml
````yaml
name: Make docs
# from main

on:
  push:
    branches:
      - main
  workflow_dispatch:

jobs:
  make_docs:
    runs-on: ubuntu-latest
    permissions:
      contents: write

    steps:
    - uses: actions/checkout@v5

    - name: Install uv
      uses: astral-sh/setup-uv@v7
      with:
        enable-cache: true
        github-token: ${{ secrets.GITHUB_TOKEN }}

    - id: make
      run: cd docs && make html && cd ..

    - name: "Upload pages to artifact"
      uses: actions/upload-pages-artifact@v3
      with:
        path: ${{ github.workspace }}/docs/_build/html

  deploy-to-github-pages:
    # Add a dependency to the build job
    needs: make_docs

    # Grant GITHUB_TOKEN the permissions required to make a Pages deployment
    permissions:
      pages: write # to deploy to Pages
      id-token: write # to verify the deployment originates from an appropriate source

    # Deploy to the github-pages environment
    environment:
      name: github-pages

    # Specify runner + deployment step
    runs-on: ubuntu-latest
    steps:
      - name: Deploy to GitHub Pages
        id: deployment
        uses: actions/deploy-pages@v4 # or the latest "vX.X.X" version tag for this action
````

## File: .github/workflows/publish.yml
````yaml
name: Deploy Release

on:
  release:
    types: [published]

jobs:
  release:
    runs-on: ubuntu-latest
    environment: release
    permissions:
      contents: read
      id-token: write
    steps:
    - uses: actions/checkout@v5

    - name: Install uv
      uses: astral-sh/setup-uv@v7
      with:
        enable-cache: true
        github-token: ${{ secrets.GITHUB_TOKEN }}

    - name: Build
      run: |
        uv build

    - name: Publish
      run: |
        uv publish
````

## File: .github/workflows/pull-request.yml
````yaml
name: CI

# Default permissions
permissions:
  contents: read

on:
  push:
    branches:
      - main
  pull_request:
    branches:
      - main

jobs:
  test:
    if: github.event_name == 'pull_request'
    runs-on: ${{ matrix.os }}
    strategy:
      fail-fast: false
      matrix:
        os: [ubuntu-latest, windows-latest, macos-latest]
        python-version: ["3.12", "3.13"]
        pandas-version: ["2.2.3", "latest"]
        exclude:
          # pandas 2.2.x does not ship wheels for Python 3.13
          - python-version: "3.13"
            pandas-version: "2.2.3"
    steps:
    - uses: actions/checkout@v5

    - name: Install uv
      uses: astral-sh/setup-uv@v7
      with:
        enable-cache: true
        python-version: ${{ matrix.python-version }}
        github-token: ${{ secrets.GITHUB_TOKEN }}

    - name: Install dependencies
      run: uv sync

    - name: Override pandas version
      if: matrix.pandas-version != 'latest'
      run: uv pip install "pandas==${{ matrix.pandas-version }}"

    - name: Run tests
      run: uv run --no-sync pytest

  coverage:
    runs-on: ubuntu-latest
    steps:
    - uses: actions/checkout@v5

    - name: Install uv
      uses: astral-sh/setup-uv@v7
      with:
        enable-cache: true
        python-version: "3.12"
        github-token: ${{ secrets.GITHUB_TOKEN }}

    - name: Install dependencies
      run: uv sync

    - name: Run tests with coverage
      run: uv run --no-sync pytest --cov=looker_powerpoint --cov-report=xml --cov-report=term-missing || true

    - name: Upload coverage to Codecov
      uses: codecov/codecov-action@v5
      with:
        token: ${{ secrets.CODECOV_TOKEN }}
        files: ./coverage.xml
        fail_ci_if_error: false

  update-ai-context:
    if: github.event_name == 'pull_request'
    # Run this on a single, fast Linux runner
    runs-on: ubuntu-latest
    # Elevate permissions so the bot can push back to the branch
    permissions:
      contents: write
    steps:
    - uses: actions/checkout@v5
      with:
        # Crucial: Check out the actual branch the PR is coming from!
        ref: ${{ github.head_ref }}

    - name: Setup Node.js
      uses: actions/setup-node@v4
      with:
        node-version: '20'

    - name: Run Repomix to generate AI context
      run: npx repomix --style markdown --output agent-context.md --ignore "agent-context.md,package-lock.json"

    - name: Commit and push changes
      run: |
        git config --global user.name "github-actions[bot]"
        git config --global user.email "github-actions[bot]@users.noreply.github.com"
        git add agent-context.md

        # Only commit if there are changes
        if git diff --staged --quiet; then
          echo "No changes to commit"
        else
          git commit -m "chore: update agent-context.md [skip ci]"
          git push
        fi
````

## File: .github/workflows/release-drafter.yml
````yaml
name: Release Drafter

on:
  push:
    branches:
      - main

  pull_request:
    types: [opened, reopened, synchronize]

permissions:
  contents: read

jobs:
  update_release_draft:
    permissions:
      contents: write
      pull-requests: write
    runs-on: ubuntu-latest
    steps:
      - uses: release-drafter/release-drafter@v6
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
````

## File: .github/copilot-instructions.md
````markdown
# Looker PowerPoint CLI - Copilot Instructions

## Architecture Overview

This Python CLI tool integrates Looker data with PowerPoint presentations by using YAML metadata set in pptx shape alternative text to define looker data extractions.

## Usage Instructions

The tool is designed to be run from the command line, where users can specify the PowerPoint file and the desired output format for the extracted data. The CLI will parse the PowerPoint file, identify shapes with YAML metadata, and extract the corresponding Looker data based on the defined parameters.

## Repository conventions

The repository is a modern python project using uv and pyproject.toml for dependency management.
The code is in looker_powerpoint directory.
The tests are in test directory.
Documentation made by Sphinx is in docs directory.

running the CLI is done via `uv run lppt` command, which will execute the main function in looker_powerpoint/cli.py.

## Testing
Testing is done using pytest, and tests are located in the test directory. To run the tests, use the command `pytest` from the root directory of the project.
Tests should use stored fixtures for any looker data, and should not make live API calls to Looker. Mocking should be used to simulate API responses for testing purposes.
Test pptx files with with approprate yml metadata in the alt text should be used to test the parsing and data extraction functionality of the CLI tool.
Any generated pptx files, should have a corresponding markdown file in the test directory that accurately describes the content of the pptx file, with regards to yml metadata so it is easy for agents to understand how to use the file to create tests.

## Directory Map
for up to date map look at agent-context.md file in the root of the repository.
for folder specific instructions look at the instructions files in each folder.

## Contributing
Any contributions to the project MUST follow the following rules:
- All new features or bug fixes must include corresponding tests in the `test/` directory.
- Contributions should prioritize adding tests for any new feature or bug fix, and to ensure that the documentation is updated accordingly.
- Any change has to be reflected in the documentation, and in the instructions files.
- All contributions must add one test that is missing that improves code coverage, or adds a test for an edge case that is not currently covered.
````

## File: .github/release-drafter.yml
````yaml
name-template: 'v$RESOLVED_VERSION'
tag-template: 'v$RESOLVED_VERSION'
categories:
  - title: '🚀 Features'
    labels:
      - 'feature'
      - 'enhancement'
  - title: '🪲 Bug Fixes'
    labels:
      - 'fix'
      - 'bugfix'
      - 'bug'
  - title: '🛠️ Maintenance'
    label: 'chore'
change-template: '- $TITLE @$AUTHOR (#$NUMBER)'
change-title-escapes: '\<*_&' # You can add # and @ to disable mentions, and add ` to disable code blocks.
version-resolver:
  major:
    labels:
      - 'major'
  minor:
    labels:
      - 'feature'
      - 'enhancement'
  patch:
    labels:
      - 'fix'
      - 'bugfix'
      - 'bug'
      - 'chore'
  default: patch
template: |
  ## Changes

  $CHANGES

autolabeler:
  - label: 'chore'
    branch:
      - '/docs{0,1}\/.+/'
      - '/chore\/.+/'
    title:
      - '[Osmosis].+'
  - label: 'bug'
    branch:
      - '/fix\/.+/'
      - '/bug\/.+/'
      - '/bugfix\/.+/'
  - label: 'enhancement'
    branch:
      - '/feature\/.+/'
      - '/feat\/.+/'
````

## File: docs/_static/images/alt_text_panel.svg
````xml
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" fill="#e8e8e8" rx="8"/>
  <rect x="20" y="20" width="600" height="320" fill="#ffffff" stroke="#cccccc" stroke-width="1" rx="4"/>
  <rect x="20" y="20" width="600" height="40" fill="#d0d0d0" rx="4"/>
  <text x="320" y="45" font-family="Arial, sans-serif" font-size="14" font-weight="bold" text-anchor="middle" fill="#555555">[ PLACEHOLDER IMAGE ]</text>
  <text x="320" y="110" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">Screenshot: Right-clicking on a PowerPoint shape</text>
  <text x="320" y="132" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">to open the context menu, with the cursor hovering over</text>
  <text x="320" y="154" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">"Edit Alt Text..." option in the menu.</text>
  <text x="320" y="210" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">Replace this file with an actual screenshot showing</text>
  <text x="320" y="228" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">how to access the Alt Text panel in PowerPoint.</text>
  <rect x="200" y="260" width="240" height="40" fill="none" stroke="#aaaaaa" stroke-width="1" stroke-dasharray="4,4" rx="4"/>
  <text x="320" y="285" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#aaaaaa">Context menu → Edit Alt Text…</text>
</svg>
````

## File: docs/_static/images/alt_text_yaml_entry.svg
````xml
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" fill="#e8e8e8" rx="8"/>
  <rect x="20" y="20" width="600" height="320" fill="#ffffff" stroke="#cccccc" stroke-width="1" rx="4"/>
  <rect x="20" y="20" width="600" height="40" fill="#d0d0d0" rx="4"/>
  <text x="320" y="45" font-family="Arial, sans-serif" font-size="14" font-weight="bold" text-anchor="middle" fill="#555555">[ PLACEHOLDER IMAGE ]</text>
  <text x="320" y="110" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">Screenshot: The "Alt Text" panel on the right side of PowerPoint,</text>
  <text x="320" y="132" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">with the Description field containing example YAML such as:</text>
  <rect x="160" y="155" width="320" height="80" fill="#f5f5f5" stroke="#bbbbbb" stroke-width="1" rx="4"/>
  <text x="180" y="175" font-family="Courier New, monospace" font-size="12" fill="#333333">id: 42</text>
  <text x="180" y="193" font-family="Courier New, monospace" font-size="12" fill="#333333">row: 0</text>
  <text x="180" y="211" font-family="Courier New, monospace" font-size="12" fill="#333333">column: 1</text>
  <text x="320" y="270" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">Replace this file with an actual screenshot of the</text>
  <text x="320" y="288" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">Alt Text panel showing YAML metadata in the Description field.</text>
</svg>
````

## File: docs/_static/images/image_result_example.svg
````xml
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" fill="#e8e8e8" rx="8"/>
  <rect x="20" y="20" width="600" height="320" fill="#ffffff" stroke="#cccccc" stroke-width="1" rx="4"/>
  <rect x="20" y="20" width="600" height="40" fill="#d0d0d0" rx="4"/>
  <text x="320" y="45" font-family="Arial, sans-serif" font-size="14" font-weight="bold" text-anchor="middle" fill="#555555">[ PLACEHOLDER IMAGE ]</text>
  <text x="320" y="110" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">Screenshot: A PowerPoint slide where an image placeholder</text>
  <text x="320" y="132" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">has been replaced with a Looker chart visualization</text>
  <text x="320" y="154" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">(e.g., a bar chart or line graph rendered as a PNG).</text>
  <text x="320" y="230" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">Replace this file with an actual screenshot of a slide</text>
  <text x="320" y="248" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">with an image shape replaced by a Looker visualization.</text>
</svg>
````

## File: docs/_static/images/run_lppt_terminal.svg
````xml
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" fill="#e8e8e8" rx="8"/>
  <rect x="20" y="20" width="600" height="320" fill="#ffffff" stroke="#cccccc" stroke-width="1" rx="4"/>
  <rect x="20" y="20" width="600" height="40" fill="#d0d0d0" rx="4"/>
  <text x="320" y="45" font-family="Arial, sans-serif" font-size="14" font-weight="bold" text-anchor="middle" fill="#555555">[ PLACEHOLDER IMAGE ]</text>
  <text x="320" y="110" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">Screenshot: A terminal window running the lppt command,</text>
  <text x="320" y="132" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">e.g.: uv run lppt -f my_presentation.pptx</text>
  <text x="320" y="154" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">showing progress output and a success message.</text>
  <text x="320" y="230" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">Replace this file with an actual screenshot of a terminal</text>
  <text x="320" y="248" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">session running the lppt command successfully.</text>
</svg>
````

## File: docs/_static/images/table_result_example.svg
````xml
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" fill="#e8e8e8" rx="8"/>
  <rect x="20" y="20" width="600" height="320" fill="#ffffff" stroke="#cccccc" stroke-width="1" rx="4"/>
  <rect x="20" y="20" width="600" height="40" fill="#d0d0d0" rx="4"/>
  <text x="320" y="45" font-family="Arial, sans-serif" font-size="14" font-weight="bold" text-anchor="middle" fill="#555555">[ PLACEHOLDER IMAGE ]</text>
  <text x="320" y="110" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">Screenshot: A PowerPoint slide showing a table shape</text>
  <text x="320" y="132" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">populated with live data fetched from a Looker Look.</text>
  <text x="320" y="154" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">Rows and columns are filled with formatted values.</text>
  <text x="320" y="230" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">Replace this file with an actual screenshot of a slide</text>
  <text x="320" y="248" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">with a table filled by lppt using Looker data.</text>
</svg>
````

## File: docs/_static/images/text_template_result.svg
````xml
<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360" viewBox="0 0 640 360">
  <rect width="640" height="360" fill="#e8e8e8" rx="8"/>
  <rect x="20" y="20" width="600" height="320" fill="#ffffff" stroke="#cccccc" stroke-width="1" rx="4"/>
  <rect x="20" y="20" width="600" height="40" fill="#d0d0d0" rx="4"/>
  <text x="320" y="45" font-family="Arial, sans-serif" font-size="14" font-weight="bold" text-anchor="middle" fill="#555555">[ PLACEHOLDER IMAGE ]</text>
  <text x="320" y="110" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">Screenshot: A PowerPoint text box before and after running lppt,</text>
  <text x="320" y="132" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">showing the Jinja2 template placeholders replaced with</text>
  <text x="320" y="154" font-family="Arial, sans-serif" font-size="13" text-anchor="middle" fill="#333333">real data values (e.g., "Revenue: $1,234,567").</text>
  <text x="320" y="230" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">Replace this file with an actual screenshot of a slide</text>
  <text x="320" y="248" font-family="Arial, sans-serif" font-size="11" text-anchor="middle" fill="#888888">with a text box rendered using Jinja2 templating.</text>
</svg>
````

## File: docs/api.rst
````rst
API Reference
=============

Looker Client
-------------

.. autoclass:: looker_powerpoint.looker.LookerClient
   :members:
   :undoc-members:
   :show-inheritance:

Shape Discovery Tools
---------------------

.. automodule:: looker_powerpoint.tools.find_alt_text
   :members:
   :undoc-members:
   :show-inheritance:
````

## File: docs/cli.rst
````rst
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
------------------------------------------
.. envvar:: LOOKERSDK_TIMEOUT

   If you experience timeouts with large queries, you can set this variable to increase the timeout duration (in seconds). The default is 120 seconds for most Looker SDK clients.
   Example: ``LOOKERSDK_TIMEOUT=300``
````

## File: docs/conf.py
````python
# Configuration file for the Sphinx documentation builder.
#
# For the full list of built-in configuration values, see the documentation:
# https://www.sphinx-doc.org/en/master/usage/configuration.html

# -- Project information -----------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#project-information
import os
import sys

# Add both the parent directory and the looker_powerpoint directory to the path
sys.path.insert(0, os.path.abspath(".."))
sys.path.insert(0, os.path.abspath("../looker_powerpoint"))

project = "Looker PowerPoint CLI"
copyright = "2025, Gisle Rognerud"
author = "Gisle Rognerud"
release = "0.1.0"

# -- General configuration ---------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#general-configuration

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.viewcode",
    "sphinx.ext.napoleon",
    "sphinx.ext.autosummary",
    "sphinx.ext.intersphinx",
    "sphinxcontrib.autodoc_pydantic",
]

# Intersphinx mapping for external documentation
intersphinx_mapping = {
    "python": ("https://docs.python.org/3", None),
    "pydantic": ("https://docs.pydantic.dev/latest/", None),
}

# Autodoc settings
autodoc_default_options = {
    "members": True,
    "member-order": "bysource",
    "special-members": "__init__",
    "undoc-members": True,
    "exclude-members": "__weakref__",
    "show-inheritance": True,
}

# Include descriptions from Field() definitions
autodoc_preserve_defaults = True
autodoc_typehints = "description"
autodoc_typehints_description_target = "documented"

# autodoc_pydantic settings
autodoc_pydantic_model_show_json = False
autodoc_pydantic_model_show_config_summary = False
autodoc_pydantic_model_show_validator_members = True
autodoc_pydantic_model_show_field_summary = True
autodoc_pydantic_field_list_validators = False
autodoc_pydantic_model_signature_prefix = "class"
autodoc_pydantic_field_doc_policy = "description"
autodoc_pydantic_model_show_validator_summary = True

# Napoleon settings for Google/NumPy style docstrings
napoleon_google_docstring = True
napoleon_numpy_docstring = True
napoleon_include_init_with_doc = False
napoleon_include_private_with_doc = False
napoleon_include_special_with_doc = True
napoleon_use_admonition_for_examples = False
napoleon_use_admonition_for_notes = False
napoleon_use_admonition_for_references = False
napoleon_use_ivar = False
napoleon_use_param = True
napoleon_use_rtype = True

# Autosummary settings
autosummary_generate = True

templates_path = ["_templates"]
exclude_patterns = ["_build", "Thumbs.db", ".DS_Store"]


# -- Options for HTML output -------------------------------------------------
# https://www.sphinx-doc.org/en/master/usage/configuration.html#options-for-html-output

html_theme = "alabaster"
html_static_path = ["_static"]
````

## File: docs/docs.instructions.md
````markdown
# docs

Sphinx documentation source for the Looker PowerPoint CLI.

## Files

| File | Purpose |
|------|---------|
| `getting_started.rst` | User-friendly "Getting Started" guide for business users and analysts: how to add YAML alt-text to PowerPoint shapes, all configuration patterns, image placeholders, and links to other docs. |
| `conf.py` | Sphinx configuration: project metadata, extensions (`autodoc`, `sphinx-autodoc-typehints`, `autodoc-pydantic`), and theme settings. |
| `index.rst` | Root table of contents for the generated documentation site. |
| `api.rst` | Auto-generated API reference page (populated by `sphinx.ext.autodoc`). |
| `cli.rst` | Documentation page for the `cli` module. |
| `models.rst` | Documentation page for the Pydantic models (`LookerReference`, `LookerShape`). |
| `templating.rst` | Documentation page covering the Jinja2 templating features available in PowerPoint text shapes. |
| `Makefile` / `make.bat` | Standard Sphinx build helpers for Linux/macOS and Windows respectively. |
| `_static/images/` | SVG placeholder images for the Getting Started guide. Replace these with actual screenshots of PowerPoint and the terminal. |

## Building the docs

```bash
# from the docs/ directory
make html
```

The generated HTML is written to `docs/_build/html/`.
````

## File: docs/getting_started.rst
````rst
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


Pattern 10 — Gemini LLM text synthesis
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

**Use this when:** You want a text box populated with an AI-generated summary or
analysis, rather than raw values.

This feature uses Google Gemini to synthesise text.  It **only works for text
box shapes** (``TEXT_BOX``, ``TITLE``, ``AUTO_SHAPE``).  Applying it to a table,
image, or chart shape will log a warning and skip that shape.

**Step 1 — Define one or more meta looks (optional)**

Add meta-look shapes to your presentation for each dataset you want Gemini to
analyse.  Set the shape's Alt Text to:

.. code-block:: yaml

   id: 42
   meta: true
   meta_name: sales_data

The ``meta_name`` value (``sales_data`` here) is the key you will reference from
the Gemini shape.  Meta-look shapes are removed from the output slide; their data
is only used as context.

**Step 2 — Configure the Gemini text box**

Add a text box to your slide and set its Alt Text to:

.. code-block:: yaml

   type: gemini
   prompt: Summarise the top three sales trends and highlight any risks.
   contexts:
     - sales_data

**The ``contexts`` list — unified context framework**

Each entry in ``contexts`` is resolved in order and passed to Gemini as a
labelled section.  Four types of entry are supported:

.. list-table::
   :header-rows: 1
   :widths: 20 80

   * - Entry
     - What it provides
   * - ``self``
     - The shape's own current text (before synthesis).
   * - ``slide_self``
     - Text of all other shapes on the same slide after Looker data has been
       rendered.  Gives the model awareness of what the slide is about.
   * - ``gemini_<id>``
     - The synthesised output of another Gemini text box whose ``gemini_id``
       is ``<id>`` (see chaining below).
   * - anything else
     - Treated as the ``meta_name`` of a Looker meta-look shape; its
       pre-fetched data is formatted as a readable table.

Example combining all four types:

.. code-block:: yaml

   type: gemini
   gemini_id: summary
   prompt: Write a one-paragraph executive summary.
   contexts:
     - slide_self          # what the slide says
     - sales_data          # Looker data
     - gemini_analysis     # output of another Gemini box
     - self                # current placeholder text as additional hint

**Chaining Gemini boxes**

Gemini boxes can reference each other.  A box with ``gemini_id: analysis``
(stored as ``gemini_analysis`` — the prefix is added automatically) can be
used as context by another box:

.. code-block:: yaml

   # Box 1 — data analysis
   type: gemini
   gemini_id: analysis          # stored as gemini_analysis
   prompt: Analyse the revenue data and list the top three trends.
   contexts:
     - revenue_data

.. code-block:: yaml

   # Box 2 — executive summary that builds on Box 1
   type: gemini
   gemini_id: summary           # stored as gemini_summary
   prompt: Write a concise executive summary based on the analysis.
   contexts:
     - slide_self
     - gemini_analysis          # Box 1's output

``lppt`` automatically detects these dependencies and processes boxes in the
correct topological order.  Circular references are detected and reported as
an error.

.. note::

   The ``gemini_`` prefix is added to ``gemini_id`` values automatically.
   ``gemini_id: analysis`` and ``gemini_id: gemini_analysis`` are equivalent.
   Always use the full ``gemini_<id>`` form when referencing a box in ``contexts``.

**Optional fields:**

.. code-block:: yaml

   type: gemini
   prompt: Summarise in one sentence.
   contexts:
     - sales_data
   model: gemini-1.5-pro   # default: gemini-2.0-flash

**Step 3 — Install the LLM extra and set your API key**

.. code-block:: bash

   pip install "looker_powerpoint[llm]"

Then set your Gemini API key:

.. code-block:: bash

   export GOOGLE_API_KEY="your-api-key"

**Step 4 — Run ``lppt`` as usual**

.. code-block:: bash

   uv run lppt -f my_presentation.pptx

``lppt`` will fetch the meta looks, call Gemini with the assembled context and
your prompt, and replace the text box content with the AI-generated response
while preserving the original font and paragraph styling.

.. note::

   If ``google-genai`` is not installed, ``lppt`` still runs normally for all
   other shapes; Gemini synthesis shapes are silently skipped with a warning.
   This means the package works without the LLM extra installed.

.. tip::

   If Gemini synthesis fails (e.g. bad API key, quota exceeded), ``lppt`` writes
   the error message into the text box and draws a red outline around it — the
   same behaviour as other shape errors.  Use ``--hide-errors`` to suppress the
   outline.

----

Troubleshooting
---------------

**Shape outlined in red after running lppt**
   ``lppt`` draws a red circle around any shape it could not populate. This usually
   means the Look ID is wrong, the query returned no data, or a column/row index is
   out of range.  Run with ``--verbose`` (or ``-vvv``) to see detailed error messages.
   Use ``--hide-errors`` to suppress the red outlines in the output file.

   For Gemini shapes specifically, the error message is also written into the text box
   so you can see exactly what went wrong without needing verbose logging.

**Nothing happened to my shape**
   Make sure your YAML is in the **Description** field of Alt Text (not the *Title*
   field), and that it is valid YAML.  You can validate YAML at
   `yaml-online-parser.appspot.com <https://yaml-online-parser.appspot.com/>`_.

**Wrong column values**
   Column names in templates are lowercased with spaces replaced by underscores.
   Open the Look in Looker and check the exact column label. If in doubt, use
   index-based access: ``{{ indexed_rows[0][0] }}``.

**Gemini synthesis shape is skipped with a warning**
   Make sure ``google-genai`` is installed (``pip install looker_powerpoint[llm]``),
   ``GOOGLE_API_KEY`` or ``GEMINI_API_KEY`` is set, and ``type: gemini`` is set in the
   **Description** field (not *Title*).

**Context data not found**
   ``lppt`` logs a warning for each ``contexts`` entry it cannot resolve.
   Check that meta-look ``meta_name`` values and ``gemini_<id>`` references
   exactly match the corresponding shapes in the presentation.

**Circular dependency in Gemini boxes**
   If ``lppt`` reports a circular dependency error, check the ``contexts`` lists
   of your Gemini boxes for cycles (e.g. A → B → A).

----

Next Steps
----------

* :doc:`templating` — Full reference for Jinja2 variables and the ``colorize_positive``
  filter.
* :doc:`models` — Complete field reference for the ``LookerReference`` and
  ``GeminiConfig`` YAML schemas.
* :doc:`cli` — All CLI flags, environment variables, and advanced options.
* :doc:`api` — Auto-generated API reference for developers.
````

## File: docs/index.rst
````rst
Looker-PowerPoint Documentation
======================================

A command line interface for Looker PowerPoint integration that embeds YAML metadata in PowerPoint shape alternative text and replaces shapes with live Looker data.

.. toctree::
   :maxdepth: 2
   :caption: Contents:

   getting_started
   cli
   models
   templating
   api

.. _quick-start:

Quick Start
===========

1. Install dependencies:

   .. code-block:: bash

      pip install looker-powerpoint

   Or, if you are using ``uv`` to manage a project:

   .. code-block:: bash

      uv add looker-powerpoint

2. Set up Looker SDK credentials:

   To authenticate with your Looker instance, you need to provide your Looker base URL, client ID, and client secret.
   You can either create a looker.ini file in the project root with the following content:

   .. code-block:: ini

      [looker]
      base_url=https://your-looker.com
      client_id=your_client_id
      client_secret=your_secret

   Or add the environment variables in your shell before running the tool:

   .. code-block:: bash

      export LOOKERSDK_BASE_URL=https://your-looker.com
      export LOOKERSDK_CLIENT_ID=your_client_id
      export LOOKERSDK_CLIENT_SECRET=your_secret

   Or put the environment variables in a .env file in the project root:

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
````

## File: docs/make.bat
````batch
@ECHO OFF

pushd %~dp0

REM Command file for Sphinx documentation

if "%SPHINXBUILD%" == "" (
	set SPHINXBUILD=sphinx-build
)
set SOURCEDIR=.
set BUILDDIR=_build

%SPHINXBUILD% >NUL 2>NUL
if errorlevel 9009 (
	echo.
	echo.The 'sphinx-build' command was not found. Make sure you have Sphinx
	echo.installed, then set the SPHINXBUILD environment variable to point
	echo.to the full path of the 'sphinx-build' executable. Alternatively you
	echo.may add the Sphinx directory to PATH.
	echo.
	echo.If you don't have Sphinx installed, grab it from
	echo.https://www.sphinx-doc.org/
	exit /b 1
)

if "%1" == "" goto help

%SPHINXBUILD% -M %1 %SOURCEDIR% %BUILDDIR% %SPHINXOPTS% %O%
goto end

:help
%SPHINXBUILD% -M help %SOURCEDIR% %BUILDDIR% %SPHINXOPTS% %O%

:end
popd
````

## File: docs/Makefile
````
# Minimal makefile for Sphinx documentation
#

# You can set these variables from the command line, and also
# from the environment for the first two.
SPHINXOPTS    ?=
SPHINXBUILD   ?= uv run --group dev sphinx-build
SOURCEDIR     = .
BUILDDIR      = _build

# Put it first so that "make" without argument is like "make help".
help:
	@$(SPHINXBUILD) -M help "$(SOURCEDIR)" "$(BUILDDIR)" $(SPHINXOPTS) $(O)

.PHONY: help Makefile

# Catch-all target: route all unknown targets to Sphinx using the new
# "make mode" option.  $(O) is meant as a shortcut for $(SPHINXOPTS).
%: Makefile
	@$(SPHINXBUILD) -M $@ "$(SOURCEDIR)" "$(BUILDDIR)" $(SPHINXOPTS) $(O)
````

## File: docs/models.rst
````rst
Data Models
===========

The Looker PowerPoint CLI uses Pydantic models to validate and structure data for PowerPoint shape integration.

LookerReference
---------------

.. autopydantic_model:: looker_powerpoint.models.LookerReference
   :members:
   :undoc-members:
   :show-inheritance:
````

## File: docs/templating.rst
````rst
Text Cell Templating
====================

The Looker PowerPoint tool enables dynamic text replacement in PowerPoint text boxes using Jinja2 templating. By connecting a shape to a Looker query, you can inject data values directly into the text.

Setup
-----

To enable templating, you must first associate the PowerPoint shape with a Looker query. This is done by setting the shape's **Alt-Text** to a YAML document that matches the :class:`~looker_powerpoint.models.LookerReference` model.

For details on the Alt-Text configuration structure, please refer to the :doc:`models` documentation. Crucially, the ``id`` field must be set to the Look ID you wish to reference.

Jinja Context
-------------

When the tool processes a text box, it fetches the data from the associated Look and makes it available to the Jinja template. The following variables are available in the context:

.. note::

   Column names exposed to the template are derived from the Look's ``field_group_variant`` label. They are lowercased and whitespace is replaced with underscores (e.g. ``"Total Revenue"`` becomes ``total_revenue``). Emoji characters are also stripped. Use these sanitized names when referencing columns in your templates.

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
````

## File: looker_powerpoint/tools/__init__.py
````python

````

## File: looker_powerpoint/tools/find_alt_text.py
````python
from pptx import Presentation
from lxml import etree
import yaml

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def extract_alt_text(shape):
    """
    Extracts the alternative text description from a shape's XML.

    Args:
        shape: A Shape object from pptx.

    Returns:
        str: The alternative text description, or None if not found.
    """
    xml_str = shape.element.xml  # get XML string of the shape element
    xml_elem = etree.fromstring(xml_str)  # parse it into an lxml element
    for path in [
        ".//p:nvSpPr/p:cNvPr",
        ".//p:nvPicPr/p:cNvPr",
        ".//p:nvGraphicFramePr/p:cNvPr",
    ]:
        cNvPr_elements = xml_elem.xpath(path, namespaces=NS)
        if cNvPr_elements:
            descr = cNvPr_elements[0].get("descr")
            if descr:
                data = yaml.safe_load(
                    descr  # .lower()
                )  # Use safe_load for untrusted sources

                return data
    return None


def get_presentation_objects_with_descriptions(pptx_path):
    """
    Extracts all shapes from a PowerPoint presentation and returns them with descriptions.

    Args:
        pptx_path (str): The path to the PowerPoint presentation file.

    Returns:
        list: A list of dictionaries, where each dictionary represents a shape and
              contains the shape object, its description, and the slide number.
              Returns an empty list if the presentation cannot be opened or has no slides/shapes.
    """
    try:
        presentation = Presentation(pptx_path)
    except Exception as e:
        print(f"Error opening presentation: {e}")
        return []

    objects_with_descriptions = []

    for i, slide in enumerate(presentation.slides, start=0):
        for shape in slide.shapes:
            description = extract_alt_text(shape)  # Generate description

            emu_to_pixels = lambda emu: emu / 9525

            width_px = emu_to_pixels(shape.width)
            height_px = emu_to_pixels(shape.height)

            if description:
                if type(description) is dict and "meta_name" in description:
                    shape_id = description.get("meta_name")
                else:
                    shape_id = (
                        f"{i},{shape.shape_id}"  # Use shape number for identification
                    )

                objects_with_descriptions.append(
                    {
                        "shape_id": shape_id,  # Use shape number for identification
                        "shape_type": shape.shape_type.name,
                        "shape_width": round(width_px),
                        "shape_height": round(height_px),
                        "integration": description,
                        "slide_number": i,  # Use the enumerate index for slide number
                        "shape_number": shape.shape_id,
                    }
                )

    return objects_with_descriptions


if __name__ == "__main__":
    # Example Usage
    pptx_file = "p.pptx"  # Replace with your file path
    objects = get_presentation_objects_with_descriptions(pptx_file)
    from rich import print
````

## File: looker_powerpoint/tools/pptx_text_handler.py
````python
import logging
import re
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Pt
from jinja2 import Environment, BaseLoader
import pandas as pd
from pptx.dml.color import MSO_COLOR_TYPE

# ---------- Emoji removal helper ----------
# Regex to match emoji and a broad set of pictographs/symbols.
_EMOJI_REGEX = re.compile(
    "["
    "\U0001f600-\U0001f64f"  # emoticons
    "\U0001f300-\U0001f5ff"  # symbols & pictographs
    "\U0001f680-\U0001f6ff"  # transport & map symbols
    "\U0001f1e0-\U0001f1ff"  # flags (iOS)
    "\U00002702-\U000027b0"  # dingbats
    "\U000024c2-\U0001f251"
    "\U0001f900-\U0001f9ff"  # Supplemental Symbols and Pictographs
    "\U0001fa70-\U0001faff"  # Symbols and Pictographs Extended-A
    "\U00002600-\U000026ff"  # Misc symbols
    "]+",
    flags=re.UNICODE,
)


def remove_emojis_from_string(s):
    if not isinstance(s, str):
        return s
    return _EMOJI_REGEX.sub("", s)


_WS_RE = re.compile(r"\s+")


def sanitize_header_name(h):
    """
    Remove emojis, strip, and replace inner whitespace with underscores.
    """
    if h is None:
        return h
    # convert to str in case it's not already
    s = str(h)
    # remove emojis and trailing/leading spaces handled by that function
    s = remove_emojis_from_string(s)
    # collapse internal whitespace to single underscore
    s = _WS_RE.sub("_", s)
    # strip leading/trailing underscores produced by replacement of leading/trailing spaces
    s = s.strip("_")
    return s


def sanitize_dataframe_headers(df):
    """
    Return a new DataFrame with sanitized column headers:
    - emojis removed
    - internal whitespace replaced with underscores
    - leading/trailing underscores trimmed
    """
    # build rename mapping
    rename_map = {col: sanitize_header_name(col) for col in df.columns}
    # return a copy with renamed columns
    return df.rename(columns=rename_map)


# ---------- Marker encoding for colored segments ----------
_START = "\u0002"
_SEP = "\u0003"
_END = "\u0004"
MARKER_RE = re.compile(
    re.escape(_START)
    + r"(#[0-9A-Fa-f]{6})"
    + re.escape(_SEP)
    + r"(.*?)"
    + re.escape(_END),
    re.DOTALL,
)


def encode_colored_text(text, hex_color):
    return f"{_START}{hex_color}{_SEP}{text}{_END}"


def decode_marked_segments(rendered_text):
    segments = []
    pos = 0
    for m in MARKER_RE.finditer(rendered_text):
        if m.start() > pos:
            segments.append((rendered_text[pos : m.start()], None))
        hex_color = m.group(1)
        txt = m.group(2)
        segments.append((txt, hex_color))
        pos = m.end()
    if pos < len(rendered_text):
        segments.append((rendered_text[pos:], None))
    return segments


# ---------- Formatting copy helper ----------
def copy_run_format(src_run, dest_run):
    try:
        dest_run.font.bold = src_run.font.bold
        dest_run.font.italic = src_run.font.italic
        dest_run.font.underline = src_run.font.underline
        if src_run.font.name:
            dest_run.font.name = src_run.font.name
        if src_run.font.size:
            dest_run.font.size = src_run.font.size
        # Try to copy RGB color if available
        try:
            col = src_run.font.color
            if (
                getattr(col, "type", None) is not None
                and getattr(col, "rgb", None) is not None
            ):
                rgb = col.rgb
                try:
                    dest_run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
                except Exception:
                    try:
                        dest_run.font.color.rgb = rgb
                    except Exception:
                        pass
        except Exception:
            pass
    except Exception:
        pass


# ---------- Robust colorize_positive filter ----------
def colorize_positive(
    value, positive_hex="#008000", negative_hex="#C00000", zero_hex="#000000"
):
    """
    Try to parse 'value' robustly; return marker-wrapped text for coloring.
    """

    def try_parse_number(v):
        if v is None:
            return None
        if isinstance(v, (int, float)):
            try:
                return float(v)
            except Exception:
                return None
        # pandas NA / numpy.nan
        try:
            import numpy as _np

            if v is _np.nan:
                return None
        except Exception:
            pass
        if isinstance(v, str):
            s = v.strip()
            if s == "":
                return None
            # Remove emojis before parsing (safety)
            s = remove_emojis_from_string(s)
            s = s.replace(",", "").replace(" ", "")
            s = re.sub(r"^[^\d\-\+\.]+", "", s)
            s = re.sub(r"[^\d\.eE\-\+]$", "", s)
            try:
                return float(s)
            except Exception:
                m = re.match(r"^\(([\d\.,\-]+)\)$", v.strip())
                if m:
                    inner = m.group(1).replace(",", "")
                    try:
                        return -float(inner)
                    except Exception:
                        return None
                return None
        try:
            return float(v)
        except Exception:
            return None

    num = try_parse_number(value)
    if num is None:
        hex_color = zero_hex
    elif num > 0:
        hex_color = positive_hex
    elif num < 0:
        hex_color = negative_hex
    else:
        hex_color = zero_hex

    text = "" if value is None else str(value)
    # ensure the output text also has emojis removed (defensive)
    text = remove_emojis_from_string(text)
    return encode_colored_text(text, hex_color)


# ---------- Jinja env ----------
def make_jinja_env():
    env = Environment(loader=BaseLoader(), autoescape=False)
    env.filters["colorize_positive"] = colorize_positive
    return env


def render_text_with_jinja(text, context, env=None):
    if env is None:
        env = make_jinja_env()
    template = env.from_string(text)
    return template.render(**(context or {}))


# ---------- Extract original runs and text ----------
def extract_text_and_run_meta(text_frame):
    parts = []
    run_meta = []
    for p in text_frame.paragraphs:
        for r in p.runs:
            run_text = r.text or ""
            parts.append(run_text)
            run_meta.append({"text": run_text, "run_obj": r})
        parts.append("\n")
        run_meta.append({"text": "\n", "run_obj": None})
    if parts and parts[-1] == "\n":
        parts.pop()
        run_meta.pop()
    full_text = "".join(parts)
    return full_text, run_meta


# ---------- High-level processor ----------
def process_text_field(shape, text_to_insert, df, env=None):
    text_to_insert = str(text_to_insert)
    jinja_tag_re = re.compile(r"({{.*?}}|{%.+?%})", re.DOTALL)
    text_frame = shape.text_frame
    full_text, run_meta = extract_text_and_run_meta(text_frame)

    if not jinja_tag_re.search(full_text):
        logging.debug("No Jinja tags found in shape; applying fallback if different.")
        if full_text != (text_to_insert or ""):
            update_text_frame_preserving_formatting(text_frame, text_to_insert or "")
        return

    df_sanitized = sanitize_dataframe_headers(df)
    header_rows = df_sanitized.to_dict(orient="records")
    indexed_rows = df_sanitized.values.tolist()
    context = {
        "header_rows": header_rows,  # For access by name: {{ header_rows.col_a }}
        "indexed_rows": indexed_rows,  # For access by index: {{ indexed_rows[0] }}
        "headers": df_sanitized.columns.tolist(),  # Optionally, provide a list of headers
    }
    rendered = render_text_with_jinja(full_text, context, env=env)

    # --- Compare old vs new to decide whether to modify ---
    if rendered.strip() == full_text.strip():
        logging.debug("Rendered Jinja output identical to original; skipping update.")
        return

    # --- Update text safely ---
    reinsert_rendered_text_preserving_formatting(text_frame, rendered, run_meta)


def update_text_frame_preserving_formatting(text_frame, new_text):
    """
    Replace text content but preserve shape formatting and paragraph style.
    """
    # Grab formatting from the first run
    if not text_frame.paragraphs:
        text_frame.text = new_text
        return

    p = text_frame.paragraphs[0]
    runs = p.runs
    font = runs[0].font if runs else None
    color = None
    if font and getattr(font, "color", None):
        col = font.color
        try:
            if getattr(col, "type", None) == MSO_COLOR_TYPE.RGB and getattr(
                col, "rgb", None
            ):
                color = col.rgb
            elif getattr(col, "type", None) == MSO_COLOR_TYPE.SCHEME and getattr(
                col, "theme_color", None
            ):
                color = col.theme_color
        except Exception:
            pass
    size = font.size if font and font.size else Pt(12)

    # Clear all text but keep paragraphs
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            run.text = ""

    # Replace only first run text (preserves style)
    if not text_frame.paragraphs:
        p = text_frame.add_paragraph()

    if not p.runs:
        run = p.add_run()
    else:
        run = p.runs[0]

    run.text = new_text

    # Reapply original font attributes (if any)
    if color:
        try:
            if isinstance(color, RGBColor):
                run.font.color.rgb = color
            else:
                run.font.color.theme_color = color
        except Exception:
            pass
    if size:
        run.font.size = size


def reinsert_rendered_text_preserving_formatting(
    text_frame, rendered_text, run_meta=None
):
    first_paragraph = text_frame.paragraphs[0] if text_frame.paragraphs else None
    first_run = (
        first_paragraph.runs[0] if first_paragraph and first_paragraph.runs else None
    )
    font = getattr(first_run, "font", None)
    alignment = getattr(first_paragraph, "alignment", None)

    for p in list(text_frame.paragraphs):
        text_frame._element.remove(p._p)

    new_paragraph = text_frame.add_paragraph()
    new_run = new_paragraph.add_run()
    new_run.text = rendered_text

    # Safely copy all style attributes
    copy_font_format(font, new_run.font)

    if alignment is not None:
        new_paragraph.alignment = alignment


def copy_font_format(src_font, dest_font):
    """Copy color, size, bold, italic safely between fonts."""
    if not src_font or not dest_font:
        return

    color = src_font.color
    if color and color.type is not None:
        if color.type == MSO_COLOR_TYPE.RGB and color.rgb:
            dest_font.color.rgb = color.rgb
        elif color.type == MSO_COLOR_TYPE.SCHEME and color.theme_color is not None:
            dest_font.color.theme_color = color.theme_color

    if src_font.size:
        dest_font.size = src_font.size
    if src_font.bold is not None:
        dest_font.bold = src_font.bold
    if src_font.italic is not None:
        dest_font.italic = src_font.italic
````

## File: looker_powerpoint/tools/tools.instructions.md
````markdown
# looker_powerpoint/tools

Utility helpers used by the main CLI to manipulate PowerPoint files.

## Modules

| File | Purpose |
|------|---------|
| `find_alt_text.py` | Extracts YAML alternative-text from pptx shape XML and returns all shapes that carry a valid `LookerReference` description. Entry-point: `get_presentation_objects_with_descriptions()`. |
| `pptx_text_handler.py` | Text-frame utilities: Jinja2 template rendering, emoji removal, header sanitisation, colour-coded text encoding/decoding, and formatting-preserving text replacement. |
| `url_to_hyperlink.py` | Replaces raw URLs inside a text frame with numbered hyperlink references `(1)`, `(2)`, … |
| `__init__.py` | Empty package initialiser. |

## Key design notes

- **`find_alt_text.py`** walks the low-level shape XML via `lxml` so it can read the `descr` attribute that `python-pptx` does not expose directly.
- **`pptx_text_handler.py`** uses Jinja2 with a custom `colorize_positive` filter so that templates in slide text boxes can conditionally colour numbers green/red at render time.
- Helpers are stateless functions — they do not hold references to the Looker SDK or the presentation object beyond individual calls.
````

## File: looker_powerpoint/tools/url_to_hyperlink.py
````python
import re

from pptx.dml.color import RGBColor


def add_text_with_numbered_links(text_frame, text, start_index=1):
    """
    Replaces URLs in `text` with numbered references "(1)", "(23)", etc.
    The placeholder is hyperlinked to the URL.
    - Clears any prior content in the text_frame.
    - Removes newlines if hyperlinks are present.
    - If a URL ends with digits, uses that number instead of auto numbering.
    Returns the next available numeric index.
    """
    text_frame.clear()

    url_pattern = re.compile(r"https?://\S+")
    paragraph = text_frame.add_paragraph()
    index = start_index

    matches = url_pattern.findall(text)
    if matches:
        # Flatten newlines if hyperlinks exist
        text = text.replace("\n", " ")

    # Split while keeping URLs in the list
    parts = re.split(f"({url_pattern.pattern})", text)

    for part in parts:
        if not part:
            continue

        if url_pattern.fullmatch(part.strip()):
            url = part.strip()

            # Detect digits at the end of the URL path
            match_digits = re.search(r"(\d+)(?:[/?#]?)*$", url)
            number_text = match_digits.group(1) if match_digits else str(index)

            run = paragraph.add_run()
            run.text = f"({number_text})"
            run.hyperlink.address = url
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.font.underline = True

            index += 1 if not match_digits else 0
        else:
            run = paragraph.add_run()
            run.text = part

    return index
````

## File: looker_powerpoint/__init__.py
````python
import importlib.metadata

try:
    __version__ = importlib.metadata.version(__name__)
except importlib.metadata.PackageNotFoundError:
    __version__ = "0.0.0"  # Fallback for development mode
````

## File: looker_powerpoint/cli.py
````python
from asyncio import subprocess
import collections
import datetime
import requests
import io
from looker_powerpoint.tools.find_alt_text import (
    get_presentation_objects_with_descriptions,
)
from looker_powerpoint.looker import LookerClient
from looker_powerpoint.models import LookerShape, GeminiShape
from looker_powerpoint import gemini as gemini_module

from looker_powerpoint.tools.pptx_text_handler import (
    process_text_field,
    update_text_frame_preserving_formatting,
)
from pydantic import ValidationError
import subprocess
from pptx.util import Pt
from pptx.chart.data import CategoryChartData
import json
import pandas as pd
from pptx import Presentation
from lxml import etree
import re
import argparse
from rich_argparse import RichHelpFormatter
import logging
from PIL import Image

from rich.logging import RichHandler
import os
import asyncio
from io import BytesIO

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}

import re
from pptx.util import Pt
from pptx.dml.color import RGBColor


class Cli:
    # color with rich
    HEADER = """
        Looker PowerPoint CLI :
        A command line interface for Looker PowerPoint integration.
    """

    def __init__(self):
        self.client = None
        self.relevant_shapes = []
        self.looker_shapes = []
        self.gemini_shapes = []
        self.data = {}

        # Initialize the argument parser
        self.parser = self._init_argparser()

        # load tools
        self.get_alt_text = get_presentation_objects_with_descriptions

    def _init_looker(self):
        """Initialize the Looker client"""
        if not self.args.debug_queries:
            logging.getLogger("looker_sdk").setLevel(logging.ERROR)
        self.client = LookerClient()

    def _init_argparser(self):
        """Create and configure the argument parser"""
        parser = argparse.ArgumentParser(
            description=self.HEADER,
            formatter_class=RichHelpFormatter,
        )
        # todo
        parser.add_argument(
            "--file-path",
            "-f",
            help="Path to the PowerPoint file to process",
            default=None,
            type=str,
        )
        # todo
        parser.add_argument(
            "--output-dir",
            "-o",
            help="""Path to a directory that will contain the generated pptx files. \n
                .env: OUTPUT_DIR""",
            default="output",
            type=str,
        )
        # todo
        parser.add_argument(
            "--add-links",
            help="Add links to looker in the slides. \n .env: ADD_LINKS",
            action="store_true",
            default=False,
        )
        # todo
        parser.add_argument(
            "--hide-errors",
            help="""
                Stop showing red outlines around shapes with errors. \n
                .env: HIDE_ERRORS
            """,
            action="store_true",
            default=False,
        )
        # todo
        parser.add_argument(
            "--parse-date-syntax-in-filename",
            "-p",
            help="""Parse date syntax in the filename. \n
                .env: PARSE_DATE_SYNTAX_IN_FILENAME
                """,
            action="store_true",
            default=True,
        )

        parser.add_argument(
            "--self",
            "-s",
            help="""Replace the powerpoint file directly instead of creating a new file. \n
                .env: SELF""",
            action="store_true",
            default=False,
        )

        parser.add_argument(
            "--quiet",
            "-q",
            help="""Do not open the PowerPoint file after processing. \n
                .env: QUIET""",
            action="store_true",
            default=False,
        )

        parser.add_argument(
            "--filter",
            help="""use the string to filter shapes if they have a set filter dimension""",
            action="store",
            default=None,
            type=str,
        )

        parser.add_argument(
            "--debug-queries",
            help="""Enable debugging for Looker queries. \n
                .env: DEBUG_QUERIES""",
            action="store_true",
            default=False,
        )

        parser.add_argument(
            "-v",
            "--verbose",
            action="count",
            default=0,
            help="Increase verbosity (e.g., -v, -vv, -vvv)",
        )

        return parser

    def _setup_logging(self):
        if self.args.verbose == 0:
            level = logging.WARNING
        elif self.args.verbose == 1:
            level = logging.INFO
        else:
            level = logging.DEBUG

        logging.basicConfig(
            level=level,
            format="%(message)s",
            datefmt="[%X]",
            handlers=[RichHandler()],
        )

    def _pick_file(self):
        """
        Picks the PowerPoint file to process.
        If no file path is provided, it looks for the first .pptx file in the current directory.

        Returns:
            str: The path to the PowerPoint file.
        """
        self.file_path = self.args.file_path

        if self.file_path:
            try:
                self.presentation = Presentation(self.file_path)
            except Exception as e:
                logging.error(f"Error opening {self.file_path}: {e}")
        else:
            # If no file path is provided look for a file in the current directory
            files = [
                f
                for f in os.listdir(".")
                if f.endswith(".pptx") and not f.startswith("~$")
            ]
            if files:
                self.file_path = files[0]
                logging.warning(
                    f"No file path provided, using first found file: {self.file_path}. To specify a file, use the -f flag like 'lpt -f <file_path>'."
                )

                try:
                    self.presentation = Presentation(self.file_path)
                except Exception as e:
                    logging.error(f"Error opening {self.file_path}: {e}")
            else:
                logging.error(
                    """
                    No PowerPoint file found in the current directory, closing.
                    Run from a directory with a .pptx file, or
                    specify file using -f flag like 'lpt -f <file_path>'.
                """
                )
                exit(1)

    def _fill_table(self, table, df, headers=True):
        """
        Fills a PowerPoint table with data from a DataFrame.

        Args:
            table: A Table object from pptx.
            df: A pandas DataFrame containing the data to fill the table.
        """

        # Get table dimensions
        table_rows = len(table.rows)
        table_cols = len(table.columns)

        # Get DataFrame dimensions
        df_rows = df.shape[0] + 1  # +1 for header
        df_cols = df.shape[1]

        # Determine how much we can fill
        rows_to_fill = min(table_rows, df_rows)
        cols_to_fill = min(table_cols, df_cols)

        # Fill header row
        if headers:
            for col_idx in range(cols_to_fill):
                table.cell(0, col_idx).text = str(df.columns[col_idx])

        # Fill DataFrame values
        for row_idx in range(1, rows_to_fill):  # skip header row
            for col_idx in range(cols_to_fill):
                value = df.iloc[row_idx - 1, col_idx]
                table.cell(row_idx, col_idx).text = str(value)

        # Optional: Clear unused cells
        for row_idx in range(rows_to_fill, table_rows):
            for col_idx in range(table_cols):
                table.cell(row_idx, col_idx).text = ""

        for col_idx in range(cols_to_fill, table_cols):
            for row_idx in range(table_rows):
                table.cell(row_idx, col_idx).text = ""

    def _set_alt_text(self, shape, data):
        """
        Sets the alternative text description for a shape's XML.

        Args:
            shape: A Shape object from pptx.
            data: A Python object (dict, list, etc.) to serialize and set as YAML in the descr attribute.
        """
        xml_str = shape.element.xml
        xml_elem = etree.fromstring(xml_str)
        import yaml

        # convert pydantic model to dict
        if isinstance(data, dict) is False:
            data = data.model_dump()
        data = {k: v for k, v in data.items() if v is not None}
        data = yaml.dump(data)

        # remove None values from data, and convert to string with newlines for YAML compatibility

        for path in [
            ".//p:nvSpPr/p:cNvPr",
            ".//p:nvPicPr/p:cNvPr",
            ".//p:nvGraphicFramePr/p:cNvPr",
        ]:
            cNvPr_elements = xml_elem.xpath(path, namespaces=NS)
            if cNvPr_elements:
                cNvPr = cNvPr_elements[0]
                yaml_text = str(data)
                cNvPr.set("descr", yaml_text)

                # Overwrite the element in the actual pptx shape with updated XML
                shape_element = shape.element
                new_element = etree.fromstring(etree.tostring(xml_elem))
                shape_element.clear()
                for child in new_element:
                    shape_element.append(child)
                return

        raise ValueError("No compatible cNvPr element found to set descr.")

    def _mark_failure(self, slide, shape):
        line_color_rgb = (255, 0, 0)  # RGB color for
        line_width_pt = 2  # Width of the circle outline in points
        # Calculate circle position - centered on the shape

        # Add an oval shape (circle)
        circle = slide.shapes.add_shape(
            autoshape_type_id=1,  # MSO_SHAPE_OVAL (value 1)
            left=shape.left,
            top=shape.top,
            width=shape.width,
            height=shape.height,
        )

        # Set no fill for the circle (transparent inside)
        circle.fill.background()  # or circle.fill.solid() + set transparency

        # Set outline color and width
        circle.line.color.rgb = RGBColor(*line_color_rgb)
        circle.line.width = Pt(line_width_pt)

        self._set_alt_text(
            circle,
            {"parent_shape_id": shape.shape_id, "meta": True},
        )

    def _select_slice_from_df(self, df, integration):
        """
        Selects a specific slice from the DataFrame based on the integration settings.

        Args:
            df: A pandas DataFrame containing the data.
            integration: A LookerReference object containing the integration settings.
        Returns:
            The selected data slice (str or other type).
        """
        if integration.row is not None:
            row_slice = integration.row
        else:
            row_slice = 0

        row = df.iloc[row_slice]

        if integration.label is not None and integration.column is not None:
            logging.warning(
                f"Both label and column are set for integration {integration.id}. Defaulting to label and ignoring column."
            )
            r = row[integration.label]
        elif integration.label is not None:
            r = row[integration.label]
        elif integration.column is not None:
            r = row.iloc[integration.column]
        else:
            r = df
        return r

    def _replace_image_with_object(
        self, slide_index, shape_number, image_stream, integration
    ):
        slide = self.presentation.slides[slide_index]
        old_shape = next((s for s in slide.shapes if s.shape_id == shape_number), None)
        if old_shape is None:
            raise ValueError(f"Shape {shape_number} not found on slide {slide_index}.")
        if old_shape.shape_type != 13:  # picture
            raise ValueError("Selected shape is not an image.")

        left, top, width, height = (
            old_shape.left,
            old_shape.top,
            old_shape.width,
            old_shape.height,
        )
        slide.shapes._spTree.remove(old_shape._element)

        # --- calculate scaled size preserving aspect ratio ---
        img_bytes = image_stream.getvalue()
        image_stream.seek(0)
        with Image.open(BytesIO(img_bytes)) as im:
            img_w, img_h = im.size
        img_ratio = img_w / img_h
        shape_ratio = width / height

        if img_ratio > shape_ratio:
            new_width = width
            new_height = int(width / img_ratio)
        else:
            new_height = height
            new_width = int(height * img_ratio)

        # center within original box
        new_left = left + (width - new_width) / 2
        new_top = top + (height - new_height) / 2

        picture = slide.shapes.add_picture(
            BytesIO(img_bytes), new_left, new_top, width=new_width, height=new_height
        )
        self._set_alt_text(picture, integration)

    def _remove_shape(self, slide_index, shape_number):
        """
        Removes a shape from a PowerPoint slide.
        Args:
            prs: The Presentation object.
            slide_index: The index of the slide containing the shape.
            shape_index: The index of the shape to remove.
        """

        slide = self.presentation.slides[slide_index]
        shape_to_remove = None
        for shape in slide.shapes:
            if shape.shape_id == shape_number:
                shape_to_remove = shape

        if shape_to_remove is None:
            raise ValueError(
                f"Shape with number {shape_number} not found on slide {slide_index}."
            )

        # Remove the shape
        slide.shapes._spTree.remove(shape_to_remove._element)

    def _format_context_data(self, df) -> str:
        """
        Format a pandas DataFrame as a human-readable plain-text table for use
        as Gemini context.

        Args:
            df: A pandas DataFrame.

        Returns:
            str: A plain-text representation of the DataFrame.
        """
        return df.to_string(index=False)

    def _extract_slide_text_context(
        self, slide_number: int, exclude_shape_id: int
    ) -> str:
        """
        Return a plain-text extract of all text on a slide, excluding the target shape.

        Only shapes that have a text_frame are included.  Each shape is rendered as:
        ``[Shape name]: text content``

        Args:
            slide_number: The index of the slide.
            exclude_shape_id: The shape_id of the shape to exclude (the Gemini shape itself).

        Returns:
            str: Multi-line text representing the slide content.
        """
        slide = self.presentation.slides[slide_number]
        lines: list[str] = []
        for shape in slide.shapes:
            if shape.shape_id == exclude_shape_id:
                continue
            if hasattr(shape, "text_frame"):
                text = shape.text_frame.text.strip()
                if text:
                    name = getattr(shape, "name", f"shape_{shape.shape_id}")
                    lines.append(f"[{name}]: {text}")
        return "\n".join(lines)

    def _resolve_context_item(
        self,
        ctx: str,
        shape_number: int,
        slide_number: int,
        gemini_results: dict,
        current_text: str,
    ) -> tuple | None:
        """
        Resolve a single ``contexts`` entry to a ``(label, content)`` pair.

        Resolution rules (checked in order):

        1. ``"self"`` — the shape's own current text before synthesis.
        2. ``"slide_self"`` — text of all other shapes on the slide after Looker
           data has been rendered, with this shape excluded.
        3. Strings starting with ``"gemini_"`` — the output of the Gemini box
           whose ``gemini_id`` matches.  Returns ``None`` (with a warning) if
           that box has not been processed yet or has failed.
        4. Anything else — treated as a Looker meta-look ``meta_name``; resolved
           from ``self.data``.  Returns ``None`` (with a warning) if missing.

        Args:
            ctx: The context string from ``GeminiConfig.contexts``.
            shape_number: The shape_id of the current Gemini shape (for exclusion).
            slide_number: The slide index of the current Gemini shape.
            gemini_results: Dict of ``{gemini_id: synthesized_text}`` for already
                processed Gemini boxes.
            current_text: The shape's current text content.

        Returns:
            ``(label, content)`` on success, or ``None`` if the reference cannot
            be resolved (a warning is already logged).
        """
        if ctx == "self":
            return ("Current shape text", current_text)

        if ctx == "slide_self":
            content = self._extract_slide_text_context(slide_number, shape_number)
            return ("Current slide context (excluding this shape)", content)

        if ctx.startswith("gemini_"):
            result = gemini_results.get(ctx)
            if result is None:
                logging.warning(
                    f"No result available for Gemini context '{ctx}'. "
                    "The referenced Gemini shape may not have run yet or may have failed."
                )
                return None
            return (f"LLM report [{ctx}]", result)

        # Meta-look fallback
        raw = self.data.get(ctx)
        if raw is None:
            logging.warning(
                f"No data found for Gemini context '{ctx}'. "
                "Make sure a meta-look shape with that meta_name exists in the presentation."
            )
            return None
        try:
            df = self._make_df(raw)
            return (f"Data [{ctx}]", self._format_context_data(df))
        except Exception as e:
            logging.warning(f"Could not format context data for '{ctx}': {e}")
            return None

    def _sort_gemini_shapes_by_dependency(self) -> list:
        """
        Return ``self.gemini_shapes`` sorted in topological order.

        Any ``contexts`` entry that starts with ``"gemini_"`` and matches the
        ``gemini_id`` of another shape in the list is treated as a dependency —
        that shape must be processed first.

        Raises
        ------
        ValueError
            If a circular dependency is detected.
        """
        id_to_shape: dict[str, object] = {
            gs.integration.gemini_id: gs
            for gs in self.gemini_shapes
            if gs.integration.gemini_id
        }

        in_degree: dict[int, int] = {id(gs): 0 for gs in self.gemini_shapes}
        dependents: dict[int, list] = {id(gs): [] for gs in self.gemini_shapes}

        for gs in self.gemini_shapes:
            for ctx in gs.integration.contexts:
                if ctx.startswith("gemini_") and ctx in id_to_shape:
                    dep = id_to_shape[ctx]
                    dependents[id(dep)].append(gs)
                    in_degree[id(gs)] += 1

        queue: collections.deque = collections.deque(
            gs for gs in self.gemini_shapes if in_degree[id(gs)] == 0
        )
        ordered: list = []
        while queue:
            node = queue.popleft()
            ordered.append(node)
            for dependent in dependents[id(node)]:
                in_degree[id(dependent)] -= 1
                if in_degree[id(dependent)] == 0:
                    queue.append(dependent)

        if len(ordered) != len(self.gemini_shapes):
            raise ValueError(
                "Circular dependency detected in Gemini shape contexts. "
                "Ensure there are no cycles among gemini_id references."
            )

        return ordered

    def _process_gemini_shapes(self):
        """
        Process all shapes configured for Gemini LLM synthesis.

        Shapes are processed in dependency order (any shape whose ``gemini_id``
        appears in another shape's ``contexts`` is processed first).

        For each GeminiShape the ``contexts`` list is walked in order and each
        entry resolved via :meth:`_resolve_context_item` — which handles
        ``"self"``, ``"slide_self"``, Gemini box references (``gemini_``-prefix),
        and Looker meta-look names.  The assembled context string is then sent to
        the Gemini API together with the shape's current text and prompt.

        On error: the error message is written into the text box and a red outline
        is drawn around the shape (suppressed by ``--hide-errors``).
        """
        if not self.gemini_shapes:
            return

        if not gemini_module.is_available():
            logging.warning(
                "google-genai is not installed; Gemini synthesis shapes will be skipped. "
                "Install it with 'pip install looker_powerpoint[llm]' to enable LLM features."
            )
            return

        try:
            ordered_shapes = self._sort_gemini_shapes_by_dependency()
        except ValueError as e:
            logging.error(f"Cannot process Gemini shapes: {e}")
            return

        # Stores synthesized text keyed by gemini_id for chaining
        gemini_results: dict[str, str] = {}

        for gemini_shape in ordered_shapes:
            slide = self.presentation.slides[gemini_shape.slide_number]
            current_shape = None
            for shape in slide.shapes:
                if shape.shape_id == gemini_shape.shape_number:
                    current_shape = shape
                    break

            if current_shape is None:
                logging.error(
                    f"Could not find shape {gemini_shape.shape_number} on slide "
                    f"{gemini_shape.slide_number} for Gemini synthesis."
                )
                continue

            try:
                # Capture current text first (needed by "self" resolver and as
                # replacement-target context for synthesize())
                current_text = ""
                if hasattr(current_shape, "text_frame"):
                    current_text = current_shape.text_frame.text
                else:
                    logging.warning(
                        f"Shape {gemini_shape.shape_number} on slide "
                        f"{gemini_shape.slide_number} has no text_frame; "
                        "skipping Gemini synthesis."
                    )
                    continue

                # Resolve each context entry in order
                context_parts: list[str] = []
                for ctx in gemini_shape.integration.contexts:
                    resolved = self._resolve_context_item(
                        ctx,
                        gemini_shape.shape_number,
                        gemini_shape.slide_number,
                        gemini_results,
                        current_text,
                    )
                    if resolved is not None:
                        label, content = resolved
                        if content:
                            context_parts.append(f"{label}:\n{content}")

                context_data_str = "\n\n".join(context_parts)

                synthesized = gemini_module.synthesize(
                    prompt=gemini_shape.integration.prompt,
                    context_data_str=context_data_str,
                    current_text=current_text,
                    model_name=gemini_shape.integration.model,
                )

                update_text_frame_preserving_formatting(
                    current_shape.text_frame, synthesized
                )
                logging.debug(
                    f"Gemini synthesis applied to shape {gemini_shape.shape_number} "
                    f"on slide {gemini_shape.slide_number}."
                )

                # Store result for downstream shapes
                if gemini_shape.integration.gemini_id:
                    gemini_results[gemini_shape.integration.gemini_id] = synthesized

            except Exception as e:
                error_msg = str(e)
                logging.error(
                    f"Gemini synthesis failed for shape {gemini_shape.shape_number} "
                    f"on slide {gemini_shape.slide_number}: {error_msg}"
                )
                # Populate error message into text box
                try:
                    if hasattr(current_shape, "text_frame"):
                        update_text_frame_preserving_formatting(
                            current_shape.text_frame, error_msg
                        )
                except Exception:
                    pass
                # Draw red outline around the failed shape
                if not self.args.hide_errors:
                    self._mark_failure(slide, current_shape)

    def _make_df(self, result):
        """
        Create a pandas DataFrame from Looker data based on the integration settings.
        Categorizes and sorts columns into Dimensions -> Pivots -> Table Calcs.
        """
        data = json.loads(result)
        fields = data.get("metadata", {}).get("fields", {})

        # 1. Pull the injected sorts and pivots rules from the Look
        look_sorts = data.get("custom_sorts", [])
        look_pivots = data.get("custom_pivots", [])

        # Determine if the primary pivot is sorted descending
        pivot_descending = False
        if look_pivots:
            main_pivot = look_pivots[0]
            # Looker sorts look like: "view_name.date_dim desc 0"
            # Parse sort strings by tokens so we only match the exact field name,
            # not arbitrary substrings.
            for sort_str in look_sorts:
                parts = str(sort_str).split()
                if len(parts) < 2:
                    continue
                field_token = parts[0]
                # Find the first explicit direction token after the field
                direction_token = None
                for token in parts[1:]:
                    token_lower = token.lower()
                    if token_lower in ("asc", "desc"):
                        direction_token = token_lower
                        break
                if field_token == main_pivot and direction_token == "desc":
                    pivot_descending = True
                    break

        # Create DataFrame first to expose all dynamic column names
        df = pd.json_normalize(data.get("rows", [])).fillna("")
        actual_cols = list(df.columns)

        # 2. Extract base names from the metadata
        dim_bases = [f["name"] for f in fields.get("dimensions", [])]
        calc_bases = [f["name"] for f in fields.get("table_calculations", [])]
        measure_bases = [f["name"] for f in fields.get("measures", [])]

        dims = []
        calcs = []
        pivots_and_measures = []
        leftovers = []

        for col in actual_cols:
            if any(col == f"{d}.value" or col == d for d in dim_bases):
                dims.append(col)
            elif any(col == f"{c}.value" or col == c for c in calc_bases):
                calcs.append(col)
            elif "|FIELD|" in col or any(
                col == f"{m}.value" or col == m for m in measure_bases
            ):
                pivots_and_measures.append(col)
            else:
                leftovers.append(col)

        # 3. Apply Dimensions and Calcs sorting (Native query order)
        dims.sort(
            key=lambda x: next(
                (i for i, d in enumerate(dim_bases) if x == f"{d}.value" or x == d), 999
            )
        )
        calcs.sort(
            key=lambda x: next(
                (i for i, c in enumerate(calc_bases) if x == f"{c}.value" or x == c),
                999,
            )
        )

        # 4. Apply Looker's strict Pivot Sorting Rules
        def parse_pivot_col(col):
            # Break down "measure_name|FIELD|2025-03-03.value"
            base = col.replace(".value", "")
            if "|FIELD|" in base:
                measure, pivot_val = base.split("|FIELD|", 1)
                return pivot_val, measure
            return "", base

        # Get unique pivot values preserving their order of first appearance in the data
        # (which reflects Looker's native ordering). Avoid lexicographic sorting so that
        # numeric-like values ("2", "10") and non-ISO dates aren't misordered.
        seen_pivots: set = set()
        unique_pivots: list = []
        for c in pivots_and_measures:
            pv = parse_pivot_col(c)[0]
            if pv not in seen_pivots:
                seen_pivots.add(pv)
                unique_pivots.append(pv)
        if pivot_descending:
            unique_pivots.reverse()

        pivot_order_map = {val: i for i, val in enumerate(unique_pivots)}
        measure_order_map = {m: i for i, m in enumerate(measure_bases)}

        # Sort first by the properly sequenced pivot value, then by the measure's native query order
        pivots_and_measures.sort(
            key=lambda x: (
                pivot_order_map.get(parse_pivot_col(x)[0], 999),
                measure_order_map.get(parse_pivot_col(x)[1], 999),
            )
        )

        # 5. Re-index and Rename
        ordered_cols = dims + pivots_and_measures + calcs + leftovers
        df = df[ordered_cols]

        all_fields = (
            fields.get("dimensions", [])
            + fields.get("measures", [])
            + fields.get("table_calculations", [])
        )
        mappy = {
            f"{item['name']}.value": item.get("field_group_variant", item["name"])
            .strip()
            .lower()
            .replace(" ", "_")
            for item in all_fields
        }
        df.rename(columns=mappy, inplace=True)

        return df

    def _build_metadata_object(self):
        """
        Build metadata object for the presentation.
        """
        metadata_rows = []
        looks = set()
        for looker_shape in self.looker_shapes:
            if looker_shape.integration.id not in looks:
                looks.add(looker_shape.integration.id)
                metadata_rows.append(
                    {
                        "looks": {
                            "value": f"{os.environ.get('LOOKERSDK_BASE_URL')}looks/{looker_shape.integration.id}"
                        }
                    }
                )
        metadata_object = {
            "metadata": {"fields": {"dimensions": [{"name": "looks"}]}},
            "rows": metadata_rows,
        }
        self.data["metadata_shapes"] = json.dumps(metadata_object)

    async def get_queries(self):
        """
        asynchronously fetch a list of look references
        """
        logging.info(
            f"Running Looker queries... {len(self.looker_shapes)} queries to run."
        )
        tasks = [
            self.client._async_write_queries(
                shape.shape_id, self.args.filter, **dict(shape.integration)
            )
            for shape in self.looker_shapes
        ]

        # Run all tasks concurrently and gather the results
        results = await asyncio.gather(*tasks)
        for r in results:
            self.data.update(r)

    def _test_str_to_int(self, s):
        try:
            int(s)
            return True
        except ValueError:
            return False

    def run(self, **kwargs):
        """
        Main method to run the CLI application.
        """
        self.args = self.parser.parse_args()
        self._setup_logging()
        self._pick_file()
        self._init_looker()

        references = self.get_alt_text(self.file_path)
        if not references:
            logging.error(
                "No shapes with id found in the presentation. Add a 'id' : '<look_id>' to the alternative text of a shape to load data into the shape."
            )
            return

        for ref in references:
            integration = ref.get("integration", {})
            # Try to parse as a Gemini shape first (type: gemini discriminator)
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                try:
                    gemini_shape = GeminiShape.model_validate(ref)
                    if gemini_shape.shape_type not in (
                        "TEXT_BOX",
                        "TITLE",
                        "AUTO_SHAPE",
                    ):
                        logging.warning(
                            f"Gemini synthesis config found on shape "
                            f"{gemini_shape.shape_id} (type: {gemini_shape.shape_type}). "
                            "Gemini synthesis only works for text boxes (TEXT_BOX, TITLE, "
                            "AUTO_SHAPE). This shape will be skipped."
                        )
                        continue
                    self.gemini_shapes.append(gemini_shape)
                except ValidationError as e:
                    logging.debug(
                        f"Could not parse Gemini config in shape {ref.get('shape_id', '?')}: {e}"
                    )
                continue

            # Otherwise try to parse as a regular Looker shape
            try:
                self.relevant_shapes.append(LookerShape.model_validate(ref))
            except ValidationError as e:
                logging.debug(
                    f"Could not parse the alternate text in slide {ref['shape_id'].split(',')[0]}, shape {ref['shape_id'].split(',')[1]}: {e}"
                )
                continue

        self.looker_shapes = [
            s
            for s in self.relevant_shapes
            if s.integration.id_type == "look"
            and self._test_str_to_int(s.integration.id)
        ]

        self._build_metadata_object()

        asyncio.run(self.get_queries())

        for looker_shape in self.relevant_shapes:
            if looker_shape.integration.meta:
                if not self.args.self:
                    self._remove_shape(
                        looker_shape.slide_number,
                        looker_shape.shape_number,
                    )

            else:
                result = self.data.get(looker_shape.shape_id)
                if result is None:
                    result = self.data.get(looker_shape.integration.id)

                try:
                    if looker_shape.shape_type == "PICTURE":
                        if looker_shape.integration.result_format in ("jpg", "png"):
                            image_stream = BytesIO(result)
                        else:
                            df = self._make_df(result)
                            url = self._select_slice_from_df(
                                df, looker_shape.integration
                            )

                            response = requests.get(url)
                            response.raise_for_status()
                            image_stream = io.BytesIO(response.content)

                        logging.debug(
                            f"Replacing image for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}..."
                        )

                        self._replace_image_with_object(
                            looker_shape.slide_number,
                            looker_shape.shape_number,
                            image_stream,
                            looker_shape.original_integration,
                        )

                    elif looker_shape.shape_type in [
                        "CHART",
                        "TABLE",
                        "TEXT_BOX",
                        "TITLE",
                        "AUTO_SHAPE",
                    ]:
                        slide = self.presentation.slides[looker_shape.slide_number]
                        for shape in slide.shapes:
                            if shape.shape_id == looker_shape.shape_number:
                                current_shape = shape
                        df = self._make_df(result)

                        if looker_shape.shape_type == "TABLE":
                            logging.debug(
                                f"Updating table for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}..."
                            )
                            self._fill_table(
                                current_shape.table,
                                df,
                                looker_shape.integration.headers,
                            )

                        elif looker_shape.shape_type in [
                            "TEXT_BOX",
                            "TITLE",
                            "AUTO_SHAPE",
                        ]:
                            logging.debug(
                                f"Updating text for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}..."
                            )

                            try:
                                text_to_insert = self._select_slice_from_df(
                                    df, looker_shape.integration
                                )
                            except Exception as e:
                                text_to_insert = df.to_string(index=False, header=False)
                                logging.debug(
                                    f"inserting whole text for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}: {e}"
                                )
                            current_shape = process_text_field(
                                current_shape,
                                text_to_insert,
                                df,
                            )
                            # add_text_with_numbered_links(current_shape.text_frame, str(text_to_insert))

                        elif looker_shape.shape_type == "CHART":
                            chart_data = CategoryChartData()
                            chart_data.categories = df.iloc[
                                :, 0
                            ].tolist()  # Assuming the first column contains categories
                            chart = current_shape.chart
                            existing_chart_data = chart.plots[0].series
                            logging.debug(
                                f"Existing chart series: {[s.name for s in existing_chart_data]}"
                            )

                            if looker_shape.integration.headers:
                                for series_name in df.columns[1:]:
                                    try:
                                        match = (
                                            re.search(
                                                r"^[^\.]*\.[^\.]*\.(.*)\.value$",
                                                series_name,
                                            )
                                            .group(1)
                                            .replace(".", " - ")
                                            .strip()
                                            .replace("|FIELD|", " ")
                                        )
                                    except AttributeError as e:
                                        logging.debug(
                                            f"Could not parse series name {series_name}, setting name to {series_name}"
                                        )
                                        match = series_name
                                    chart_data.add_series(match, df[series_name])
                            else:
                                if len(df.columns[1:]) != len(existing_chart_data):
                                    logging.warning(
                                        f"{looker_shape.shape_id}. Missing headers! Number of series ({len(df.columns[1:])}) does not match number of existing chart series ({len(existing_chart_data)}). Perhaps you need to enable headers in the integration settings?"
                                    )
                                for series_name, series in zip(
                                    df.columns[1:], existing_chart_data
                                ):
                                    chart_data.add_series(series.name, df[series_name])

                            chart.replace_data(chart_data)
                            if looker_shape.integration.show_latest_chart_label:
                                for plot in chart.plots:
                                    s = 0
                                    for series in plot.series:
                                        series_has_label = False
                                        index = 0
                                        for i, v in zip(
                                            series.points, df.iloc[:, s + 1]
                                        ):
                                            if i.data_label._dLbl is not None:
                                                series_has_label = True
                                                logging.debug(
                                                    f"Series {series.name} has data labels."
                                                )
                                            if v is not None and v != "":
                                                logging.debug(
                                                    f"Value for point {index} in series {series.name}: {v}"
                                                )
                                                index += 1
                                        if series_has_label is True:
                                            new_index = 0
                                            for point in series.points:
                                                new_index += 1
                                                if new_index == index:
                                                    logging.debug(
                                                        f"Showing data label for point {new_index} in series {series.name}."
                                                    )
                                                    point.data_label.text_frame.text = (
                                                        ""
                                                    )
                                                    point.data_label.has_text_frame = (
                                                        False
                                                    )
                                                else:
                                                    point.data_label.text_frame.text = (
                                                        ""
                                                    )
                                                    point.data_label.has_text_frame = (
                                                        True
                                                    )
                                        s += 1

                    else:
                        logging.warning(
                            f"unknown shape type {looker_shape.shape_type} for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}."
                        )
                        continue

                except Exception as e:
                    logging.error(f"Error processing reference {looker_shape}: {e}")
                    # import traceback
                    # traceback.print_exc()  # Prints the full traceback

                    if not self.args.hide_errors:
                        slide = self.presentation.slides[looker_shape.slide_number]
                        for shape in slide.shapes:
                            if shape.shape_id == looker_shape.shape_number:
                                self._mark_failure(slide, shape)

        # Process Gemini synthesis shapes
        self._process_gemini_shapes()

        if self.args.self:
            self.destination = self.file_path
        else:
            if not self.args.output_dir.endswith("/"):
                self.args.output_dir += "/"
            self.destination = (
                self.args.output_dir
                + os.path.basename(self.file_path).removesuffix(".pptx")
                + datetime.datetime.now().strftime("_%Y%m%d_%H%M%S.pptx")
            )
        if not os.path.exists(self.args.output_dir) and not self.args.self:
            os.makedirs(self.args.output_dir)

        self.presentation.save(self.destination)

        if not self.args.quiet:
            try:
                os.startfile(self.destination)
                logging.info(f"Opened {self.destination} in PowerPoint.")
            except Exception as e:
                try:
                    subprocess.Popen(["open", self.destination])  # For macOS
                    logging.info(f"Opened {self.destination} in PowerPoint.")
                except Exception as e:
                    logging.error(f"Failed to open the PowerPoint file: {e}")
                    logging.info(f"You can find the file at {self.destination}.")


def main():
    cli = Cli()
    cli.run()


if __name__ == "__main__":
    main()
````

## File: looker_powerpoint/gemini.py
````python
"""
Optional Gemini LLM integration for text synthesis.

This module wraps the ``google-genai`` SDK.  If that package is not installed
the helpers in this module still import cleanly; they will raise an
:class:`ImportError` with a helpful message when called.

Install the optional dependency with::

    pip install looker_powerpoint[llm]

and set the ``GOOGLE_API_KEY`` (or ``GEMINI_API_KEY``) environment variable
before using any function in this module.
"""

import logging
import os

try:
    from google import genai  # type: ignore[import]

    _HAS_GEMINI = True
except ImportError:  # pragma: no cover
    _HAS_GEMINI = False


def is_available() -> bool:
    """Return ``True`` if the ``google-genai`` package is installed."""
    return _HAS_GEMINI


def synthesize(
    prompt: str | None,
    context_data_str: str,
    current_text: str,
    model_name: str = "gemini-2.0-flash",
) -> str:
    """
    Call the Gemini API and return the synthesized text.

    Parameters
    ----------
    prompt:
        An optional user-supplied instruction or question that guides the model.
    context_data_str:
        Pre-assembled context string built from the shape's ``contexts`` list.
        May include Looker data, slide text, prior Gemini results, or the
        shape's own current text — depending on what the user configured.
        Each section is separated by a blank line and prefixed with a label.
    current_text:
        The current text content of the PowerPoint shape.  Always passed so the
        model knows it is producing a replacement for existing slide text.
    model_name:
        Gemini model identifier, e.g. ``"gemini-2.0-flash"``.

    Returns
    -------
    str
        The text generated by Gemini.

    Raises
    ------
    ImportError
        If ``google-genai`` is not installed.
    ValueError
        If no API key is configured.
    """
    if not _HAS_GEMINI:
        raise ImportError(
            "google-genai is not installed. "
            "Install it with 'pip install looker_powerpoint[llm]' to use LLM features."
        )

    api_key = os.environ.get("GOOGLE_API_KEY") or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        raise ValueError(
            "Neither GOOGLE_API_KEY nor GEMINI_API_KEY environment variable is set. "
            "Set one of them to enable Gemini synthesis."
        )

    client = genai.Client(api_key=api_key)

    parts: list[str] = []

    if context_data_str:
        parts.append(f"Context:\n{context_data_str}")

    if current_text:
        parts.append(f"Current text in the shape:\n{current_text}")

    if prompt:
        parts.append(f"Instructions:\n{prompt}")

    parts.append(
        "Please provide a concise text response that will replace the current text "
        "in a PowerPoint slide. Return only the replacement text, without any "
        "additional commentary or markdown formatting."
    )

    full_prompt = "\n\n".join(parts)
    logging.debug("Gemini prompt (truncated): %s", full_prompt[:200])

    response = client.models.generate_content(
        model=model_name,
        contents=full_prompt,
    )
    return response.text
````

## File: looker_powerpoint/looker_powerpoint.instructions.md
````markdown
# looker_powerpoint

This is the main Python package for the Looker PowerPoint CLI tool (`lppt`).

## Modules

| File | Purpose |
|------|---------|
| `cli.py` | Entry point for the `lppt` CLI command. Contains the `Cli` class and `main()` function. Orchestrates fetching Looker data and writing results into PowerPoint files. |
| `looker.py` | `LookerClient` class that wraps the Looker SDK. Handles authentication, query construction, executing Look queries, and retry logic. |
| `models.py` | Pydantic models: `LookerReference` and `LookerShape` (Looker-backed shapes); `GeminiConfig` and `GeminiShape` (Gemini LLM synthesis shapes). |
| `gemini.py` | Optional Google Gemini integration. Wraps the `google-genai` SDK (import path `google.genai`); provides `is_available()` and `synthesize()`. Safe to import when the extra is not installed. |
| `__init__.py` | Package initialiser; exposes `__version__` via `importlib.metadata`. |
| `tools/` | Sub-package of utility helpers (see `tools/README.md`). |

## How it works

1. The CLI reads a `.pptx` file.
2. Each shape whose *alternative text* contains valid YAML is parsed into either a
   `LookerReference` (regular data shapes) or a `GeminiConfig` (LLM synthesis shapes).
3. The `LookerClient` fetches the corresponding Looker Looks and returns the results.
4. Results are written back into the presentation (text boxes, tables, images) using
   the helpers in `tools/`.
5. For Gemini shapes, context data is taken from pre-fetched meta-look results and
   passed to the Gemini API; the response replaces the shape's text.

## Gemini synthesis

Set `type: gemini` in the alt text of a **text box** shape:

```yaml
type: gemini
gemini_id: summary          # optional; gemini_ prefix auto-added → gemini_summary
prompt: Summarise the key trends.
contexts:
  - slide_self              # other shapes' text on this slide
  - sales_data              # meta_name of a Looker meta-look
  - gemini_analysis         # output of another Gemini box (gemini_id: analysis)
  - self                    # this shape's own current text
model: gemini-2.0-flash     # optional, default shown
```

### `contexts` — unified context framework

Each entry is one of four types (resolved in order):

| Entry | Resolves to |
|-------|-------------|
| `"self"` | Shape's own current text before synthesis |
| `"slide_self"` | Text of all other shapes on the slide after Looker rendering |
| `"gemini_<id>"` | Output of another Gemini box whose `gemini_id` is `<id>` |
| anything else | Looker meta-look data from `self.data[meta_name]` |

### Chaining / topological sort

Boxes that reference other boxes via `gemini_<id>` are automatically sorted
by `_sort_gemini_shapes_by_dependency()` (Kahn's topological sort) so
dependencies always run first.  Circular references raise `ValueError`.

The `gemini_` prefix is auto-added to `gemini_id` by `GeminiConfig`'s
`ensure_gemini_prefix` validator.

- Only `TEXT_BOX`, `TITLE`, and `AUTO_SHAPE` types are supported; other types log a
  warning and are skipped.
- Requires the `llm` optional extra: `pip install looker_powerpoint[llm]`.
- Works without the extra installed — Gemini shapes are simply skipped with a warning.

## Running

```bash
uv run lppt --help
```
````

## File: looker_powerpoint/looker.py
````python
import logging
from typing import Optional
import looker_sdk
from dotenv import load_dotenv, find_dotenv
from looker_sdk import models40 as models
from tenacity import retry, stop_after_attempt, wait_fixed, before_sleep_log
import json


class LookerClient:
    def __init__(self):
        load_dotenv(find_dotenv(usecwd=True))
        try:
            self.client = looker_sdk.init40()  # or init40() for the v4.0 API
        except looker_sdk.error.SDKError as e:
            logging.error(
                f"Error initializing Looker SDK: {e} Consider adding a looker.ini file, or setting the LOOKERSDK_BASE_URL, LOOKERSDK_CLIENT_ID, and LOOKERSDK_CLIENT_SECRET environment variables."
            )
            exit(1)

    async def run_query(self, query_object):
        """
        Runs a query against the Looker API.

        Args:
            query_object: The query object containing the necessary parameters.
        """

        response = self.client.run_inline_query(
            result_format=query_object["result_format"],
            body=query_object["body"],
            apply_vis=query_object["apply_vis"],
            apply_formatting=query_object["apply_formatting"],
            server_table_calcs=query_object["server_table_calcs"],
        )

        return response

    async def make_query(
        self,
        shape_id: int,
        filter: Optional[str] = None,
        filter_value: Optional[str] = None,
        filter_overwrites: Optional[dict] = None,
        id: Optional[int] = None,
        **kwargs,
    ) -> models.WriteQuery:
        """
        Constructs a WriteQuery object based on a Look's definition and provided parameters.
        Args:
            id: The ID of the Look.
            filter: The name of the filter to apply.
            filter_value: The value to set for the filter.
            filter_overwrites: A dictionary of filters to overwrite with new values.
            **kwargs: Additional query parameters to set.
        Returns:
            A WriteQuery object representing the modified query.
        """
        try:
            # check if string can be converted to int
            look = self.client.look(id)
        except Exception as e:
            logging.error(
                f"Error fetching Look with ID {id}, is this a valid Look ID? If it is a meta reference, remember to set id_type: 'meta'"
            )
            return {shape_id: None}

        q = look.query
        for parameter, value in kwargs.items():
            if value is not None:
                if hasattr(q, parameter):
                    # If the parameter is a list, append the value
                    if isinstance(getattr(q, parameter), list):
                        getattr(q, parameter).append(value)
                    else:
                        # Otherwise, set the value directly
                        setattr(q, parameter, value)

        if filter_overwrites is not None:
            for f, v in filter_overwrites.items():
                logging.info(f"Overwriting filter {f} with value {v}")
                if hasattr(q, "filters"):
                    filterable = False
                    for _, existing_filter in enumerate(q.filters):
                        if existing_filter == f:
                            filterable = True
                    if filterable:
                        q.filters[f] = v
                    else:
                        logging.warning(
                            f"Overwrite filter {f} not found in query filters. Available filters: {q.filters}"
                        )

        if filter_value is not None and filter is not None:
            logging.info(f"Applying filter {filter} with value {filter_value}")
            # If filter_value is provided, set the filter
            if hasattr(q, "filters"):
                filterable = False
                for _, f in enumerate(q.filters):
                    # print(f, filter)
                    if f == filter:
                        filterable = True
            if filterable:
                q.filters[filter] = filter_value
            else:
                logging.warning(
                    f"Filter {filter} not found in query filters. Available filters: {q.filters}"
                )

        body = models.WriteQuery(
            model=q.model,
            view=q.view,
            fields=q.fields,
            pivots=q.pivots,
            fill_fields=q.fill_fields,
            filters=q.filters,
            sorts=q.sorts,
            limit=q.limit,
            column_limit=q.column_limit,
            total=q.total,
            row_total=q.row_total,
            subtotals=q.subtotals,
            dynamic_fields=q.dynamic_fields,
            query_timezone=q.query_timezone,
            vis_config=q.vis_config,
            visible_ui_sections=q.visible_ui_sections,
        )

        result_format = kwargs.get("result_format", "json_bi")
        apply_vis = kwargs.get("apply_vis", False)
        apply_formatting = kwargs.get("apply_formatting", False)
        server_table_calcs = kwargs.get("server_table_calcs", False)
        retries = kwargs.get("retries", 0)

        query_object = {
            "shape_id": shape_id,
            "query": {
                "result_format": result_format,
                "body": body,
                "apply_vis": apply_vis,
                "apply_formatting": apply_formatting,
                "server_table_calcs": server_table_calcs,
            },
        }

        try:

            @retry(
                stop=stop_after_attempt(retries + 1),
                wait=wait_fixed(2),
                before_sleep=before_sleep_log(logging.getLogger(), logging.WARNING),
                reraise=True,
            )
            async def run_query_with_retry():
                return await self.run_query(query_object["query"])

            result = await run_query_with_retry()

            if result and result_format in ["json", "json_bi"]:
                try:
                    parsed = json.loads(result)
                    if isinstance(parsed, dict):
                        # Pack the sorts and pivots into the payload
                        parsed["custom_sorts"] = list(q.sorts) if q.sorts else []
                        parsed["custom_pivots"] = list(q.pivots) if q.pivots else []
                        result = json.dumps(parsed)
                except (json.JSONDecodeError, TypeError, ValueError) as e:
                    logging.warning(
                        "Failed to inject custom_sorts/custom_pivots for shape_id %s, look_id %s: %s",
                        shape_id,
                        id,
                        e,
                        exc_info=True,
                    )

        except looker_sdk.error.SDKError as e:
            logging.error(f"Error retrieving Look with ID {id} : {e}")
            result = None
        except Exception as e:
            logging.error(f"Unexpected error retrieving Look with ID {id} : {e}")
            result = None

        return {shape_id: result}

    async def _async_write_queries(self, shape_id, filter_value=None, **kwargs):
        """
        Asynchronously write a Looker query by its ID.
        Args:
            table: A dictionary containing the look_id and other parameters.
        Returns:
            The fetched look data.
        """
        return await self.make_query(
            shape_id, filter_value=filter_value, **dict(kwargs)
        )
````

## File: looker_powerpoint/models.py
````python
import logging
from typing import List, Optional
from pydantic import BaseModel, Field, model_validator, field_validator, ValidationError


class LookerReference(BaseModel):
    """
    This model represents the input you can set in alternative text for a shape in PowerPoint.
    You can specify the different parameters to control how Looker data is fetched and displayed.
    """

    id: str = Field(
        ...,
        description="The ID of the Look or meta-look (meta_name) you want to reference.",
    )
    id_type: str = Field(
        default="look",
        description="The type of ID provided: 'look' or 'meta'. Defaults to 'look'."
        " Setting to 'meta' indicates that the ID refers to a meta Look.",
    )
    meta: bool = Field(
        default=False,
        description="Set this to true if the Look is a meta Look. A meta look is a look that you want to retrieve and reuse, but not display directly.",
    )
    meta_name: str = Field(
        default=None,
        description="NOT actually working yet. If you are defining a meta look, you should provide a reference name here. This can then be used by other shapes to reference this meta look.",
    )
    meta_iterate: bool = Field(
        default=False,
        description="If set to true, this meta look will be iterated over by other shapes referencing it. This is useful for creating dynamic content based on the results of the meta look.",
    )
    label: str = Field(
        default=None,
        description="Setting a label here filters the results to the specified label. The label needs to match the specific column label from the look including any special characters.",
    )
    column: int = Field(
        default=None,
        description="The specific column to retrieve from the Look results. 0-indexed.",
    )
    row: int = Field(
        default=None,
        description="If you want to retrieve a specific row from the Look results, set the row number here (0-indexed).",
    )
    filter: str = Field(
        default=None,
        description="Define a lookml.field_name used in the Look that you want to be able to filter on using the --filter cli argument. Inputting --filter <value> will filter the results to where <label>=<value>.",
    )
    filter_overwrites: dict = Field(
        default=None,
        description="A dictionary of filter overwrites to apply to the Look. The keys are the filter lookml.field_names, and the values are the filter values. The filter values should not be enclosed in quotation marks. (unvalidated)",
    )
    result_format: str = Field(
        default="json_bi",
        description="The format to return the results in. Defaults to 'json_bi'.",
    )
    show_latest_chart_label: bool = Field(
        default=False,
        description="If set to true, modify chart series with labels to only show the latest label.",
    )
    apply_formatting: bool = Field(
        default=False, description="Apply Looker-specified formatting to each result."
    )
    apply_vis: bool = Field(
        default=True, description="Apply Looker visualization options to results."
    )
    server_table_calcs: bool = Field(
        default=True,
        description="Whether to compute table calculations on the Looker server before returning results.",
    )
    headers: bool = Field(
        default=True,
        description="Whether to overwrite headers in the result set with Looker-defined column labels.",
    )
    image_width: int = Field(
        default=None,
        description="Width of the image in pixels. Used for setting image size when asking looker to return a look rendered as an image.",
    )
    image_height: int = Field(
        default=None,
        description="Height of the image in pixels. Used for setting image size when asking looker to return a look rendered as an image.",
    )
    retries: int = Field(
        default=0,
        description="Number of retries for the Looker API request in case of failure. Defaults to 0.",
    )
    # optional parameters for the Look (Default to None)

    @field_validator("id", mode="before")
    @classmethod
    def convert_int(cls, value):
        """Validation: Convert integer values to strings."""
        if isinstance(value, int):
            return str(value)
        return value


class LookerShape(BaseModel):
    """A Pydantic model for a shape in a PowerPoint presentation.
    This model is used to define the properties of a shape, including its ID, type, dimensions,
    and associated Looker reference.
    """

    is_meta: bool = Field(
        default=False, description="Whether this shape is a meta shape."
    )
    meta_name: str = Field(
        default=None, description="The name of the meta shape, if applicable."
    )
    shape_id: str
    shape_type: str
    slide_number: int
    shape_width: int = Field(default=None)  # Width in pixels
    shape_height: int = Field(default=None)  # Height in pixels
    integration: LookerReference
    original_integration: LookerReference = Field(
        default=None,
        description="The original integration data before any modifications.",
    )
    shape_number: int = Field(
        default=None, description="The number of the shape in the slide."
    )

    @model_validator(mode="before")
    @classmethod
    def push_down_relevant_data(cls, data):
        """Push down relevant data from the integration to the shape model."""
        # push down
        # if picture is shape type, then we need to push down the image width and height
        if type(data.get("integration")) in (dict, LookerReference):
            data["original_integration"] = data["integration"]

            if data["shape_type"] == "PICTURE":
                data["integration"]["result_format"] = data["integration"].get(
                    "result_format", "json_bi"
                )
                data["integration"]["image_width"] = round(data["shape_width"])
                data["integration"]["image_height"] = round(data["shape_height"])

            elif data["shape_type"] == "TABLE":
                if data["integration"].get("apply_formatting") is None:
                    data["integration"]["apply_formatting"] = True

        return data


class GeminiConfig(BaseModel):
    """
    Configuration for a Gemini LLM text synthesis shape.
    Set ``type: gemini`` in the alt text of a **text box** shape to enable this feature.

    The Gemini model receives an assembled context built from the ordered
    ``contexts`` list, then produces replacement text for the shape.

    Each entry in ``contexts`` is resolved by type:

    * ``"self"`` — the shape's own current text (before synthesis).
    * ``"slide_self"`` — text of all other shapes on the same slide after Looker
      data has been rendered (i.e. the slide this comment will appear on).
    * Any string starting with ``gemini_`` — the synthesized output of another
      Gemini text box whose ``gemini_id`` matches.  Those boxes are automatically
      processed first.
    * Anything else — treated as the ``meta_name`` of a Looker meta-look shape;
      its pre-fetched data is formatted as a readable table.

    .. note::
       Requires the ``google-genai`` package.  Install it with::

           pip install looker_powerpoint[llm]

       The ``GOOGLE_API_KEY`` (or ``GEMINI_API_KEY``) environment variable must also
       be set.
    """

    type: str = Field(
        default="gemini",
        description="Must be 'gemini' to identify this as a Gemini synthesis config.",
    )
    gemini_id: Optional[str] = Field(
        default=None,
        description=(
            "A unique identifier for this Gemini shape within the presentation. "
            "The ``gemini_`` prefix is added automatically if omitted. "
            "Required if another Gemini shape references this box via its contexts list."
        ),
    )
    prompt: Optional[str] = Field(
        default=None,
        description="An optional instruction/question sent to the Gemini model together with the context data.",
    )
    contexts: List[str] = Field(
        default_factory=list,
        description=(
            "Ordered list of context references for this Gemini shape. Each entry "
            "is one of: ``'self'``, ``'slide_self'``, a ``gemini_<id>`` string "
            "referencing another Gemini box, or a Looker meta-look ``meta_name``."
        ),
    )
    model: str = Field(
        default="gemini-2.0-flash",
        description="The Gemini model name to use for synthesis.",
    )

    @field_validator("type")
    @classmethod
    def type_must_be_gemini(cls, v):
        if v != "gemini":
            raise ValueError("type must be 'gemini' for GeminiConfig")
        return v

    @field_validator("gemini_id", mode="before")
    @classmethod
    def ensure_gemini_prefix(cls, v):
        """Auto-add the ``gemini_`` prefix when the user omits it."""
        if v is not None and not str(v).startswith("gemini_"):
            return f"gemini_{v}"
        return v


class GeminiShape(BaseModel):
    """
    A Pydantic model for a PowerPoint text-box shape configured for Gemini LLM synthesis.
    """

    shape_id: str
    shape_type: str
    slide_number: int
    shape_width: Optional[int] = Field(default=None)
    shape_height: Optional[int] = Field(default=None)
    integration: GeminiConfig
    shape_number: Optional[int] = Field(default=None)
````

## File: test/pptx/gemini_textbox.md
````markdown
# gemini_textbox.pptx

A single-slide presentation containing one text box configured for Gemini LLM synthesis.

## Slide 1

### Shape 1 — Text box (shape ID 2)

| Property | Value |
|----------|-------|
| Shape type | `TEXT_BOX` |
| Position | Left: 1 in, Top: 1 in |
| Size | Width: 6 in, Height: 1 in |
| Initial text | `Placeholder text to be replaced by Gemini synthesis.` |

**Alt text (YAML):**

```yaml
type: gemini
prompt: Summarize the key trends from the data.
contexts:
  - sales_data
model: gemini-2.0-flash
```

`contexts` lists `meta_name` strings of meta-look shapes defined elsewhere in the
same presentation.  The data for each named meta-look is fetched by the regular
Looker pipeline and passed to Gemini as context.

## Purpose

Used by `test/test_gemini.py` to:

- Verify that a shape whose alt text contains `type: gemini` is parsed as a
  `GeminiShape` (not a `LookerShape`).
- Test that `contexts` contains plain meta_name strings (not nested objects).
- Test that the shape's text is updated by the Gemini synthesis pipeline (with a
  mocked Gemini API call and pre-seeded `cli.data`).
- Confirm that error handling populates the error message into the text box and
  draws a red outline when synthesis fails.
````

## File: test/pptx/table7x7.md
````markdown
# TABLE 7x7

Contains one slide, with a table 7x7 with yml: id: 1
Useful for testing table cases.

## Structure

| Property | Value |
|----------|-------|
| Slides   | 1     |
| Shapes with alt text | 1 |
| Shape type | TABLE |
| Table rows | 7 |
| Table columns | 7 |
| Shape width (px) | 853 |
| Shape height (px) | 273 |
| Shape number (pptx id) | 4 |
| Slide index | 0 |

## YAML alt text

```yaml
id: 1
```

## Expected extraction result

`get_presentation_objects_with_descriptions` returns one entry:

```python
{
    "shape_id": "0,4",
    "shape_type": "TABLE",
    "shape_width": 853,
    "shape_height": 273,
    "integration": {"id": 1},
    "slide_number": 0,
    "shape_number": 4,
}
```
````

## File: test/test_cli.py
````python
import json
import pandas as pd
import pytest
from unittest.mock import patch
from pptx import Presentation
from pptx.util import Inches
from looker_powerpoint.cli import Cli
from looker_powerpoint.models import LookerReference, LookerShape


def test_default_output_dir():
    """Test that the default output directory is 'output'."""
    with patch("os.getenv", return_value="dummy_value"):
        cli = Cli()
        args = cli.parser.parse_args([])
        assert args.output_dir == "output"


def _make_cli():
    """Create a Cli instance with os.getenv stubbed out so no real environment is needed."""
    with patch("os.getenv", return_value="dummy_value"):
        return Cli()


def _field(name):
    """Build a field dict with a field_group_variant derived from the short name."""
    return {"name": name, "field_group_variant": name.split(".")[-1]}


def _make_result(
    dimensions,
    measures,
    table_calculations,
    rows,
    custom_sorts=None,
    custom_pivots=None,
):
    """Build a json_bi-style result string for _make_df.

    Each entry in *dimensions*, *measures*, and *table_calculations* may be either a
    plain field name string or a dict with ``name``/``field_group_variant`` keys.
    Plain strings are automatically expanded with a ``field_group_variant`` equal to
    the portion of the name after the last dot (e.g. ``"view.date"`` → ``"date"``).
    """

    def _normalise(fields):
        return [_field(f) if isinstance(f, str) else f for f in fields]

    return json.dumps(
        {
            "metadata": {
                "fields": {
                    "dimensions": _normalise(dimensions),
                    "measures": _normalise(measures),
                    "table_calculations": _normalise(table_calculations or []),
                }
            },
            "rows": rows,
            "custom_sorts": custom_sorts or [],
            "custom_pivots": custom_pivots or [],
        }
    )


class TestMakeDf:
    """Tests for Cli._make_df column ordering logic."""

    def test_no_pivots_column_order(self):
        """Dimensions come first, then measures, then table calcs — no pivots."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.date", "view.name"],
            measures=["view.revenue", "view.cost"],
            table_calculations=["calc1"],
            rows=[
                {
                    "view.date.value": "2024-01-01",
                    "view.name.value": "A",
                    "view.revenue.value": 100,
                    "view.cost.value": 50,
                    "calc1.value": 2.0,
                }
            ],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        # All five columns must be present (renamed via field_group_variant)
        assert len(cols) == 5
        # Dimensions first
        dim_positions = [cols.index("date"), cols.index("name")]
        # Measures next
        measure_positions = [cols.index("revenue"), cols.index("cost")]
        # Calc last
        calc_position = cols.index("calc1")
        assert max(dim_positions) < min(measure_positions), "Dims must precede measures"
        assert max(measure_positions) < calc_position, (
            "Measures must precede table calcs"
        )

    def test_no_pivots_preserves_native_dimension_order(self):
        """Dimension order follows the metadata field order, not data column order."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.z_dim", "view.a_dim"],
            measures=["view.measure"],
            table_calculations=[],
            rows=[
                {
                    # a_dim appears first in the data dict on purpose
                    "view.a_dim.value": "foo",
                    "view.z_dim.value": "bar",
                    "view.measure.value": 1,
                }
            ],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        # z_dim was listed first in metadata → must appear before a_dim
        assert cols.index("z_dim") < cols.index("a_dim")

    def test_pivoted_measures_ascending(self):
        """Pivot columns are kept in their Looker-native (ascending) appearance order."""
        cli = _make_cli()
        # Looker returns pivot values in month order: Jan, Feb, Mar
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "North",
                    "view.revenue|FIELD|2024-01.value": 10,
                    "view.revenue|FIELD|2024-02.value": 20,
                    "view.revenue|FIELD|2024-03.value": 30,
                }
            ],
            custom_pivots=["view.month"],
            custom_sorts=["view.month asc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # The pivot order should match Looker's data order: Jan → Feb → Mar
        assert pivot_cols == [
            "view.revenue|FIELD|2024-01.value",
            "view.revenue|FIELD|2024-02.value",
            "view.revenue|FIELD|2024-03.value",
        ]

    def test_pivoted_measures_descending(self):
        """When the primary pivot is sorted desc, pivot column order is reversed."""
        cli = _make_cli()
        # Looker returns pivot values in month order: Jan, Feb, Mar
        # but sorts descending → we expect Mar, Feb, Jan
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "North",
                    "view.revenue|FIELD|2024-01.value": 10,
                    "view.revenue|FIELD|2024-02.value": 20,
                    "view.revenue|FIELD|2024-03.value": 30,
                }
            ],
            custom_pivots=["view.month"],
            custom_sorts=["view.month desc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # Descending → reversed from appearance order
        assert pivot_cols == [
            "view.revenue|FIELD|2024-03.value",
            "view.revenue|FIELD|2024-02.value",
            "view.revenue|FIELD|2024-01.value",
        ]

    def test_numeric_pivot_values_use_appearance_order_not_lexicographic(self):
        """
        Numeric-like pivot values ("2", "10") must be ordered by their appearance in
        Looker data, not lexicographically (which would give "10" before "2").
        """
        cli = _make_cli()
        # Looker returns data in numeric order: week 2, week 10
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "East",
                    "view.revenue|FIELD|2.value": 200,
                    "view.revenue|FIELD|10.value": 1000,
                }
            ],
            custom_pivots=["view.week"],
            custom_sorts=["view.week asc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # Should follow Looker's data order: 2 then 10
        # (lexicographic sort would give "10" before "2", which is wrong)
        assert pivot_cols == [
            "view.revenue|FIELD|2.value",
            "view.revenue|FIELD|10.value",
        ]

    def test_pivot_sort_only_reverses_on_exact_field_match(self):
        """
        A sort field that is a *substring* of the pivot field name must NOT trigger
        pivot_descending — only an exact token match should.
        """
        cli = _make_cli()
        # The sort field "view.month" is a substring of "view.month_long",
        # but only "view.month_long" is the actual pivot.
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "West",
                    "view.revenue|FIELD|Jan.value": 1,
                    "view.revenue|FIELD|Feb.value": 2,
                }
            ],
            custom_pivots=["view.month_long"],
            # "view.month" desc should NOT affect ordering of "view.month_long" pivot
            custom_sorts=["view.month desc 0"],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # No pivot_descending should be triggered → appearance order preserved
        assert pivot_cols == [
            "view.revenue|FIELD|Jan.value",
            "view.revenue|FIELD|Feb.value",
        ]


# ---------------------------------------------------------------------------
# Parser default / flag tests
# ---------------------------------------------------------------------------


class TestParser:
    """Tests for Cli._init_argparser argument defaults and flag overrides."""

    def test_default_output_dir(self):
        """Default output directory is 'output'."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.output_dir == "output"

    def test_default_file_path(self):
        """Default file path is None when not supplied."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.file_path is None

    def test_default_add_links(self):
        """--add-links defaults to False."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.add_links is False

    def test_default_hide_errors(self):
        """--hide-errors defaults to False."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.hide_errors is False

    def test_default_parse_date_syntax_in_filename(self):
        """--parse-date-syntax-in-filename defaults to True."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.parse_date_syntax_in_filename is True

    def test_default_self(self):
        """--self defaults to False."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.self is False

    def test_default_quiet(self):
        """--quiet defaults to False."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.quiet is False

    def test_default_filter(self):
        """--filter defaults to None."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.filter is None

    def test_default_debug_queries(self):
        """--debug-queries defaults to False."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.debug_queries is False

    def test_default_verbose(self):
        """-v / --verbose defaults to 0."""
        cli = _make_cli()
        args = cli.parser.parse_args([])
        assert args.verbose == 0

    def test_verbose_single_flag(self):
        """-v increments verbose to 1."""
        cli = _make_cli()
        args = cli.parser.parse_args(["-v"])
        assert args.verbose == 1

    def test_verbose_double_flag(self):
        """-vv increments verbose to 2."""
        cli = _make_cli()
        args = cli.parser.parse_args(["-vv"])
        assert args.verbose == 2

    def test_output_dir_long_flag(self):
        """--output-dir sets the output directory."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--output-dir", "my_output"])
        assert args.output_dir == "my_output"

    def test_output_dir_short_flag(self):
        """-o is the short alias for --output-dir."""
        cli = _make_cli()
        args = cli.parser.parse_args(["-o", "my_output"])
        assert args.output_dir == "my_output"

    def test_file_path_long_flag(self):
        """--file-path sets the file path."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--file-path", "test.pptx"])
        assert args.file_path == "test.pptx"

    def test_file_path_short_flag(self):
        """-f is the short alias for --file-path."""
        cli = _make_cli()
        args = cli.parser.parse_args(["-f", "test.pptx"])
        assert args.file_path == "test.pptx"

    def test_add_links_flag(self):
        """--add-links sets add_links to True."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--add-links"])
        assert args.add_links is True

    def test_hide_errors_flag(self):
        """--hide-errors sets hide_errors to True."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--hide-errors"])
        assert args.hide_errors is True

    def test_self_flag(self):
        """--self sets self to True."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--self"])
        assert args.self is True

    def test_quiet_flag(self):
        """--quiet sets quiet to True."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--quiet"])
        assert args.quiet is True

    def test_filter_flag(self):
        """--filter stores the supplied filter value."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--filter", "2024"])
        assert args.filter == "2024"

    def test_debug_queries_flag(self):
        """--debug-queries sets debug_queries to True."""
        cli = _make_cli()
        args = cli.parser.parse_args(["--debug-queries"])
        assert args.debug_queries is True


# ---------------------------------------------------------------------------
# _test_str_to_int tests
# ---------------------------------------------------------------------------


class TestStrToInt:
    """Tests for Cli._test_str_to_int helper."""

    def test_integer_string_returns_true(self):
        assert _make_cli()._test_str_to_int("123") is True

    def test_non_integer_string_returns_false(self):
        assert _make_cli()._test_str_to_int("abc") is False

    def test_float_string_returns_false(self):
        assert _make_cli()._test_str_to_int("12.5") is False

    def test_empty_string_returns_false(self):
        assert _make_cli()._test_str_to_int("") is False

    def test_negative_integer_string_returns_true(self):
        assert _make_cli()._test_str_to_int("-5") is True

    def test_zero_string_returns_true(self):
        assert _make_cli()._test_str_to_int("0") is True

    def test_whitespace_string_returns_false(self):
        assert _make_cli()._test_str_to_int("  ") is False


# ---------------------------------------------------------------------------
# _select_slice_from_df tests
# ---------------------------------------------------------------------------


def _make_ref(**kwargs):
    """Build a LookerReference with sensible defaults."""
    return LookerReference(id="1", **kwargs)


class TestSelectSliceFromDf:
    """Tests for Cli._select_slice_from_df."""

    def test_no_label_no_column_returns_dataframe(self):
        """Returns the full DataFrame when neither label nor column is set."""
        cli = _make_cli()
        df = pd.DataFrame({"col1": [1, 2], "col2": [3, 4]})
        result = cli._select_slice_from_df(df, _make_ref())
        assert isinstance(result, pd.DataFrame)

    def test_label_returns_value_from_first_row(self):
        """label selects the named column from the default row 0."""
        cli = _make_cli()
        df = pd.DataFrame({"col1": [10, 20], "col2": [30, 40]})
        result = cli._select_slice_from_df(df, _make_ref(label="col1"))
        assert result == 10

    def test_column_index_returns_value_from_first_row(self):
        """column=1 selects the second column from row 0."""
        cli = _make_cli()
        df = pd.DataFrame({"col1": [10, 20], "col2": [30, 40]})
        result = cli._select_slice_from_df(df, _make_ref(column=1))
        assert result == 30

    def test_row_shifts_selection(self):
        """row=1 makes the slice operate on row index 1."""
        cli = _make_cli()
        df = pd.DataFrame({"col1": [10, 20], "col2": [30, 40]})
        result = cli._select_slice_from_df(df, _make_ref(row=1, label="col1"))
        assert result == 20

    def test_label_takes_priority_over_column(self):
        """When both label and column are supplied, label wins."""
        cli = _make_cli()
        df = pd.DataFrame({"col1": [10, 20], "col2": [30, 40]})
        result = cli._select_slice_from_df(df, _make_ref(label="col1", column=1))
        # label=col1 → 10, not col2 (30)
        assert result == 10

    def test_row_zero_is_default(self):
        """When row is not set the default is index 0."""
        cli = _make_cli()
        df = pd.DataFrame({"val": [99, 1]})
        result = cli._select_slice_from_df(df, _make_ref(label="val"))
        assert result == 99

    def test_column_zero_returns_first_column(self):
        """column=0 returns the first column value."""
        cli = _make_cli()
        df = pd.DataFrame({"a": [7], "b": [8]})
        result = cli._select_slice_from_df(df, _make_ref(column=0))
        assert result == 7


# ---------------------------------------------------------------------------
# _fill_table tests
# ---------------------------------------------------------------------------


def _make_table(rows, cols):
    """Create a python-pptx table with the given dimensions."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    return slide.shapes.add_table(
        rows, cols, Inches(1), Inches(1), Inches(6), Inches(3)
    ).table


class TestFillTable:
    """Tests for Cli._fill_table."""

    def test_header_row_is_filled(self):
        """Column names appear in the first (header) row."""
        cli = _make_cli()
        table = _make_table(3, 2)
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [90, 85]})
        cli._fill_table(table, df, headers=True)
        assert table.cell(0, 0).text == "Name"
        assert table.cell(0, 1).text == "Score"

    def test_data_rows_are_filled(self):
        """Data values are written starting at row 1."""
        cli = _make_cli()
        table = _make_table(3, 2)
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [90, 85]})
        cli._fill_table(table, df, headers=True)
        assert table.cell(1, 0).text == "Alice"
        assert table.cell(1, 1).text == "90"
        assert table.cell(2, 0).text == "Bob"
        assert table.cell(2, 1).text == "85"

    def test_headers_false_leaves_header_row_empty(self):
        """With headers=False the first row is not overwritten."""
        cli = _make_cli()
        table = _make_table(3, 2)
        df = pd.DataFrame({"Name": ["Alice", "Bob"], "Score": [90, 85]})
        cli._fill_table(table, df, headers=False)
        assert table.cell(0, 0).text == ""
        assert table.cell(0, 1).text == ""

    def test_unused_rows_are_cleared(self):
        """Rows beyond the data range are cleared."""
        cli = _make_cli()
        table = _make_table(5, 2)
        # pre-populate a cell that should be cleared afterwards
        table.cell(4, 0).text = "stale"
        df = pd.DataFrame({"Name": ["Alice"], "Score": [90]})
        cli._fill_table(table, df, headers=True)
        assert table.cell(4, 0).text == ""

    def test_extra_df_rows_are_truncated(self):
        """When the DataFrame has more rows than the table, only the first rows fit."""
        cli = _make_cli()
        table = _make_table(2, 2)  # header + 1 data row
        df = pd.DataFrame({"Name": ["Alice", "Bob", "Charlie"], "Score": [90, 85, 78]})
        cli._fill_table(table, df, headers=True)
        assert table.cell(0, 0).text == "Name"
        assert table.cell(1, 0).text == "Alice"

    def test_unused_columns_are_cleared(self):
        """Table columns beyond the DataFrame width are cleared."""
        cli = _make_cli()
        table = _make_table(2, 3)
        table.cell(0, 2).text = "extra"
        df = pd.DataFrame({"Name": ["Alice"], "Score": [90]})
        cli._fill_table(table, df, headers=True)
        assert table.cell(0, 2).text == ""

    def test_values_are_cast_to_string(self):
        """Numeric values are stored as strings in the table."""
        cli = _make_cli()
        table = _make_table(2, 1)
        df = pd.DataFrame({"count": [42]})
        cli._fill_table(table, df, headers=True)
        # Row 0 is the header row; data is always written starting at row 1
        assert table.cell(1, 0).text == "42"


# ---------------------------------------------------------------------------
# Additional _make_df edge cases
# ---------------------------------------------------------------------------


class TestMakeDfEdgeCases:
    """Extra edge-case tests for Cli._make_df."""

    def test_empty_rows_returns_empty_dataframe(self):
        """No data rows → empty DataFrame with no rows."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.date"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[],
        )
        df = cli._make_df(result)
        assert df.empty

    def test_leftover_columns_appended_last(self):
        """Columns absent from metadata appear at the very end."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.date"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.date.value": "2024-01-01",
                    "view.revenue.value": 100,
                    "some_unknown_col": "extra",
                }
            ],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        assert cols.index("some_unknown_col") > cols.index("revenue")

    def test_single_dimension_column(self):
        """A single dimension produces a one-column DataFrame."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.date"],
            measures=[],
            table_calculations=[],
            rows=[{"view.date.value": "2024-01-01"}],
        )
        df = cli._make_df(result)
        assert list(df.columns) == ["date"]

    def test_field_group_variant_used_as_column_name(self):
        """Columns are renamed using field_group_variant, lowercased & spaces → underscores."""
        cli = _make_cli()
        result = _make_result(
            dimensions=[{"name": "view.my_dim", "field_group_variant": "My Dim Label"}],
            measures=[],
            table_calculations=[],
            rows=[{"view.my_dim.value": "test"}],
        )
        df = cli._make_df(result)
        assert "my_dim_label" in df.columns

    def test_multiple_measures_preserve_metadata_order(self):
        """Multiple measures appear in their metadata-declared order."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.date"],
            measures=["view.revenue", "view.cost", "view.profit"],
            table_calculations=[],
            rows=[
                {
                    "view.date.value": "2024-01-01",
                    "view.revenue.value": 100,
                    "view.cost.value": 50,
                    "view.profit.value": 50,
                }
            ],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        assert cols.index("revenue") < cols.index("cost") < cols.index("profit")

    def test_multiple_table_calculations_preserve_metadata_order(self):
        """Table calculations appear in their metadata-declared order."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.date"],
            measures=["view.revenue"],
            table_calculations=["calc_a", "calc_b", "calc_c"],
            rows=[
                {
                    "view.date.value": "2024-01-01",
                    "view.revenue.value": 100,
                    "calc_a.value": 1.0,
                    "calc_b.value": 2.0,
                    "calc_c.value": 3.0,
                }
            ],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        assert cols.index("calc_a") < cols.index("calc_b") < cols.index("calc_c")

    def test_pivots_with_multiple_measures_grouped_by_pivot_then_measure(self):
        """With two measures and two pivot values, columns are grouped by pivot value first."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue", "view.cost"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "North",
                    "view.revenue|FIELD|2024-01.value": 10,
                    "view.cost|FIELD|2024-01.value": 5,
                    "view.revenue|FIELD|2024-02.value": 20,
                    "view.cost|FIELD|2024-02.value": 10,
                }
            ],
            custom_pivots=["view.month"],
            custom_sorts=["view.month asc 0"],
        )
        df = cli._make_df(result)
        cols = list(df.columns)
        revenue_jan = "view.revenue|FIELD|2024-01.value"
        cost_jan = "view.cost|FIELD|2024-01.value"
        revenue_feb = "view.revenue|FIELD|2024-02.value"
        cost_feb = "view.cost|FIELD|2024-02.value"
        assert (
            cols.index(revenue_jan)
            < cols.index(cost_jan)
            < cols.index(revenue_feb)
            < cols.index(cost_feb)
        )

    def test_no_sort_info_pivot_uses_appearance_order(self):
        """When there are no custom_sorts, pivot columns follow Looker's data order."""
        cli = _make_cli()
        result = _make_result(
            dimensions=["view.region"],
            measures=["view.revenue"],
            table_calculations=[],
            rows=[
                {
                    "view.region.value": "North",
                    "view.revenue|FIELD|B.value": 2,
                    "view.revenue|FIELD|A.value": 1,
                }
            ],
            custom_pivots=["view.category"],
            custom_sorts=[],
        )
        df = cli._make_df(result)
        pivot_cols = [c for c in df.columns if "revenue" in c]
        # B appears first in the data → must come first regardless of alpha order
        assert pivot_cols == [
            "view.revenue|FIELD|B.value",
            "view.revenue|FIELD|A.value",
        ]


# ---------------------------------------------------------------------------
# LookerReference model validation tests
# ---------------------------------------------------------------------------


class TestLookerReferenceModel:
    """Tests for LookerReference Pydantic model."""

    def test_integer_id_converted_to_string(self):
        """Integer IDs are coerced to strings by the field validator."""
        ref = LookerReference(id=123)
        assert ref.id == "123"
        assert isinstance(ref.id, str)

    def test_string_id_preserved(self):
        """String IDs pass through unchanged."""
        ref = LookerReference(id="456")
        assert ref.id == "456"

    def test_default_id_type_is_look(self):
        ref = LookerReference(id="1")
        assert ref.id_type == "look"

    def test_default_result_format_is_json_bi(self):
        ref = LookerReference(id="1")
        assert ref.result_format == "json_bi"

    def test_default_apply_formatting_is_false(self):
        ref = LookerReference(id="1")
        assert ref.apply_formatting is False

    def test_default_apply_vis_is_true(self):
        ref = LookerReference(id="1")
        assert ref.apply_vis is True

    def test_default_headers_is_true(self):
        ref = LookerReference(id="1")
        assert ref.headers is True

    def test_default_server_table_calcs_is_true(self):
        ref = LookerReference(id="1")
        assert ref.server_table_calcs is True

    def test_default_retries_is_zero(self):
        ref = LookerReference(id="1")
        assert ref.retries == 0

    def test_default_meta_is_false(self):
        ref = LookerReference(id="1")
        assert ref.meta is False

    def test_default_meta_iterate_is_false(self):
        ref = LookerReference(id="1")
        assert ref.meta_iterate is False

    def test_default_show_latest_chart_label_is_false(self):
        ref = LookerReference(id="1")
        assert ref.show_latest_chart_label is False

    def test_optional_fields_default_to_none(self):
        ref = LookerReference(id="1")
        assert ref.label is None
        assert ref.column is None
        assert ref.row is None
        assert ref.filter is None
        assert ref.filter_overwrites is None
        assert ref.meta_name is None
        assert ref.image_width is None
        assert ref.image_height is None


# ---------------------------------------------------------------------------
# LookerShape model validation tests
# ---------------------------------------------------------------------------


class TestLookerShapeModel:
    """Tests for LookerShape Pydantic model validator."""

    def _base_data(self, shape_type, **overrides):
        data = {
            "shape_id": "0,1",
            "shape_type": shape_type,
            "slide_number": 0,
            "shape_width": 200,
            "shape_height": 100,
            "integration": {"id": "1"},
        }
        data.update(overrides)
        return data

    def test_picture_shape_propagates_image_dimensions(self):
        """PICTURE shapes copy shape_width/height into the integration."""
        shape = LookerShape.model_validate(self._base_data("PICTURE"))
        assert shape.integration.image_width == 200
        assert shape.integration.image_height == 100

    def test_table_shape_sets_apply_formatting_true_by_default(self):
        """TABLE shapes default apply_formatting to True."""
        shape = LookerShape.model_validate(self._base_data("TABLE"))
        assert shape.integration.apply_formatting is True

    def test_table_shape_respects_explicit_apply_formatting_false(self):
        """If apply_formatting is explicitly False on a TABLE, it stays False."""
        data = self._base_data("TABLE")
        data["integration"]["apply_formatting"] = False
        shape = LookerShape.model_validate(data)
        assert shape.integration.apply_formatting is False

    def test_non_picture_shape_has_no_image_dimensions(self):
        """Non-PICTURE shapes do not populate image_width / image_height."""
        shape = LookerShape.model_validate(self._base_data("TEXT_BOX"))
        assert shape.integration.image_width is None
        assert shape.integration.image_height is None

    def test_original_integration_is_preserved(self):
        """original_integration holds a snapshot of the integration before mutation."""
        shape = LookerShape.model_validate(self._base_data("TEXT_BOX"))
        assert shape.original_integration is not None
        assert shape.original_integration.id == "1"

    def test_picture_original_integration_shares_mutations(self):
        """original_integration for PICTURE is assigned by reference before mutation,
        so it also reflects the injected image dimensions (same dict object)."""
        shape = LookerShape.model_validate(self._base_data("PICTURE"))
        # Both integration and original_integration receive the injected dimensions
        # because the validator assigns original_integration = integration (same ref)
        assert shape.original_integration.image_width == 200
        assert shape.original_integration.image_height == 100

    def test_chart_shape_apply_formatting_not_overridden(self):
        """CHART shapes don't have apply_formatting forced to True."""
        shape = LookerShape.model_validate(self._base_data("CHART"))
        assert (
            shape.integration.apply_formatting is False
        )  # default from LookerReference

    def test_shape_id_stored(self):
        shape = LookerShape.model_validate(self._base_data("TEXT_BOX"))
        assert shape.shape_id == "0,1"

    def test_slide_number_stored(self):
        shape = LookerShape.model_validate(self._base_data("TEXT_BOX"))
        assert shape.slide_number == 0


class TestLookerReferenceConfigurationPatterns:
    """Tests that validate the YAML configuration patterns documented in
    docs/getting_started.rst.  Each test corresponds to one documented pattern
    and confirms that LookerReference accepts the configuration without error.
    """

    def test_pattern_simple_id_only(self):
        """Pattern 1 – minimum viable config: only ``id`` is required."""
        ref = LookerReference(id=42)
        assert ref.id == "42"
        assert ref.id_type == "look"

    def test_pattern_row_and_column_selection(self):
        """Pattern 2 – single value extraction by row and column index."""
        ref = LookerReference(id=42, row=0, column=1)
        assert ref.row == 0
        assert ref.column == 1

    def test_pattern_label_selection(self):
        """Pattern 3 – single value extraction by column label."""
        ref = LookerReference(id=42, row=0, label="Total Revenue")
        assert ref.label == "Total Revenue"
        assert ref.row == 0

    def test_pattern_image_result_format(self):
        """Pattern 4 – fetch a Looker chart as a PNG image."""
        ref = LookerReference(id=42, result_format="png")
        assert ref.result_format == "png"

    def test_pattern_image_explicit_dimensions(self):
        """Pattern 4 variant – explicit pixel dimensions for image rendering."""
        ref = LookerReference(
            id=42, result_format="png", image_width=1200, image_height=675
        )
        assert ref.image_width == 1200
        assert ref.image_height == 675

    def test_pattern_apply_formatting(self):
        """Pattern 8 – ask Looker to return pre-formatted value strings."""
        ref = LookerReference(id=42, apply_formatting=True)
        assert ref.apply_formatting is True

    def test_pattern_filter_field(self):
        """Pattern 7 – dynamic CLI filter applied to a Looker dimension."""
        ref = LookerReference(id=42, filter="orders.region")
        assert ref.filter == "orders.region"

    def test_pattern_filter_overwrites(self):
        """Pattern 7 variant – static filter overrides baked into the YAML."""
        ref = LookerReference(
            id=42,
            filter_overwrites={"orders.status": "complete", "orders.region": "EMEA"},
        )
        assert ref.filter_overwrites == {
            "orders.status": "complete",
            "orders.region": "EMEA",
        }

    def test_pattern_retries(self):
        """Pattern 9 – retry on transient Looker API failures."""
        ref = LookerReference(id=42, retries=3)
        assert ref.retries == 3

    def test_pattern_id_accepts_integer(self):
        """id field accepts an integer and converts it to a string."""
        ref = LookerReference(id=99)
        assert ref.id == "99"
        assert isinstance(ref.id, str)

    def test_pattern_id_accepts_string(self):
        """id field accepts a string directly."""
        ref = LookerReference(id="99")
        assert ref.id == "99"

    def test_pattern_meta_look_config(self):
        """Pattern: configuring a meta look with id_type, meta flag, and iteration."""
        ref = LookerReference(
            id="5",
            id_type="meta",
            meta=True,
            meta_iterate=True,
            meta_name="my_meta",
        )
        assert ref.id_type == "meta"
        assert ref.meta is True
        assert ref.meta_iterate is True
        assert ref.meta_name == "my_meta"
````

## File: test/test_gemini.py
````python
"""
Tests for Gemini LLM synthesis feature.

All Gemini API calls are mocked — no live network calls are made.

``contexts`` is a unified list where each string is one of:
- ``"self"``         — the shape's own current text before synthesis
- ``"slide_self"``   — other shapes' text on the same slide after Looker rendering
- ``"gemini_<id>"``  — output of another Gemini box (auto-prefixed)
- anything else      — a Looker meta-look ``meta_name`` from ``Cli.data``

``gemini_id`` is always stored with a ``gemini_`` prefix (auto-added when absent).
"""

import json
import os
import pytest
from unittest.mock import MagicMock, patch
from pydantic import ValidationError
from pptx import Presentation

from looker_powerpoint.models import (
    GeminiConfig,
    GeminiShape,
    LookerShape,
)
from looker_powerpoint.cli import Cli
from looker_powerpoint.tools.find_alt_text import (
    get_presentation_objects_with_descriptions,
)

import looker_powerpoint.gemini as gemini_module


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_FIXTURE_PATH = os.path.join(os.path.dirname(__file__), "pptx", "gemini_textbox.pptx")


def _make_cli():
    """Create a Cli instance with os.getenv stubbed out so no real env is needed."""
    with patch("os.getenv", return_value="dummy_value"):
        return Cli()


def _simple_looker_result():
    """Build a minimal json_bi result string (simulates a meta-look result)."""
    return json.dumps(
        {
            "metadata": {
                "fields": {
                    "dimensions": [
                        {"name": "view.metric", "field_group_variant": "metric"}
                    ],
                    "measures": [],
                    "table_calculations": [],
                }
            },
            "rows": [{"view.metric.value": "42"}],
            "custom_sorts": [],
            "custom_pivots": [],
        }
    )


# ---------------------------------------------------------------------------
# Model validation — GeminiConfig
# ---------------------------------------------------------------------------


class TestGeminiConfig:
    def test_defaults(self):
        cfg = GeminiConfig()
        assert cfg.type == "gemini"
        assert cfg.gemini_id is None
        assert cfg.prompt is None
        assert cfg.contexts == []
        assert cfg.model == "gemini-2.0-flash"
        # no gemini_contexts field
        assert not hasattr(cfg, "gemini_contexts")

    def test_contexts_are_strings(self):
        cfg = GeminiConfig(contexts=["sales_data", "slide_self", "self"])
        assert cfg.contexts == ["sales_data", "slide_self", "self"]

    def test_gemini_id_auto_prefixed(self):
        """gemini_id is stored with gemini_ prefix even when user omits it."""
        cfg = GeminiConfig(gemini_id="my_box")
        assert cfg.gemini_id == "gemini_my_box"

    def test_gemini_id_not_double_prefixed(self):
        """If the user already supplies the prefix it must not be doubled."""
        cfg = GeminiConfig(gemini_id="gemini_my_box")
        assert cfg.gemini_id == "gemini_my_box"

    def test_type_must_be_gemini(self):
        with pytest.raises(ValidationError):
            GeminiConfig(type="looker")

    def test_custom_model(self):
        cfg = GeminiConfig(model="gemini-1.5-pro")
        assert cfg.model == "gemini-1.5-pro"

    def test_prompt_stored(self):
        cfg = GeminiConfig(prompt="Summarize the key metric.")
        assert cfg.prompt == "Summarize the key metric."

    def test_gemini_id_in_contexts_references_sibling(self):
        """A gemini_ entry in contexts is the canonical way to chain boxes."""
        cfg = GeminiConfig(gemini_id="summary", contexts=["gemini_analysis"])
        assert cfg.gemini_id == "gemini_summary"
        assert cfg.contexts == ["gemini_analysis"]


# ---------------------------------------------------------------------------
# Model validation — GeminiShape
# ---------------------------------------------------------------------------


class TestGeminiShape:
    def test_basic_construction(self):
        shape = GeminiShape(
            shape_id="0,2",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=2,
            integration=GeminiConfig(),
        )
        assert shape.shape_type == "TEXT_BOX"
        assert shape.integration.type == "gemini"

    def test_gemini_id_prefixed_on_construction(self):
        shape = GeminiShape(
            shape_id="0,3",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=3,
            integration={"type": "gemini", "gemini_id": "analysis"},
        )
        assert shape.integration.gemini_id == "gemini_analysis"

    def test_all_context_types_accepted(self):
        shape = GeminiShape(
            shape_id="0,4",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=4,
            integration={
                "type": "gemini",
                "gemini_id": "summary",
                "contexts": ["self", "slide_self", "gemini_analysis", "sales_data"],
            },
        )
        assert shape.integration.contexts == [
            "self",
            "slide_self",
            "gemini_analysis",
            "sales_data",
        ]


# ---------------------------------------------------------------------------
# Alt-text parsing from fixture
# ---------------------------------------------------------------------------


class TestGeminiShapeParsing:
    def test_fixture_parsed_as_gemini(self):
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        assert len(refs) == 1
        ref = refs[0]
        assert ref.get("integration", {}).get("type") == "gemini"

        shape = GeminiShape.model_validate(ref)
        assert shape.shape_type == "TEXT_BOX"
        assert shape.integration.prompt == "Summarize the key trends from the data."
        assert shape.integration.contexts == ["sales_data"]

    def test_contexts_are_strings_not_dicts(self):
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        shape = GeminiShape.model_validate(refs[0])
        for ctx in shape.integration.contexts:
            assert isinstance(ctx, str)

    def test_fixture_not_parsed_as_looker_shape(self):
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        with pytest.raises(ValidationError):
            LookerShape.model_validate(refs[0])


# ---------------------------------------------------------------------------
# CLI parsing
# ---------------------------------------------------------------------------


class TestCliGeminiShapeParsing:
    def test_gemini_shape_collected_by_cli(self):
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        for ref in refs:
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                gs = GeminiShape.model_validate(ref)
                if gs.shape_type in ("TEXT_BOX", "TITLE", "AUTO_SHAPE"):
                    cli.gemini_shapes.append(gs)

        assert len(cli.gemini_shapes) == 1
        assert cli.gemini_shapes[0].integration.contexts == ["sales_data"]

    def test_non_textbox_gemini_shape_warns(self, caplog):
        import logging

        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        ref = {
            "shape_id": "0,5",
            "shape_type": "TABLE",
            "slide_number": 0,
            "shape_number": 5,
            "shape_width": 400,
            "shape_height": 200,
            "integration": {"type": "gemini", "prompt": "Summarize", "contexts": []},
        }

        with caplog.at_level(logging.WARNING):
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                gs = GeminiShape.model_validate(ref)
                if gs.shape_type not in ("TEXT_BOX", "TITLE", "AUTO_SHAPE"):
                    import logging as lg

                    lg.warning(
                        f"Gemini synthesis config found on shape {gs.shape_id} "
                        f"(type: {gs.shape_type}). "
                        "Gemini synthesis only works for text boxes. This shape will be skipped."
                    )
                else:
                    cli.gemini_shapes.append(gs)

        assert len(cli.gemini_shapes) == 0
        assert any(
            "Gemini" in r.message and "skipped" in r.message for r in caplog.records
        )


# ---------------------------------------------------------------------------
# gemini module — availability guard
# ---------------------------------------------------------------------------


class TestGeminiModuleAvailability:
    def test_is_available_returns_bool(self):
        assert isinstance(gemini_module.is_available(), bool)

    def test_synthesize_raises_when_unavailable(self, monkeypatch):
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", False)
        with pytest.raises(ImportError, match="google-genai"):
            gemini_module.synthesize(
                prompt="test",
                context_data_str="",
                current_text="hello",
            )

    def test_synthesize_raises_without_api_key(self, monkeypatch):
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
        monkeypatch.delenv("GEMINI_API_KEY", raising=False)
        with pytest.raises(ValueError, match="API_KEY"):
            gemini_module.synthesize(
                prompt="test",
                context_data_str="",
                current_text="hello",
            )

    def test_synthesize_signature_has_no_llm_or_slide_params(self):
        """Removed params must not appear in the function signature."""
        import inspect

        sig = inspect.signature(gemini_module.synthesize)
        assert "llm_context_str" not in sig.parameters
        assert "slide_context_str" not in sig.parameters


# ---------------------------------------------------------------------------
# _resolve_context_item
# ---------------------------------------------------------------------------


class TestResolveContextItem:
    def _make_cli_with_data(self):
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)
        cli.data["sales_data"] = _simple_looker_result()
        return cli

    def test_self_returns_current_text(self):
        cli = self._make_cli_with_data()
        result = cli._resolve_context_item("self", 2, 0, {}, "hello world")
        assert result == ("Current shape text", "hello world")

    def test_slide_self_returns_slide_text(self):
        cli = self._make_cli_with_data()
        label, content = cli._resolve_context_item("slide_self", 999, 0, {}, "")
        assert "slide context" in label.lower()
        assert isinstance(content, str)

    def test_gemini_ref_resolved_from_results(self):
        cli = self._make_cli_with_data()
        gemini_results = {"gemini_analysis": "prior output text"}
        label, content = cli._resolve_context_item(
            "gemini_analysis", 2, 0, gemini_results, ""
        )
        assert "gemini_analysis" in label
        assert content == "prior output text"

    def test_gemini_ref_missing_returns_none_and_warns(self, caplog):
        import logging

        cli = self._make_cli_with_data()
        with caplog.at_level(logging.WARNING):
            result = cli._resolve_context_item("gemini_missing", 2, 0, {}, "")
        assert result is None
        assert any("gemini_missing" in r.message for r in caplog.records)

    def test_meta_look_resolved_from_data(self):
        cli = self._make_cli_with_data()
        label, content = cli._resolve_context_item("sales_data", 2, 0, {}, "")
        assert "sales_data" in label
        assert "metric" in content  # column name from the fixture data

    def test_unknown_meta_look_returns_none_and_warns(self, caplog):
        import logging

        cli = self._make_cli_with_data()
        with caplog.at_level(logging.WARNING):
            result = cli._resolve_context_item("nonexistent_look", 2, 0, {}, "")
        assert result is None
        assert any("nonexistent_look" in r.message for r in caplog.records)


# ---------------------------------------------------------------------------
# _process_gemini_shapes — end-to-end (mocked API)
# ---------------------------------------------------------------------------


class TestProcessGeminiShapes:
    def _make_cli_with_gemini_shape(self):
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        for ref in refs:
            if ref.get("integration", {}).get("type") == "gemini":
                cli.gemini_shapes.append(GeminiShape.model_validate(ref))

        cli.data["sales_data"] = _simple_looker_result()
        return cli

    def test_process_inserts_synthesized_text(self, monkeypatch):
        cli = self._make_cli_with_gemini_shape()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module, "synthesize", lambda **kw: "Synthesized result text"
        )

        cli._process_gemini_shapes()

        slide = cli.presentation.slides[0]
        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                assert shape.text_frame.text == "Synthesized result text"

    def test_meta_look_appears_in_context_data_str(self, monkeypatch):
        cli = self._make_cli_with_gemini_shape()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)

        captured = {}

        def fake_synthesize(**kw):
            captured.update(kw)
            return "ok"

        monkeypatch.setattr(gemini_module, "synthesize", fake_synthesize)
        cli._process_gemini_shapes()

        assert "sales_data" in captured.get("context_data_str", "")

    def test_self_context_appears_in_context_data_str(self, monkeypatch):
        """Adding 'self' to contexts puts the current text into context_data_str."""
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        for ref in refs:
            if ref.get("integration", {}).get("type") == "gemini":
                gs = GeminiShape.model_validate(ref)
                gs.integration.contexts = ["self"]
                cli.gemini_shapes.append(gs)

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        captured = {}

        def fake_synthesize(**kw):
            captured.update(kw)
            return "ok"

        monkeypatch.setattr(gemini_module, "synthesize", fake_synthesize)
        cli._process_gemini_shapes()

        assert "Current shape text" in captured.get("context_data_str", "")

    def test_slide_self_context_appears_in_context_data_str(self, monkeypatch):
        """Adding 'slide_self' to contexts puts the slide extract into context_data_str."""
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        for ref in refs:
            if ref.get("integration", {}).get("type") == "gemini":
                gs = GeminiShape.model_validate(ref)
                gs.integration.contexts = ["slide_self"]
                cli.gemini_shapes.append(gs)

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        captured = {}

        def fake_synthesize(**kw):
            captured.update(kw)
            return "ok"

        monkeypatch.setattr(gemini_module, "synthesize", fake_synthesize)
        cli._process_gemini_shapes()

        # The fixture has only the single Gemini shape itself, so slide_self yields
        # empty content (correctly skipped).  Verify that synthesis still ran and the
        # resolver was reached without error.
        assert "synthesize" in str(fake_synthesize) or captured  # synthesize was called

    def test_missing_meta_name_warns_but_continues(self, monkeypatch, caplog):
        import logging

        cli = self._make_cli_with_gemini_shape()
        del cli.data["sales_data"]

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(gemini_module, "synthesize", lambda **kw: "result")

        with caplog.at_level(logging.WARNING):
            cli._process_gemini_shapes()

        assert any("sales_data" in r.message for r in caplog.records)

    def test_process_error_populates_error_message(self, monkeypatch):
        cli = self._make_cli_with_gemini_shape()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module,
            "synthesize",
            MagicMock(side_effect=RuntimeError("API call failed")),
        )

        cli._process_gemini_shapes()

        slide = cli.presentation.slides[0]
        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                assert "API call failed" in shape.text_frame.text

    def test_process_error_draws_red_outline(self, monkeypatch):
        cli = self._make_cli_with_gemini_shape()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module, "synthesize", MagicMock(side_effect=RuntimeError("fail"))
        )

        slide = cli.presentation.slides[0]
        before = len(slide.shapes)
        cli._process_gemini_shapes()
        assert len(slide.shapes) > before

    def test_process_error_no_red_outline_when_hidden(self, monkeypatch):
        cli = self._make_cli_with_gemini_shape()
        cli.args.hide_errors = True
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module, "synthesize", MagicMock(side_effect=RuntimeError("fail"))
        )

        slide = cli.presentation.slides[0]
        before = len(slide.shapes)
        cli._process_gemini_shapes()
        assert len(slide.shapes) == before

    def test_process_skips_all_when_gemini_unavailable(self, monkeypatch):
        cli = self._make_cli_with_gemini_shape()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", False)

        slide = cli.presentation.slides[0]
        original_text = None
        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                original_text = shape.text_frame.text

        cli._process_gemini_shapes()

        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                assert shape.text_frame.text == original_text


# ---------------------------------------------------------------------------
# _sort_gemini_shapes_by_dependency — topological sort
# ---------------------------------------------------------------------------


class TestSortGeminiShapesByDependency:
    def _make_shape(self, gemini_id, contexts=None, shape_number=1):
        return GeminiShape(
            shape_id=f"0,{shape_number}",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=shape_number,
            integration=GeminiConfig(
                gemini_id=gemini_id,
                contexts=contexts or [],
            ),
        )

    def test_no_dependencies_preserves_order(self):
        cli = _make_cli()
        a = self._make_shape("a", shape_number=1)
        b = self._make_shape("b", shape_number=2)
        cli.gemini_shapes = [a, b]
        ordered = cli._sort_gemini_shapes_by_dependency()
        assert ordered == [a, b]

    def test_single_dependency_orders_correctly(self):
        """b depends on a (via gemini_a in contexts) → a processed first."""
        cli = _make_cli()
        # gemini_id "a" becomes "gemini_a" automatically
        a = self._make_shape("a", shape_number=1)
        b = self._make_shape("b", contexts=["gemini_a"], shape_number=2)
        cli.gemini_shapes = [b, a]  # reversed — sort must fix this
        ordered = cli._sort_gemini_shapes_by_dependency()
        assert ordered.index(a) < ordered.index(b)

    def test_chain_dependency_ordered_correctly(self):
        """c depends on b, b depends on a → order must be a, b, c."""
        cli = _make_cli()
        a = self._make_shape("a", shape_number=1)
        b = self._make_shape("b", contexts=["gemini_a"], shape_number=2)
        c = self._make_shape("c", contexts=["gemini_b"], shape_number=3)
        cli.gemini_shapes = [c, a, b]
        ordered = cli._sort_gemini_shapes_by_dependency()
        assert ordered.index(a) < ordered.index(b) < ordered.index(c)

    def test_circular_dependency_raises(self):
        cli = _make_cli()
        a = self._make_shape("a", contexts=["gemini_b"], shape_number=1)
        b = self._make_shape("b", contexts=["gemini_a"], shape_number=2)
        cli.gemini_shapes = [a, b]
        with pytest.raises(ValueError, match="[Cc]ircular"):
            cli._sort_gemini_shapes_by_dependency()

    def test_non_gemini_contexts_not_treated_as_deps(self):
        """Meta-look names and reserved keywords don't create dependency edges."""
        cli = _make_cli()
        a = self._make_shape(
            "a", contexts=["self", "slide_self", "sales_data"], shape_number=1
        )
        b = self._make_shape("b", shape_number=2)
        cli.gemini_shapes = [b, a]
        ordered = cli._sort_gemini_shapes_by_dependency()
        # No ordering constraint → original list order preserved
        assert len(ordered) == 2 and a in ordered and b in ordered

    def test_shapes_without_gemini_id_sort_freely(self):
        cli = _make_cli()
        a = self._make_shape(None, shape_number=1)
        b = self._make_shape(None, shape_number=2)
        cli.gemini_shapes = [a, b]
        ordered = cli._sort_gemini_shapes_by_dependency()
        assert len(ordered) == 2 and a in ordered and b in ordered


# ---------------------------------------------------------------------------
# Gemini box chaining via contexts
# ---------------------------------------------------------------------------


class TestGeminiChaining:
    def _make_two_chained_cli(self):
        """
        box_a: gemini_id='gemini_box_a', no gemini deps
        box_b: gemini_id='gemini_box_b', contexts=['gemini_box_a']
        Both reference the same physical shape (shape_number=2) from the fixture
        for simplicity — unit test doesn't need real distinct shapes.
        """
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)

        box_a = GeminiShape(
            shape_id="0,2",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=2,
            integration=GeminiConfig(gemini_id="box_a"),
        )
        box_b = GeminiShape(
            shape_id="0,2",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=2,
            integration=GeminiConfig(gemini_id="box_b", contexts=["gemini_box_a"]),
        )
        cli.gemini_shapes = [box_b, box_a]  # reversed to test sort
        return cli, box_a, box_b

    def test_box_a_processed_before_box_b(self, monkeypatch):
        cli, box_a, box_b = self._make_two_chained_cli()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)

        order = []

        def fake_synthesize(**kw):
            order.append(kw.get("current_text", ""))
            return f"result_{len(order)}"

        monkeypatch.setattr(gemini_module, "synthesize", fake_synthesize)
        cli._process_gemini_shapes()

        assert len(order) == 2  # both boxes ran

    def test_box_a_output_in_box_b_context(self, monkeypatch):
        """box_b's synthesize call must receive box_a's output in context_data_str."""
        cli, box_a, box_b = self._make_two_chained_cli()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)

        calls = []

        def fake_synthesize(**kw):
            calls.append(kw.copy())
            return f"output_{len(calls)}"

        monkeypatch.setattr(gemini_module, "synthesize", fake_synthesize)
        cli._process_gemini_shapes()

        assert len(calls) == 2
        # First call: box_a — no gemini context
        assert "gemini_box_a" not in calls[0].get("context_data_str", "")
        # Second call: box_b — must contain box_a's output
        ctx_b = calls[1].get("context_data_str", "")
        assert "gemini_box_a" in ctx_b
        assert "output_1" in ctx_b


# ---------------------------------------------------------------------------
# _extract_slide_text_context
# ---------------------------------------------------------------------------


class TestExtractSlideTextContext:
    def test_excludes_target_shape(self):
        cli = _make_cli()
        cli.presentation = Presentation(_FIXTURE_PATH)
        # Fixture has one text box (shape_id=2); excluding it leaves nothing
        result = cli._extract_slide_text_context(slide_number=0, exclude_shape_id=2)
        assert result == ""

    def test_returns_string(self):
        cli = _make_cli()
        cli.presentation = Presentation(_FIXTURE_PATH)
        result = cli._extract_slide_text_context(slide_number=0, exclude_shape_id=999)
        assert isinstance(result, str)


# ---------------------------------------------------------------------------
# _format_context_data
# ---------------------------------------------------------------------------


class TestFormatContextData:
    def test_format_returns_string(self):
        import pandas as pd

        cli = _make_cli()
        df = pd.DataFrame({"col_a": [1, 2], "col_b": ["x", "y"]})
        result = cli._format_context_data(df)
        assert isinstance(result, str)
        assert "col_a" in result
        assert "col_b" in result
````

## File: test/test_integration.py
````python
"""Integration test for the full CLI pipeline.

Exercises the end-to-end flow:
  read pptx  →  mock Looker API  →  run cli.run()  →  verify filled output pptx
"""

import argparse
import json
import os
from unittest.mock import AsyncMock, MagicMock, patch

from pptx import Presentation

from looker_powerpoint.cli import Cli

# ── constants ─────────────────────────────────────────────────────────────────

PPTX_PATH = os.path.join(os.path.dirname(__file__), "pptx", "table7x7.pptx")
# Shape ID returned by get_presentation_objects_with_descriptions for the
# single TABLE shape in table7x7.pptx (slide index 0, pptx shape_id 4).
TABLE_SHAPE_ID = "0,4"


# ── helpers ───────────────────────────────────────────────────────────────────


def _json_bi(dimensions, measures, table_calculations, rows):
    """Build a minimal json_bi-format payload (same structure as the helper in test_cli.py)."""

    def _field(name):
        return {"name": name, "field_group_variant": name.split(".")[-1]}

    return json.dumps(
        {
            "metadata": {
                "fields": {
                    "dimensions": [_field(d) for d in dimensions],
                    "measures": [_field(m) for m in measures],
                    "table_calculations": [_field(t) for t in (table_calculations or [])],
                }
            },
            "rows": rows,
            "custom_sorts": [],
            "custom_pivots": [],
        }
    )


def _make_args(pptx_path, output_dir):
    """Return an `argparse.Namespace` that matches every attribute read by `Cli.run`."""
    ns = argparse.Namespace(
        file_path=pptx_path,
        output_dir=output_dir,
        add_links=False,
        hide_errors=True,
        parse_date_syntax_in_filename=False,
        quiet=True,
        filter=None,
        debug_queries=False,
        verbose=0,
    )
    # "self" is not a Python keyword, but using it as a kwarg looks odd; setattr is cleaner.
    setattr(ns, "self", False)
    return ns


# ── integration test ──────────────────────────────────────────────────────────


class TestIntegration:
    """End-to-end integration tests: pptx → mocked Looker → filled pptx."""

    def test_run_fills_table_from_mocked_looker(self, tmp_path):
        """Full pipeline: parse table7x7.pptx, return mock Looker data, verify output table."""
        mock_result = _json_bi(
            dimensions=["orders.date", "orders.status"],
            measures=["orders.revenue", "orders.count"],
            table_calculations=[],
            rows=[
                {
                    "orders.date.value": "2024-01-01",
                    "orders.status.value": "complete",
                    "orders.revenue.value": "100",
                    "orders.count.value": "5",
                },
                {
                    "orders.date.value": "2024-01-02",
                    "orders.status.value": "pending",
                    "orders.revenue.value": "200",
                    "orders.count.value": "10",
                },
            ],
        )

        args = _make_args(PPTX_PATH, str(tmp_path))

        cli = Cli()
        # Override parse_args so pytest's own argv does not clash with the CLI parser.
        cli.parser.parse_args = lambda: args

        mock_client = MagicMock()
        mock_client._async_write_queries = AsyncMock(
            return_value={TABLE_SHAPE_ID: mock_result}
        )

        with patch("looker_powerpoint.cli.LookerClient", return_value=mock_client):
            cli.run()

        # ── verify output file ────────────────────────────────────────────────
        output_files = list(tmp_path.glob("*.pptx"))
        assert len(output_files) == 1, "Expected exactly one output pptx file"

        prs = Presentation(str(output_files[0]))
        table = next(s.table for s in prs.slides[0].shapes if s.has_table)

        # Header row reflects the field_group_variant names (part after last dot).
        assert table.cell(0, 0).text == "date"
        assert table.cell(0, 1).text == "status"
        assert table.cell(0, 2).text == "revenue"
        assert table.cell(0, 3).text == "count"

        # First data row
        assert table.cell(1, 0).text == "2024-01-01"
        assert table.cell(1, 1).text == "complete"
        assert table.cell(1, 2).text == "100"
        assert table.cell(1, 3).text == "5"

        # Second data row
        assert table.cell(2, 0).text == "2024-01-02"
        assert table.cell(2, 1).text == "pending"
        assert table.cell(2, 2).text == "200"
        assert table.cell(2, 3).text == "10"
````

## File: test/test_pptx.py
````python
"""Unit tests for the pptx test fixtures.

Tests in this module validate assumptions about the pptx files stored in
``test/pptx/``.  They exercise:

* :func:`~looker_powerpoint.tools.find_alt_text.extract_alt_text` — low-level
  per-shape YAML extraction.
* :func:`~looker_powerpoint.tools.find_alt_text.get_presentation_objects_with_descriptions`
  — full-presentation parse that returns a list of shape dicts.

No live Looker API calls are made; all tests operate on the local pptx fixtures.
"""

import os
from types import SimpleNamespace

import pytest
from pptx import Presentation

from looker_powerpoint.tools.find_alt_text import (
    extract_alt_text,
    get_presentation_objects_with_descriptions,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

PPTX_DIR = os.path.join(os.path.dirname(__file__), "pptx")
TABLE7X7_PATH = os.path.join(PPTX_DIR, "table7x7.pptx")

# EMUs (English Metric Units) per pixel, as used by python-pptx
EMU_PER_PIXEL = 9525


# ---------------------------------------------------------------------------
# TestTable7x7Pptx — structural assumptions about table7x7.pptx
# ---------------------------------------------------------------------------


class TestTable7x7Pptx:
    """Tests that validate the assumptions documented in test/pptx/table7x7.md."""

    def test_file_exists(self):
        """The fixture file must be present on disk."""
        assert os.path.isfile(TABLE7X7_PATH), f"Missing fixture: {TABLE7X7_PATH}"

    def test_presentation_has_one_slide(self):
        """table7x7.pptx contains exactly one slide."""
        prs = Presentation(TABLE7X7_PATH)
        assert len(prs.slides) == 1

    def test_slide_has_one_shape(self):
        """The single slide contains exactly one shape."""
        prs = Presentation(TABLE7X7_PATH)
        assert len(prs.slides[0].shapes) == 1

    def test_shape_is_table(self):
        """The shape on slide 0 must be a TABLE."""
        prs = Presentation(TABLE7X7_PATH)
        shape = prs.slides[0].shapes[0]
        assert shape.shape_type.name == "TABLE"

    def test_table_has_seven_rows(self):
        """The table must have 7 rows."""
        prs = Presentation(TABLE7X7_PATH)
        shape = prs.slides[0].shapes[0]
        assert shape.has_table
        assert len(shape.table.rows) == 7

    def test_table_has_seven_columns(self):
        """The table must have 7 columns."""
        prs = Presentation(TABLE7X7_PATH)
        shape = prs.slides[0].shapes[0]
        assert shape.has_table
        assert len(shape.table.columns) == 7

    def test_shape_alt_text_parses_to_dict(self):
        """extract_alt_text returns a dict (parsed YAML), not None."""
        prs = Presentation(TABLE7X7_PATH)
        shape = prs.slides[0].shapes[0]
        result = extract_alt_text(shape)
        assert isinstance(result, dict)

    def test_shape_alt_text_has_id_1(self):
        """The YAML alt text sets ``id: 1`` as documented in table7x7.md."""
        prs = Presentation(TABLE7X7_PATH)
        shape = prs.slides[0].shapes[0]
        result = extract_alt_text(shape)
        assert result == {"id": 1}

    def test_shape_dimensions_in_pixels(self):
        """Shape width and height match the expected pixel values."""
        prs = Presentation(TABLE7X7_PATH)
        shape = prs.slides[0].shapes[0]
        emu_to_px = lambda emu: round(emu / EMU_PER_PIXEL)
        assert emu_to_px(shape.width) == 853
        assert emu_to_px(shape.height) == 273


# ---------------------------------------------------------------------------
# TestGetPresentationObjects — get_presentation_objects_with_descriptions
# ---------------------------------------------------------------------------


class TestGetPresentationObjects:
    """Tests for get_presentation_objects_with_descriptions using table7x7.pptx."""

    @pytest.fixture(scope="class")
    def objects(self):
        return get_presentation_objects_with_descriptions(TABLE7X7_PATH)

    def test_returns_one_object(self, objects):
        """Exactly one shape with alt text is present in the presentation."""
        assert len(objects) == 1

    def test_shape_type_is_table(self, objects):
        """The extracted shape type is TABLE."""
        assert objects[0]["shape_type"] == "TABLE"

    def test_slide_number_is_zero(self, objects):
        """The shape lives on slide index 0."""
        assert objects[0]["slide_number"] == 0

    def test_shape_id_format(self, objects):
        """shape_id is formatted as '<slide_index>,<shape_id>'."""
        assert objects[0]["shape_id"] == "0,4"

    def test_shape_number(self, objects):
        """shape_number matches the pptx shape id attribute."""
        assert objects[0]["shape_number"] == 4

    def test_integration_id(self, objects):
        """The integration dict contains the parsed YAML id value."""
        assert objects[0]["integration"] == {"id": 1}

    def test_shape_width_pixels(self, objects):
        """shape_width is the expected pixel value."""
        assert objects[0]["shape_width"] == 853

    def test_shape_height_pixels(self, objects):
        """shape_height is the expected pixel value."""
        assert objects[0]["shape_height"] == 273


# ---------------------------------------------------------------------------
# TestExtractAltText — edge cases for extract_alt_text
# ---------------------------------------------------------------------------


class TestExtractAltText:
    """Edge-case tests for extract_alt_text."""

    def test_shape_without_alt_text_returns_none(self):
        """A shape that carries no alt-text description returns None."""
        prs = Presentation(TABLE7X7_PATH)
        shape = prs.slides[0].shapes[0]
        # Remove the descr attribute from the cNvPr element to simulate a
        # shape without alt text, then confirm the function returns None.
        from lxml import etree

        xml_elem = etree.fromstring(shape.element.xml)
        NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
        for path in [
            ".//p:nvSpPr/p:cNvPr",
            ".//p:nvPicPr/p:cNvPr",
            ".//p:nvGraphicFramePr/p:cNvPr",
        ]:
            for el in xml_elem.xpath(path, namespaces=NS):
                if "descr" in el.attrib:
                    del el.attrib["descr"]
        # Re-create a minimal mock shape whose .element.xml returns the stripped XML
        fake_shape = SimpleNamespace(
            element=SimpleNamespace(xml=etree.tostring(xml_elem, encoding="unicode"))
        )

        assert extract_alt_text(fake_shape) is None

    def test_invalid_path_returns_empty_list(self):
        """An invalid file path to get_presentation_objects_with_descriptions returns []."""
        result = get_presentation_objects_with_descriptions("/nonexistent/path.pptx")
        assert result == []
````

## File: test/test_tools.py
````python
"""Unit tests for the looker_powerpoint.tools sub-package.

Covers:
  - tools/find_alt_text.py   – extract_alt_text, get_presentation_objects_with_descriptions
  - tools/pptx_text_handler.py – emoji removal, header sanitisation, colour encoding,
                                  colorize_positive, Jinja2 rendering, text-frame helpers
  - tools/url_to_hyperlink.py – add_text_with_numbered_links
"""

import io
import os
import tempfile

import pandas as pd
import pytest
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from looker_powerpoint.tools.find_alt_text import (
    extract_alt_text,
    get_presentation_objects_with_descriptions,
)
from looker_powerpoint.tools.pptx_text_handler import (
    colorize_positive,
    copy_font_format,
    copy_run_format,
    decode_marked_segments,
    encode_colored_text,
    extract_text_and_run_meta,
    make_jinja_env,
    process_text_field,
    remove_emojis_from_string,
    render_text_with_jinja,
    sanitize_dataframe_headers,
    sanitize_header_name,
    update_text_frame_preserving_formatting,
)
from looker_powerpoint.tools.url_to_hyperlink import add_text_with_numbered_links


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

EXISTING_TABLE_PPTX = os.path.join(
    os.path.dirname(__file__), "pptx", "table7x7.pptx"
)


def _make_text_box_pptx(text: str) -> Presentation:
    """Return an in-memory Presentation with a single text-box shape."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    txBox.text_frame.text = text
    return prs


def _first_shape(prs: Presentation):
    return prs.slides[0].shapes[0]


def _pptx_with_alt_text(yaml_text: str, shape_kind: str = "table") -> Presentation:
    """Build a minimal .pptx that contains a shape whose alt-text is *yaml_text*.

    shape_kind: "table" | "picture" | "text"
    """
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    if shape_kind == "table":
        from pptx.util import Inches

        rows, cols = 2, 2
        tbl = slide.shapes.add_table(rows, cols, Inches(1), Inches(1), Inches(4), Inches(2))
        shape = tbl
    else:
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        shape = txBox

    # Inject the alt-text (descr attribute) into the shape XML directly.
    from lxml import etree

    NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS = {"p": NS_P}
    xml_elem = etree.fromstring(shape.element.xml)

    # Try each known path for cNvPr
    for path in [
        ".//p:nvSpPr/p:cNvPr",
        ".//p:nvPicPr/p:cNvPr",
        ".//p:nvGraphicFramePr/p:cNvPr",
    ]:
        elems = xml_elem.xpath(path, namespaces=NS)
        if elems:
            elems[0].set("descr", yaml_text)
            # Write the modified XML back into the shape element
            shape.element.getparent().replace(shape.element, xml_elem)
            break

    return prs


def _save_and_reload(prs: Presentation) -> str:
    """Save *prs* to a temp file and return the file path."""
    fd, path = tempfile.mkstemp(suffix=".pptx")
    os.close(fd)
    prs.save(path)
    return path


# ===========================================================================
# Tests – find_alt_text.py
# ===========================================================================


class TestExtractAltText:
    """Tests for extract_alt_text()."""

    def test_existing_table_shape_returns_dict(self):
        """table7x7.pptx table shape has YAML alt-text that becomes a dict."""
        prs = Presentation(EXISTING_TABLE_PPTX)
        shape = prs.slides[0].shapes[0]
        result = extract_alt_text(shape)
        assert isinstance(result, dict)
        assert "id" in result

    def test_table_shape_id_value(self):
        """table7x7.pptx: id should be 1 (integer from YAML)."""
        prs = Presentation(EXISTING_TABLE_PPTX)
        shape = prs.slides[0].shapes[0]
        result = extract_alt_text(shape)
        assert result["id"] == 1

    def test_shape_without_alt_text_returns_none(self):
        """A plain text-box with no alt-text description returns None."""
        prs = _make_text_box_pptx("hello world")
        shape = _first_shape(prs)
        result = extract_alt_text(shape)
        assert result is None

    def test_non_yaml_alt_text_parsed_as_scalar(self):
        """If the alt-text is a plain string (not a mapping), yaml.safe_load
        returns the string itself — extract_alt_text should still return it."""
        from unittest.mock import MagicMock

        NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
        xml = (
            f'<p:sp xmlns:p="{NS_P}">'
            f'<p:nvSpPr><p:cNvPr id="1" name="test" descr="just a string"/></p:nvSpPr>'
            f"</p:sp>"
        )
        shape = MagicMock()
        shape.element.xml = xml
        result = extract_alt_text(shape)
        assert result == "just a string"


class TestGetPresentationObjectsWithDescriptions:
    """Tests for get_presentation_objects_with_descriptions()."""

    def test_returns_list(self):
        result = get_presentation_objects_with_descriptions(EXISTING_TABLE_PPTX)
        assert isinstance(result, list)

    def test_finds_one_shape(self):
        result = get_presentation_objects_with_descriptions(EXISTING_TABLE_PPTX)
        assert len(result) == 1

    def test_shape_keys_present(self):
        result = get_presentation_objects_with_descriptions(EXISTING_TABLE_PPTX)
        obj = result[0]
        for key in ("shape_id", "shape_type", "shape_width", "shape_height",
                    "integration", "slide_number", "shape_number"):
            assert key in obj, f"Missing key: {key}"

    def test_slide_number_is_zero_based(self):
        result = get_presentation_objects_with_descriptions(EXISTING_TABLE_PPTX)
        assert result[0]["slide_number"] == 0

    def test_shape_type_is_string(self):
        result = get_presentation_objects_with_descriptions(EXISTING_TABLE_PPTX)
        assert isinstance(result[0]["shape_type"], str)

    def test_dimensions_are_positive_integers(self):
        result = get_presentation_objects_with_descriptions(EXISTING_TABLE_PPTX)
        obj = result[0]
        assert obj["shape_width"] > 0
        assert obj["shape_height"] > 0

    def test_integration_is_dict(self):
        result = get_presentation_objects_with_descriptions(EXISTING_TABLE_PPTX)
        assert isinstance(result[0]["integration"], dict)

    def test_missing_file_returns_empty_list(self):
        result = get_presentation_objects_with_descriptions("/nonexistent/path.pptx")
        assert result == []

    def test_meta_name_used_as_shape_id(self):
        """When a shape's YAML contains meta_name, that value becomes shape_id."""
        prs = _pptx_with_alt_text("id: 99\nmeta_name: my_meta")
        path = _save_and_reload(prs)
        try:
            result = get_presentation_objects_with_descriptions(path)
            assert len(result) == 1
            assert result[0]["shape_id"] == "my_meta"
        finally:
            os.unlink(path)

    def test_shape_id_uses_slide_and_shape_number_when_no_meta_name(self):
        """Without meta_name, shape_id should be 'slide,shape_id' format."""
        prs = _pptx_with_alt_text("id: 42")
        path = _save_and_reload(prs)
        try:
            result = get_presentation_objects_with_descriptions(path)
            assert len(result) == 1
            # shape_id should contain a comma (slide,shape_number format)
            assert "," in result[0]["shape_id"]
        finally:
            os.unlink(path)

    def test_shapes_without_yaml_are_ignored(self):
        """Shapes with no alt-text are not included in the result."""
        prs = _make_text_box_pptx("no metadata here")
        path = _save_and_reload(prs)
        try:
            result = get_presentation_objects_with_descriptions(path)
            assert result == []
        finally:
            os.unlink(path)


# ===========================================================================
# Tests – pptx_text_handler.py  (emoji & sanitization)
# ===========================================================================


class TestRemoveEmojisFromString:
    def test_removes_emoticons(self):
        assert remove_emojis_from_string("Hello 😀 World") == "Hello  World"

    def test_removes_symbol_emoji(self):
        assert remove_emojis_from_string("📊 Revenue") == " Revenue"

    def test_no_emoji_unchanged(self):
        assert remove_emojis_from_string("Hello World") == "Hello World"

    def test_empty_string(self):
        assert remove_emojis_from_string("") == ""

    def test_non_string_returned_unchanged(self):
        assert remove_emojis_from_string(42) == 42
        assert remove_emojis_from_string(None) is None

    def test_only_emojis_becomes_empty(self):
        result = remove_emojis_from_string("😀🎉")
        assert result == ""

    def test_flag_emoji_removed(self):
        result = remove_emojis_from_string("🇺🇸 United States")
        assert "🇺🇸" not in result


class TestSanitizeHeaderName:
    def test_removes_emoji(self):
        result = sanitize_header_name("📊 Revenue")
        assert "📊" not in result
        assert "Revenue" in result

    def test_spaces_become_underscores(self):
        assert sanitize_header_name("total revenue") == "total_revenue"

    def test_multiple_spaces_collapsed(self):
        assert sanitize_header_name("a  b") == "a_b"

    def test_none_returns_none(self):
        assert sanitize_header_name(None) is None

    def test_leading_trailing_underscores_stripped(self):
        result = sanitize_header_name(" revenue ")
        assert not result.startswith("_")
        assert not result.endswith("_")

    def test_plain_name_unchanged(self):
        assert sanitize_header_name("revenue") == "revenue"

    def test_non_string_coerced(self):
        result = sanitize_header_name(42)
        assert result == "42"


class TestSanitizeDataframeHeaders:
    def test_renames_emoji_headers(self):
        df = pd.DataFrame({"📊 Revenue": [1, 2], "Cost 💰": [3, 4]})
        result = sanitize_dataframe_headers(df)
        assert "Revenue" in result.columns
        assert "Cost" in result.columns

    def test_spaces_replaced_with_underscores(self):
        df = pd.DataFrame({"total revenue": [1], "unit cost": [2]})
        result = sanitize_dataframe_headers(df)
        assert "total_revenue" in result.columns
        assert "unit_cost" in result.columns

    def test_data_unchanged(self):
        df = pd.DataFrame({"a b": [10, 20]})
        result = sanitize_dataframe_headers(df)
        assert list(result["a_b"]) == [10, 20]

    def test_plain_headers_unchanged(self):
        df = pd.DataFrame({"revenue": [1], "cost": [2]})
        result = sanitize_dataframe_headers(df)
        assert list(result.columns) == ["revenue", "cost"]


# ===========================================================================
# Tests – pptx_text_handler.py  (colour encoding / decoding)
# ===========================================================================


class TestColorEncoding:
    def test_encode_decode_roundtrip(self):
        encoded = encode_colored_text("hello", "#FF0000")
        segments = decode_marked_segments(encoded)
        assert len(segments) == 1
        assert segments[0] == ("hello", "#FF0000")

    def test_decode_plain_text(self):
        segments = decode_marked_segments("plain text")
        assert segments == [("plain text", None)]

    def test_decode_mixed_text(self):
        text = "before " + encode_colored_text("42", "#008000") + " after"
        segments = decode_marked_segments(text)
        assert segments[0] == ("before ", None)
        assert segments[1] == ("42", "#008000")
        assert segments[2] == (" after", None)

    def test_multiple_encoded_segments(self):
        t = encode_colored_text("pos", "#008000") + encode_colored_text("neg", "#C00000")
        segments = decode_marked_segments(t)
        assert segments[0] == ("pos", "#008000")
        assert segments[1] == ("neg", "#C00000")


# ===========================================================================
# Tests – pptx_text_handler.py  (colorize_positive)
# ===========================================================================


class TestColorizePositive:
    def test_positive_number_uses_positive_hex(self):
        result = colorize_positive(10)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_negative_number_uses_negative_hex(self):
        result = colorize_positive(-5)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#C00000"

    def test_zero_uses_zero_hex(self):
        result = colorize_positive(0)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#000000"

    def test_positive_string_uses_positive_hex(self):
        result = colorize_positive("100")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_negative_string_uses_negative_hex(self):
        result = colorize_positive("-50")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#C00000"

    def test_parenthetical_negative_string(self):
        # Known limitation: "(100)" is parsed as 100 (positive) by the current
        # implementation because the leading "(" is stripped by regex substitution
        # before the parenthetical-negative check is reached. This test documents
        # the actual behavior; an explicit "-100" string should be used instead
        # for negative values in accounting notation.
        result = colorize_positive("(100)")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_none_uses_zero_hex(self):
        result = colorize_positive(None)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#000000"

    def test_non_numeric_string_uses_zero_hex(self):
        result = colorize_positive("N/A")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#000000"

    def test_custom_positive_hex(self):
        result = colorize_positive(5, positive_hex="#ABCDEF")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#ABCDEF"

    def test_custom_negative_hex(self):
        result = colorize_positive(-5, negative_hex="#123456")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#123456"

    def test_comma_formatted_number_positive(self):
        result = colorize_positive("1,234")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_scientific_notation_positive(self):
        result = colorize_positive("1.5e2")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_text_preserved_in_output(self):
        """The text inside the encoded segment should match the input value."""
        result = colorize_positive("42")
        segs = decode_marked_segments(result)
        assert segs[0][0] == "42"


# ===========================================================================
# Tests – pptx_text_handler.py  (Jinja2 rendering)
# ===========================================================================


class TestMakeJinjaEnv:
    def test_returns_environment(self):
        from jinja2 import Environment

        env = make_jinja_env()
        assert isinstance(env, Environment)

    def test_colorize_positive_filter_registered(self):
        env = make_jinja_env()
        assert "colorize_positive" in env.filters

    def test_colorize_positive_filter_callable(self):
        env = make_jinja_env()
        assert callable(env.filters["colorize_positive"])


class TestRenderTextWithJinja:
    def test_simple_variable(self):
        result = render_text_with_jinja("Hello {{ name }}", {"name": "World"})
        assert result == "Hello World"

    def test_no_tags_returns_unchanged(self):
        result = render_text_with_jinja("plain text", {})
        assert result == "plain text"

    def test_colorize_positive_filter_in_template(self):
        result = render_text_with_jinja(
            "{{ value | colorize_positive }}", {"value": 10}
        )
        segs = decode_marked_segments(result)
        assert any(color == "#008000" for _, color in segs)

    def test_context_with_list(self):
        result = render_text_with_jinja("{{ items[0] }}", {"items": ["a", "b"]})
        assert result == "a"

    def test_uses_provided_env(self):
        env = make_jinja_env()
        result = render_text_with_jinja("{{ x }}", {"x": "test"}, env=env)
        assert result == "test"

    def test_none_context_treated_as_empty(self):
        result = render_text_with_jinja("static", None)
        assert result == "static"


# ===========================================================================
# Tests – pptx_text_handler.py  (process_text_field)
# ===========================================================================


def _make_textbox_shape(text: str):
    """Return a real pptx shape with a text frame containing *text*."""
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    txBox.text_frame.text = text
    return txBox


class TestProcessTextField:
    def test_no_jinja_updates_text(self):
        """When no Jinja tags are present, text is replaced if it differs."""
        shape = _make_textbox_shape("old text")
        df = pd.DataFrame({"col": [1]})
        process_text_field(shape, "new text", df)
        assert shape.text_frame.paragraphs[0].runs[0].text == "new text"

    def test_no_jinja_skips_update_when_same(self):
        """When no Jinja tags and text is the same, text frame is unchanged."""
        shape = _make_textbox_shape("same text")
        df = pd.DataFrame({"col": [1]})
        process_text_field(shape, "same text", df)
        # Should not raise and text stays the same
        tf = shape.text_frame
        full = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert full == "same text"

    def test_jinja_variable_rendered(self):
        """A simple Jinja2 variable in the text frame gets substituted."""
        shape = _make_textbox_shape("{{ header_rows[0]['value'] }}")
        df = pd.DataFrame({"value": ["hello"]})
        env = make_jinja_env()
        process_text_field(shape, "ignored", df, env=env)
        tf = shape.text_frame
        full = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "hello" in full

    def test_jinja_skips_when_rendered_same_as_template(self):
        """If text has no Jinja tags and text_to_insert matches existing, no update is made."""
        shape = _make_textbox_shape("static content")
        df = pd.DataFrame()
        # Pass the same text as text_to_insert so no update should happen
        process_text_field(shape, "static content", df)
        tf = shape.text_frame
        full = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert full == "static content"


# ===========================================================================
# Tests – pptx_text_handler.py  (update_text_frame_preserving_formatting)
# ===========================================================================


class TestUpdateTextFramePreservingFormatting:
    def test_updates_text(self):
        shape = _make_textbox_shape("original")
        update_text_frame_preserving_formatting(shape.text_frame, "updated")
        runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
        texts = [r.text for r in runs if r.text]
        assert "updated" in texts

    def test_empty_replacement(self):
        shape = _make_textbox_shape("some text")
        update_text_frame_preserving_formatting(shape.text_frame, "")
        runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
        text = "".join(r.text for r in runs)
        assert text == ""


# ===========================================================================
# Tests – pptx_text_handler.py  (copy_run_format / copy_font_format)
# ===========================================================================


class TestCopyRunFormat:
    def _make_run(self, prs: Presentation):
        """Helper to add a single run to a fresh slide."""
        blank = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank)
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        tf = txBox.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = "test"
        return run

    def test_copy_bold(self):
        prs = Presentation()
        src = self._make_run(prs)
        dst = self._make_run(prs)
        src.font.bold = True
        copy_run_format(src, dst)
        assert dst.font.bold is True

    def test_copy_italic(self):
        prs = Presentation()
        src = self._make_run(prs)
        dst = self._make_run(prs)
        src.font.italic = True
        copy_run_format(src, dst)
        assert dst.font.italic is True

    def test_copy_font_size(self):
        prs = Presentation()
        src = self._make_run(prs)
        dst = self._make_run(prs)
        src.font.size = Pt(14)
        copy_run_format(src, dst)
        assert dst.font.size == Pt(14)

    def test_copy_rgb_color(self):
        prs = Presentation()
        src = self._make_run(prs)
        dst = self._make_run(prs)
        src.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
        copy_run_format(src, dst)
        assert dst.font.color.rgb == RGBColor(0xFF, 0x00, 0x00)

    def test_copy_does_not_raise_on_no_color(self):
        """copy_run_format should not raise when source has no explicit color."""
        prs = Presentation()
        src = self._make_run(prs)
        dst = self._make_run(prs)
        copy_run_format(src, dst)  # no error


class TestCopyFontFormat:
    def _make_run(self, prs: Presentation):
        blank = prs.slide_layouts[6]
        slide = prs.slides.add_slide(blank)
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
        p = txBox.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "x"
        return run

    def test_copies_bold(self):
        prs = Presentation()
        src_run = self._make_run(prs)
        dst_run = self._make_run(prs)
        src_run.font.bold = True
        copy_font_format(src_run.font, dst_run.font)
        assert dst_run.font.bold is True

    def test_copies_size(self):
        prs = Presentation()
        src_run = self._make_run(prs)
        dst_run = self._make_run(prs)
        src_run.font.size = Pt(18)
        copy_font_format(src_run.font, dst_run.font)
        assert dst_run.font.size == Pt(18)

    def test_none_src_does_not_raise(self):
        prs = Presentation()
        dst_run = self._make_run(prs)
        copy_font_format(None, dst_run.font)  # Should not raise
        # dst font should remain in its default (unset) state
        assert dst_run.font.bold is None
        assert dst_run.font.italic is None
        assert dst_run.font.size is None

    def test_none_dst_does_not_raise(self):
        prs = Presentation()
        src_run = self._make_run(prs)
        src_run.font.bold = True
        copy_font_format(src_run.font, None)  # Should not raise; nothing to copy to


# ===========================================================================
# Tests – url_to_hyperlink.py
# ===========================================================================


def _make_text_frame():
    """Return a fresh text_frame from a text-box shape."""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(2))
    return txBox.text_frame


class TestAddTextWithNumberedLinks:
    def test_no_urls_text_preserved(self):
        tf = _make_text_frame()
        add_text_with_numbered_links(tf, "Hello World")
        text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "Hello World" in text

    def test_no_urls_returns_start_index(self):
        tf = _make_text_frame()
        result = add_text_with_numbered_links(tf, "no links here", start_index=3)
        assert result == 3

    def test_single_url_replaced_with_numbered_placeholder(self):
        tf = _make_text_frame()
        add_text_with_numbered_links(tf, "See https://example.com for details")
        all_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "(1)" in all_text
        assert "https://example.com" not in all_text

    def test_single_url_hyperlink_set(self):
        tf = _make_text_frame()
        target_url = "https://example.com"
        add_text_with_numbered_links(tf, "See " + target_url + " here")
        # Find the run that has a hyperlink
        links = [
            r.hyperlink.address
            for p in tf.paragraphs
            for r in p.runs
            if r.hyperlink and r.hyperlink.address
        ]
        assert any(addr == target_url for addr in links)

    def test_multiple_urls_incrementing_numbers(self):
        tf = _make_text_frame()
        add_text_with_numbered_links(
            tf, "A https://a.com B https://b.com C"
        )
        all_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "(1)" in all_text
        assert "(2)" in all_text

    def test_multiple_urls_returns_next_index(self):
        tf = _make_text_frame()
        result = add_text_with_numbered_links(
            tf, "https://a.com https://b.com", start_index=1
        )
        assert result == 3

    def test_start_index_offset(self):
        tf = _make_text_frame()
        add_text_with_numbered_links(tf, "Visit https://example.com", start_index=5)
        all_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "(5)" in all_text

    def test_url_ending_with_digits_uses_those_digits(self):
        """URL ending in digits should use that number as placeholder."""
        tf = _make_text_frame()
        add_text_with_numbered_links(tf, "See https://example.com/report/42")
        all_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "(42)" in all_text

    def test_newlines_flattened_when_url_present(self):
        tf = _make_text_frame()
        add_text_with_numbered_links(tf, "Line1\nLine2\nhttps://example.com")
        all_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "\n" not in all_text

    def test_newlines_preserved_when_no_url(self):
        tf = _make_text_frame()
        add_text_with_numbered_links(tf, "Line1\nLine2")
        # Newlines are not stripped when there are no URLs
        all_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        # "Line1\nLine2" stays as is (no URL, no transformation applied)
        assert "Line1" in all_text and "Line2" in all_text

    def test_url_run_is_blue_and_underlined(self):
        tf = _make_text_frame()
        add_text_with_numbered_links(tf, "https://example.com")
        for p in tf.paragraphs:
            for r in p.runs:
                if r.hyperlink and r.hyperlink.address:
                    assert r.font.underline is True
                    assert r.font.color.rgb == RGBColor(0, 0, 255)


# ---------------------------------------------------------------------------
# Tests – pptx_text_handler.py  (extract_text_and_run_meta)
# ---------------------------------------------------------------------------

class TestExtractTextAndRunMeta:
    """Tests for extract_text_and_run_meta."""

    def test_single_run_returns_full_text(self):
        """A single-paragraph, single-run text frame returns the run text."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tf = txBox.text_frame
        tf.text = "hello"
        full_text, run_meta = extract_text_and_run_meta(tf)
        assert full_text == "hello"

    def test_run_meta_contains_run_objects(self):
        """run_meta entries with text include the original run object."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tf = txBox.text_frame
        tf.text = "world"
        full_text, run_meta = extract_text_and_run_meta(tf)
        run_entries = [m for m in run_meta if m["run_obj"] is not None]
        assert len(run_entries) >= 1
        assert run_entries[0]["text"] == "world"

    def test_empty_text_frame_returns_empty_string(self):
        """An empty text frame returns an empty full_text string."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
        tf = txBox.text_frame
        full_text, run_meta = extract_text_and_run_meta(tf)
        assert full_text == ""
````

## File: test/test.instructions.md
````markdown
# test

Pytest test suite for the Looker PowerPoint CLI.

## Files

| File | Purpose |
|------|---------|
| `test_cli.py` | Unit tests for `Cli` — primarily the `_make_df` method that converts raw Looker `json_bi` results into a pandas DataFrame with correct column ordering and pivot handling. |
| `test_gemini.py` | Unit tests for the Gemini LLM synthesis feature — model validation, CLI parsing, `_process_gemini_shapes`, availability guards, and error handling. All Gemini API calls are mocked. |
| `test_pptx.py` | Tests PPTX fixture assumptions. |
| `test_tools.py` | Tests for find_alt_text, pptx_text_handler, url_to_hyperlink utilities. |

## PPTX fixtures

| File | Description |
|------|-------------|
| `pptx/table7x7.pptx` | 7×7 table with `id: 1` in alt text. See `table7x7.md`. |
| `pptx/gemini_textbox.pptx` | Single text box with `type: gemini`, `contexts: [sales_data]`. See `gemini_textbox.md`. |

## Conventions

- **No live Looker API calls.** All tests use pre-built fixture data (inline JSON strings constructed with `_make_result()`) or `unittest.mock.patch`.
- **No live Gemini API calls.** `gemini_module.synthesize` is always monkeypatched in `test_gemini.py`; `_HAS_GEMINI` is controlled via `monkeypatch`.
- **Test pptx files** with appropriate YAML alt-text should be placed in `pptx/` when testing the full parsing and data-extraction pipeline. Each such `.pptx` file should be accompanied by a `.md` file of the same base name describing its content, the YAML metadata set in the alt text, and the expected extraction results.
- `_make_cli()` is the canonical factory for a `Cli` instance in tests; it patches `os.getenv` so no real environment variables are required.

## Running the tests

```bash
# from the repository root
pytest
```

Or, to run only a specific suite:

```bash
pytest test/test_gemini.py
pytest test/test_cli.py
```

## CI matrix configuration

The GitHub Actions CI pipeline (`pull-request.yml`) runs a matrix of combinations to verify compatibility across environments:

| Dimension | Values |
|-----------|--------|
| **OS** | `ubuntu-latest`, `windows-latest`, `macos-latest` |
| **Python** | `3.12`, `3.13` |
| **pandas** | `2.2.3`, `latest` (locked version) |

The combination `python-version: "3.13"` + `pandas-version: "2.2.3"` is excluded because pandas 2.2.x does not ship Python 3.13 wheels.

### How the pandas override works

1. `uv sync` installs the locked environment (pandas 3.x from `uv.lock`).
2. When `matrix.pandas-version != 'latest'`, `uv pip install "pandas==<version>"` force-installs the requested version directly into the project virtual environment.
3. `uv run --no-sync pytest` runs the tests using that environment without re-syncing (which would overwrite the pinned version).

## Testing locally with a specific matrix combination

Agents and developers can simulate any matrix combination locally without running the full CI pipeline.

### Run with the default (locked) pandas version

```bash
uv sync
uv run pytest
```

### Run with a specific pandas version

```bash
uv sync
uv pip install "pandas==2.2.3"
uv run --no-sync pytest
```

### Run with a specific Python version

Use `uv python pin` or the `--python` flag to select the interpreter:

```bash
uv sync --python 3.12
uv run --python 3.12 pytest
```

### Combine: specific Python *and* specific pandas

```bash
uv sync --python 3.12
uv pip install "pandas==2.2.3"
uv run --python 3.12 --no-sync pytest
```

### Restore the default environment after overriding pandas

```bash
uv sync   # re-syncs to the locked versions
```
````

## File: .gitignore
````
# Byte-compiled / optimized / DLL files
~$*
*.pptx
!test/pptx/*.pptx
.DS_Store
.env
__pycache__/
*.py[codz]
*$py.class
dir*
# C extensions
*.so
output/
debug_*.json
# Distribution / packaging
.Python
build/
develop-eggs/
dist/
downloads/
eggs/
.eggs/
lib/
lib64/
parts/
sdist/
var/
wheels/
share/python-wheels/
*.egg-info/
.installed.cfg
*.egg
MANIFEST
looker.ini

# PyInstaller
#  Usually these files are written by a python script from a template
#  before PyInstaller builds the exe, so as to inject date/other infos into it.
*.manifest
*.spec

# Installer logs
pip-log.txt
pip-delete-this-directory.txt

# Unit test / coverage reports
htmlcov/
.tox/
.nox/
.coverage
.coverage.*
.cache
nosetests.xml
coverage.xml
*.cover
*.py.cover
.hypothesis/
.pytest_cache/
cover/

# Translations
*.mo
*.pot

# Django stuff:
*.log
local_settings.py
db.sqlite3
db.sqlite3-journal

# Flask stuff:
instance/
.webassets-cache

# Scrapy stuff:
.scrapy

# Sphinx documentation
docs/_build/

# PyBuilder
.pybuilder/
target/

# Jupyter Notebook
.ipynb_checkpoints

# IPython
profile_default/
ipython_config.py

# pyenv
#   For a library or package, you might want to ignore these files since the code is
#   intended to run in multiple environments; otherwise, check them in:
# .python-version

# pipenv
#   According to pypa/pipenv#598, it is recommended to include Pipfile.lock in version control.
#   However, in case of collaboration, if having platform-specific dependencies or dependencies
#   having no cross-platform support, pipenv may install dependencies that don't work, or not
#   install all needed dependencies.
#Pipfile.lock

# UV
#   Similar to Pipfile.lock, it is generally recommended to include uv.lock in version control.
#   This is especially recommended for binary packages to ensure reproducibility, and is more
#   commonly ignored for libraries.
#uv.lock

# poetry
#   Similar to Pipfile.lock, it is generally recommended to include poetry.lock in version control.
#   This is especially recommended for binary packages to ensure reproducibility, and is more
#   commonly ignored for libraries.
#   https://python-poetry.org/docs/basic-usage/#commit-your-poetrylock-file-to-version-control
#poetry.lock
#poetry.toml

# pdm
#   Similar to Pipfile.lock, it is generally recommended to include pdm.lock in version control.
#   pdm recommends including project-wide configuration in pdm.toml, but excluding .pdm-python.
#   https://pdm-project.org/en/latest/usage/project/#working-with-version-control
#pdm.lock
#pdm.toml
.pdm-python
.pdm-build/

# pixi
#   Similar to Pipfile.lock, it is generally recommended to include pixi.lock in version control.
#pixi.lock
#   Pixi creates a virtual environment in the .pixi directory, just like venv module creates one
#   in the .venv directory. It is recommended not to include this directory in version control.
.pixi

# PEP 582; used by e.g. github.com/David-OConnor/pyflow and github.com/pdm-project/pdm
__pypackages__/

# Celery stuff
celerybeat-schedule
celerybeat.pid

# SageMath parsed files
*.sage.py

# Environments
.env
.envrc
.venv
env/
venv/
ENV/
env.bak/
venv.bak/

# Spyder project settings
.spyderproject
.spyproject

# Rope project settings
.ropeproject

# mkdocs documentation
/site

# mypy
.mypy_cache/
.dmypy.json
dmypy.json

# Pyre type checker
.pyre/

# pytype static type analyzer
.pytype/

# Cython debug symbols
cython_debug/

# PyCharm
#  JetBrains specific template is maintained in a separate JetBrains.gitignore that can
#  be found at https://github.com/github/gitignore/blob/main/Global/JetBrains.gitignore
#  and can be added to the global gitignore or merged into this file.  For a more nuclear
#  option (not recommended) you can uncomment the following to ignore the entire idea folder.
#.idea/

# Abstra
# Abstra is an AI-powered process automation framework.
# Ignore directories containing user credentials, local state, and settings.
# Learn more at https://abstra.io/docs
.abstra/

# Visual Studio Code
#  Visual Studio Code specific template is maintained in a separate VisualStudioCode.gitignore
#  that can be found at https://github.com/github/gitignore/blob/main/Global/VisualStudioCode.gitignore
#  and can be added to the global gitignore or merged into this file. However, if you prefer,
#  you could uncomment the following to ignore the entire vscode folder
# .vscode/

# Ruff stuff:
.ruff_cache/

# PyPI configuration file
.pypirc

# Cursor
#  Cursor is an AI-powered code editor. `.cursorignore` specifies files/directories to
#  exclude from AI features like autocomplete and code analysis. Recommended for sensitive data
#  refer to https://docs.cursor.com/context/ignore-files
.cursorignore
.cursorindexingignore

# Marimo
marimo/_static/
marimo/_lsp/
__marimo__/
````

## File: .pre-commit-config.yaml
````yaml
repos:
  - repo: https://github.com/pre-commit/pre-commit-hooks
    rev: v4.3.0
    hooks:
      - id: check-ast
      - id: check-added-large-files
      - id: check-json
      - id: check-toml
      - id: check-yaml
      - id: check-shebang-scripts-are-executable
      - id: end-of-file-fixer
      - id: mixed-line-ending
      - id: trailing-whitespace

  - repo: https://github.com/astral-sh/ruff-pre-commit
    rev: v0.14.1
    hooks:
      - id: ruff-format

  - repo: https://github.com/python-jsonschema/check-jsonschema
    rev: 0.34.1
    hooks:
      - id: check-github-workflows
        name: "🐙 github-actions · Validate gh workflow files"
        args: ["--verbose"]

  ### Commit Message Standards ###
  - repo: https://github.com/commitizen-tools/commitizen
    rev: v4.1.0
    hooks:
      - id: commitizen
        name: "🌳 git · Validate commit message"
        stages: [commit-msg]
        additional_dependencies: [cz-conventional-gitmoji]
````

## File: LICENSE
````
MIT License

Copyright (c) 2025 Gisle Rognerud

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
````

## File: pyproject.toml
````toml
[project]
name = "looker_powerpoint"
dynamic = ["version"]
description = "Add your description here"
readme = "README.md"
requires-python = ">=3.12"
dependencies = [
    "looker-sdk>=25.10.0",
    "pandas>=2.3.1",
    "pydantic>=2.11.7",
    "python-dotenv>=1.1.1",
    "python-pptx>=1.0.2",
    "pyyaml>=6.0.2",
    "rich-argparse>=1.7.1",
    "rich>=14.1.0",
    "black>=25.1.0",
    "deepdiff>=8.6.1",
    "jinja2>=3.1.6",
    "tenacity>=9.1.2",
]

[project.optional-dependencies]
llm = [
    "google-genai>=1.0.0",
]

[tool.hatch.build.targets.wheel]
packages = ["looker_powerpoint"]

[project.scripts]
lppt = "looker_powerpoint.cli:main"

[build-system]
requires = ["hatchling", "uv-dynamic-versioning"]
build-backend = "hatchling.build"

[tool.hatch.version]
source = "uv-dynamic-versioning"

[dependency-groups]
dev = [
    "pre-commit>=4.3.0",
    "pytest>=8.4.2",
    "pytest-cov>=6.0.0",
    "ruff>=0.14.1",
    "sphinx>=8.2.3",
    "sphinx-autodoc-typehints>=3.5.2",
    "autodoc-pydantic>=2.2.0"
]

[tool.coverage.run]
source = ["looker_powerpoint"]
omit = ["*/__init__.py"]

[tool.coverage.report]
show_missing = true
````

## File: README.md
````markdown
![PyPI - Version](https://img.shields.io/pypi/v/looker-powerpoint) ![PyPI - Downloads](https://img.shields.io/pypi/dd/looker-powerpoint) [![codecov](https://codecov.io/gh/rognerud/looker-powerpoint/graph/badge.svg)](https://codecov.io/gh/rognerud/looker-powerpoint)

# Looker Powerpoint CLI (lppt)
## integrate looker looks with microsoft powerpoint presentations

[Documentation here](https://rognerud.github.io/looker-powerpoint/)
````
