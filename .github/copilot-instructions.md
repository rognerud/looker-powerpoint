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
Any generated pptx files, should have a corresponding markdown file in the test directory that accurately describes the content of the pptx file, with regards to yml metadata and expected data extraction results.

## Contributing
Contributions should prioritize adding tests for any new feature or bug fix, and to ensure that the documentation is updated accordingly.
In addition, adding markdown files for agentic workflow for directories, and updating this file with a updated mapping of directories and their purpose is also a good contribution.
