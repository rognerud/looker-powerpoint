# test

Pytest test suite for the Looker PowerPoint CLI.

## Files

| File | Purpose |
|------|---------|
| `test_cli.py` | Unit tests for `Cli` — primarily the `_make_df` method that converts raw Looker `json_bi` results into a pandas DataFrame with correct column ordering and pivot handling. |

## Conventions

- **No live Looker API calls.** All tests use pre-built fixture data (inline JSON strings constructed with `_make_result()`) or `unittest.mock.patch`.
- **Test pptx files** with appropriate YAML alt-text should be placed in this directory when testing the full parsing and data-extraction pipeline. Each such `.pptx` file should be accompanied by a `.md` file of the same base name describing its content, the YAML metadata set in the alt text, and the expected extraction results.
- `_make_cli()` is the canonical factory for a `Cli` instance in tests; it patches `os.getenv` so no real environment variables are required.

## Running the tests

```bash
# from the repository root
pytest
```

Or, to run only this suite:

```bash
pytest test/test_cli.py
```
