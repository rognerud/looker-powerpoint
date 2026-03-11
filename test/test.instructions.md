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
