# looker_powerpoint

This is the main Python package for the Looker PowerPoint CLI tool (`lppt`).

## Modules

| File | Purpose |
|------|---------|
| `cli.py` | Entry point for the `lppt` CLI command. Contains the `Cli` class and `main()` function. Orchestrates fetching Looker data and writing results into PowerPoint files. |
| `looker.py` | `LookerClient` class that wraps the Looker SDK. Handles authentication, query construction, executing Look queries, and retry logic. |
| `models.py` | Pydantic models: `LookerReference` and `LookerShape` (Looker-backed shapes); `GeminiConfig` and `GeminiShape` (Gemini LLM synthesis shapes). |
| `gemini.py` | Optional Google Gemini integration. Wraps `google-generativeai`; provides `is_available()` and `synthesize()`. Safe to import when the extra is not installed. |
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
prompt: Summarise the key trends.
contexts:
  - sales_data      # meta_name of a meta-look shape in the same presentation
model: gemini-2.0-flash   # optional, default shown
```

- `contexts` is a list of `meta_name` strings; each name must match a meta-look
  shape (a shape with `meta: true` and `meta_name: <name>` in its alt text).
- Only `TEXT_BOX`, `TITLE`, and `AUTO_SHAPE` types are supported; other types log a
  warning and are skipped.
- Requires the `llm` optional extra: `pip install looker_powerpoint[llm]`.
- Works without the extra installed — Gemini shapes are simply skipped with a warning.

## Running

```bash
uv run lppt --help
```
