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
