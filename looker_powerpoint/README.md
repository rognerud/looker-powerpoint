# looker_powerpoint

This is the main Python package for the Looker PowerPoint CLI tool (`lppt`).

## Modules

| File | Purpose |
|------|---------|
| `cli.py` | Entry point for the `lppt` CLI command. Contains the `Cli` class and `main()` function. Orchestrates fetching Looker data and writing results into PowerPoint files. |
| `looker.py` | `LookerClient` class that wraps the Looker SDK. Handles authentication, query construction, executing Look queries, and retry logic. |
| `models.py` | Pydantic models: `LookerReference` (YAML alt-text schema for a shape) and `LookerShape` (shape metadata + its `LookerReference`). |
| `__init__.py` | Package initialiser; exposes `__version__` via `importlib.metadata`. |
| `tools/` | Sub-package of utility helpers (see `tools/README.md`). |

## How it works

1. The CLI reads a `.pptx` file.
2. Each shape whose *alternative text* contains valid YAML is parsed into a `LookerReference`.
3. The `LookerClient` fetches the corresponding Looker Look and returns the result.
4. Results are written back into the presentation (text boxes, tables, images) using the helpers in `tools/`.

## Running

```bash
uv run lppt --help
```
