# docs

Sphinx documentation source for the Looker PowerPoint CLI.

## Files

| File | Purpose |
|------|---------|
| `conf.py` | Sphinx configuration: project metadata, extensions (`autodoc`, `sphinx-autodoc-typehints`, `autodoc-pydantic`), and theme settings. |
| `index.rst` | Root table of contents for the generated documentation site. |
| `api.rst` | Auto-generated API reference page (populated by `sphinx.ext.autodoc`). |
| `cli.rst` | Documentation page for the `cli` module. |
| `models.rst` | Documentation page for the Pydantic models (`LookerReference`, `LookerShape`). |
| `templating.rst` | Documentation page covering the Jinja2 templating features available in PowerPoint text shapes. |
| `Makefile` / `make.bat` | Standard Sphinx build helpers for Linux/macOS and Windows respectively. |

## Building the docs

```bash
# from the docs/ directory
make html
```

The generated HTML is written to `docs/_build/html/`.
