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
