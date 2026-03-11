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
  - id: 1
model: gemini-2.0-flash
```

## Purpose

Used by `test/test_gemini.py` to:

- Verify that a shape whose alt text contains `type: gemini` is parsed as a
  `GeminiShape` (not a `LookerShape`).
- Test that the shape's text is updated by the Gemini synthesis pipeline (with a
  mocked Gemini API call).
- Confirm that error handling populates the error message into the text box and
  draws a red outline when synthesis fails.
