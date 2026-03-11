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
