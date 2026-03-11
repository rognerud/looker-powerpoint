# TABLE 7x7

Contains one slide, with a table 7x7 with yml: id: 1
Useful for testing table cases.

## Structure

| Property | Value |
|----------|-------|
| Slides   | 1     |
| Shapes with alt text | 1 |
| Shape type | TABLE |
| Table rows | 7 |
| Table columns | 7 |
| Shape width (px) | 853 |
| Shape height (px) | 273 |
| Shape number (pptx id) | 4 |
| Slide index | 0 |

## YAML alt text

```yaml
id: 1
```

## Expected extraction result

`get_presentation_objects_with_descriptions` returns one entry:

```python
{
    "shape_id": "0,4",
    "shape_type": "TABLE",
    "shape_width": 853,
    "shape_height": 273,
    "integration": {"id": 1},
    "slide_number": 0,
    "shape_number": 4,
}
```
