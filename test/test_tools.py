"""
Unit tests for looker_powerpoint/tools/:
  - find_alt_text.py
  - pptx_text_handler.py
  - url_to_hyperlink.py
"""
import io
import pytest
import pandas as pd
from unittest.mock import MagicMock, patch, PropertyMock
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from lxml import etree

from looker_powerpoint.tools.find_alt_text import (
    extract_alt_text,
    get_presentation_objects_with_descriptions,
)
from looker_powerpoint.tools.pptx_text_handler import (
    remove_emojis_from_string,
    sanitize_header_name,
    sanitize_dataframe_headers,
    encode_colored_text,
    decode_marked_segments,
    colorize_positive,
    make_jinja_env,
    render_text_with_jinja,
    extract_text_and_run_meta,
    copy_run_format,
    copy_font_format,
    reinsert_rendered_text_preserving_formatting,
    process_text_field,
    _START,
    _SEP,
    _END,
)
from looker_powerpoint.tools.url_to_hyperlink import add_text_with_numbered_links


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_pptx_bytes():
    """Return a minimal in-memory pptx with one blank slide."""
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[6])  # blank layout
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def _pptx_with_alt_text(descr_yaml: str):
    """
    Return an in-memory pptx with a text box whose alt-text (descr attribute)
    is set to *descr_yaml*.
    """
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(2), Inches(1))
    # Set the alt-text description directly on the cNvPr XML element
    sp_elem = txBox.element
    # Find cNvPr (non-visual properties)
    ns = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
    pptx_ns = "http://schemas.openxmlformats.org/presentationml/2006/main"
    draw_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    # Use the lxml tree to set descr
    for elem in sp_elem.iter():
        if elem.tag.endswith("}cNvPr") or elem.tag == "cNvPr":
            elem.set("descr", descr_yaml)
            break
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ---------------------------------------------------------------------------
# Tests: find_alt_text.py
# ---------------------------------------------------------------------------

class TestExtractAltText:
    """Tests for extract_alt_text()."""

    def test_returns_none_when_no_alt_text(self):
        """A shape without a descr attribute should return None."""
        buf = _make_pptx_bytes()
        prs = Presentation(buf)
        slide = prs.slides[0]
        # Add a shape without alt text
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(1), Inches(1))
        result = extract_alt_text(txBox)
        assert result is None

    def test_returns_dict_for_yaml_alt_text(self):
        """A shape with valid YAML in alt-text should return a parsed dict."""
        buf = _pptx_with_alt_text("meta_name: my_shape\nlook_id: 42")
        prs = Presentation(buf)
        slide = prs.slides[0]
        shape = slide.shapes[0]
        result = extract_alt_text(shape)
        assert isinstance(result, dict)
        assert result.get("meta_name") == "my_shape"
        assert result.get("look_id") == 42

    def test_returns_none_for_empty_descr(self):
        """A shape with an empty descr string returns None (yaml.safe_load of '' is None)."""
        buf = _pptx_with_alt_text("")
        prs = Presentation(buf)
        slide = prs.slides[0]
        shape = slide.shapes[0]
        # Empty descr yields None from yaml.safe_load; function returns None
        result = extract_alt_text(shape)
        assert result is None

    def test_returns_scalar_for_plain_string_yaml(self):
        """Alt-text containing a bare string should return that string."""
        buf = _pptx_with_alt_text("just_a_string")
        prs = Presentation(buf)
        slide = prs.slides[0]
        shape = slide.shapes[0]
        result = extract_alt_text(shape)
        assert result == "just_a_string"


class TestGetPresentationObjectsWithDescriptions:
    """Tests for get_presentation_objects_with_descriptions()."""

    def test_returns_empty_list_for_presentation_without_descriptions(self):
        buf = _make_pptx_bytes()
        # Save to a temp file path via bytes (pass BytesIO directly)
        result = get_presentation_objects_with_descriptions(buf)
        assert result == []

    def test_returns_shapes_with_descriptions(self):
        buf = _pptx_with_alt_text("meta_name: slide_shape\nlook_id: 1")
        result = get_presentation_objects_with_descriptions(buf)
        assert len(result) == 1
        entry = result[0]
        assert entry["shape_id"] == "slide_shape"
        assert entry["integration"]["look_id"] == 1
        assert entry["slide_number"] == 0

    def test_returns_empty_list_on_bad_path(self):
        result = get_presentation_objects_with_descriptions("/nonexistent/path/file.pptx")
        assert result == []

    def test_shape_id_fallback_when_no_meta_name(self):
        """When alt-text dict has no 'meta_name', shape_id is '{slide_idx},{shape_id}'."""
        buf = _pptx_with_alt_text("look_id: 99")
        result = get_presentation_objects_with_descriptions(buf)
        assert len(result) == 1
        # shape_id should be fallback "0,<shape_id>" style
        assert "," in result[0]["shape_id"]

    def test_result_contains_expected_keys(self):
        buf = _pptx_with_alt_text("meta_name: my_shape")
        result = get_presentation_objects_with_descriptions(buf)
        assert len(result) == 1
        keys = set(result[0].keys())
        assert keys == {
            "shape_id",
            "shape_type",
            "shape_width",
            "shape_height",
            "integration",
            "slide_number",
            "shape_number",
        }


# ---------------------------------------------------------------------------
# Tests: pptx_text_handler.py — pure helpers
# ---------------------------------------------------------------------------

class TestRemoveEmojisFromString:
    def test_removes_emoticon(self):
        assert remove_emojis_from_string("Hello 😀 World") == "Hello  World"

    def test_keeps_plain_text(self):
        assert remove_emojis_from_string("no emojis here") == "no emojis here"

    def test_non_string_passthrough(self):
        assert remove_emojis_from_string(42) == 42
        assert remove_emojis_from_string(None) is None

    def test_removes_flag_emoji(self):
        result = remove_emojis_from_string("flag 🇺🇸 here")
        assert "🇺🇸" not in result

    def test_empty_string(self):
        assert remove_emojis_from_string("") == ""


class TestSanitizeHeaderName:
    def test_none_returns_none(self):
        assert sanitize_header_name(None) is None

    def test_strips_whitespace(self):
        assert sanitize_header_name("  hello  ") == "hello"

    def test_replaces_inner_whitespace_with_underscore(self):
        assert sanitize_header_name("hello world") == "hello_world"

    def test_removes_emojis(self):
        result = sanitize_header_name("Revenue 💰 Total")
        assert "💰" not in result
        assert "Revenue" in result

    def test_strips_leading_trailing_underscores(self):
        # spaces become underscores; leading/trailing stripped
        assert sanitize_header_name("  hello  ") == "hello"

    def test_non_string_input(self):
        # Should convert to str first
        result = sanitize_header_name(123)
        assert result == "123"

    def test_multiple_spaces_become_single_underscore(self):
        assert sanitize_header_name("a  b") == "a_b"


class TestSanitizeDataframeHeaders:
    def test_renames_columns_with_spaces(self):
        df = pd.DataFrame(columns=["hello world", "foo bar"])
        result = sanitize_dataframe_headers(df)
        assert list(result.columns) == ["hello_world", "foo_bar"]

    def test_removes_emoji_from_columns(self):
        df = pd.DataFrame(columns=["Revenue 💰", "Cost"])
        result = sanitize_dataframe_headers(df)
        assert "💰" not in result.columns[0]

    def test_preserves_data(self):
        df = pd.DataFrame({"a b": [1, 2, 3]})
        result = sanitize_dataframe_headers(df)
        assert list(result["a_b"]) == [1, 2, 3]

    def test_no_change_needed(self):
        df = pd.DataFrame(columns=["clean", "column"])
        result = sanitize_dataframe_headers(df)
        assert list(result.columns) == ["clean", "column"]


class TestEncodeColoredText:
    def test_basic_encoding(self):
        result = encode_colored_text("hello", "#FF0000")
        assert result == f"{_START}#FF0000{_SEP}hello{_END}"

    def test_encoded_text_decoded_correctly(self):
        encoded = encode_colored_text("world", "#008000")
        segments = decode_marked_segments(encoded)
        assert segments == [("world", "#008000")]


class TestDecodeMarkedSegments:
    def test_plain_text_no_markers(self):
        segments = decode_marked_segments("plain text")
        assert segments == [("plain text", None)]

    def test_single_colored_segment(self):
        encoded = f"{_START}#FF0000{_SEP}red text{_END}"
        segments = decode_marked_segments(encoded)
        assert segments == [("red text", "#FF0000")]

    def test_mixed_plain_and_colored(self):
        encoded = f"prefix {_START}#00FF00{_SEP}green{_END} suffix"
        segments = decode_marked_segments(encoded)
        assert len(segments) == 3
        assert segments[0] == ("prefix ", None)
        assert segments[1] == ("green", "#00FF00")
        assert segments[2] == (" suffix", None)

    def test_multiple_colored_segments(self):
        a = encode_colored_text("pos", "#008000")
        b = encode_colored_text("neg", "#C00000")
        segments = decode_marked_segments(a + b)
        assert segments == [("pos", "#008000"), ("neg", "#C00000")]

    def test_empty_string(self):
        assert decode_marked_segments("") == []


class TestColorizePositive:
    def test_positive_number(self):
        result = colorize_positive(10)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"  # green

    def test_negative_number(self):
        result = colorize_positive(-5)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#C00000"  # red

    def test_zero(self):
        result = colorize_positive(0)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#000000"  # black

    def test_positive_string(self):
        result = colorize_positive("42.5")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_negative_string(self):
        result = colorize_positive("-3.14")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#C00000"

    def test_none_value(self):
        result = colorize_positive(None)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#000000"

    def test_empty_string(self):
        result = colorize_positive("")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#000000"

    def test_parenthesized_negative(self):
        # The number-stripping regexes remove the surrounding parens before the
        # explicit parenthesized-negative branch is reached, so "(100)" ends up
        # being parsed as positive 100.  This test documents the current behavior.
        # If accounting-convention parenthesized negatives need to be supported,
        # the try_parse_number logic should be adjusted.
        result = colorize_positive("(100)")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"  # treated as positive 100
    def test_custom_positive_color(self):
        result = colorize_positive(5, positive_hex="#AABBCC")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#AABBCC"

    def test_string_with_comma(self):
        """'1,000' should parse as 1000 (positive)."""
        result = colorize_positive("1,000")
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_emoji_removed_from_output(self):
        result = colorize_positive("5 😀")
        segs = decode_marked_segments(result)
        assert "😀" not in segs[0][0]


class TestRenderTextWithJinja:
    def test_simple_variable_substitution(self):
        result = render_text_with_jinja("Hello {{ name }}", {"name": "World"})
        assert result == "Hello World"

    def test_empty_context(self):
        result = render_text_with_jinja("no vars", {})
        assert result == "no vars"

    def test_colorize_positive_filter_available(self):
        env = make_jinja_env()
        result = render_text_with_jinja("{{ 5 | colorize_positive }}", {}, env=env)
        segs = decode_marked_segments(result)
        assert segs[0][1] == "#008000"

    def test_loop_rendering(self):
        result = render_text_with_jinja(
            "{% for row in header_rows %}{{ row.x }}{% endfor %}",
            {"header_rows": [{"x": "A"}, {"x": "B"}]},
        )
        assert result == "AB"


# ---------------------------------------------------------------------------
# Tests: pptx_text_handler.py — shape-based helpers (using real pptx objects)
# ---------------------------------------------------------------------------

def _text_frame_with_text(text: str):
    """Return a real pptx text_frame populated with *text*."""
    prs = Presentation()
    blank = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank)
    txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    tf = txBox.text_frame
    tf.text = text
    return tf


class TestExtractTextAndRunMeta:
    def test_extracts_text(self):
        tf = _text_frame_with_text("hello world")
        full_text, run_meta = extract_text_and_run_meta(tf)
        assert "hello world" in full_text

    def test_no_trailing_newline(self):
        tf = _text_frame_with_text("single line")
        full_text, _ = extract_text_and_run_meta(tf)
        assert not full_text.endswith("\n")

    def test_run_meta_has_correct_structure(self):
        tf = _text_frame_with_text("test")
        _, run_meta = extract_text_and_run_meta(tf)
        # Each item must have 'text' and 'run_obj' keys
        for item in run_meta:
            assert "text" in item
            assert "run_obj" in item


class TestCopyRunFormat:
    def test_copies_bold_and_italic(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        src_box = slide.shapes.add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
        dst_box = slide.shapes.add_textbox(Inches(2), Inches(0), Inches(1), Inches(1))
        src_tf = src_box.text_frame
        dst_tf = dst_box.text_frame
        src_tf.text = "source"
        dst_tf.text = "dest"
        src_run = src_tf.paragraphs[0].runs[0]
        dst_run = dst_tf.paragraphs[0].runs[0]
        src_run.font.bold = True
        src_run.font.italic = True
        copy_run_format(src_run, dst_run)
        assert dst_run.font.bold is True
        assert dst_run.font.italic is True

    def test_copies_font_size(self):
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        src_box = slide.shapes.add_textbox(Inches(0), Inches(0), Inches(1), Inches(1))
        dst_box = slide.shapes.add_textbox(Inches(2), Inches(0), Inches(1), Inches(1))
        src_tf = src_box.text_frame
        dst_tf = dst_box.text_frame
        src_tf.text = "source"
        dst_tf.text = "dest"
        src_run = src_tf.paragraphs[0].runs[0]
        dst_run = dst_tf.paragraphs[0].runs[0]
        src_run.font.size = Pt(24)
        copy_run_format(src_run, dst_run)
        assert dst_run.font.size == Pt(24)


class TestReinsertRenderedTextPreservingFormatting:
    def test_updates_text_frame_text(self):
        tf = _text_frame_with_text("original text")
        _, run_meta = extract_text_and_run_meta(tf)
        reinsert_rendered_text_preserving_formatting(tf, "new text", run_meta)
        new_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "new text" in new_text

    def test_handles_empty_run_meta(self):
        tf = _text_frame_with_text("text")
        reinsert_rendered_text_preserving_formatting(tf, "replaced", run_meta=None)
        new_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "replaced" in new_text


class TestProcessTextField:
    def test_no_jinja_tag_no_change_when_same(self):
        """If no Jinja tags and text matches, no modification occurs."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        tf = txBox.text_frame
        tf.text = "static text"
        df = pd.DataFrame()
        # text_to_insert matches; no Jinja; should be a no-op (no crash)
        process_text_field(txBox, "static text", df)
        assert tf.text == "static text"

    def test_no_jinja_tag_updates_when_different(self):
        """If no Jinja tags but text differs from current, the text is updated."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        tf = txBox.text_frame
        tf.text = "old text"
        df = pd.DataFrame()
        process_text_field(txBox, "new text", df)
        # The update should have been applied
        text_content = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "new text" in text_content

    def test_jinja_renders_template(self):
        """Jinja template in alt-text is rendered with the DataFrame context."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1))
        tf = txBox.text_frame
        tf.text = "{{ header_rows[0].value }}"
        df = pd.DataFrame({"value": ["hello"]})
        process_text_field(txBox, "ignored_text_to_insert", df)
        text_content = "".join(r.text for p in tf.paragraphs for r in p.runs)
        assert "hello" in text_content


# ---------------------------------------------------------------------------
# Tests: url_to_hyperlink.py
# ---------------------------------------------------------------------------

class TestAddTextWithNumberedLinks:
    def _fresh_text_frame(self):
        """Create a fresh in-memory presentation and return a text frame."""
        prs = Presentation()
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(5), Inches(1))
        return txBox.text_frame

    def _all_runs(self, tf):
        """Return all runs across all paragraphs in the text frame."""
        return [r for p in tf.paragraphs for r in p.runs]

    def test_plain_text_no_url(self):
        tf = self._fresh_text_frame()
        result = add_text_with_numbered_links(tf, "no url here")
        assert result == 1  # index unchanged
        runs = self._all_runs(tf)
        assert any("no url here" in r.text for r in runs)

    def test_single_url_replaced_with_numbered_reference(self):
        tf = self._fresh_text_frame()
        result = add_text_with_numbered_links(tf, "see https://example.com for details")
        assert result == 2  # one URL consumed, next index is 2
        texts = [r.text for r in self._all_runs(tf)]
        # The reference "(1)" should appear somewhere
        assert any("(1)" in t for t in texts)
        # The raw URL should not appear as a run text
        assert not any("https://example.com" == t.strip() for t in texts)

    def test_hyperlink_set_on_url_run(self):
        tf = self._fresh_text_frame()
        add_text_with_numbered_links(tf, "https://example.com")
        runs = self._all_runs(tf)
        url_runs = [r for r in runs if "(1)" in r.text]
        assert len(url_runs) == 1
        assert url_runs[0].hyperlink.address == "https://example.com"

    def test_multiple_urls(self):
        tf = self._fresh_text_frame()
        text = "first https://a.com then https://b.com"
        result = add_text_with_numbered_links(tf, text)
        assert result == 3  # two URLs
        texts = [r.text for r in self._all_runs(tf)]
        assert any("(1)" in t for t in texts)
        assert any("(2)" in t for t in texts)

    def test_start_index_respected(self):
        tf = self._fresh_text_frame()
        result = add_text_with_numbered_links(tf, "https://example.com", start_index=5)
        assert result == 6
        texts = [r.text for r in self._all_runs(tf)]
        assert any("(5)" in t for t in texts)

    def test_url_with_trailing_digits_uses_those_digits(self):
        """A URL ending with digits uses those digits as the reference number,
        and the index counter is NOT incremented (since the number came from the URL)."""
        tf = self._fresh_text_frame()
        result = add_text_with_numbered_links(tf, "https://example.com/items/42")
        texts = [r.text for r in self._all_runs(tf)]
        # The reference should use the trailing digits from the URL
        assert any("(42)" in t for t in texts)
        # index is NOT incremented when trailing digits are used
        assert result == 1

    def test_newlines_flattened_when_url_present(self):
        tf = self._fresh_text_frame()
        add_text_with_numbered_links(tf, "line1\nhttps://example.com\nline2")
        # All content should be in one paragraph (newlines flattened)
        full_text = " ".join(r.text for r in self._all_runs(tf))
        assert "\n" not in full_text

    def test_returns_start_index_when_no_urls(self):
        tf = self._fresh_text_frame()
        result = add_text_with_numbered_links(tf, "no urls", start_index=3)
        assert result == 3

    def test_url_run_is_blue_and_underlined(self):
        tf = self._fresh_text_frame()
        add_text_with_numbered_links(tf, "https://example.com")
        runs = self._all_runs(tf)
        url_runs = [r for r in runs if "(1)" in r.text]
        assert url_runs[0].font.color.rgb == RGBColor(0, 0, 255)
        assert url_runs[0].font.underline is True
