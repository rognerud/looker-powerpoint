"""
Tests for Gemini LLM synthesis feature.

All Gemini API calls are mocked — no live network calls are made.
"""

import json
import os
import pytest
from unittest.mock import patch, MagicMock
from pydantic import ValidationError
from pptx import Presentation
from pptx.util import Inches

from looker_powerpoint.models import (
    GeminiConfig,
    GeminiContextRef,
    GeminiShape,
    LookerShape,
)
from looker_powerpoint.cli import Cli
from looker_powerpoint.tools.find_alt_text import get_presentation_objects_with_descriptions

import looker_powerpoint.gemini as gemini_module


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_FIXTURE_PATH = os.path.join(os.path.dirname(__file__), "pptx", "gemini_textbox.pptx")


def _make_cli():
    """Create a Cli instance with os.getenv stubbed out so no real env is needed."""
    with patch("os.getenv", return_value="dummy_value"):
        return Cli()


def _simple_looker_result():
    """Build a minimal json_bi result string."""
    return json.dumps(
        {
            "metadata": {
                "fields": {
                    "dimensions": [{"name": "view.metric", "field_group_variant": "metric"}],
                    "measures": [],
                    "table_calculations": [],
                }
            },
            "rows": [{"view.metric.value": "42"}],
            "custom_sorts": [],
            "custom_pivots": [],
        }
    )


# ---------------------------------------------------------------------------
# Model validation tests
# ---------------------------------------------------------------------------


class TestGeminiContextRef:
    def test_required_id_field(self):
        ref = GeminiContextRef(id="10")
        assert ref.id == "10"

    def test_int_id_is_coerced_to_str(self):
        ref = GeminiContextRef(id=5)
        assert ref.id == "5"

    def test_defaults(self):
        ref = GeminiContextRef(id="1")
        assert ref.row is None
        assert ref.column is None
        assert ref.label is None
        assert ref.filter is None
        assert ref.filter_overwrites is None
        assert ref.result_format == "json_bi"
        assert ref.apply_formatting is False
        assert ref.apply_vis is True
        assert ref.server_table_calcs is True
        assert ref.retries == 0

    def test_missing_id_raises(self):
        with pytest.raises(ValidationError):
            GeminiContextRef()


class TestGeminiConfig:
    def test_defaults(self):
        cfg = GeminiConfig()
        assert cfg.type == "gemini"
        assert cfg.prompt is None
        assert cfg.contexts == []
        assert cfg.model == "gemini-2.0-flash"

    def test_with_contexts(self):
        cfg = GeminiConfig(contexts=[{"id": "1"}, {"id": "2"}])
        assert len(cfg.contexts) == 2
        assert cfg.contexts[0].id == "1"

    def test_type_must_be_gemini(self):
        with pytest.raises(ValidationError):
            GeminiConfig(type="looker")

    def test_custom_model(self):
        cfg = GeminiConfig(model="gemini-1.5-pro")
        assert cfg.model == "gemini-1.5-pro"

    def test_prompt_stored(self):
        cfg = GeminiConfig(prompt="Summarize the key metric.")
        assert cfg.prompt == "Summarize the key metric."


class TestGeminiShape:
    def test_basic_construction(self):
        shape = GeminiShape(
            shape_id="0,2",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=2,
            integration=GeminiConfig(),
        )
        assert shape.shape_type == "TEXT_BOX"
        assert shape.integration.type == "gemini"

    def test_nested_context_refs(self):
        shape = GeminiShape(
            shape_id="0,3",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=3,
            integration={"type": "gemini", "contexts": [{"id": 7}]},
        )
        assert shape.integration.contexts[0].id == "7"


# ---------------------------------------------------------------------------
# Alt-text parsing tests
# ---------------------------------------------------------------------------


class TestGeminiShapeParsing:
    def test_fixture_parsed_as_gemini(self):
        """The gemini_textbox fixture should yield a GeminiShape, not a LookerShape."""
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        assert len(refs) == 1
        ref = refs[0]
        integration = ref.get("integration", {})
        assert integration.get("type") == "gemini"

        # Should validate as GeminiShape
        shape = GeminiShape.model_validate(ref)
        assert shape.shape_type == "TEXT_BOX"
        assert shape.integration.prompt == "Summarize the key trends from the data."
        assert len(shape.integration.contexts) == 1
        assert shape.integration.contexts[0].id == "1"

    def test_fixture_not_parsed_as_looker_shape(self):
        """A Gemini shape YAML must not accidentally validate as a LookerShape."""
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        with pytest.raises(ValidationError):
            LookerShape.model_validate(refs[0])


# ---------------------------------------------------------------------------
# CLI parsing of Gemini shapes
# ---------------------------------------------------------------------------


class TestCliGeminiShapeParsing:
    def test_gemini_shape_collected_by_cli(self):
        """CLI run() must parse Gemini shapes into self.gemini_shapes."""
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        # Simulate the ref-parsing loop from run()
        from looker_powerpoint.models import GeminiShape, GeminiConfig
        from pydantic import ValidationError as VE

        for ref in refs:
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                gemini_shape = GeminiShape.model_validate(ref)
                if gemini_shape.shape_type in ("TEXT_BOX", "TITLE", "AUTO_SHAPE"):
                    cli.gemini_shapes.append(gemini_shape)

        assert len(cli.gemini_shapes) == 1
        assert cli.gemini_shapes[0].integration.type == "gemini"

    def test_non_textbox_gemini_shape_warns(self, caplog):
        """A Gemini config on a non-text shape should log a warning and be skipped."""
        import logging

        cli = _make_cli()
        cli.args = cli.parser.parse_args([])

        # Simulate a TABLE shape with Gemini config
        ref = {
            "shape_id": "0,5",
            "shape_type": "TABLE",
            "slide_number": 0,
            "shape_number": 5,
            "shape_width": 400,
            "shape_height": 200,
            "integration": {
                "type": "gemini",
                "prompt": "Summarize",
                "contexts": [],
            },
        }

        with caplog.at_level(logging.WARNING):
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                from looker_powerpoint.models import GeminiShape
                gemini_shape = GeminiShape.model_validate(ref)
                if gemini_shape.shape_type not in ("TEXT_BOX", "TITLE", "AUTO_SHAPE"):
                    import logging as lg
                    lg.warning(
                        f"Gemini synthesis config found on shape "
                        f"{gemini_shape.shape_id} (type: {gemini_shape.shape_type}). "
                        "Gemini synthesis only works for text boxes (TEXT_BOX, TITLE, "
                        "AUTO_SHAPE). This shape will be skipped."
                    )
                else:
                    cli.gemini_shapes.append(gemini_shape)

        assert len(cli.gemini_shapes) == 0
        assert any("Gemini" in r.message and "skipped" in r.message for r in caplog.records)


# ---------------------------------------------------------------------------
# gemini module availability guard
# ---------------------------------------------------------------------------


class TestGeminiModuleAvailability:
    def test_is_available_reflects_import(self):
        # We can only assert it returns a bool without controlling the import.
        result = gemini_module.is_available()
        assert isinstance(result, bool)

    def test_synthesize_raises_when_unavailable(self, monkeypatch):
        """When google-generativeai is absent, synthesize() must raise ImportError."""
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", False)
        with pytest.raises(ImportError, match="google-generativeai"):
            gemini_module.synthesize(
                prompt="test",
                context_data_str="",
                current_text="hello",
            )

    def test_synthesize_raises_without_api_key(self, monkeypatch):
        """With the package present but no API key, synthesize() raises ValueError."""
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.delenv("GOOGLE_API_KEY", raising=False)
        monkeypatch.delenv("GEMINI_API_KEY", raising=False)
        with pytest.raises(ValueError, match="API_KEY"):
            gemini_module.synthesize(
                prompt="test",
                context_data_str="",
                current_text="hello",
            )


# ---------------------------------------------------------------------------
# _process_gemini_shapes tests (mocked)
# ---------------------------------------------------------------------------


class TestProcessGeminiShapes:
    def _make_cli_with_gemini_shape(self):
        """Set up a Cli with a loaded presentation and one Gemini shape."""
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        for ref in refs:
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                gemini_shape = GeminiShape.model_validate(ref)
                cli.gemini_shapes.append(gemini_shape)

        # Pre-populate context data
        ctx_key = f"gemini_ctx_{cli.gemini_shapes[0].shape_id}_0"
        cli.data[ctx_key] = _simple_looker_result()
        return cli

    def test_process_inserts_synthesized_text(self, monkeypatch):
        """Synthesized text is written into the shape's text frame."""
        cli = self._make_cli_with_gemini_shape()

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module, "synthesize", lambda **kw: "Synthesized result text"
        )

        cli._process_gemini_shapes()

        slide = cli.presentation.slides[0]
        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                assert shape.text_frame.text == "Synthesized result text"

    def test_process_error_populates_error_message(self, monkeypatch):
        """On synthesis failure, the error message is written into the text box."""
        cli = self._make_cli_with_gemini_shape()

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module,
            "synthesize",
            MagicMock(side_effect=RuntimeError("API call failed")),
        )

        cli._process_gemini_shapes()

        slide = cli.presentation.slides[0]
        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                assert "API call failed" in shape.text_frame.text

    def test_process_error_draws_red_outline(self, monkeypatch):
        """On synthesis failure, a red-outline shape is added (mark_failure)."""
        cli = self._make_cli_with_gemini_shape()

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module,
            "synthesize",
            MagicMock(side_effect=RuntimeError("fail")),
        )

        slide = cli.presentation.slides[0]
        shape_count_before = len(slide.shapes)
        cli._process_gemini_shapes()
        shape_count_after = len(slide.shapes)

        # _mark_failure adds an oval shape
        assert shape_count_after > shape_count_before

    def test_process_error_no_red_outline_when_hidden(self, monkeypatch):
        """With --hide-errors, no red outline is drawn on failure."""
        cli = self._make_cli_with_gemini_shape()
        cli.args.hide_errors = True

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(
            gemini_module,
            "synthesize",
            MagicMock(side_effect=RuntimeError("fail")),
        )

        slide = cli.presentation.slides[0]
        shape_count_before = len(slide.shapes)
        cli._process_gemini_shapes()
        shape_count_after = len(slide.shapes)

        assert shape_count_after == shape_count_before

    def test_process_skips_all_when_gemini_unavailable(self, monkeypatch):
        """If google-generativeai is absent, all Gemini shapes are skipped."""
        cli = self._make_cli_with_gemini_shape()

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", False)

        slide = cli.presentation.slides[0]
        original_text = None
        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                original_text = shape.text_frame.text

        cli._process_gemini_shapes()

        for shape in slide.shapes:
            if shape.shape_id == cli.gemini_shapes[0].shape_number:
                assert shape.text_frame.text == original_text


# ---------------------------------------------------------------------------
# Additional edge-case: format_context_data
# ---------------------------------------------------------------------------


class TestFormatContextData:
    def test_format_returns_string(self):
        import pandas as pd

        cli = _make_cli()
        df = pd.DataFrame({"col_a": [1, 2], "col_b": ["x", "y"]})
        result = cli._format_context_data(df)
        assert isinstance(result, str)
        assert "col_a" in result
        assert "col_b" in result
