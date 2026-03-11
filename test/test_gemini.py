"""
Tests for Gemini LLM synthesis feature.

All Gemini API calls are mocked — no live network calls are made.

The ``contexts`` field in ``GeminiConfig`` is a list of **meta_name strings**.
These names reference meta-look shapes defined elsewhere in the same presentation;
their data is already pre-fetched by the regular Looker query pipeline and stored
in ``Cli.data`` keyed by the meta_name.
"""

import json
import os
import pytest
from unittest.mock import MagicMock, patch
from pydantic import ValidationError
from pptx import Presentation
from pptx.util import Inches

from looker_powerpoint.models import (
    GeminiConfig,
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
    """Build a minimal json_bi result string (simulates a meta-look result)."""
    return json.dumps(
        {
            "metadata": {
                "fields": {
                    "dimensions": [
                        {"name": "view.metric", "field_group_variant": "metric"}
                    ],
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
# Model validation — GeminiConfig
# ---------------------------------------------------------------------------


class TestGeminiConfig:
    def test_defaults(self):
        cfg = GeminiConfig()
        assert cfg.type == "gemini"
        assert cfg.prompt is None
        assert cfg.contexts == []
        assert cfg.model == "gemini-2.0-flash"

    def test_contexts_are_strings(self):
        """contexts must be a list of meta_name strings, not nested objects."""
        cfg = GeminiConfig(contexts=["sales_data", "revenue_data"])
        assert cfg.contexts == ["sales_data", "revenue_data"]

    def test_type_must_be_gemini(self):
        with pytest.raises(ValidationError):
            GeminiConfig(type="looker")

    def test_custom_model(self):
        cfg = GeminiConfig(model="gemini-1.5-pro")
        assert cfg.model == "gemini-1.5-pro"

    def test_prompt_stored(self):
        cfg = GeminiConfig(prompt="Summarize the key metric.")
        assert cfg.prompt == "Summarize the key metric."

    def test_single_context(self):
        cfg = GeminiConfig(contexts=["my_meta_look"])
        assert len(cfg.contexts) == 1
        assert cfg.contexts[0] == "my_meta_look"


# ---------------------------------------------------------------------------
# Model validation — GeminiShape
# ---------------------------------------------------------------------------


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

    def test_contexts_are_plain_strings(self):
        shape = GeminiShape(
            shape_id="0,3",
            shape_type="TEXT_BOX",
            slide_number=0,
            shape_number=3,
            integration={"type": "gemini", "contexts": ["kpi_data", "trend_data"]},
        )
        assert shape.integration.contexts == ["kpi_data", "trend_data"]


# ---------------------------------------------------------------------------
# Alt-text parsing from fixture
# ---------------------------------------------------------------------------


class TestGeminiShapeParsing:
    def test_fixture_parsed_as_gemini(self):
        """The gemini_textbox fixture must yield a GeminiShape, not a LookerShape."""
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        assert len(refs) == 1
        ref = refs[0]
        integration = ref.get("integration", {})
        assert integration.get("type") == "gemini"

        shape = GeminiShape.model_validate(ref)
        assert shape.shape_type == "TEXT_BOX"
        assert shape.integration.prompt == "Summarize the key trends from the data."
        assert shape.integration.contexts == ["sales_data"]

    def test_contexts_are_strings_not_dicts(self):
        """Parsed contexts must be plain strings (meta_name references)."""
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        shape = GeminiShape.model_validate(refs[0])
        for ctx in shape.integration.contexts:
            assert isinstance(ctx, str), f"Expected str, got {type(ctx)}"

    def test_fixture_not_parsed_as_looker_shape(self):
        """A Gemini shape must not accidentally validate as a LookerShape."""
        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        with pytest.raises(ValidationError):
            LookerShape.model_validate(refs[0])


# ---------------------------------------------------------------------------
# CLI parsing of Gemini shapes
# ---------------------------------------------------------------------------


class TestCliGeminiShapeParsing:
    def test_gemini_shape_collected_by_cli(self):
        """Shapes with type:gemini are collected into cli.gemini_shapes."""
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        for ref in refs:
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                gemini_shape = GeminiShape.model_validate(ref)
                if gemini_shape.shape_type in ("TEXT_BOX", "TITLE", "AUTO_SHAPE"):
                    cli.gemini_shapes.append(gemini_shape)

        assert len(cli.gemini_shapes) == 1
        assert cli.gemini_shapes[0].integration.type == "gemini"
        assert cli.gemini_shapes[0].integration.contexts == ["sales_data"]

    def test_non_textbox_gemini_shape_warns(self, caplog):
        """A Gemini config on a non-text shape must log a warning and be skipped."""
        import logging

        cli = _make_cli()
        cli.args = cli.parser.parse_args([])

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
                "contexts": ["my_meta"],
            },
        }

        with caplog.at_level(logging.WARNING):
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                gemini_shape = GeminiShape.model_validate(ref)
                if gemini_shape.shape_type not in ("TEXT_BOX", "TITLE", "AUTO_SHAPE"):
                    import logging as lg
                    lg.warning(
                        f"Gemini synthesis config found on shape "
                        f"{gemini_shape.shape_id} (type: {gemini_shape.shape_type}). "
                        "Gemini synthesis only works for text boxes. This shape will be skipped."
                    )
                else:
                    cli.gemini_shapes.append(gemini_shape)

        assert len(cli.gemini_shapes) == 0
        assert any("Gemini" in r.message and "skipped" in r.message for r in caplog.records)


# ---------------------------------------------------------------------------
# gemini module availability guard
# ---------------------------------------------------------------------------


class TestGeminiModuleAvailability:
    def test_is_available_returns_bool(self):
        assert isinstance(gemini_module.is_available(), bool)

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
# _process_gemini_shapes — context data looked up via meta_name from self.data
# ---------------------------------------------------------------------------


class TestProcessGeminiShapes:
    def _make_cli_with_gemini_shape(self):
        """
        Set up a Cli with a loaded presentation and one Gemini shape.
        Pre-seeds ``cli.data`` with a meta-look result keyed by the meta_name
        referenced in the fixture's ``contexts`` list ('sales_data').
        """
        cli = _make_cli()
        cli.args = cli.parser.parse_args([])
        cli.presentation = Presentation(_FIXTURE_PATH)

        refs = get_presentation_objects_with_descriptions(_FIXTURE_PATH)
        for ref in refs:
            integration = ref.get("integration", {})
            if isinstance(integration, dict) and integration.get("type") == "gemini":
                cli.gemini_shapes.append(GeminiShape.model_validate(ref))

        # Pre-populate data keyed by meta_name — simulates meta-look pre-fetch
        cli.data["sales_data"] = _simple_looker_result()
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

    def test_context_meta_name_passed_to_synthesize(self, monkeypatch):
        """The context data for 'sales_data' (the meta_name) is included in the call."""
        cli = self._make_cli_with_gemini_shape()
        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)

        captured = {}

        def fake_synthesize(**kw):
            captured.update(kw)
            return "ok"

        monkeypatch.setattr(gemini_module, "synthesize", fake_synthesize)
        cli._process_gemini_shapes()

        # The context string should mention the meta_name label
        assert "sales_data" in captured.get("context_data_str", "")

    def test_missing_meta_name_warns_but_continues(self, monkeypatch, caplog):
        """If a referenced meta_name has no data, a warning is logged and synthesis proceeds."""
        import logging

        cli = self._make_cli_with_gemini_shape()
        # Remove the pre-seeded data to simulate a missing meta-look
        del cli.data["sales_data"]

        monkeypatch.setattr(gemini_module, "_HAS_GEMINI", True)
        monkeypatch.setattr(gemini_module, "synthesize", lambda **kw: "result")

        with caplog.at_level(logging.WARNING):
            cli._process_gemini_shapes()

        assert any("sales_data" in r.message for r in caplog.records)

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

        assert len(slide.shapes) > shape_count_before

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

        assert len(slide.shapes) == shape_count_before

    def test_process_skips_all_when_gemini_unavailable(self, monkeypatch):
        """If google-generativeai is absent, all Gemini shapes are skipped unchanged."""
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
# _format_context_data
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
