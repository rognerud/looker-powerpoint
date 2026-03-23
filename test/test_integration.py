"""Integration test for the full CLI pipeline.

Exercises the end-to-end flow:
  read pptx  →  mock Looker API  →  run cli.run()  →  verify filled output pptx
"""

import argparse
import asyncio
import json
import os
from types import SimpleNamespace
from unittest.mock import AsyncMock, MagicMock, patch

from pptx import Presentation

from looker_powerpoint.cli import Cli
from looker_powerpoint.looker import LookerClient

# ── constants ─────────────────────────────────────────────────────────────────

PPTX_PATH = os.path.join(os.path.dirname(__file__), "pptx", "table7x7.pptx")
# Shape ID returned by get_presentation_objects_with_descriptions for the
# single TABLE shape in table7x7.pptx (slide index 0, pptx shape_id 4).
TABLE_SHAPE_ID = "0,4"


# ── helpers ───────────────────────────────────────────────────────────────────


def _json_bi(dimensions, measures, table_calculations, rows):
    """Build a minimal json_bi-format payload (same structure as the helper in test_cli.py)."""

    def _field(name):
        return {"name": name, "field_group_variant": name.split(".")[-1]}

    return json.dumps(
        {
            "metadata": {
                "fields": {
                    "dimensions": [_field(d) for d in dimensions],
                    "measures": [_field(m) for m in measures],
                    "table_calculations": [
                        _field(t) for t in (table_calculations or [])
                    ],
                }
            },
            "rows": rows,
            "custom_sorts": [],
            "custom_pivots": [],
        }
    )


def _make_args(pptx_path, output_dir):
    """Return an `argparse.Namespace` that matches every attribute read by `Cli.run`."""
    ns = argparse.Namespace(
        file_path=pptx_path,
        output_dir=output_dir,
        add_links=False,
        hide_errors=True,
        parse_date_syntax_in_filename=False,
        quiet=True,
        filter=None,
        debug_queries=False,
        verbose=0,
    )
    # "self" is not a Python keyword, but using it as a kwarg looks odd; setattr is cleaner.
    setattr(ns, "self", False)
    return ns


# ── integration test ──────────────────────────────────────────────────────────


class TestIntegration:
    """End-to-end integration tests: pptx → mocked Looker → filled pptx."""

    def test_run_fills_table_from_mocked_looker(self, tmp_path):
        """Full pipeline: parse table7x7.pptx, return mock Looker data, verify output table."""
        mock_result = _json_bi(
            dimensions=["orders.date", "orders.status"],
            measures=["orders.revenue", "orders.count"],
            table_calculations=[],
            rows=[
                {
                    "orders.date.value": "2024-01-01",
                    "orders.status.value": "complete",
                    "orders.revenue.value": "100",
                    "orders.count.value": "5",
                },
                {
                    "orders.date.value": "2024-01-02",
                    "orders.status.value": "pending",
                    "orders.revenue.value": "200",
                    "orders.count.value": "10",
                },
            ],
        )

        args = _make_args(PPTX_PATH, str(tmp_path))

        cli = Cli()
        # Override parse_args so pytest's own argv does not clash with the CLI parser.
        cli.parser.parse_args = lambda: args

        mock_client = MagicMock()
        mock_client._async_write_queries = AsyncMock(
            return_value={TABLE_SHAPE_ID: mock_result}
        )

        with patch("looker_powerpoint.cli.LookerClient", return_value=mock_client):
            cli.run()

        # ── verify output file ────────────────────────────────────────────────
        output_files = list(tmp_path.glob("*.pptx"))
        assert len(output_files) == 1, "Expected exactly one output pptx file"

        prs = Presentation(str(output_files[0]))
        table = next(s.table for s in prs.slides[0].shapes if s.has_table)

        # Header row reflects the field_group_variant names (part after last dot).
        assert table.cell(0, 0).text == "date"
        assert table.cell(0, 1).text == "status"
        assert table.cell(0, 2).text == "revenue"
        assert table.cell(0, 3).text == "count"

        # First data row
        assert table.cell(1, 0).text == "2024-01-01"
        assert table.cell(1, 1).text == "complete"
        assert table.cell(1, 2).text == "100"
        assert table.cell(1, 3).text == "5"

        # Second data row
        assert table.cell(2, 0).text == "2024-01-02"
        assert table.cell(2, 1).text == "pending"
        assert table.cell(2, 2).text == "200"
        assert table.cell(2, 3).text == "10"


# ── make_query unit tests ─────────────────────────────────────────────────────


def _make_mock_look(existing_filters=None):
    """Build a minimal mock Look whose .query.filters matches *existing_filters*."""
    query = SimpleNamespace(
        model="my_model",
        view="my_view",
        fields=["orders.status", "orders.count"],
        pivots=None,
        fill_fields=None,
        filters=existing_filters if existing_filters is not None else {},
        sorts=None,
        limit="500",
        column_limit=None,
        total=None,
        row_total=None,
        subtotals=None,
        dynamic_fields=None,
        query_timezone=None,
        vis_config=None,
        visible_ui_sections=None,
    )
    look = SimpleNamespace(query=query)
    return look


class TestMakeQueryFilterOverwrites:
    """Unit tests for the filter_overwrites behaviour inside LookerClient.make_query."""

    def _run(self, coro):
        """Run a coroutine synchronously."""
        return asyncio.run(coro)

    def _make_client_with_look(self, look):
        """Return a LookerClient whose underlying SDK is fully mocked."""
        with patch("looker_powerpoint.looker.looker_sdk"):
            client = LookerClient.__new__(LookerClient)
            sdk_mock = MagicMock()
            sdk_mock.look.return_value = look
            # run_inline_query returns minimal JSON so result parsing doesn't fail
            sdk_mock.run_inline_query.return_value = json.dumps({"rows": []})
            client.client = sdk_mock
            return client

    def test_filter_overwrites_overwrites_existing_filter(self):
        """filter_overwrites replaces an existing filter value."""
        look = _make_mock_look({"orders.status": "pending"})
        client = self._make_client_with_look(look)
        result = self._run(
            client.make_query(
                shape_id="s1",
                id=1,
                filter_overwrites={"orders.status": "complete"},
            )
        )
        # Result dict maps shape_id to the query result (non-None means query succeeded)
        assert "s1" in result
        assert result["s1"] is not None
        assert look.query.filters["orders.status"] == "complete"

    def test_filter_overwrites_adds_new_filter_key(self):
        """filter_overwrites can add a filter that was not in the original query."""
        look = _make_mock_look({"orders.status": "pending"})
        client = self._make_client_with_look(look)
        self._run(
            client.make_query(
                shape_id="s2",
                id=1,
                filter_overwrites={"orders.region": "EMEA"},
            )
        )
        assert look.query.filters["orders.region"] == "EMEA"
        # Existing filter is untouched
        assert look.query.filters["orders.status"] == "pending"

    def test_filter_overwrites_handles_none_filters_dict(self):
        """filter_overwrites works even when the query has no filters (None)."""
        look = _make_mock_look(None)
        client = self._make_client_with_look(look)
        self._run(
            client.make_query(
                shape_id="s3",
                id=1,
                filter_overwrites={"orders.status": "complete"},
            )
        )
        assert look.query.filters["orders.status"] == "complete"

    def test_filter_overwrites_comma_separated_value_applied(self):
        """Comma-separated multi-value strings (from list normalization) are applied correctly."""
        look = _make_mock_look({"orders.status": "pending"})
        client = self._make_client_with_look(look)
        self._run(
            client.make_query(
                shape_id="s4",
                id=1,
                # Value is already normalised to a CSV string by the model validator
                filter_overwrites={"orders.status": "complete,pending,shipped"},
            )
        )
        assert look.query.filters["orders.status"] == "complete,pending,shipped"
