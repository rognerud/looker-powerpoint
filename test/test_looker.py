"""Unit tests for looker_powerpoint.looker (LookerClient).

All tests are fully isolated – no live Looker API calls are made.
The Looker SDK is mocked at the module level so that LookerClient can be
instantiated without any environment variables or network access.
"""

import asyncio
import json
import logging
import sys

import pytest
from unittest.mock import AsyncMock, MagicMock, call, patch

import looker_sdk
from looker_sdk import models40 as models

from looker_powerpoint.looker import LookerClient


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_client() -> LookerClient:
    """Return a LookerClient with the Looker SDK fully mocked out."""
    with (
        patch("looker_powerpoint.looker.load_dotenv"),
        patch("looker_powerpoint.looker.find_dotenv", return_value=""),
        patch("looker_powerpoint.looker.looker_sdk.init40") as mock_init,
    ):
        mock_sdk = MagicMock()
        mock_init.return_value = mock_sdk
        client = LookerClient()
    # client.client is already set to mock_sdk
    return client


def _make_mock_look(filters=None, sorts=None, pivots=None):
    """Return a mock Look object whose .query attribute has sensible defaults."""
    mock_look = MagicMock()
    mock_query = MagicMock()
    mock_query.model = "test_model"
    mock_query.view = "test_view"
    mock_query.fields = ["field1", "field2"]
    mock_query.pivots = pivots if pivots is not None else []
    mock_query.fill_fields = None
    mock_query.filters = filters if filters is not None else {"date": "7 days"}
    mock_query.sorts = sorts if sorts is not None else ["field1"]
    mock_query.limit = "100"
    mock_query.column_limit = None
    mock_query.total = False
    mock_query.row_total = None
    mock_query.subtotals = None
    mock_query.dynamic_fields = None
    mock_query.query_timezone = None
    mock_query.vis_config = None
    mock_query.visible_ui_sections = None
    mock_look.query = mock_query
    return mock_look


def _json_bi_result(rows=None, sorts=None, pivots=None):
    """Build a minimal json_bi result string."""
    return json.dumps(
        {
            "metadata": {"fields": {"dimensions": [], "measures": []}},
            "rows": rows or [],
            "custom_sorts": sorts or [],
            "custom_pivots": pivots or [],
        }
    )


# ---------------------------------------------------------------------------
# __init__
# ---------------------------------------------------------------------------


class TestLookerClientInit:
    def test_init_success(self):
        """Happy path: SDK initialises without error."""
        with (
            patch("looker_powerpoint.looker.load_dotenv"),
            patch("looker_powerpoint.looker.find_dotenv", return_value=""),
            patch("looker_powerpoint.looker.looker_sdk.init40") as mock_init,
        ):
            mock_sdk = MagicMock()
            mock_init.return_value = mock_sdk
            client = LookerClient()
        assert client.client is mock_sdk

    def test_init_sdk_error_exits(self):
        """SDKError during init should call exit(1) → SystemExit."""
        with (
            patch("looker_powerpoint.looker.load_dotenv"),
            patch("looker_powerpoint.looker.find_dotenv", return_value=""),
            patch(
                "looker_powerpoint.looker.looker_sdk.init40",
                side_effect=looker_sdk.error.SDKError("bad"),
            ),
        ):
            with pytest.raises(SystemExit):
                LookerClient()


# ---------------------------------------------------------------------------
# run_query
# ---------------------------------------------------------------------------


class TestRunQuery:
    def test_run_query_delegates_to_sdk(self):
        """run_query passes all fields from the query_object to run_inline_query."""
        client = _make_client()
        mock_body = MagicMock()
        mock_response = '{"rows":[]}'
        client.client.run_inline_query.return_value = mock_response

        query_obj = {
            "result_format": "json_bi",
            "body": mock_body,
            "apply_vis": True,
            "apply_formatting": True,
            "server_table_calcs": True,
        }
        result = asyncio.run(client.run_query(query_obj))

        assert result == mock_response
        client.client.run_inline_query.assert_called_once_with(
            result_format="json_bi",
            body=mock_body,
            apply_vis=True,
            apply_formatting=True,
            server_table_calcs=True,
        )

    def test_run_query_returns_sdk_response(self):
        """run_query returns exactly what the SDK returns."""
        client = _make_client()
        client.client.run_inline_query.return_value = "some bytes"
        query_obj = {
            "result_format": "png",
            "body": MagicMock(),
            "apply_vis": False,
            "apply_formatting": False,
            "server_table_calcs": False,
        }
        result = asyncio.run(client.run_query(query_obj))
        assert result == "some bytes"


# ---------------------------------------------------------------------------
# make_query
# ---------------------------------------------------------------------------


class TestMakeQuery:
    def test_look_fetch_fails_returns_none(self):
        """If client.look() raises, make_query returns {shape_id: None}."""
        client = _make_client()
        client.client.look.side_effect = Exception("not found")

        result = asyncio.run(client.make_query(shape_id=1, id=999))
        assert result == {1: None}

    def test_basic_query_no_filters(self):
        """Happy path: no filter args, result is returned with injected sorts/pivots."""
        client = _make_client()
        mock_look = _make_mock_look(
            filters={"date": "7 days"}, sorts=["field1"], pivots=[]
        )
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        result = asyncio.run(client.make_query(shape_id=1, id=42))
        assert 1 in result
        payload = json.loads(result[1])
        assert payload["custom_sorts"] == ["field1"]
        assert payload["custom_pivots"] == []

    def test_result_format_not_json_skips_injection(self):
        """Non-json/json_bi formats are returned as-is (no sort/pivot injection)."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = b"\x89PNG..."

        result = asyncio.run(client.make_query(shape_id=2, id=42, result_format="png"))
        assert result[2] == b"\x89PNG..."

    def test_json_result_is_list_not_injected(self):
        """If the JSON result is a list (not a dict), sorts/pivots are not injected."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        raw = json.dumps([{"row": 1}])
        client.client.run_inline_query.return_value = raw

        result = asyncio.run(client.make_query(shape_id=3, id=42))
        assert result[3] == raw

    def test_filter_overwrite_applies_existing_filter(self):
        """filter_overwrites replaces an existing filter value."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days", "status": "all"})
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        asyncio.run(
            client.make_query(
                shape_id=1,
                id=42,
                filter_overwrites={"date": "30 days"},
            )
        )
        assert mock_look.query.filters["date"] == "30 days"

    def test_filter_overwrite_unknown_filter_warns(self, caplog):
        """filter_overwrites with an unknown key logs a warning."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days"})
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        with caplog.at_level(logging.WARNING):
            asyncio.run(
                client.make_query(
                    shape_id=1,
                    id=42,
                    filter_overwrites={"unknown_filter": "value"},
                )
            )
        assert any("not found" in r.message.lower() for r in caplog.records)

    def test_single_filter_applies_to_existing_key(self):
        """filter + filter_value sets the matching filter."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days"})
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        asyncio.run(
            client.make_query(
                shape_id=1,
                id=42,
                filter="date",
                filter_value="90 days",
            )
        )
        assert mock_look.query.filters["date"] == "90 days"

    def test_single_filter_unknown_key_warns(self, caplog):
        """filter + filter_value with an unknown filter key logs a warning."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days"})
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        with caplog.at_level(logging.WARNING):
            asyncio.run(
                client.make_query(
                    shape_id=1,
                    id=42,
                    filter="nonexistent",
                    filter_value="value",
                )
            )
        assert any("not found" in r.message.lower() for r in caplog.records)

    def test_sdk_error_during_run_returns_none(self, caplog):
        """SDKError while running the query returns {shape_id: None}."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        client.client.run_inline_query.side_effect = looker_sdk.error.SDKError("boom")

        with caplog.at_level(logging.ERROR):
            result = asyncio.run(client.make_query(shape_id=5, id=42))
        assert result == {5: None}

    def test_unexpected_error_during_run_returns_none(self, caplog):
        """Unexpected exceptions while running the query return {shape_id: None}."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        client.client.run_inline_query.side_effect = RuntimeError("unexpected")

        with caplog.at_level(logging.ERROR):
            result = asyncio.run(client.make_query(shape_id=6, id=42))
        assert result == {6: None}

    def test_result_none_returned_as_none(self):
        """If run_inline_query returns None the result dict maps shape_id to None."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = None

        result = asyncio.run(client.make_query(shape_id=7, id=42))
        assert result == {7: None}

    def test_json_decode_error_returns_raw(self, caplog):
        """If json.loads fails on the result, the raw value is returned unchanged."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        # Return invalid JSON so json.loads raises JSONDecodeError
        client.client.run_inline_query.return_value = "not-valid-json"

        with caplog.at_level(logging.WARNING):
            result = asyncio.run(client.make_query(shape_id=8, id=42))
        # The raw string should be returned without raising
        assert result[8] == "not-valid-json"

    def test_kwargs_applied_to_query(self):
        """Extra kwargs that match query attributes are applied to the query object."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        asyncio.run(
            client.make_query(
                shape_id=1,
                id=42,
                limit="500",
            )
        )
        # limit is a plain string attribute, so setattr should have been called
        assert mock_look.query.limit == "500"

    def test_list_attribute_appended_not_overwritten(self):
        """When the existing attribute is a list, the kwarg value is appended."""
        client = _make_client()
        mock_look = _make_mock_look(sorts=["field1"])
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        asyncio.run(
            client.make_query(
                shape_id=1,
                id=42,
                sorts="field2",
            )
        )
        assert "field2" in mock_look.query.sorts

    def test_apply_vis_and_formatting_flags_passed(self):
        """apply_vis, apply_formatting, server_table_calcs kwargs are forwarded to run_inline_query."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        asyncio.run(
            client.make_query(
                shape_id=1,
                id=42,
                apply_vis=True,
                apply_formatting=True,
                server_table_calcs=True,
            )
        )
        _, call_kwargs = client.client.run_inline_query.call_args
        assert call_kwargs["apply_vis"] is True
        assert call_kwargs["apply_formatting"] is True
        assert call_kwargs["server_table_calcs"] is True

    def test_pivots_injected_into_json_result(self):
        """custom_pivots is populated from the query's pivots field."""
        client = _make_client()
        mock_look = _make_mock_look(pivots=["dim1", "dim2"])
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        result = asyncio.run(client.make_query(shape_id=1, id=42))
        payload = json.loads(result[1])
        assert payload["custom_pivots"] == ["dim1", "dim2"]

    def test_none_sorts_yields_empty_list(self):
        """When q.sorts is None, custom_sorts is an empty list."""
        client = _make_client()
        mock_look = _make_mock_look(sorts=None, pivots=None)
        mock_look.query.sorts = None
        mock_look.query.pivots = None
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        result = asyncio.run(client.make_query(shape_id=1, id=42))
        payload = json.loads(result[1])
        assert payload["custom_sorts"] == []
        assert payload["custom_pivots"] == []


# ---------------------------------------------------------------------------
# _async_write_queries
# ---------------------------------------------------------------------------


class TestAsyncWriteQueries:
    def test_delegates_to_make_query(self):
        """_async_write_queries is a thin wrapper around make_query."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        raw = json.dumps({"metadata": {}, "rows": []})
        client.client.run_inline_query.return_value = raw

        result = asyncio.run(
            client._async_write_queries(
                shape_id=10,
                filter_value="30 days",
                id=99,
            )
        )
        assert 10 in result

    def test_filter_value_propagated(self):
        """filter_value kwarg is propagated from _async_write_queries to make_query."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days"})
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        asyncio.run(
            client._async_write_queries(
                shape_id=11,
                filter_value="90 days",
                filter="date",
                id=99,
            )
        )
        assert mock_look.query.filters["date"] == "90 days"
