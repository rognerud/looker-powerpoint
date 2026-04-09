"""Unit tests for looker_powerpoint.looker (LookerClient).

This suite provides:
  1. Focused unit tests for each extracted static/instance helper method,
     achieving 100% line coverage on every code path.
  2. End-to-end behavioural tests for ``make_query`` and
     ``_async_write_queries`` that exercise the full async pipeline.

All tests are fully isolated – no live Looker API calls are made.
The Looker SDK is mocked at the module level so that LookerClient can be
instantiated without any environment variables or network access.
"""

import asyncio
import json
import logging

import pytest
from unittest.mock import MagicMock, patch

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


# ---------------------------------------------------------------------------
# __init__ / _initialize_client
# ---------------------------------------------------------------------------


class TestInit:
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
        """SDKError during init should call sys.exit(1) → SystemExit."""
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

    def test_initialize_client_sets_client(self):
        """_initialize_client() can be called on an already-constructed instance."""
        client = _make_client()
        new_sdk = MagicMock()
        with patch("looker_powerpoint.looker.looker_sdk.init40", return_value=new_sdk):
            client._initialize_client()
        assert client.client is new_sdk

    def test_initialize_client_sdk_error(self):
        """_initialize_client() calls sys.exit on SDKError."""
        client = _make_client()
        with patch(
            "looker_powerpoint.looker.looker_sdk.init40",
            side_effect=looker_sdk.error.SDKError("fail"),
        ):
            with pytest.raises(SystemExit):
                client._initialize_client()


# ---------------------------------------------------------------------------
# _apply_kwargs_to_query
# ---------------------------------------------------------------------------


class TestApplyKwargsToQuery:
    def test_sets_scalar_attribute(self):
        """A scalar attribute is replaced when value is not None."""
        q = MagicMock()
        q.limit = "100"
        LookerClient._apply_kwargs_to_query(q, {"limit": "500"})
        assert q.limit == "500"

    def test_appends_to_list_attribute(self):
        """When the existing attribute is a list, the value is appended."""
        q = MagicMock()
        q.sorts = ["field1"]
        LookerClient._apply_kwargs_to_query(q, {"sorts": "field2"})
        assert "field2" in q.sorts

    def test_ignores_none_value(self):
        """kwargs with value None are skipped."""
        q = MagicMock()
        q.limit = "100"
        LookerClient._apply_kwargs_to_query(q, {"limit": None})
        assert q.limit == "100"

    def test_ignores_unknown_attribute(self):
        """kwargs for attributes that don't exist on q are silently ignored."""
        q = MagicMock(spec=[])  # no attributes
        LookerClient._apply_kwargs_to_query(q, {"nonexistent": "value"})


# ---------------------------------------------------------------------------
# _apply_filter_overwrites
# ---------------------------------------------------------------------------


class TestApplyFilterOverwrites:
    def test_overwrites_existing_filter(self):
        """An existing filter key is updated."""
        q = MagicMock()
        q.filters = {"date": "7 days", "status": "all"}
        LookerClient._apply_filter_overwrites(q, {"date": "30 days"})
        assert q.filters["date"] == "30 days"

    def test_unknown_filter_warns(self, caplog):
        """An unknown filter key triggers a WARNING."""
        q = MagicMock()
        q.filters = {"date": "7 days"}
        with caplog.at_level(logging.WARNING):
            LookerClient._apply_filter_overwrites(q, {"unknown": "value"})
        assert any("not found" in r.message.lower() for r in caplog.records)

    def test_no_filters_attribute_noop(self):
        """If q has no 'filters' attribute, the method returns without error."""
        q = MagicMock(spec=[])  # no attributes
        LookerClient._apply_filter_overwrites(q, {"date": "30 days"})

    def test_filters_is_none_noop(self):
        """If q.filters is None, the method returns without error."""
        q = MagicMock()
        q.filters = None
        LookerClient._apply_filter_overwrites(q, {"date": "30 days"})


# ---------------------------------------------------------------------------
# _apply_single_filter
# ---------------------------------------------------------------------------


class TestApplySingleFilter:
    def test_sets_existing_filter(self):
        """A filter key that exists in q.filters is updated."""
        q = MagicMock()
        q.filters = {"date": "7 days"}
        LookerClient._apply_single_filter(q, "date", "90 days")
        assert q.filters["date"] == "90 days"

    def test_unknown_filter_warns(self, caplog):
        """An unknown filter key triggers a WARNING."""
        q = MagicMock()
        q.filters = {"date": "7 days"}
        with caplog.at_level(logging.WARNING):
            LookerClient._apply_single_filter(q, "nonexistent", "value")
        assert any("not found" in r.message.lower() for r in caplog.records)

    def test_no_filters_attribute_warns(self, caplog):
        """If q has no 'filters' attribute, a warning is logged."""
        q = MagicMock(spec=[])  # no attributes
        with caplog.at_level(logging.WARNING):
            LookerClient._apply_single_filter(q, "date", "90 days")
        assert caplog.records

    def test_filters_is_none_warns(self, caplog):
        """If q.filters is None, a warning is logged."""
        q = MagicMock()
        q.filters = None
        with caplog.at_level(logging.WARNING):
            LookerClient._apply_single_filter(q, "date", "90 days")
        assert caplog.records


# ---------------------------------------------------------------------------
# _build_write_query
# ---------------------------------------------------------------------------


class TestBuildWriteQuery:
    def test_returns_write_query_with_correct_fields(self):
        """All query fields are transferred to the WriteQuery."""
        q = MagicMock()
        q.model = "m"
        q.view = "v"
        q.fields = ["f1"]
        q.pivots = []
        q.fill_fields = None
        q.filters = {"date": "7 days"}
        q.sorts = ["f1"]
        q.limit = "100"
        q.column_limit = None
        q.total = False
        q.row_total = None
        q.subtotals = None
        q.dynamic_fields = None
        q.query_timezone = "UTC"
        q.vis_config = None
        q.visible_ui_sections = None

        wq = LookerClient._build_write_query(q)
        assert isinstance(wq, models.WriteQuery)
        assert wq.model == "m"
        assert wq.view == "v"
        assert wq.query_timezone == "UTC"


# ---------------------------------------------------------------------------
# _post_process_result
# ---------------------------------------------------------------------------


class TestPostProcessResult:
    def _q(self, sorts=None, pivots=None):
        q = MagicMock()
        q.sorts = sorts
        q.pivots = pivots
        return q

    def test_injects_sorts_and_pivots(self):
        """custom_sorts and custom_pivots are added to a dict JSON result."""
        q = self._q(sorts=["field1"], pivots=["dim1"])
        raw = json.dumps({"metadata": {}, "rows": []})
        result = LookerClient._post_process_result(raw, "json_bi", q, 1, 42)
        payload = json.loads(result)
        assert payload["custom_sorts"] == ["field1"]
        assert payload["custom_pivots"] == ["dim1"]

    def test_json_format_also_injected(self):
        """'json' format also receives injection."""
        q = self._q(sorts=["s1"])
        raw = json.dumps({"rows": []})
        result = LookerClient._post_process_result(raw, "json", q, 1, 42)
        assert "custom_sorts" in json.loads(result)

    def test_non_json_format_returned_unchanged(self):
        """PNG / CSV formats are returned as-is."""
        q = self._q()
        result = LookerClient._post_process_result(b"\x89PNG", "png", q, 1, 42)
        assert result == b"\x89PNG"

    def test_none_result_returned_as_none(self):
        """A None result is returned unchanged."""
        q = self._q()
        assert LookerClient._post_process_result(None, "json_bi", q, 1, 42) is None

    def test_list_result_not_injected(self):
        """If the JSON value is a list (not a dict), result is returned unchanged."""
        q = self._q()
        raw = json.dumps([{"row": 1}])
        result = LookerClient._post_process_result(raw, "json_bi", q, 1, 42)
        assert result == raw

    def test_invalid_json_logs_warning_returns_raw(self, caplog):
        """Invalid JSON logs a warning and returns the original string."""
        q = self._q()
        with caplog.at_level(logging.WARNING):
            result = LookerClient._post_process_result("not-json", "json_bi", q, 1, 42)
        assert result == "not-json"
        assert caplog.records

    def test_none_sorts_and_pivots_yield_empty_lists(self):
        """When q.sorts / q.pivots is None, the injected lists are empty."""
        q = self._q(sorts=None, pivots=None)
        raw = json.dumps({"rows": []})
        result = LookerClient._post_process_result(raw, "json_bi", q, 1, 42)
        payload = json.loads(result)
        assert payload["custom_sorts"] == []
        assert payload["custom_pivots"] == []


# ---------------------------------------------------------------------------
# run_query
# ---------------------------------------------------------------------------


class TestRunQuery:
    def test_delegates_to_sdk(self):
        """run_query passes all fields from the query_object to run_inline_query."""
        client = _make_client()
        mock_body = MagicMock()
        client.client.run_inline_query.return_value = '{"rows":[]}'

        query_obj = {
            "result_format": "json_bi",
            "body": mock_body,
            "apply_vis": True,
            "apply_formatting": True,
            "server_table_calcs": True,
        }
        result = asyncio.run(client.run_query(query_obj))

        assert result == '{"rows":[]}'
        client.client.run_inline_query.assert_called_once_with(
            result_format="json_bi",
            body=mock_body,
            apply_vis=True,
            apply_formatting=True,
            server_table_calcs=True,
        )

    def test_returns_sdk_response(self):
        """run_query returns exactly what the SDK returns."""
        client = _make_client()
        client.client.run_inline_query.return_value = b"binary data"
        query_obj = {
            "result_format": "png",
            "body": MagicMock(),
            "apply_vis": False,
            "apply_formatting": False,
            "server_table_calcs": False,
        }
        assert asyncio.run(client.run_query(query_obj)) == b"binary data"


# ---------------------------------------------------------------------------
# _run_query_with_retry
# ---------------------------------------------------------------------------


class TestRunQueryWithRetry:
    def test_calls_run_query(self):
        """_run_query_with_retry invokes run_query once when no failure."""
        client = _make_client()
        client.client.run_inline_query.return_value = '{"rows":[]}'

        query_obj = {
            "result_format": "json_bi",
            "body": MagicMock(),
            "apply_vis": False,
            "apply_formatting": False,
            "server_table_calcs": False,
        }
        result = asyncio.run(client._run_query_with_retry(query_obj, retries=0))
        assert result == '{"rows":[]}'

    def test_reraises_on_failure(self):
        """With retries=0, a failing run_query propagates the exception."""
        client = _make_client()
        client.client.run_inline_query.side_effect = RuntimeError("fail")

        query_obj = {
            "result_format": "json_bi",
            "body": MagicMock(),
            "apply_vis": False,
            "apply_formatting": False,
            "server_table_calcs": False,
        }
        with pytest.raises(RuntimeError):
            asyncio.run(client._run_query_with_retry(query_obj, retries=0))


# ---------------------------------------------------------------------------
# make_query
# ---------------------------------------------------------------------------


class TestMakeQuery:
    def test_look_fetch_fails_returns_none(self):
        """If client.look() raises, make_query returns {shape_id: None}."""
        client = _make_client()
        client.client.look.side_effect = Exception("not found")
        assert asyncio.run(client.make_query(shape_id=1, id=999)) == {1: None}

    def test_basic_query_no_filters(self):
        """Happy path: sorts/pivots injected into json_bi result."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days"}, sorts=["field1"])
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps(
            {"metadata": {}, "rows": []}
        )

        result = asyncio.run(client.make_query(shape_id=1, id=42))
        payload = json.loads(result[1])
        assert payload["custom_sorts"] == ["field1"]

    def test_result_format_not_json_returned_as_is(self):
        """Non-JSON formats bypass the sort/pivot injection."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look()
        client.client.run_inline_query.return_value = b"\x89PNG"

        result = asyncio.run(client.make_query(shape_id=2, id=42, result_format="png"))
        assert result[2] == b"\x89PNG"

    def test_json_result_is_list_not_injected(self):
        """A top-level list result is returned without injection."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look()
        raw = json.dumps([{"row": 1}])
        client.client.run_inline_query.return_value = raw

        result = asyncio.run(client.make_query(shape_id=3, id=42))
        assert result[3] == raw

    def test_filter_overwrite_applies(self):
        """filter_overwrites updates the filter before the query runs."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days"})
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        asyncio.run(
            client.make_query(shape_id=1, id=42, filter_overwrites={"date": "30 days"})
        )
        assert mock_look.query.filters["date"] == "30 days"

    def test_filter_overwrite_unknown_warns(self, caplog):
        """Unknown filter_overwrites key emits a warning."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look(filters={"date": "7 days"})
        client.client.run_inline_query.return_value = json.dumps({})

        with caplog.at_level(logging.WARNING):
            asyncio.run(
                client.make_query(
                    shape_id=1, id=42, filter_overwrites={"bad_filter": "x"}
                )
            )
        assert any("not found" in r.message.lower() for r in caplog.records)

    def test_single_filter_applies(self):
        """filter + filter_value updates the matching filter."""
        client = _make_client()
        mock_look = _make_mock_look(filters={"date": "7 days"})
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        asyncio.run(
            client.make_query(shape_id=1, id=42, filter="date", filter_value="90 days")
        )
        assert mock_look.query.filters["date"] == "90 days"

    def test_single_filter_unknown_warns(self, caplog):
        """Unknown single filter emits a warning."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look(filters={"date": "7 days"})
        client.client.run_inline_query.return_value = json.dumps({})

        with caplog.at_level(logging.WARNING):
            asyncio.run(
                client.make_query(
                    shape_id=1, id=42, filter="nope", filter_value="value"
                )
            )
        assert any("not found" in r.message.lower() for r in caplog.records)

    def test_sdk_error_during_run_returns_none(self):
        """SDKError while executing the query returns {shape_id: None}."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look()
        client.client.run_inline_query.side_effect = looker_sdk.error.SDKError("boom")

        result = asyncio.run(client.make_query(shape_id=5, id=42))
        assert result == {5: None}

    def test_unexpected_error_during_run_returns_none(self):
        """Unexpected exceptions return {shape_id: None}."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look()
        client.client.run_inline_query.side_effect = RuntimeError("unexpected")

        result = asyncio.run(client.make_query(shape_id=6, id=42))
        assert result == {6: None}

    def test_result_none_returned_as_none(self):
        """None from run_inline_query is mapped to {shape_id: None}."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look()
        client.client.run_inline_query.return_value = None

        result = asyncio.run(client.make_query(shape_id=7, id=42))
        assert result == {7: None}

    def test_json_decode_error_returns_raw(self):
        """Invalid JSON from run_inline_query is returned as-is."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look()
        client.client.run_inline_query.return_value = "not-valid-json"

        result = asyncio.run(client.make_query(shape_id=8, id=42))
        assert result[8] == "not-valid-json"

    def test_kwargs_applied_to_query(self):
        """Extra kwargs matching query attributes override them."""
        client = _make_client()
        mock_look = _make_mock_look()
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        asyncio.run(client.make_query(shape_id=1, id=42, limit="500"))
        assert mock_look.query.limit == "500"

    def test_list_attribute_appended(self):
        """List-type query attributes have the value appended."""
        client = _make_client()
        mock_look = _make_mock_look(sorts=["field1"])
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({})

        asyncio.run(client.make_query(shape_id=1, id=42, sorts="field2"))
        assert "field2" in mock_look.query.sorts

    def test_apply_vis_formatting_flags_forwarded(self):
        """apply_vis, apply_formatting, server_table_calcs reach run_inline_query."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look()
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
        _, kw = client.client.run_inline_query.call_args
        assert kw["apply_vis"] is True
        assert kw["apply_formatting"] is True
        assert kw["server_table_calcs"] is True

    def test_pivots_injected(self):
        """custom_pivots is populated from the query's pivots field."""
        client = _make_client()
        client.client.look.return_value = _make_mock_look(pivots=["dim1", "dim2"])
        client.client.run_inline_query.return_value = json.dumps({"rows": []})

        result = asyncio.run(client.make_query(shape_id=1, id=42))
        payload = json.loads(result[1])
        assert payload["custom_pivots"] == ["dim1", "dim2"]

    def test_none_sorts_and_pivots(self):
        """None sorts/pivots become empty lists in the result."""
        client = _make_client()
        mock_look = _make_mock_look()
        mock_look.query.sorts = None
        mock_look.query.pivots = None
        client.client.look.return_value = mock_look
        client.client.run_inline_query.return_value = json.dumps({"rows": []})

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
        client.client.look.return_value = _make_mock_look()
        client.client.run_inline_query.return_value = json.dumps({})

        result = asyncio.run(
            client._async_write_queries(shape_id=10, filter_value="30 days", id=99)
        )
        assert 10 in result

    def test_filter_value_propagated(self):
        """filter_value kwarg is propagated all the way to the query."""
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
