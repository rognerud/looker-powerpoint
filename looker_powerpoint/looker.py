"""Looker SDK wrapper.

``LookerClient`` wraps the Looker SDK and exposes async helpers for
fetching and executing Looker Looks from a PowerPoint pipeline.

The class is structured so that every code path can be independently
unit-tested:
- SDK initialization is delegated to ``_initialize_client()`` so it can be
  patched without reconstructing the whole object.
- Filter application, query building, and result post-processing are each
  extracted into focused static / instance methods.
- The nested retry closure is wrapped in ``_run_query_with_retry()`` so
  retry behaviour is testable on its own.
"""

import json
import logging
import sys
from typing import Optional

import looker_sdk
from dotenv import find_dotenv, load_dotenv
from looker_sdk import models40 as models
from tenacity import (
    before_sleep_log,
    retry,
    stop_after_attempt,
    wait_fixed,
)


class LookerClient:
    """Looker SDK wrapper with full unit-test coverage."""

    # ------------------------------------------------------------------
    # Construction
    # ------------------------------------------------------------------

    def __init__(self) -> None:
        load_dotenv(find_dotenv(usecwd=True))
        self._initialize_client()

    def _initialize_client(self) -> None:
        """Initialize the Looker SDK client.

        Separated from ``__init__`` so that tests can patch only this method.
        """
        try:
            self.client = looker_sdk.init40()
        except looker_sdk.error.SDKError as e:
            logging.error(
                "Error initializing Looker SDK: %s  Consider adding a looker.ini "
                "file, or setting the LOOKERSDK_BASE_URL, LOOKERSDK_CLIENT_ID, and "
                "LOOKERSDK_CLIENT_SECRET environment variables.",
                e,
            )
            sys.exit(1)

    # ------------------------------------------------------------------
    # Static helpers – each independently testable
    # ------------------------------------------------------------------

    @staticmethod
    def _apply_kwargs_to_query(q, kwargs: dict) -> None:
        """Apply extra *kwargs* as attribute overrides on a query object *q*.

        If the existing attribute value is a ``list``, the new value is
        *appended*; otherwise the attribute is set directly.
        Unknown / ``None``-valued kwargs are silently ignored.
        """
        for parameter, value in kwargs.items():
            if value is not None and hasattr(q, parameter):
                existing = getattr(q, parameter)
                if isinstance(existing, list):
                    existing.append(value)
                else:
                    setattr(q, parameter, value)

    @staticmethod
    def _apply_filter_overwrites(q, filter_overwrites: dict) -> None:
        """Overwrite specific filter values on query *q*.

        Only keys that already exist in *q.filters* are updated; unknown keys
        trigger a warning log.  If *q* has no ``filters`` attribute the method
        returns immediately.
        """
        if not hasattr(q, "filters") or q.filters is None:
            return
        for f, v in filter_overwrites.items():
            logging.info("Overwriting filter %s with value %s", f, v)
            if f in q.filters:
                q.filters[f] = v
            else:
                logging.warning(
                    "Overwrite filter %s not found in query filters. "
                    "Available filters: %s",
                    f,
                    q.filters,
                )

    @staticmethod
    def _apply_single_filter(q, filter_name: str, filter_value: str) -> None:
        """Apply a single *filter_name* / *filter_value* pair to query *q*.

        The filter must already exist in *q.filters*; unknown filters trigger a
        warning log.  If *q* has no ``filters`` attribute the method returns
        immediately.
        """
        logging.info("Applying filter %s with value %s", filter_name, filter_value)
        if not hasattr(q, "filters") or q.filters is None:
            logging.warning(
                "Filter %s not found in query filters (no filters on query).",
                filter_name,
            )
            return
        if filter_name in q.filters:
            q.filters[filter_name] = filter_value
        else:
            logging.warning(
                "Filter %s not found in query filters. Available filters: %s",
                filter_name,
                q.filters,
            )

    @staticmethod
    def _build_write_query(q) -> models.WriteQuery:
        """Construct a ``WriteQuery`` from a Look's query object."""
        return models.WriteQuery(
            model=q.model,
            view=q.view,
            fields=q.fields,
            pivots=q.pivots,
            fill_fields=q.fill_fields,
            filters=q.filters,
            sorts=q.sorts,
            limit=q.limit,
            column_limit=q.column_limit,
            total=q.total,
            row_total=q.row_total,
            subtotals=q.subtotals,
            dynamic_fields=q.dynamic_fields,
            query_timezone=q.query_timezone,
            vis_config=q.vis_config,
            visible_ui_sections=q.visible_ui_sections,
        )

    @staticmethod
    def _post_process_result(
        result: str,
        result_format: str,
        q,
        shape_id,
        look_id,
    ) -> str:
        """Inject ``custom_sorts`` and ``custom_pivots`` into JSON results.

        For ``json`` and ``json_bi`` result formats where the parsed value is a
        ``dict``, the query's sort and pivot fields are embedded so that
        downstream consumers can access them without making another API call.

        Returns *result* unchanged for non-JSON formats, ``None`` results, or
        results whose top-level type is not ``dict``.
        """
        if result is None or result_format not in ("json", "json_bi"):
            return result
        try:
            parsed = json.loads(result)
            if isinstance(parsed, dict):
                parsed["custom_sorts"] = list(q.sorts) if q.sorts else []
                parsed["custom_pivots"] = list(q.pivots) if q.pivots else []
                return json.dumps(parsed)
        except (json.JSONDecodeError, TypeError, ValueError) as e:
            logging.warning(
                "Failed to inject custom_sorts/custom_pivots for shape_id %s, "
                "look_id %s: %s",
                shape_id,
                look_id,
                e,
                exc_info=True,
            )
        return result

    # ------------------------------------------------------------------
    # Core async API
    # ------------------------------------------------------------------

    async def run_query(self, query_object: dict):
        """Execute a query via the Looker SDK and return the raw response."""
        return self.client.run_inline_query(
            result_format=query_object["result_format"],
            body=query_object["body"],
            apply_vis=query_object["apply_vis"],
            apply_formatting=query_object["apply_formatting"],
            server_table_calcs=query_object["server_table_calcs"],
        )

    async def _run_query_with_retry(self, query_object: dict, retries: int):
        """Run *query_object* with up to *retries* additional attempts on failure."""

        @retry(
            stop=stop_after_attempt(retries + 1),
            wait=wait_fixed(2),
            before_sleep=before_sleep_log(logging.getLogger(), logging.WARNING),
            reraise=True,
        )
        async def _attempt():
            return await self.run_query(query_object)

        return await _attempt()

    async def make_query(
        self,
        shape_id: int,
        filter: Optional[str] = None,
        filter_value: Optional[str] = None,
        filter_overwrites: Optional[dict] = None,
        id: Optional[int] = None,
        **kwargs,
    ) -> dict:
        """Fetch a Look, build a query from it, execute it, and return the result.

        Args:
            shape_id: Identifier of the PowerPoint shape this result is for.
            filter: Name of a single filter to override.
            filter_value: Value for the single filter override.
            filter_overwrites: Mapping of filter names → new values.
            id: Looker Look ID.
            **kwargs: Extra query parameters forwarded to the ``WriteQuery``
                      constructor (e.g. ``result_format``, ``limit``).

        Returns:
            ``{shape_id: <result string or None>}``
        """
        try:
            look = self.client.look(id)
        except Exception:
            logging.error(
                "Error fetching Look with ID %s, is this a valid Look ID? "
                "If it is a meta reference, remember to set id_type: 'meta'",
                id,
            )
            return {shape_id: None}

        q = look.query
        self._apply_kwargs_to_query(q, kwargs)

        if filter_overwrites is not None:
            self._apply_filter_overwrites(q, filter_overwrites)

        if filter_value is not None and filter is not None:
            self._apply_single_filter(q, filter, filter_value)

        body = self._build_write_query(q)

        result_format = kwargs.get("result_format", "json_bi")
        query_object = {
            "result_format": result_format,
            "body": body,
            "apply_vis": kwargs.get("apply_vis", False),
            "apply_formatting": kwargs.get("apply_formatting", False),
            "server_table_calcs": kwargs.get("server_table_calcs", False),
        }
        retries = kwargs.get("retries", 0)

        try:
            result = await self._run_query_with_retry(query_object, retries)
            result = self._post_process_result(result, result_format, q, shape_id, id)
        except looker_sdk.error.SDKError as e:
            logging.error("Error retrieving Look with ID %s : %s", id, e)
            result = None
        except Exception as e:
            logging.error("Unexpected error retrieving Look with ID %s : %s", id, e)
            result = None

        return {shape_id: result}

    async def _async_write_queries(
        self,
        shape_id,
        filter_value=None,
        **kwargs,
    ):
        """Thin private wrapper around ``make_query`` used by the CLI."""
        return await self.make_query(
            shape_id,
            filter_value=filter_value,
            **dict(kwargs),
        )
