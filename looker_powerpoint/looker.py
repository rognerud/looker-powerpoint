import logging
from typing import Optional
import looker_sdk
from dotenv import load_dotenv, find_dotenv
from looker_sdk import models40 as models


class LookerClient:
    def __init__(self):
        load_dotenv(find_dotenv(usecwd=True))
        try:
            self.client = looker_sdk.init40()  # or init40() for the v4.0 API
        except looker_sdk.error.SDKError as e:
            logging.error(
                f"Error initializing Looker SDK: {e} Consider adding a looker.ini file, or setting the LOOKERSDK_BASE_URL, LOOKERSDK_CLIENT_ID, and LOOKERSDK_CLIENT_SECRET environment variables."
            )
            exit(1)

    async def run_query(self, query_object):
        """
        Runs a query against the Looker API.

        Args:
            query_object: The query object containing the necessary parameters.
        """

        try:
            response = self.client.run_inline_query(
                result_format=query_object["result_format"],
                body=query_object["body"],
                apply_vis=query_object["apply_vis"],
                apply_formatting=query_object["apply_formatting"],
                server_table_calcs=query_object["server_table_calcs"],
            )

        except looker_sdk.error.SDKError as e:
            logging.error(f"Error retrieving Look with ID {id} : {e}")
            return None
        except Exception as e:
            logging.error(f"Unexpected error retrieving Look with ID {id} : {e}")
            return None

        return response

    async def make_query(
        self,
        shape_id: int,
        filter: Optional[str] = None,
        filter_value: Optional[str] = None,
        filter_overwrites: Optional[dict] = None,
        id: Optional[int] = None,
        **kwargs,
    ) -> models.WriteQuery:
        """
        Constructs a WriteQuery object based on a Look's definition and provided parameters.
        Args:
            id: The ID of the Look.
            filter: The name of the filter to apply.
            filter_value: The value to set for the filter.
            filter_overwrites: A dictionary of filters to overwrite with new values.
            **kwargs: Additional query parameters to set.
        Returns:
            A WriteQuery object representing the modified query.
        """
        try:
            # check if string can be converted to int
            look = self.client.look(id)
        except Exception as e:
            logging.error(
                f"Error fetching Look with ID {id}, is this a valid Look ID? If it is a meta reference, remember to set id_type: 'meta'"
            )
            return {shape_id: None}

        q = look.query
        for parameter, value in kwargs.items():
            if value is not None:
                if hasattr(q, parameter):
                    # If the parameter is a list, append the value
                    if isinstance(getattr(q, parameter), list):
                        getattr(q, parameter).append(value)
                    else:
                        # Otherwise, set the value directly
                        setattr(q, parameter, value)

        if filter_overwrites is not None:
            for f, v in filter_overwrites.items():
                logging.info(f"Overwriting filter {f} with value {v}")
                if hasattr(q, "filters"):
                    filterable = False
                    for _, existing_filter in enumerate(q.filters):
                        if existing_filter == f:
                            filterable = True
                    if filterable:
                        q.filters[f] = v
                    else:
                        logging.warning(
                            f"Overwrite filter {f} not found in query filters. Available filters: {q.filters}"
                        )

        if filter_value is not None and filter is not None:
            logging.info(f"Applying filter {filter} with value {filter_value}")
            # If filter_value is provided, set the filter
            if hasattr(q, "filters"):
                filterable = False
                for _, f in enumerate(q.filters):
                    # print(f, filter)
                    if f == filter:
                        filterable = True
            if filterable:
                q.filters[filter] = filter_value
            else:
                logging.warning(
                    f"Filter {filter} not found in query filters. Available filters: {q.filters}"
                )

        body = models.WriteQuery(
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

        result_format = kwargs.get("result_format", "json_bi")
        apply_vis = kwargs.get("apply_vis", False)
        apply_formatting = kwargs.get("apply_formatting", False)
        server_table_calcs = kwargs.get("server_table_calcs", False)

        query_object = {
            "shape_id": shape_id,
            "query": {
                "result_format": result_format,
                "body": body,
                "apply_vis": apply_vis,
                "apply_formatting": apply_formatting,
                "server_table_calcs": server_table_calcs,
            },
        }
        result = await self.run_query(query_object["query"])

        return {shape_id: result}

    async def _async_write_queries(self, shape_id, filter_value=None, **kwargs):
        """
        Asynchronously write a Looker query by its ID.
        Args:
            table: A dictionary containing the look_id and other parameters.
        Returns:
            The fetched look data.
        """
        return await self.make_query(
            shape_id, filter_value=filter_value, **dict(kwargs)
        )
