import logging
import looker_sdk
from rich import print

from looker_sdk import models40 as models


class LookerClient:
    def __init__(self):
        self.client = looker_sdk.init40()  # or init40() for the v4.0 API

    async def get_look(self, shape_id, filter_value, look_id: str, filter: str = None, **kwargs):
        """
        Retrieves a Look's definition by its ID.

        Args:
            look_id: The ID of the Look.

        Returns:
            A Look object containing the Look's definition.
        """

        try:
        #     response = self.client.run_look(look_id, **kwargs)

            look = self.client.look(look_id)
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

            if filter_value is not None and filter is not None:
                # If filter_value is provided, set the filter
                if hasattr(q, 'filters'):
                    filterable = False
                    for _, f in enumerate(q.filters):
                        print(f, filter)
                        if f == filter:
                            filterable = True
                if filterable:
                    q.filters[filter] = filter_value

            request = models.WriteQuery(
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

            result_format = "json_bi"

            response = self.client.run_inline_query(
                result_format=result_format, body=request
            )

        except looker_sdk.error.SDKError as e:
            logging.error(f"Error retrieving Look with ID {look_id}")
            return {shape_id: None}

        return {shape_id: response}
