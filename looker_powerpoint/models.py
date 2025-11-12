import logging
from pydantic import BaseModel, Field, model_validator, field_validator, ValidationError


class LookerReference(BaseModel):
    """
    This model represents the input you can set in alternative text for a shape in PowerPoint.
    You can specify the different parameters to control how Looker data is fetched and displayed.
    """

    id: str = Field(
        ...,
        description="The ID of the Look or meta-look (meta_name) you want to reference.",
    )
    id_type: str = Field(
        default="look",
        description="The type of ID provided: 'look' or 'meta'. Defaults to 'look'."
        " Setting to 'meta' indicates that the ID refers to a meta Look.",
    )
    meta: bool = Field(
        default=False,
        description="Set this to true if the Look is a meta Look. A meta look is a look that you want to retrieve and reuse, but not display directly.",
    )
    meta_name: str = Field(
        default=None,
        description="NOT actually working yet. If you are defining a meta look, you should provide a reference name here. This can then be used by other shapes to reference this meta look.",
    )
    meta_iterate: bool = Field(
        default=False,
        description="If set to true, this meta look will be iterated over by other shapes referencing it. This is useful for creating dynamic content based on the results of the meta look.",
    )
    label: str = Field(
        default=None,
        description="Setting a label here filters the results to the specified label. The label needs to match the specific column label from the look including any special characters.",
    )
    column: int = Field(
        default=None,
        description="The specific column to retrieve from the Look results. 0-indexed.",
    )
    row: int = Field(
        default=None,
        description="If you want to retrieve a specific row from the Look results, set the row number here (0-indexed).",
    )
    filter: str = Field(
        default=None,
        description="Define a lookml.field_name used in the Look that you want to be able to filter on using the --filter cli argument. Inputting --filter <value> will filter the results to where <label>=<value>.",
    )
    filter_overwrites: dict = Field(
        default=None,
        description="A dictionary of filter overwrites to apply to the Look. The keys are the filter lookml.field_names, and the values are the filter values. The filter values should not be enclosed in quotation marks. (unvalidated)",
    )
    result_format: str = Field(
        default="json_bi",
        description="The format to return the results in. Defaults to 'json_bi'.",
    )
    show_latest_chart_label: bool = Field(
        default=False,
        description="If set to true, modify chart series with labels to only show the latest label.",
    )
    apply_formatting: bool = Field(
        default=False, description="Apply Looker-specified formatting to each result."
    )
    apply_vis: bool = Field(
        default=True, description="Apply Looker visualization options to results."
    )
    server_table_calcs: bool = Field(
        default=True,
        description="Whether to compute table calculations on the Looker server before returning results.",
    )
    headers: bool = Field(
        default=True,
        description="Whether to overwrite headers in the result set with Looker-defined column labels.",
    )
    image_width: int = Field(
        default=None,
        description="Width of the image in pixels. Used for setting image size when asking looker to return a look rendered as an image.",
    )
    image_height: int = Field(
        default=None,
        description="Height of the image in pixels. Used for setting image size when asking looker to return a look rendered as an image.",
    )
    # optional parameters for the Look (Default to None)

    @field_validator("id", mode="before")
    @classmethod
    def convert_int(cls, value):
        """Validation: Convert integer values to strings."""
        if isinstance(value, int):
            return str(value)
        return value


class LookerShape(BaseModel):
    """A Pydantic model for a shape in a PowerPoint presentation.
    This model is used to define the properties of a shape, including its ID, type, dimensions,
    and associated Looker reference.
    """

    is_meta: bool = Field(
        default=False, description="Whether this shape is a meta shape."
    )
    meta_name: str = Field(
        default=None, description="The name of the meta shape, if applicable."
    )
    shape_id: str
    shape_type: str
    slide_number: int
    shape_width: int = Field(default=None)  # Width in pixels
    shape_height: int = Field(default=None)  # Height in pixels
    integration: LookerReference
    original_integration: LookerReference = Field(
        default=None,
        description="The original integration data before any modifications.",
    )
    shape_number: int = Field(
        default=None, description="The number of the shape in the slide."
    )

    @model_validator(mode="before")
    @classmethod
    def push_down_relevant_data(cls, data):
        """Push down relevant data from the integration to the shape model."""
        # push down
        # if picture is shape type, then we need to push down the image width and height
        if type(data.get("integration")) in (dict, LookerReference):
            data["original_integration"] = data["integration"]

            if data["shape_type"] == "PICTURE":
                data["integration"]["result_format"] = data["integration"].get(
                    "result_format", "json_bi"
                )
                data["integration"]["image_width"] = round(data["shape_width"])
                data["integration"]["image_height"] = round(data["shape_height"])

            elif data["shape_type"] == "TABLE":
                if data["integration"].get("apply_formatting") is None:
                    data["integration"]["apply_formatting"] = True

        return data
