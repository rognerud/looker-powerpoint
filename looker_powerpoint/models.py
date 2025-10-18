from pydantic import BaseModel, Field, model_validator, field_validator, ValidationError



class LookerReference(BaseModel):
    """A Pydantic model for Looker reference integration.
    This model is used to define the parameters required to create a Looker model
    for a specific shape in a PowerPoint presentation.
    """

    id: str
    id_type: str = Field(
        default="look", description="The type of ID provided: 'look' or 'meta'."
    )
    meta: bool = Field(
        default=False, description="Whether this Looker reference is a meta Look."
    )
    meta_name: str = Field(
        default=None, description="The name of the meta Look, if applicable."
    )
    label: str = Field(
        default=None, description="An optional label for the Looker reference."
    )
    filter: str = Field(default=None, description="dimension to expose for filtering on")
    filter_overwrites: dict = Field(
        default=None, description="A dictionary of filter overwrites to apply to the Look."
    )
    result_format: str = Field(default="json_bi")  # Default result format
    apply_formatting: bool = Field(
        default=False, description="Apply model-specified formatting to each result."
    )
    apply_vis: bool = Field(
        default=True, description="Apply visualization options to results."
    )
    server_table_calcs: bool = Field(
        default=True, description="Whether to compute table calculations on the server."
    )
    headers: bool = Field(
        default=True, description="Whether to include column headers in the results."
    )
    image_width: int = Field(default=None, description="Width of the image in pixels")
    image_height: int = Field(default=None, description="Height of the image in pixels") 
    # optional parameters for the Look (Default to None)

    @field_validator("id", mode="before")
    @classmethod
    def convert_int(cls, value):
        """Convert integer values to strings."""
        if isinstance(value, int):
            return str(value)
        return value


class LookerShape(BaseModel):
    """A Pydantic model for a shape in a PowerPoint presentation.
    This model is used to define the properties of a shape, including its ID, type, dimensions,
    and associated Looker reference.
    """
    is_meta: bool = Field(default=False, description="Whether this shape is a meta shape.")
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
        data["original_integration"] = data["integration"]

        if data["shape_type"] == "PICTURE":
            if data["integration"].get("result_format") is None:
                data["integration"]["result_format"] = "jpg"
            data["integration"]["image_width"] = round(data["shape_width"])
            data["integration"]["image_height"] = round(data["shape_height"])

        elif data["shape_type"] == "TABLE":
            if data["integration"].get("apply_formatting") is None:
                data["integration"]["apply_formatting"] = True

        return data
