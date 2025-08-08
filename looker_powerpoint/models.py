from pydantic import BaseModel, Field, model_validator, field_validator, ValidationError


class LookerReference(BaseModel):
    """A Pydantic model for Looker reference integration.
    This model is used to define the parameters required to create a Looker model
    for a specific shape in a PowerPoint presentation.
    """

    look_id: str
    result_format: str = "json_bi"  # Default result format
    apply_formatting: bool = Field(
        default=False, description="Apply model-specified formatting to each result."
    )
    apply_vis: bool = Field(
        default=True, description="Apply visualization options to results."
    )
    image_width: int = Field(default=None, description="Width of the image in pixels")
    image_height: int = Field(default=None, description="Height of the image in pixels")
    # optional parameters for the Look (Default to None)

    @field_validator("look_id", mode="before")
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

    shape_id: str
    shape_type: str
    slide_number: int
    shape_width: int = Field(default=None)  # Width in pixels
    shape_height: int = Field(default=None)  # Height in pixels
    integration: LookerReference
    shape_number: int = Field(
        default=None, description="The number of the shape in the slide."
    )

    @model_validator(mode="before")
    @classmethod
    def push_down_relevant_data(cls, data):
        """Push down relevant data from the integration to the shape model."""
        # push down
        # if picture is shape type, then we need to push down the image width and height

        if data["shape_type"] == "PICTURE":
            data["original_integration"] = data["integration"]
            data["integration"]["result_format"] = "jpg"
            data["integration"]["image_width"] = round(data["shape_width"])
            data["integration"]["image_height"] = round(data["shape_height"])

        elif data["shape_type"] == "TABLE":
            if data["integration"].get("apply_formatting") is None:
                data["integration"]["apply_formatting"] = True

        return data
