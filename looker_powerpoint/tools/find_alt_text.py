from pptx import Presentation
from lxml import etree
import yaml

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


def extract_alt_text(shape):
    """
    Extracts the alternative text description from a shape's XML.

    Args:
        shape: A Shape object from pptx.

    Returns:
        str: The alternative text description, or None if not found.
    """
    xml_str = shape.element.xml  # get XML string of the shape element
    xml_elem = etree.fromstring(xml_str)  # parse it into an lxml element
    for path in [
        ".//p:nvSpPr/p:cNvPr",
        ".//p:nvPicPr/p:cNvPr",
        ".//p:nvGraphicFramePr/p:cNvPr",
    ]:
        cNvPr_elements = xml_elem.xpath(path, namespaces=NS)
        if cNvPr_elements:
            descr = cNvPr_elements[0].get("descr")
            if descr:
                data = yaml.safe_load(
                    descr  # .lower()
                )  # Use safe_load for untrusted sources

                return data
    return None


def get_presentation_objects_with_descriptions(pptx_path):
    """
    Extracts all shapes from a PowerPoint presentation and returns them with descriptions.

    Args:
        pptx_path (str): The path to the PowerPoint presentation file.

    Returns:
        list: A list of dictionaries, where each dictionary represents a shape and
              contains the shape object, its description, and the slide number.
              Returns an empty list if the presentation cannot be opened or has no slides/shapes.
    """
    try:
        presentation = Presentation(pptx_path)
    except Exception as e:
        print(f"Error opening presentation: {e}")
        return []

    objects_with_descriptions = []

    for i, slide in enumerate(presentation.slides, start=0):
        for shape in slide.shapes:
            description = extract_alt_text(shape)  # Generate description

            emu_to_pixels = lambda emu: emu / 9525

            width_px = emu_to_pixels(shape.width)
            height_px = emu_to_pixels(shape.height)

            if description:
                if type(description) is dict and "meta_name" in description:
                    shape_id = description.get("meta_name")
                else:
                    shape_id = (
                        f"{i},{shape.shape_id}"  # Use shape number for identification
                    )

                objects_with_descriptions.append(
                    {
                        "shape_id": shape_id,  # Use shape number for identification
                        "shape_type": shape.shape_type.name,
                        "shape_width": round(width_px),
                        "shape_height": round(height_px),
                        "integration": description,
                        "slide_number": i,  # Use the enumerate index for slide number
                        "shape_number": shape.shape_id,
                    }
                )

    return objects_with_descriptions


if __name__ == "__main__":
    # Example Usage
    pptx_file = "p.pptx"  # Replace with your file path
    objects = get_presentation_objects_with_descriptions(pptx_file)
    from rich import print
