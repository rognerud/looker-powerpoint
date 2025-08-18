from asyncio import subprocess
from looker_powerpoint.find_alt_text import get_presentation_objects_with_descriptions
from looker_powerpoint.looker import LookerClient
from looker_powerpoint.models import LookerShape
from pydantic import ValidationError
import subprocess
from pptx.util import Pt
from dotenv import load_dotenv
from pptx.chart.data import CategoryChartData
import json
import pandas as pd
from pptx import Presentation
from lxml import etree
import re
import argparse
from rich_argparse import RichHelpFormatter
import logging

from rich.logging import RichHandler
import os
import asyncio
from io import BytesIO
from pptx.dml.color import RGBColor

NS = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}


class Cli:

    # color with rich
    HEADER = """
        Looker PowerPoint CLI : 
        A command line interface for Looker PowerPoint integration.
    """

    def __init__(self):
        logging.basicConfig(
            level=logging.INFO,
            format="%(message)s",
            datefmt="[%X]",
            handlers=[RichHandler()],
        )

        load_dotenv()
        # check if the required environment variables are set
        required_env_vars = [
            "LOOKERSDK_BASE_URL",
            "LOOKERSDK_CLIENT_ID",
            "LOOKERSDK_CLIENT_SECRET",
        ]
        for var in required_env_vars:
            if not os.getenv(var):
                logging.error(
                    f"Environment variable {var} is not set. Please set it before running the CLI. (e.g. export {var}=<value> or create a .env file and set it there: {var}=<value>)"
                )
                exit(1)

        # Initialize the argument parser
        self._args_parser = self._init_argparser()
        self.args = self._args_parser.parse_args()

        self.client = LookerClient()

        self.looker_shapes = []
        self.data = {}

    def _init_argparser(self):
        """Create and configure the argument parser"""
        parser = argparse.ArgumentParser(
            description=self.HEADER,
            formatter_class=RichHelpFormatter,
        )
        # todo
        parser.add_argument(
            "--file-path",
            "-f",
            help="Path to the PowerPoint file to process",
            default=None,
            type=str,
        )
        # todo
        parser.add_argument(
            "--output-dir",
            "-o",
            help="""Path to a directory that will contain the generated pptx files. \n
                .env: OUTPUT_DIR""",
            default="output",
            type=str,
        )
        # todo
        parser.add_argument(
            "--add-links",
            help="Add links to looker in the slides. \n .env: ADD_LINKS",
            action="store_true",
            default=False,
        )
        # todo
        parser.add_argument(
            "--hide-errors",
            help="""
                Stop showing red outlines around shapes with errors. \n
                .env: HIDE_ERRORS
            """,
            action="store_true",
            default=False,
        )
        # todo
        parser.add_argument(
            "--parse-date-syntax-in-filename",
            "-p",
            help="""Parse date syntax in the filename. \n
                .env: PARSE_DATE_SYNTAX_IN_FILENAME
                """,
            action="store_true",
            default=True,
        )

        parser.add_argument(
            "--self",
            "-s",
            help="""Replace the powerpoint file directly instead of creating a new file. \n
                .env: SELF""",
            action="store_true",
            default=False,
        )

        parser.add_argument(
            "--quiet",
            "-q",
            help="""Do not open the PowerPoint file after processing. \n
                .env: QUIET""",
            action="store_true",
            default=False,
        )

        return parser

    def _fill_table(self, table, df):
        """
        Fills a PowerPoint table with data from a DataFrame.

        Args:
            table: A Table object from pptx.
            df: A pandas DataFrame containing the data to fill the table.
        """

        # Get table dimensions
        table_rows = len(table.rows)
        table_cols = len(table.columns)

        # Get DataFrame dimensions
        df_rows = df.shape[0] + 1  # +1 for header
        df_cols = df.shape[1]

        # Determine how much we can fill
        rows_to_fill = min(table_rows, df_rows)
        cols_to_fill = min(table_cols, df_cols)

        # Fill header row
        for col_idx in range(cols_to_fill):
            table.cell(0, col_idx).text = str(df.columns[col_idx])

        # Fill DataFrame values
        for row_idx in range(1, rows_to_fill):  # skip header row
            for col_idx in range(cols_to_fill):
                value = df.iloc[row_idx - 1, col_idx]
                table.cell(row_idx, col_idx).text = str(value)

        # Optional: Clear unused cells
        for row_idx in range(rows_to_fill, table_rows):
            for col_idx in range(table_cols):
                table.cell(row_idx, col_idx).text = ""

        for col_idx in range(cols_to_fill, table_cols):
            for row_idx in range(table_rows):
                table.cell(row_idx, col_idx).text = ""

    def _set_alt_text(self, shape, data):
        """
        Sets the alternative text description for a shape's XML.

        Args:
            shape: A Shape object from pptx.
            data: A Python object (dict, list, etc.) to serialize and set as YAML in the descr attribute.
        """
        xml_str = shape.element.xml
        xml_elem = etree.fromstring(xml_str)

        for path in [
            ".//p:nvSpPr/p:cNvPr",
            ".//p:nvPicPr/p:cNvPr",
            ".//p:nvGraphicFramePr/p:cNvPr",
        ]:
            cNvPr_elements = xml_elem.xpath(path, namespaces=NS)
            if cNvPr_elements:
                cNvPr = cNvPr_elements[0]
                yaml_text = str(data)
                cNvPr.set("descr", yaml_text)

                # Overwrite the element in the actual pptx shape with updated XML
                shape_element = shape.element
                new_element = etree.fromstring(etree.tostring(xml_elem))
                shape_element.clear()
                for child in new_element:
                    shape_element.append(child)
                return

        raise ValueError("No compatible cNvPr element found to set descr.")

    def _mark_failure(self, slide, shape):
        line_color_rgb = (255, 0, 0)  # RGB color for
        line_width_pt = 2  # Width of the circle outline in points
        # Calculate circle position - centered on the shape

        # Add an oval shape (circle)
        circle = slide.shapes.add_shape(
            autoshape_type_id=1,  # MSO_SHAPE_OVAL (value 1)
            left=shape.left,
            top=shape.top,
            width=shape.width,
            height=shape.height,
        )

        # Set no fill for the circle (transparent inside)
        circle.fill.background()  # or circle.fill.solid() + set transparency

        # Set outline color and width
        circle.line.color.rgb = RGBColor(*line_color_rgb)
        circle.line.width = Pt(line_width_pt)

    def _replace_image_with_object(
        self, prs, slide_index, shape_index, image_stream, integration
    ):
        """
        Replaces an existing image in a PowerPoint slide with a new image from a stream.
        Args:
            prs: The Presentation object.
            slide_index: The index of the slide containing the image.
            shape_index: The index of the shape to replace.
            image_stream: A BytesIO stream containing the new image data.
            integration: A dictionary containing integration details.
        """

        slide = prs.slides[slide_index]
        old_shape = None
        for shape in slide.shapes:
            if shape.shape_id == shape_index:
                old_shape = shape

        if not old_shape.shape_type == 13:  # 13 = PICTURE
            raise ValueError("Selected shape is not an image.")

        # Save original position and size
        left = old_shape.left
        top = old_shape.top
        width = old_shape.width
        height = old_shape.height

        # Remove the old image
        slide.shapes._spTree.remove(old_shape._element)

        # Insert new image from memory (image_stream must be a BytesIO object)
        picture = slide.shapes.add_picture(
            image_stream, left, top, width=width, height=height
        )
        self._set_alt_text(picture, integration)

    async def _async_fetch_look(self, shape_id, **kwargs):
        """
        Asynchronously fetch a Looker look by its ID.
        Args:
            table: A dictionary containing the look_id and other parameters.
        Returns:
            The fetched look data.
        """
        return await self.client.get_look(shape_id, **dict(kwargs))

    async def get_looks(self):
        """
        asyncronously fetch a list of look references
        """
        tasks = [
            self._async_fetch_look(shape.shape_id, **dict(shape.integration))
            for shape in self.looker_shapes
        ]

        # Run all tasks concurrently and gather the results
        self.results = await asyncio.gather(*tasks)

    def run(self, file_path: str = None):
        """
        Main method to run the CLI application.
        Args:
            file_path: Optional path to a PowerPoint file. If not provided, the first .pptx file in the current directory will be used.
        """

        if file_path:
            try:
                presentation = Presentation(file_path)
            except Exception as e:
                logging.error(f"Error opening {file_path}: {e}")
        else:
            # If no file path is provided look for a file in the current directory
            files = [
                f
                for f in os.listdir(".")
                if f.endswith(".pptx") and not f.startswith("~$")
            ]
            if files:
                file_path = files[0]
                logging.warning(
                    f"""
                    No file path provided, using first found file: {file_path}. 
                    To specify a file, use the -f flag like 'lpt -f <file_path>'.
                """
                )
                presentation = Presentation(file_path)
            else:
                logging.error(
                    """
                    No PowerPoint file found in the current directory, closing. 
                    Specify file using -f flag like 'lpt -f <file_path>'.
                """
                )
                return

        references = get_presentation_objects_with_descriptions(file_path)
        if not references:
            logging.error(
                "No shapes with look_id found in the presentation. Add a look_id : <look_id> to the alternative text of a shape to load data into the shape."
            )
            return

        for ref in references:
            try:
                self.looker_shapes.append(LookerShape.model_validate(ref))
            except ValidationError as e:
                logging.error(f"Validation error when loading alternative text for shape {ref['shape_id']}: {e}")
                continue

        asyncio.run(self.get_looks())
        for d in self.results:
            self.data.update(d)

        for looker_shape in self.looker_shapes:
            result = self.data.get(looker_shape.shape_id)
            try:
                if looker_shape.shape_type == "PICTURE":
                    image_stream = BytesIO(result)
                    self._replace_image_with_object(
                        presentation,
                        looker_shape.slide_number,
                        looker_shape.id,
                        image_stream,
                        looker_shape.original_integration,
                    )

                df = pd.json_normalize(json.loads(result).get("rows", [])).fillna(
                    ""
                )

                if looker_shape.shape_type == "TABLE":
                    slide = presentation.slides[looker_shape.slide_number]

                    for shape in slide.shapes:
                        if shape.shape_id == looker_shape.shape_number:
                            chart_shape = shape

                    self._fill_table(chart_shape.table, df)

                elif looker_shape.shape_type in ["TEXT_BOX", "TITLE", "AUTO_SHAPE"]:
                    slide = presentation.slides[looker_shape.slide_number]

                    for shape in slide.shapes:
                        if shape.shape_id == looker_shape.shape_number:
                            text_shape = shape
                            text_shape.text = df[0][0]

                elif looker_shape.shape_type == "CHART":
                        chart_data = CategoryChartData()
                        chart_data.categories = df.iloc[
                            :, 0
                        ].tolist()  # Assuming the first column contains categories

                        for series_name in df.columns[1:]:
                            match = (
                                re.search(r"^[^\.]*\.[^\.]*\.(.*)\.value$", series_name)
                                .group(1)
                                .replace(".", " - ")
                            )

                            chart_data.add_series(match, df[series_name])
                            # print(match)

                        slide = presentation.slides[looker_shape.slide_number]
                        for shape in slide.shapes:
                            if shape.shape_id == looker_shape.shape_number:
                                chart_shape = shape

                        # Replace chart data (assumes the shape is a chart)
                        chart = chart_shape.chart
                        chart.replace_data(chart_data)
                
                else:
                    logging.error(
                        f"Unsupported shape type {looker_shape.shape_type} for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}."
                    )
                    continue

            except Exception as e:
                logging.error(f"Error processing reference {looker_shape}: {e}")
                if not self.args.hide_errors:
                    slide = presentation.slides[looker_shape.slide_number]
                    for shape in slide.shapes:
                        if shape.shape_id == looker_shape.shape_number:
                            self._mark_failure(slide, shape)

        if self.args.self:
            self.destination = file_path
        else:
            if not self.args.output_dir.endswith("/"):
                self.args.output_dir += "/"
            self.destination = self.args.output_dir + os.path.basename(file_path)

        if not os.path.exists(self.args.output_dir) and not self.args.self:
            os.makedirs(self.args.output_dir)

        presentation.save(self.destination)

        if not self.args.quiet:
            try:
                os.startfile(self.destination)
                logging.info(f"Opened {self.destination} in PowerPoint.")
            except Exception as e:
                try:
                    subprocess.Popen(["open", self.destination])  # For macOS
                    logging.info(f"Opened {self.destination} in PowerPoint.")
                except Exception as e:
                    logging.error(f"Failed to open the PowerPoint file: {e}")
                    logging.info(f"You can find the file at {self.destination}.")


def main():
    cli = Cli()
    cli.run()


if __name__ == "__main__":
    main()
