from asyncio import subprocess
import requests
import io
from looker_powerpoint.tools.find_alt_text import (
    get_presentation_objects_with_descriptions,
)
from looker_powerpoint.looker import LookerClient
from looker_powerpoint.models import LookerShape
from looker_powerpoint.tools.group_queries import group_queries_by_identity
from pydantic import ValidationError
import subprocess
from pptx.util import Pt
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

import re
from pptx.dml.color import RGBColor

import re
from pptx.dml.color import RGBColor


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

        self.client = None
        self.relevant_shapes = []
        self.looker_shapes = []
        self.data = {}

        # Initialize the argument parser
        self.parser = self._init_argparser()

        # load tools
        self.get_alt_text = get_presentation_objects_with_descriptions

    def _init_looker(self):
        """Initialize the Looker client"""
        self.client = LookerClient()

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

        parser.add_argument(
            "--filter",
            help="""use the string to filter shapes if they have a set filter dimension""",
            action="store",
            default=None,
            type=str,
        )

        parser.add_argument(
            "--debug-queries",
            help="""Enable debugging for Looker queries. \n
                .env: DEBUG_QUERIES""",
            action="store_true",
            default=False,
        )

        return parser

    def _pick_file(self):
        """
        Picks the PowerPoint file to process.
        If no file path is provided, it looks for the first .pptx file in the current directory.

        Returns:
            str: The path to the PowerPoint file.
        """
        self.file_path = self.args.file_path

        if self.file_path:
            try:
                self.presentation = Presentation(self.file_path)
            except Exception as e:
                logging.error(f"Error opening {self.file_path}: {e}")
        else:
            # If no file path is provided look for a file in the current directory
            files = [
                f
                for f in os.listdir(".")
                if f.endswith(".pptx") and not f.startswith("~$")
            ]
            if files:
                self.file_path = files[0]
                logging.warning(
                    f"No file path provided, using first found file: {self.file_path}. To specify a file, use the -f flag like 'lpt -f <file_path>'."
                )

                try:
                    self.presentation = Presentation(self.file_path)
                except Exception as e:
                    logging.error(f"Error opening {self.file_path}: {e}")
            else:
                logging.error(
                    """
                    No PowerPoint file found in the current directory, closing.
                    Run from a directory with a .pptx file, or
                    specify file using -f flag like 'lpt -f <file_path>'.
                """
                )
                exit(1)

    def _fill_table(self, table, df, headers=True):
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
        if headers:
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
        import yaml

        # convert pydantic model to dict
        if isinstance(data, dict) is False:
            data = data.model_dump()
        data = {k: v for k, v in data.items() if v is not None}
        data = yaml.dump(data)

        # remove None values from data, and convert to string with newlines for YAML compatibility

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

        self._set_alt_text(
            circle,
            {"parent_shape_id": shape.shape_id, "meta": True},
        )

    def _replace_image_with_object(
        self, slide_index, shape_number, image_stream, integration
    ):
        """
        Replaces an existing image in a PowerPoint slide with a new image from a stream.
        Args:
            slide_index: The index of the slide containing the image.
            shape_index: The index of the shape to replace.
            image_stream: A BytesIO stream containing the new image data.
            integration: A dictionary containing integration details.
        """

        slide = self.presentation.slides[slide_index]
        old_shape = None
        for shape in slide.shapes:
            if shape.shape_id == shape_number:
                old_shape = shape

        if old_shape is None:
            for shape in slide.shapes:
                logging.warning(
                    f"Shape numbers on slide {slide_index}: {shape.shape_number} (type: {shape.shape_type})"
                )
            raise ValueError(
                f"Shape with number {shape_number} not found on slide {slide_index}."
            )

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

    def _remove_shape(self, slide_index, shape_number):
        """
        Removes a shape from a PowerPoint slide.
        Args:
            prs: The Presentation object.
            slide_index: The index of the slide containing the shape.
            shape_index: The index of the shape to remove.
        """

        slide = self.presentation.slides[slide_index]
        shape_to_remove = None
        for shape in slide.shapes:
            if shape.shape_id == shape_number:
                shape_to_remove = shape

        if shape_to_remove is None:
            raise ValueError(
                f"Shape with number {shape_number} not found on slide {slide_index}."
            )

        # Remove the shape
        slide.shapes._spTree.remove(shape_to_remove._element)

    def _make_df(self, result):
        """
        Create a pandas DataFrame from Looker data based on the integration settings.
        Args:
            result: The Looker data to convert.
        """
        data = json.loads(result)

        fields = data.get("metadata", {}).get("fields", {})

        # Combine all relevant field groups
        all_fields = (
            fields.get("dimensions", [])
            + fields.get("measures", [])
            + fields.get("table_calculations", [])
        )

        # Build the mapping
        mappy = {
            f"{item['name']}.value": item.get("field_group_variant", item["name"])
            .strip()
            .lower()
            for item in all_fields
        }
        logging.debug(f"Header mapping: {mappy}")
        # Create DataFrame
        df = pd.json_normalize(data.get("rows", [])).fillna("")

        # Apply rename
        df.rename(columns=mappy, inplace=True)

        return df

    async def get_queries(self):
        """
        asyncronously fetch a list of look references
        """
        logging.info(
            f"Fetching Looker queries... {len(self.looker_shapes)} queries to fetch."
        )
        tasks = [
            self.client._async_write_queries(
                shape.shape_id, self.args.filter, **dict(shape.integration)
            )
            for shape in self.looker_shapes
        ]

        # Run all tasks concurrently and gather the results
        results = await asyncio.gather(*tasks)
        form_queries = {r["shape_id"]: r["query"] for r in results}
        self.groups = group_queries_by_identity(form_queries)
        self.queries = [g["query"] for g in self.groups]

    async def run_queries(self):
        """
        asyncronously fetch a list of look references
        """
        logging.info(f"Running Looker queries... {len(self.queries)} queries to run.")
        tasks = [self.client._async_run_queries(query) for query in self.queries]

        # Run all tasks concurrently and gather the results
        self.query_results = await asyncio.gather(*tasks)

    def run(self, **kwargs):
        """
        Main method to run the CLI application.
        """
        self.args = self.parser.parse_args()
        self._pick_file()
        self._init_looker()

        references = self.get_alt_text(self.file_path)
        if not references:
            logging.error(
                "No shapes with id found in the presentation. Add a 'id' : '<look_id>' to the alternative text of a shape to load data into the shape."
            )
            return

        for ref in references:
            try:
                self.relevant_shapes.append(LookerShape.model_validate(ref))
            except ValidationError as e:
                logging.error(
                    f"Validation error when loading alternative text for shape {ref['shape_id']}: {e}"
                )
                continue

        self.looker_shapes = [
            s for s in self.relevant_shapes if s.integration.id_type == "look"
        ]
        metadata_rows = []
        looks = set()
        for looker_shape in self.looker_shapes:
            if looker_shape.integration.id not in looks:
                looks.add(looker_shape.integration.id)
                metadata_rows.append(
                    {
                        "looks": {
                            "value": f"{os.environ.get('LOOKERSDK_BASE_URL')}looks/{looker_shape.integration.id}"
                        }
                    }
                )
        metadata_object = {
            "metadata": {"fields": {"dimensions": [{"name": "looks"}]}},
            "rows": metadata_rows,
        }

        self.data["metadata_shapes"] = json.dumps(metadata_object)

        asyncio.run(self.get_queries())
        asyncio.run(self.run_queries())

        for group, result in zip(self.groups, self.query_results):
            for sid in group["shapes"]:
                self.data[sid] = result
                if self.args.debug_queries:
                    # dump result for debugging
                    w = json.loads(result)
                    with open(f"debug_{sid}.json", "w", encoding="utf-8") as f:
                        json.dump(w, f, indent=4, ensure_ascii=False)
        for looker_shape in self.relevant_shapes:
            if looker_shape.integration.meta:
                if not self.args.self:
                    self._remove_shape(
                        looker_shape.slide_number,
                        looker_shape.shape_number,
                    )

            else:
                result = self.data.get(looker_shape.shape_id)
                if result is None:
                    result = self.data.get(looker_shape.integration.id)

                try:
                    if looker_shape.shape_type == "PICTURE":
                        if (
                            looker_shape.integration.result_format == "jpg"
                            or looker_shape.integration.result_format == "png"
                        ) and looker_shape.integration.id_type == "look":
                            image_stream = BytesIO(result)
                        else:
                            df = self._make_df(result)
                            url = df[looker_shape.integration.label][0]
                            logging.info(
                                f"Fetching image from URL: {url} for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}..."
                            )
                            response = requests.get(url)
                            response.raise_for_status()
                            image_stream = io.BytesIO(response.content)

                        self._replace_image_with_object(
                            looker_shape.slide_number,
                            looker_shape.shape_number,
                            image_stream,
                            looker_shape.original_integration,
                        )

                    elif looker_shape.shape_type in [
                        "CHART",
                        "TABLE",
                        "TEXT_BOX",
                        "TITLE",
                        "AUTO_SHAPE",
                    ]:
                        slide = self.presentation.slides[looker_shape.slide_number]
                        for shape in slide.shapes:
                            if shape.shape_id == looker_shape.shape_number:
                                current_shape = shape
                        df = self._make_df(result)

                        if looker_shape.shape_type == "TABLE":
                            logging.debug(
                                f"Updating table for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}..."
                            )
                            self._fill_table(
                                current_shape.table,
                                df,
                                looker_shape.integration.headers,
                            )

                        elif looker_shape.shape_type in [
                            "TEXT_BOX",
                            "TITLE",
                            "AUTO_SHAPE",
                        ]:
                            logging.debug(
                                f"Updating text for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}..."
                            )

                            try:
                                text_to_insert = df[looker_shape.integration.label][0]
                            except Exception as e:
                                text_to_insert = df.to_string(index=False, header=False)
                                logging.warning(
                                    f"inserting whole text for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}: {e}"
                                )
                            current_shape.text = str(text_to_insert)
                            # add_text_with_numbered_links(current_shape.text_frame, str(text_to_insert))

                        elif looker_shape.shape_type == "CHART":
                            chart_data = CategoryChartData()
                            chart_data.categories = df.iloc[
                                :, 0
                            ].tolist()  # Assuming the first column contains categories
                            chart = current_shape.chart
                            existing_chart_data = chart.plots[0].series
                            logging.debug(
                                f"Existing chart series: {[s.name for s in existing_chart_data]}"
                            )

                            if looker_shape.integration.headers:
                                for series_name in df.columns[1:]:
                                    try:
                                        match = (
                                            re.search(
                                                r"^[^\.]*\.[^\.]*\.(.*)\.value$",
                                                series_name,
                                            )
                                            .group(1)
                                            .replace(".", " - ")
                                            .strip()
                                            .replace("|FIELD|", " ")
                                        )
                                    except Exception as e:
                                        logging.error(
                                            f"Error parsing series name {series_name}: {e}"
                                        )
                                        match = series_name
                                    chart_data.add_series(match, df[series_name])
                            else:
                                if len(df.columns[1:]) != len(existing_chart_data):
                                    logging.warning(
                                        f"{looker_shape.shape_id}. Missing headers! Number of series ({len(df.columns[1:])}) does not match number of existing chart series ({len(existing_chart_data)}). Perhaps you need to enable headers in the integration settings?"
                                    )
                                for series_name, series in zip(
                                    df.columns[1:], existing_chart_data
                                ):
                                    chart_data.add_series(series.name, df[series_name])
                            chart.replace_data(chart_data)

                    else:
                        logging.warning(
                            f"unknown shape type {looker_shape.shape_type} for shape {looker_shape.shape_number} on slide {looker_shape.slide_number}."
                        )
                        continue

                except Exception as e:
                    logging.error(f"Error processing reference {looker_shape}: {e}")
                    import traceback

                    traceback.print_exc()  # Prints the full traceback

                    if not self.args.hide_errors:
                        slide = self.presentation.slides[looker_shape.slide_number]
                        for shape in slide.shapes:
                            if shape.shape_id == looker_shape.shape_number:
                                self._mark_failure(slide, shape)

        if self.args.self:
            self.destination = self.file_path
        else:
            if not self.args.output_dir.endswith("/"):
                self.args.output_dir += "/"
            self.destination = self.args.output_dir + os.path.basename(self.file_path)

        if not os.path.exists(self.args.output_dir) and not self.args.self:
            os.makedirs(self.args.output_dir)

        self.presentation.save(self.destination)

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
