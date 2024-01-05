#!/usr/bin/env python3
"""Example of an executable data-loader plugin for the Rerun Viewer for docx files."""
from __future__ import annotations

import argparse
import os

from docx import Document
from docx.opc.exceptions import PackageNotFoundError
from docx.oxml import CT_P, CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table
import xml.etree.ElementTree as ET
import rerun as rr  # pip install rerun-sdk


def docx_to_markdown(docx_path):
    """
    Convert a .docx file to a Markdown formatted string, placing tables correctly.
    """
    try:
        # Load the docx file
        doc = Document(docx_path)
    except PackageNotFoundError:
        return "Invalid .docx file or file not found."

    markdown_content = ""
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            markdown_content += process_paragraph(block)
        elif isinstance(block, Table):
            markdown_content += process_table(block)

    return markdown_content


def iter_block_items(parent):
    """
    Generate a reference to each paragraph and table child within parent, in document order.
    """
    if hasattr(parent, "element"):
        parent_elm = parent.element.body
    else:
        raise ValueError("parent must be a Document object")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


def process_paragraph(para):
    """
    Process a single paragraph, applying Markdown formatting.
    """
    text = para.text.strip()
    if not text:  # Skip empty paragraphs
        return ""

    # Parse the XML to find the numId value
    numId = get_numId(para._element.xml)

    if numId:
        if numId == "1":
            return f"1. {text}\n"  # Assuming '1' indicates a numbered list
        elif numId == "2":
            return f"* {text}\n"  # Assuming '2' indicates a bulleted list
        # Add additional conditions if there are more list types
    elif para.style.name.startswith("Heading"):
        level = len(para.style.name.replace("Heading", ""))
        return "#" * level + " " + text + "\n\n"
    else:
        return text + "\n\n"


def get_numId(xml_string):
    """
    Extract numId value from paragraph XML.
    """
    root = ET.fromstring(xml_string)
    namespace = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    numPr = root.find(".//w:numPr", namespace)
    if numPr is not None:
        numId = numPr.find("./w:numId", namespace)
        if numId is not None:
            return numId.attrib.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val")
    return None


def process_table(table):
    """
    Convert a docx table to GitHub Flavored Markdown table format.
    """
    markdown_table = ""
    for i, row in enumerate(table.rows):
        row_data = [cell.text.strip() for cell in row.cells]

        # Add the row to the table
        markdown_table += "| " + " | ".join(row_data) + " |\n"

        # After the first row (header), add the separator row
        if i == 0:
            markdown_table += "|---" * len(row.cells) + "|\n"

    return markdown_table + "\n"


# The Rerun Viewer will always pass these two pieces of information:
# 1. The path to be loaded, as a positional arg.
# 2. A shared recording ID, via the `--recording-id` flag.
#
# It is up to you whether you make use of that shared recording ID or not.
# If you use it, the data will end up in the same recording as all other plugins interested in
# that file, otherwise you can just create a dedicated recording for it. Or both.
parser = argparse.ArgumentParser(
    description="""
This is an example executable data-loader plugin for the Rerun Viewer.
Any executable on your `$PATH` with a name that starts with `rerun-loader-` will be
treated as an external data-loader.

This example will load .docx files and log them as simplified as markdown documents,
and return a special exit code to indicate that it doesn't support anything else.

To try it out, copy it in your $PATH as `rerun-loader-python-example-docx`, 
then open a .docx source file with Rerun (`rerun file.docx`).
"""
)
parser.add_argument("filepath", type=str)
parser.add_argument("--recording-id", type=str)
args = parser.parse_args()


def main() -> None:
    is_file = os.path.isfile(args.filepath)
    is_docx_file = os.path.splitext(args.filepath)[1].lower() == ".docx"

    # Inform the Rerun Viewer that we do not support that kind of file.
    if not is_file or not is_docx_file:
        exit(rr.EXTERNAL_DATA_LOADER_INCOMPATIBLE_EXIT_CODE)

    rr.init("rerun_example_external_data_loader_docx", recording_id=args.recording_id)
    # The most important part of this: log to standard output so the Rerun Viewer can ingest it!
    rr.stdout()

    markdown = docx_to_markdown(args.filepath)
    rr.log(args.filepath, rr.TextDocument(markdown, media_type=rr.MediaType.MARKDOWN), timeless=True)


if __name__ == "__main__":
    main()
