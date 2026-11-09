"""Table cells, which is where a document stops being a list of paragraphs.

Both Word and PowerPoint store a table as `<w:tbl>` rows of `<w:tc>` cells, and a cell holds
paragraphs, which hold runs. Nothing about that requires a DOM: a parse-and-reserialise round
trip reorders namespace declarations and rewrites self-closing tags, which is the fidelity loss
this project exists to avoid, so the cell is found by string surgery and only its text changes.

The subtlety is the one paragraphs have: a cell's text is the concatenation of its runs, so
replacing it means writing the new text into the first run and emptying the rest rather than
deleting them, because deleting a run drops the formatting the cell's look depends on.
"""

from __future__ import annotations

import re
from typing import List, Optional

from .package import OoxmlPackage
from .paragraphs import _unescape_xml, escape_xml

TABLE = re.compile(r"<w:tbl>.*?</w:tbl>", re.DOTALL)
ROW = re.compile(r"<w:tr[ >].*?</w:tr>", re.DOTALL)
CELL = re.compile(r"<w:tc>.*?</w:tc>", re.DOTALL)
TEXT = re.compile(r"<w:t(?: [^>]*)?>(.*?)</w:t>", re.DOTALL)


def table_count(xml: str) -> int:
    """How many tables the part contains, so a caller can refuse an index instead of guessing."""
    return len(TABLE.findall(xml))


def table_cells(xml: str) -> List[List[List[str]]]:
    """Every table as rows of flattened cell text: `table_cells(xml)[table][row][cell]`."""
    tables: List[List[List[str]]] = []
    for table in TABLE.findall(xml):
        rows: List[List[str]] = []
        for row in ROW.findall(table):
            cells = CELL.findall(row)
            rows.append(["".join(_unescape_xml(t) for t in TEXT.findall(cell)) for cell in cells])
        tables.append(rows)
    return tables


def table_rows(xml: str, table_index: int = 0) -> List[List[str]]:
    """One table's rows, which is what a caller usually wants."""
    tables = table_cells(xml)
    return tables[table_index] if table_index < len(tables) else []


def _cell_span(xml: str, table_index: int, row_index: int, column_index: int) -> Optional[tuple]:
    tables = list(TABLE.finditer(xml))
    if table_index >= len(tables):
        return None
    table = tables[table_index]
    rows = list(ROW.finditer(xml, table.start(), table.end()))
    if row_index >= len(rows):
        return None
    row = rows[row_index]
    cells = list(CELL.finditer(xml, row.start(), row.end()))
    if column_index >= len(cells):
        return None
    cell = cells[column_index]
    return cell.start(), cell.end()


def set_table_cell_text(
    package: OoxmlPackage,
    part_name: str,
    table_index: int,
    row_index: int,
    column_index: int,
    text: str,
) -> Optional[str]:
    """Set one cell's text, keeping its formatting. Returns the previous text, or None if absent."""
    xml = package.text(part_name)
    span = _cell_span(xml, table_index, row_index, column_index)
    if span is None:
        return None
    start, end = span
    cell_xml = xml[start:end]
    runs = list(TEXT.finditer(cell_xml))
    previous = "".join(_unescape_xml(run.group(1)) for run in runs)

    if not runs:
        # A cell with no text element at all: put one inside its first paragraph.
        paragraph = re.search(r"<w:p[ >].*?</w:p>", cell_xml, re.DOTALL)
        if paragraph is None:
            return None
        addition = '<w:r><w:t xml:space="preserve">' + escape_xml(text) + "</w:t></w:r>"
        closing = paragraph.end() - len("</w:p>")
        updated_cell = cell_xml[:closing] + addition + cell_xml[closing:]
        package.replace(part_name, xml[:start] + updated_cell + xml[end:])
        return previous

    updated_cell = cell_xml
    # later runs first, so the first run's offsets stay valid while the others are emptied
    for run in reversed(runs[1:]):
        updated_cell = updated_cell[: run.start(1)] + "" + updated_cell[run.end(1) :]

    first = runs[0]
    opener = updated_cell[first.start() : first.start(1)]
    if text != text.strip() and "xml:space" not in opener:
        opener = opener.replace("<w:t", '<w:t xml:space="preserve"', 1)
    updated_cell = updated_cell[: first.start()] + opener + escape_xml(text) + updated_cell[first.end(1) :]

    package.replace(part_name, xml[:start] + updated_cell + xml[end:])
    return previous
