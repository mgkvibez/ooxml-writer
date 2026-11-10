"""Table cells, which is where a document stops being a list of paragraphs.

A table is rows of cells, and a cell holds paragraphs, which hold runs. Nothing about that requires a DOM: a
parse-and-reserialise round trip reorders namespace declarations and rewrites self-closing tags, which is the
fidelity
loss this project exists to avoid, so the cell is found by string surgery and only its text changes.

Two namespaces, one shape. WordprocessingML calls a table ``<w:tbl>`` and a run ``<w:r><w:t>``; DrawingML,
which is what a PPTX slide uses, calls them ``<a:tbl>`` and ``<a:r><a:t>``. They are structurally identical
and textually different, so the patterns here capture the prefix and every written fragment reuses it. An
earlier version of this module hard-coded ``w:`` and wrote WordprocessingML runs into PowerPoint tables, which
produced a slide PowerPoint would repair on open.

The subtlety is the one paragraphs have: a cell's text is the concatenation of its runs, so replacing it means
writing the new text into the first run and emptying the rest rather than deleting them, because deleting a
run drops the formatting
the cell's look depends on.
"""

from __future__ import annotations

import re
from typing import List, NamedTuple, Optional

from .package import OoxmlPackage
from .paragraphs import _unescape_xml, escape_xml

TABLE = re.compile(r"<(?P<ns>[wa]):tbl[ >].*?</(?P=ns):tbl>", re.DOTALL)
ROW = re.compile(r"<(?P<ns>[wa]):tr[ >].*?</(?P=ns):tr>", re.DOTALL)
CELL = re.compile(r"<(?P<ns>[wa]):tc[ >].*?</(?P=ns):tc>", re.DOTALL)
CELL_TEXT = re.compile(r"<(?P<ns>[wa]):t(?: [^>]*)?>(.*?)</(?P=ns):t>", re.DOTALL)
CELL_PARAGRAPH = re.compile(r"<(?P<ns>[wa]):p[ >].*?</(?P=ns):p>", re.DOTALL)


class CellSpan(NamedTuple):
    """Where a cell lives in the part, and which namespace its tags use."""

    start: int
    end: int
    namespace: str


def table_count(xml: str) -> int:
    """How many tables the part contains, so a caller can refuse an index instead of guessing."""
    return len(TABLE.findall(xml))


def table_cells(xml: str) -> List[List[List[str]]]:
    """Every table as rows of flattened cell text: `table_cells(xml)[table][row][cell]`."""
    tables: List[List[List[str]]] = []
    for table in TABLE.finditer(xml):
        rows: List[List[str]] = []
        for row in ROW.finditer(table.group(0)):
            cells = CELL.finditer(row.group(0))
            texts = [
                "".join(_unescape_xml(match.group(2)) for match in CELL_TEXT.finditer(cell.group(0)))
                for cell in cells
            ]
            rows.append(texts)
        tables.append(rows)
    return tables


def table_rows(xml: str, table_index: int = 0) -> List[List[str]]:
    """One table's rows, which is what a caller usually wants."""
    tables = table_cells(xml)
    return tables[table_index] if table_index < len(tables) else []


def cell_span(xml: str, table_index: int, row_index: int, column_index: int) -> Optional[CellSpan]:
    """Byte span of one cell, or None when any of the three indices does not exist.

    Returning None rather than clamping is deliberate: a caller that guessed the wrong table deserves a
    refusal, not a
    silent edit to a different one.
    """
    tables = list(TABLE.finditer(xml))
    if table_index >= len(tables):
        return None
    table = tables[table_index]
    table_xml = table.group(0)
    rows = list(ROW.finditer(table_xml))
    if row_index >= len(rows):
        return None
    row = rows[row_index]
    row_xml = row.group(0)
    cells = list(CELL.finditer(row_xml))
    if column_index >= len(cells):
        return None
    cell = cells[column_index]
    start = table.start() + row.start() + cell.start()
    end = table.start() + row.start() + cell.end()
    return CellSpan(start, end, cell.group("ns"))


def _inline_run(namespace: str, text: str) -> str:
    """A run carrying text, in whichever namespace the cell already uses."""
    preserve = ' xml:space="preserve"' if text != text.strip() else ""
    return f"<{namespace}:r><{namespace}:t{preserve}>{escape_xml(text)}</{namespace}:t></{namespace}:r>"


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
    span = cell_span(xml, table_index, row_index, column_index)
    if span is None:
        return None
    cell_xml = xml[span.start : span.end]
    runs = list(CELL_TEXT.finditer(cell_xml))
    previous = "".join(_unescape_xml(run.group(2)) for run in runs)

    if not runs:
        # A cell with no text element at all: put one inside its first paragraph, in the cell's own namespace.
        paragraph = CELL_PARAGRAPH.search(cell_xml)
        if paragraph is None:
            return None
        addition = _inline_run(span.namespace, text)
        closing = paragraph.end() - len(f"</{span.namespace}:p>")
        updated_cell = cell_xml[:closing] + addition + cell_xml[closing:]
        package.replace(part_name, xml[: span.start] + updated_cell + xml[span.end :])
        return previous

    updated_cell = cell_xml
    # later runs first, so the first run's offsets stay valid while the others are emptied
    for run in reversed(runs[1:]):
        updated_cell = updated_cell[: run.start(2)] + "" + updated_cell[run.end(2) :]

    first = runs[0]
    opener = updated_cell[first.start() : first.start(2)]
    if text != text.strip() and "xml:space" not in opener:
        opener = opener.replace(f"<{span.namespace}:t", f'<{span.namespace}:t xml:space="preserve"', 1)
    updated_cell = updated_cell[: first.start()] + opener + escape_xml(text) + updated_cell[first.end(2) :]

    package.replace(part_name, xml[: span.start] + updated_cell + xml[span.end :])
    return previous


def cell_namespace(xml: str, table_index: int = 0) -> Optional[str]:
    """Which namespace a table uses, which is what tells a PPTX table from a DOCX one."""
    tables = list(TABLE.finditer(xml))
    if table_index >= len(tables):
        return None
    return tables[table_index].group("ns")
