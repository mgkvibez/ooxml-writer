"""Tables: read every cell, write one cell, and leave the rest of the part byte-identical."""

from __future__ import annotations

import zipfile

from ooxml_writer.package import OoxmlPackage
from ooxml_writer.tables import set_table_cell_text, table_cells, table_count, table_rows

CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/></Types>'
)

TABLE_XML = (
    '<w:tbl><w:tblPr/>'
    "<w:tr>"
    "<w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Name</w:t></w:r></w:p></w:tc>"
    "<w:tc><w:p><w:r><w:t>Role</w:t></w:r></w:p></w:tc>"
    "</w:tr>"
    "<w:tr>"
    "<w:tc><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Ada</w:t></w:r></w:p></w:tc>"
    "<w:tc><w:p><w:r><w:t xml:space=\"preserve\">Engineer </w:t></w:r><w:r><w:t>II</w:t></w:r></w:p></w:tc>"
    "</w:tr>"
    "</w:tbl>"
)

DOCUMENT_TAIL = "<w:p><w:r><w:t>After the table</w:t></w:r></w:p></w:body></w:document>"
PART = '<?xml version="1.0"?><w:document><w:body>' + TABLE_XML + DOCUMENT_TAIL


def build(tmp_path):
    path = tmp_path / "doc.docx"
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES_XML)
        archive.writestr("word/document.xml", PART)
    return path


def test_reads_every_cell_including_a_text_split_across_runs(tmp_path):
    package = OoxmlPackage.open(build(tmp_path))
    assert table_count(package.text("word/document.xml")) == 1
    assert table_cells(package.text("word/document.xml")) == [
        [["Name", "Role"], ["Ada", "Engineer II"]]
    ]
    assert table_rows(package.text("word/document.xml"), 0)[1] == ["Ada", "Engineer II"]


def test_sets_one_cell_and_returns_what_was_there(tmp_path):
    package = OoxmlPackage.open(build(tmp_path))
    previous = set_table_cell_text(package, "word/document.xml", 0, 1, 0, "Grace")
    assert previous == "Ada"
    assert table_rows(package.text("word/document.xml"))[1][0] == "Grace"


def test_keeps_the_cells_run_properties(tmp_path):
    package = OoxmlPackage.open(build(tmp_path))
    set_table_cell_text(package, "word/document.xml", 0, 1, 0, "Grace")
    xml = package.text("word/document.xml")
    # the italic run property belonged to the cell's only run and must survive the edit
    assert "<w:i/>" in xml
    assert "Grace" in xml


def test_collapses_a_split_cell_into_one_run(tmp_path):
    package = OoxmlPackage.open(build(tmp_path))
    previous = set_table_cell_text(package, "word/document.xml", 0, 1, 1, "Engineer III")
    assert previous == "Engineer II"
    cells = table_rows(package.text("word/document.xml"))
    assert cells[1][1] == "Engineer III"
    xml = package.text("word/document.xml")
    assert xml.count("<w:t>") + xml.count("<w:t ") == 6


def test_preserves_leading_and_trailing_spaces(tmp_path):
    package = OoxmlPackage.open(build(tmp_path))
    set_table_cell_text(package, "word/document.xml", 0, 0, 0, " padded ")
    xml = package.text("word/document.xml")
    assert '<w:t xml:space="preserve"> padded </w:t>' in xml
    assert table_rows(xml)[0][0] == " padded "


def test_refuses_a_cell_that_does_not_exist(tmp_path):
    package = OoxmlPackage.open(build(tmp_path))
    assert set_table_cell_text(package, "word/document.xml", 0, 9, 0, "x") is None
    assert set_table_cell_text(package, "word/document.xml", 0, 0, 9, "x") is None
    assert set_table_cell_text(package, "word/document.xml", 9, 0, 0, "x") is None


def test_leaves_everything_outside_the_cell_untouched(tmp_path):
    package = OoxmlPackage.open(build(tmp_path))
    before = package.text("word/document.xml")
    set_table_cell_text(package, "word/document.xml", 0, 1, 0, "Grace")
    after = package.text("word/document.xml")
    assert before.replace("Ada", "Grace") == after
    assert "After the table" in after
