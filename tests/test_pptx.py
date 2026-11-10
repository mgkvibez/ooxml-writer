"""A real PowerPoint deck, edited and written back.

The DOCX-shaped tests prove the string surgery; this file proves the thing the README claims, on a file
produced by python-pptx: every part that was not edited is copied byte for byte, and the parts that were
edited stay well-formed.
"""

import xml.etree.ElementTree as ElementTree
import zipfile

import pytest

from ooxml_writer import OoxmlPackage
from ooxml_writer.paragraphs import find_paragraph, paragraphs, replace_paragraph_text, replace_phrase
from ooxml_writer.tables import cell_namespace, set_table_cell_text, table_rows

FIXTURE = "tests/fixtures/deck.pptx"
SLIDE = "ppt/slides/slide1.xml"


@pytest.fixture()
def package():
    return OoxmlPackage.open(FIXTURE)


def well_formed(xml, part="part"):
    """Every edited part must still parse; a string splice that unbalances a tag is a corrupt deck."""
    return ElementTree.fromstring(xml.encode("utf-8"))


def test_opening_a_real_deck_and_saving_it_unchanged_changes_nothing(package, tmp_path):
    out = tmp_path / "untouched.pptx"
    package.save(out)
    with zipfile.ZipFile(FIXTURE) as before, zipfile.ZipFile(out) as after:
        assert before.namelist() == after.namelist()
        for name in before.namelist():
            assert before.read(name) == after.read(name), name


def test_the_fixture_really_does_split_a_phrase_across_runs(package):
    """The trap this project exists for: a naive `xml.replace("Hello world")` finds nothing on a real
    slide."""
    xml = package.text(SLIDE)
    assert "Hello world" not in xml
    assert "<a:t>Hello </a:t>" in xml
    assert find_paragraph(xml, "Hello world") is not None


def test_a_pptx_table_is_drawingml_not_wordprocessingml(package):
    assert cell_namespace(package.text(SLIDE)) == "a"
    assert table_rows(package.text(SLIDE)) == [["Region", "Revenue"], ["North", "120"], ["South", "98"]]


def test_setting_a_drawingml_cell_uses_the_a_namespace(package):
    previous = set_table_cell_text(package, SLIDE, 0, 1, 1, "137")
    assert previous == "120"
    xml = package.text(SLIDE)
    assert table_rows(xml) == [["Region", "Revenue"], ["North", "137"], ["South", "98"]]
    # the fragment written back must be DrawingML: a <w:r> here is what makes
    # PowerPoint offer to repair the file
    assert "<w:" not in xml
    well_formed(xml)


def test_setting_a_cell_keeps_every_other_part_byte_identical(package, tmp_path):
    set_table_cell_text(package, SLIDE, 0, 2, 1, "101")
    out = tmp_path / "edited.pptx"
    package.save(out)
    with zipfile.ZipFile(FIXTURE) as before, zipfile.ZipFile(out) as after:
        for name in before.namelist():
            if name == SLIDE:
                assert before.read(name) != after.read(name)
            else:
                assert before.read(name) == after.read(name), name


def test_replacing_a_phrase_keeps_the_runs_on_either_side(package):
    """`Hello world` is three runs; changing it must edit those three and not the paragraph around them."""
    index = replace_phrase(package, SLIDE, "Hello world", "Hello there")
    assert index is not None
    xml = package.text(SLIDE)
    well_formed(xml)
    paragraph = paragraphs(xml)[index]
    assert paragraph.text == "Hello there"
    # the first run carries the text, the rest are emptied so their properties survive
    assert [run.text for run in paragraph.runs] == ["Hello there", "", ""]


def test_replacing_a_phrase_inside_one_run_is_a_substring_splice(package):
    assert replace_phrase(package, SLIDE, "steady", "improving") is not None
    xml = package.text(SLIDE)
    assert "Outlook is improving." in xml
    assert "steady" not in xml
    well_formed(xml)


def test_replacing_a_missing_phrase_reports_absence_instead_of_writing_something(package):
    before = package.text(SLIDE)
    assert replace_phrase(package, SLIDE, "not in this deck", "x") is None
    assert package.text(SLIDE) == before


def test_replacing_a_whole_paragraph_still_works_on_a_real_slide(package):
    previous = replace_paragraph_text(package, SLIDE, 0, "Quarterly review, revised")
    assert previous == "Quarterly review"
    assert "Quarterly review, revised" in package.text(SLIDE)
    well_formed(package.text(SLIDE))


def test_absent_table_or_cell_is_none_rather_than_a_neighbouring_edit(package):
    assert set_table_cell_text(package, SLIDE, 4, 0, 0, "x") is None
    assert set_table_cell_text(package, SLIDE, 0, 9, 0, "x") is None
    assert set_table_cell_text(package, SLIDE, 0, 0, 9, "x") is None


def test_leading_spaces_survive_because_xml_would_strip_them(package):
    set_table_cell_text(package, SLIDE, 0, 0, 0, " padded ")
    xml = package.text(SLIDE)
    assert 'xml:space="preserve"> padded </a:t>' in xml
    assert table_rows(xml)[0][0] == " padded "
    well_formed(xml)
