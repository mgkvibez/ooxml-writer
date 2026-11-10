"""Shape editing by name, on the real deck, checked by an independent reader where one is available."""

import shutil
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation

from ooxml_writer import OoxmlPackage
from ooxml_writer.shapes import ShapeError, find_shape, set_shape_text, shape_text, shapes

DECK = Path("tests/fixtures/deck.pptx")
SLIDE = "ppt/slides/slide1.xml"
OTHER_SLIDE = "ppt/slides/slide2.xml"


@pytest.fixture
def package(tmp_path):
    copy = tmp_path / "deck.pptx"
    shutil.copyfile(DECK, copy)
    return OoxmlPackage.open(copy), copy


class TestListing:
    def test_finds_every_shape_with_the_name_powerpoint_shows(self, package):
        opened, _ = package
        listed = shapes(opened.text(SLIDE))
        assert [shape.name for shape in listed] == ["Title 1", "Content Placeholder 2", "TextBox 3"]
        assert [shape.shape_id for shape in listed] == ["2", "3", "4"]

    def test_flattens_each_body_to_one_line_per_paragraph(self, package):
        opened, _ = package
        body = find_shape(opened.text(SLIDE), "Content Placeholder 2")
        assert body is not None
        assert body.paragraphs == 3
        # a run-on string would read "Revenue is up.Costs are flat." and hide the paragraph boundaries
        assert body.text == "Revenue is up.\nCosts are flat.\nOutlook is steady."

    def test_a_slide_that_is_not_there_answers_none_rather_than_raising(self, package):
        opened, _ = package
        assert find_shape(opened.text(SLIDE), "No Such Shape") is None
        assert shape_text(opened, SLIDE, "No Such Shape") is None

    def test_matches_ignoring_case_when_exactly_one_shape_would_answer(self, package):
        opened, _ = package
        assert find_shape(opened.text(SLIDE), "title 1") is not None
        assert find_shape(opened.text(SLIDE), "TITLE 1").name == "Title 1"


class TestEditing:
    def test_replaces_a_body_and_returns_what_it_said(self, package):
        opened, _ = package
        previous = set_shape_text(opened, SLIDE, "Content Placeholder 2", "One line instead of three.")
        assert previous == "Revenue is up.\nCosts are flat.\nOutlook is steady."
        edited = find_shape(opened.text(SLIDE), "Content Placeholder 2")
        assert edited is not None
        assert edited.text == "One line instead of three."
        assert edited.paragraphs == 1

    def test_leaves_every_other_shape_alone(self, package):
        opened, _ = package
        set_shape_text(opened, SLIDE, "Content Placeholder 2", "Replaced.")
        assert shape_text(opened, SLIDE, "Title 1") == "Quarterly review"
        assert shape_text(opened, SLIDE, "TextBox 3") == "Hello world"
        assert [shape.name for shape in shapes(opened.text(SLIDE))] == [
            "Title 1",
            "Content Placeholder 2",
            "TextBox 3",
        ]

    def test_leaves_other_parts_byte_identical(self, package):
        opened, written = package
        untouched_before = opened.part(OTHER_SLIDE)
        set_shape_text(opened, SLIDE, "Title 1", "Annual review")
        assert opened.part(OTHER_SLIDE) == untouched_before
        opened.save(written)
        with zipfile.ZipFile(written) as archive:
            assert archive.read(OTHER_SLIDE) == untouched_before

    def test_keeps_the_first_runs_formatting_and_drops_the_rest_of_the_paragraphs(self, package):
        opened, _ = package
        # a body whose first run carries properties the look depends on, and whose later paragraph carries
        # its own
        formatted = (
            "<p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"Runs 7\"/></p:nvSpPr><p:txBody><a:bodyPr/>"
            '<a:p><a:r><a:rPr lang="en-GB" sz="1800" b="1"/><a:t>old one</a:t></a:r>'
            '<a:r><a:rPr lang="en-GB" sz="1800"/><a:t> and more</a:t></a:r></a:p>'
            '<a:p><a:r><a:rPr sz="1800" i="1"/><a:t>old two</a:t></a:r></a:p>'
            "</p:txBody></p:sp>"
        )
        opened.replace(SLIDE, opened.text(SLIDE).replace("</p:spTree>", formatted + "</p:spTree>"))
        previous = set_shape_text(opened, SLIDE, "Runs 7", "new text")
        assert previous == "old one and more\nold two"
        edited = opened.text(SLIDE)
        assert '<a:t>new text</a:t>' in edited
        assert 'sz="1800" b="1"' in edited  # the first run's properties survive
        assert edited.count('i="1"') == 0  # the paragraph that was dropped went with its properties
        assert "old two" not in edited

    def test_escapes_markup_in_the_new_text(self, package):
        opened, _ = package
        set_shape_text(opened, SLIDE, "Title 1", "R&D <draft>")
        assert shape_text(opened, SLIDE, "Title 1") == "R&D <draft>"
        assert "R&amp;D &lt;draft&gt;" in opened.text(SLIDE)

    def test_refuses_a_shape_with_no_text_body_instead_of_doing_nothing(self, package):
        opened, _ = package
        decoration = (
            '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Decoration 9"/></p:nvSpPr>'
            '<p:spPr><a:prstGeom prst="rect"/></p:spPr></p:sp>'
        )
        opened.replace(SLIDE, opened.text(SLIDE).replace("</p:spTree>", decoration + "</p:spTree>"))
        with pytest.raises(ShapeError, match="no text body"):
            set_shape_text(opened, SLIDE, "Decoration 9", "text")

    def test_refuses_a_name_that_is_not_there_and_says_what_is(self, package):
        opened, _ = package
        with pytest.raises(ShapeError, match="Title 1"):
            set_shape_text(opened, SLIDE, "Footer", "text")

    def test_refuses_a_name_two_shapes_would_answer(self, package):
        opened, _ = package
        twin = (
            '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Title 1"/></p:nvSpPr>'
            '<p:txBody><a:p><a:r><a:t>twin</a:t></a:r></a:p></p:txBody></p:sp>'
        )
        opened.replace(SLIDE, opened.text(SLIDE).replace("</p:spTree>", twin + "</p:spTree>"))
        with pytest.raises(ShapeError, match="names have to be unique"):
            set_shape_text(opened, SLIDE, "Title 1", "text")


class TestReopenedByPowerPointLibrary:
    def test_python_pptx_sees_the_new_text_on_the_right_shape(self, package):
        opened, written = package
        set_shape_text(opened, SLIDE, "Content Placeholder 2", "Reopened by python-pptx.")
        opened.save(written)

        reading = Presentation(str(written))
        texts = {
            shape.name: shape.text_frame.text if shape.has_text_frame else ""
            for shape in reading.slides[0].shapes
        }
        assert texts["Content Placeholder 2"] == "Reopened by python-pptx."
        assert texts["Title 1"] == "Quarterly review"

    def test_the_saved_file_is_still_a_valid_archive_with_the_same_parts_in_order(self, package):
        opened, written = package
        set_shape_text(opened, SLIDE, "Title 1", "Annual review")
        opened.save(written)
        with zipfile.ZipFile(written) as archive:
            assert archive.testzip() is None
            with zipfile.ZipFile(DECK) as original:
                assert archive.namelist() == original.namelist()
