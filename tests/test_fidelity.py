"""The scoreboard, including the case it deliberately cannot see.

The last class here is the important one: a part whose bytes changed while its element census did not. A
scoreboard that claimed to notice that would be lying, and a test that pins the blind spot is how the claim
stays honest.
"""

import shutil
from pathlib import Path

import pytest

from ooxml_writer import OoxmlPackage
from ooxml_writer.fidelity import census, compare, main, scoreboard
from ooxml_writer.paragraphs import replace_phrase
from ooxml_writer.shapes import set_shape_text

DECK = Path("tests/fixtures/deck.pptx")
SLIDE = "ppt/slides/slide1.xml"


@pytest.fixture
def pair(tmp_path):
    copies = []
    for name in ("before.pptx", "after.pptx"):
        target = tmp_path / name
        shutil.copyfile(DECK, target)
        copies.append(OoxmlPackage.open(target))
    return copies[0], copies[1]


class TestCensus:
    def test_counts_elements_by_the_prefix_the_part_actually_writes(self, pair):
        before, _ = pair
        slide = census(before.part(SLIDE))
        # the prefixes come from the part's own xmlns declarations, not from a table in this module
        assert slide["p:sp"] == 3
        assert slide["p:txBody"] == 3
        assert slide["a:p"] == 11  # paragraphs anywhere in the slide, including the ones inside the table
        assert slide["a:t"] == 13

    def test_a_binary_part_is_an_empty_census_rather_than_an_exception(self, pair):
        before, _ = pair
        assert census(before.part("ppt/media/image1.png")) == {}

    def test_counts_a_whole_part_rather_than_only_its_top_level(self, pair):
        before, _ = pair
        # the table sits inside a graphicFrame, three levels below the shape tree, and every cell is counted
        slide_parts = [name for name in before.order if name.startswith("ppt/slides/slide")]
        counted = [census(before.part(name)) for name in slide_parts]
        assert sum(part.get("a:tbl", 0) for part in counted) == 1
        assert sum(part.get("a:tc", 0) for part in counted) == 6


class TestBaseline:
    def test_an_untouched_save_keeps_every_part_structure(self):
        summary = scoreboard(DECK)
        assert summary.startswith("41 of 41 parts kept their structure, 2643 elements counted")

    def test_nothing_is_reported_as_changed_when_nothing_changed(self, pair):
        before, after = pair
        assert compare(before, after).changed == []

    def test_usage_is_an_error_rather_than_a_default(self):
        assert main(["fidelity"]) == 2


class TestAfterAnEdit:
    def test_a_collapsed_body_names_what_it_removed(self, pair):
        before, after = pair
        set_shape_text(after, SLIDE, "Content Placeholder 2", "One line instead of three.")
        report = compare(before, after)
        assert [part.name for part in report.changed] == [SLIDE]
        changed = report.changed[0]
        # two paragraphs went, and with each of them one run and one piece of text
        assert changed.lost == {"a:p": 2, "a:r": 2, "a:t": 2}
        assert changed.gained == {}
        assert "-2 a:p" in changed.describe()

    def test_one_changed_part_still_leaves_the_other_forty_alone(self, pair):
        before, after = pair
        set_shape_text(after, SLIDE, "Content Placeholder 2", "One line instead of three.")
        report = compare(before, after)
        assert len(report.parts) == 41
        assert len(report.kept) == 40
        assert SLIDE not in [part.name for part in report.kept]

    def test_adding_a_shape_is_gained_elements(self, pair):
        before, after = pair
        extra = (
            '<p:sp><p:nvSpPr><p:cNvPr id="9" name="Extra 9"/></p:nvSpPr>'
            "<p:txBody><a:p><a:r><a:t>added</a:t></a:r></a:p></p:txBody></p:sp>"
        )
        after.replace(SLIDE, after.text(SLIDE).replace("</p:spTree>", extra + "</p:spTree>"))
        changed = compare(before, after).changed
        assert len(changed) == 1
        assert changed[0].gained["p:sp"] == 1
        assert changed[0].lost == {}


class TestWhatItCannotSee:
    def test_a_phrase_replacement_keeps_the_census_and_changes_the_bytes(self, pair):
        before, after = pair
        replace_phrase(after, SLIDE, "Hello world", "Hello there")
        # the honest limit: same elements, different text, so structure is kept and this scoreboard says
        # nothing about the sentence itself. The round-trip report is what sees the difference.
        assert before.part(SLIDE) != after.part(SLIDE)
        assert compare(before, after).changed == []

    def test_new_text_in_the_same_run_keeps_the_census(self, pair):
        before, after = pair
        set_shape_text(after, SLIDE, "Title 1", "Annual review")
        assert before.part(SLIDE) != after.part(SLIDE)
        assert compare(before, after).changed == []

    def test_a_moved_shape_keeps_the_census_too(self, pair):
        before, after = pair
        after.replace(SLIDE, after.text(SLIDE).replace('<a:off x="0" y="0"/>', '<a:off x="500000" y="0"/>'))
        assert compare(before, after).changed == []
