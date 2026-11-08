"""Paragraph editing: the split-run trap, XML escaping, and byte-faithfulness of everything else."""

import io
import xml.etree.ElementTree as ElementTree
import zipfile

from ooxml_writer.package import OoxmlPackage
from ooxml_writer.paragraphs import (
    all_text,
    escape_xml,
    find_paragraph,
    paragraphs,
    replace_paragraph_text,
    set_first_matching_paragraph,
)

SLIDE = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>
    <p:sp><p:txBody>
      <a:p><a:r><a:rPr lang="en-US"/><a:t>Hello </a:t></a:r><a:r><a:rPr b="1"/><a:t>wor</a:t></a:r>\
<a:r><a:t>ld</a:t></a:r></a:p>
      <a:p><a:r><a:t>Second paragraph</a:t></a:r></a:p>
      <a:p></a:p>
    </p:txBody></p:sp>
  </p:spTree></p:cSld>
</p:sld>
"""


def build_package(slide_xml=SLIDE, extra=None):
    """A real zip in memory: the module edits packages, so the tests must not hand it a bare string."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", "<Types/>")
        archive.writestr("ppt/slides/slide1.xml", slide_xml)
        for name, data in (extra or {}).items():
            archive.writestr(name, data)
    buffer.seek(0)
    return OoxmlPackage.open(buffer)


def slide_text(package):
    return package.text("ppt/slides/slide1.xml")


# --- reading across runs ----------------------------------------------------


def test_paragraphs_flatten_split_runs():
    flattened = paragraphs(SLIDE)
    assert flattened[0].text == "Hello world"
    assert [run.text for run in flattened[0].runs] == ["Hello ", "wor", "ld"]
    assert flattened[1].text == "Second paragraph"


def test_a_split_phrase_is_findable_even_though_the_raw_xml_never_contains_it():
    assert "Hello world" not in SLIDE  # the whole point: the bytes do not contain the phrase
    found = find_paragraph(SLIDE, "Hello world")
    assert found is not None
    assert found.index == 0


def test_a_missing_phrase_returns_none():
    assert find_paragraph(SLIDE, "not in this slide") is None


# --- writing ----------------------------------------------------------------


def test_replacing_a_paragraph_keeps_first_run_formatting_and_empties_the_rest():
    package = build_package()
    previous = set_first_matching_paragraph(package, "ppt/slides/slide1.xml", "Hello world", "Goodbye")
    assert previous == "Hello world"

    xml = slide_text(package)
    # The first run's properties survive.
    assert '<a:rPr lang="en-US"/>' in xml
    # The bold run's properties also survive, but its text is now empty.
    assert '<a:rPr b="1"/><a:t></a:t>' in xml
    assert list(all_text(package, "ppt/slides/slide1.xml"))[0] == "Goodbye"


def test_only_the_targeted_paragraph_changes():
    package = build_package()
    replace_paragraph_text(package, "ppt/slides/slide1.xml", 0, "Changed")
    texts = list(all_text(package, "ppt/slides/slide1.xml"))
    assert texts[0] == "Changed"
    assert texts[1] == "Second paragraph"


def test_untouched_parts_are_byte_identical_after_a_save():
    package = build_package(extra={"ppt/media/image1.png": b"\x89PNG\r\n\x1a\nbinary"})
    replace_paragraph_text(package, "ppt/slides/slide1.xml", 1, "Edited")
    output = io.BytesIO()
    package.save(output)
    output.seek(0)
    with zipfile.ZipFile(output) as archive:
        assert archive.read("ppt/media/image1.png") == b"\x89PNG\r\n\x1a\nbinary"
        assert archive.read("[Content_Types].xml") == b"<Types/>"


def test_text_with_xml_metacharacters_is_escaped_not_injected():
    package = build_package()
    replace_paragraph_text(package, "ppt/slides/slide1.xml", 1, "a < b & c > d")
    xml = slide_text(package)
    assert "a &lt; b &amp; c &gt; d" in xml
    # And it reads back as the original characters.
    assert list(all_text(package, "ppt/slides/slide1.xml"))[1] == "a < b & c > d"
    # The result must still parse as XML, which is the real proof escaping worked.
    ElementTree.fromstring(xml)


def test_leading_and_trailing_spaces_are_preserved_explicitly():
    package = build_package()
    replace_paragraph_text(package, "ppt/slides/slide1.xml", 1, "  padded  ")
    xml = slide_text(package)
    assert 'xml:space="preserve"' in xml
    assert list(all_text(package, "ppt/slides/slide1.xml"))[1] == "  padded  "


def test_an_empty_paragraph_gains_a_run():
    package = build_package()
    previous = replace_paragraph_text(package, "ppt/slides/slide1.xml", 2, "filled in")
    assert previous == ""
    assert list(all_text(package, "ppt/slides/slide1.xml"))[2] == "filled in"


def test_out_of_range_paragraph_is_refused():
    package = build_package()
    assert replace_paragraph_text(package, "ppt/slides/slide1.xml", 99, "x") is None
    assert set_first_matching_paragraph(package, "ppt/slides/slide1.xml", "absent", "x") is None


def test_escape_xml_covers_the_three_characters_that_matter():
    assert escape_xml("a&b") == "a&amp;b"
    assert escape_xml("<t>") == "&lt;t&gt;"
    assert escape_xml("plain") == "plain"
