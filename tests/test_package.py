"""Round-trip fidelity, which is the only property that separates this from a renderer."""

import zipfile

import pytest

from ooxml_writer import OoxmlPackage, PartMissing
from ooxml_writer.package import replace_text_run

CONTENT_TYPES_XML = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="xml" ContentType="application/xml"/></Types>'
)

SLIDE_XML = (
    '<?xml version="1.0" encoding="UTF-8"?>'
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    "<p:cSld><p:spTree><p:sp><p:txBody><a:p>"
    "<a:r><a:t>Hello</a:t></a:r><a:r><a:t>World</a:t></a:r>"
    "</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"
)


def build(path, extra_parts=None):
    parts = {"[Content_Types].xml": CONTENT_TYPES_XML, "ppt/slides/slide1.xml": SLIDE_XML}
    parts.update(extra_parts or {})
    with zipfile.ZipFile(path, "w") as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return path


def test_open_then_save_without_edits_is_byte_identical(tmp_path):
    source = build(tmp_path / "in.pptx")
    before = OoxmlPackage.open(source)
    after_path = tmp_path / "out.pptx"
    before.save(after_path)

    comparison = OoxmlPackage.open(after_path)
    assert comparison.diff(before) == []


def test_replacing_a_text_run_leaves_the_rest_of_the_document_untouched(tmp_path):
    package = OoxmlPackage.open(build(tmp_path / "in.pptx"))
    original = package.text("ppt/slides/slide1.xml")

    replaced = replace_text_run(package, "ppt/slides/slide1.xml", 0, "Goodbye")

    assert replaced == "Hello"
    updated = package.text("ppt/slides/slide1.xml")
    assert "Goodbye" in updated and "World" in updated
    # Byte-for-byte the same apart from the run that was asked to change.
    assert updated.replace("Goodbye", "Hello") == original


def test_part_order_is_preserved(tmp_path):
    source = build(tmp_path / "in.pptx", {"docProps/app.xml": b"<app/>"})
    package = OoxmlPackage.open(source)
    output = tmp_path / "out.pptx"
    package.save(output)
    with zipfile.ZipFile(output) as archive:
        assert archive.namelist() == package.order


def test_missing_part_raises_rather_than_returning_empty(tmp_path):
    package = OoxmlPackage.open(build(tmp_path / "in.pptx"))
    with pytest.raises(PartMissing):
        package.part("ppt/slides/slide99.xml")


def test_a_package_without_content_types_is_refused(tmp_path):
    broken = tmp_path / "broken.pptx"
    with zipfile.ZipFile(broken, "w") as archive:
        archive.writestr("ppt/slides/slide1.xml", SLIDE_XML)
    with pytest.raises(PartMissing):
        OoxmlPackage.open(broken)


def test_diff_names_the_parts_that_changed(tmp_path):
    first = OoxmlPackage.open(build(tmp_path / "a.pptx"))
    second = OoxmlPackage.open(build(tmp_path / "b.pptx"))
    replace_text_run(second, "ppt/slides/slide1.xml", 1, "There")
    assert first.diff(second) == ["ppt/slides/slide1.xml"]


def test_replacing_a_run_that_does_not_exist_reports_rather_than_corrupting(tmp_path):
    package = OoxmlPackage.open(build(tmp_path / "in.pptx"))
    assert replace_text_run(package, "ppt/slides/slide1.xml", 9, "nope") is None
