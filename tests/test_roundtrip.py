"""The round-trip report, checked against a real deck and against the failures it exists to catch.

A report that only ever says "identical" is worthless, so every classification is tested by constructing the
thing it describes: a changed part, an added one, a dropped one, and an archive whose order moved while its
contents did not.
"""

import zipfile

import pytest

from ooxml_writer import OoxmlPackage
from ooxml_writer.paragraphs import replace_paragraph_text
from ooxml_writer.roundtrip import (
    ADDED,
    CHANGED,
    DROPPED,
    IDENTICAL,
    compare,
    main,
    roundtrip,
)

FIXTURE = "tests/fixtures/deck.pptx"
SLIDE = "ppt/slides/slide1.xml"


@pytest.fixture()
def package():
    return OoxmlPackage.open(FIXTURE)


def test_the_real_deck_round_trips_part_for_part():
    report = roundtrip(FIXTURE)
    assert report.outcome == "identical"
    assert report.order_preserved
    assert len(report.parts) == len(package_parts(FIXTURE))
    assert report.changed == []
    assert report.added == []
    assert report.dropped == []


def package_parts(path):
    with zipfile.ZipFile(path) as archive:
        return archive.namelist()


def test_an_edited_slide_is_named_and_everything_else_is_still_identical(package):
    index = next(number for number, entry in enumerate(_paragraphs(package)) if entry.text.strip())
    replace_paragraph_text(package, SLIDE, index, "Edited for the round-trip report")
    saved = _save_and_reopen(package)
    report = compare(OoxmlPackage.open(FIXTURE), saved)
    assert [part.name for part in report.changed] == [SLIDE]
    assert report.outcome == "differs"
    assert len(report.identical) == len(report.parts) - 1
    changed = report.changed[0]
    assert changed.original_sha256 != changed.saved_sha256
    assert changed.original_size is not None and changed.saved_size is not None


def test_a_changed_part_reports_both_sizes_rather_than_only_that_it_exists(package):
    package.replace(SLIDE, package.text(SLIDE).replace("</p:spTree>", "<!-- note --></p:spTree>"))
    report = compare(OoxmlPackage.open(FIXTURE), _save_and_reopen(package))
    changed = report.changed[0]
    assert changed.status == CHANGED
    assert changed.saved_size > changed.original_size
    assert "->" in changed.describe()


def test_an_added_part_is_reported_as_added(package):
    package.replace("ppt/notesSlides/notesSlide1.xml", "<p:notes/>")
    report = compare(OoxmlPackage.open(FIXTURE), _save_and_reopen(package))
    added = report.added
    assert [part.name for part in added] == ["ppt/notesSlides/notesSlide1.xml"]
    assert added[0].original_size is None
    assert added[0].status == ADDED


def test_a_dropped_part_is_reported_as_dropped_not_as_missing_from_the_report(package):
    package.parts.pop("ppt/printerSettings/printerSettings1.bin")
    package.order.remove("ppt/printerSettings/printerSettings1.bin")
    report = compare(OoxmlPackage.open(FIXTURE), _save_and_reopen(package))
    dropped = report.dropped
    assert [part.name for part in dropped] == ["ppt/printerSettings/printerSettings1.bin"]
    assert dropped[0].saved_size is None
    assert dropped[0].status == DROPPED


def test_a_reordered_archive_is_caught_even_though_every_part_is_identical(package):
    # the pathological case: byte-identical parts, different order, so every content check passes and a person
    # comparing the two archives sees a different file
    reordered = OoxmlPackage()
    for name in reversed(package.order):
        reordered.replace(name, package.part(name))
    report = compare(package, reordered)
    assert all(part.status == IDENTICAL for part in report.parts)
    assert not report.order_preserved
    assert report.outcome == "differs"


def test_the_summary_states_the_counts_and_the_order_verdict(package):
    report = compare(package, _save_and_reopen(package))
    summary = report.summarise()
    assert "identical" in summary
    assert "archive order: preserved" in summary
    assert SLIDE in summary


def test_the_report_keeps_the_original_order_and_appends_what_is_new(package):
    package.replace("docProps/custom.xml", "<Properties/>")
    report = compare(OoxmlPackage.open(FIXTURE), _save_and_reopen(package))
    assert report.parts[-1].name == "docProps/custom.xml"


def test_the_cli_refuses_to_run_without_a_file(capsys):
    assert main([]) == 2
    assert "usage:" in capsys.readouterr().err


def test_the_cli_exits_zero_on_a_deck_that_comes_back_unchanged(capsys):
    assert main([FIXTURE]) == 0
    output = capsys.readouterr().out
    assert "0 changed" in output
    assert "archive order: preserved" in output


def _save_and_reopen(package):
    import tempfile

    with tempfile.TemporaryDirectory() as directory:
        target = f"{directory}/out.pptx"
        package.save(target)
        return OoxmlPackage.open(target)


def _paragraphs(package):
    from ooxml_writer.paragraphs import paragraphs

    return paragraphs(package.text(SLIDE))
