"""Paragraph-level text editing that survives PowerPoint's habit of splitting a word across runs.

The trap this module exists for: PowerPoint stores "Hello world" as *several* ``<a:r>`` runs, each carrying
its own formatting and the editor splits wherever the formatting or the editing history changed. So a naive
``xml.replace("Hello world", ...)`` on a real slide usually finds nothing, and a reader searching raw XML
concludes the text is not there. Both must work on the paragraph concatenated text, while the write
still has to happen at run level so formatting of the first run survives.

Everything here is string-surgical on purpose: a DOM round-trip rewrites namespace declarations and collapses
self-closing tags, which is precisely the fidelity loss this project exists to avoid.
"""

import re
from typing import Iterable, List, NamedTuple, Optional

from .package import OoxmlPackage

PARAGRAPH = re.compile(r"<a:p>(.*?)</a:p>", re.DOTALL)
RUN = re.compile(r"<a:r>.*?</a:r>", re.DOTALL)
TEXT = re.compile(r"<a:t(?:\s[^>]*)?>(.*?)</a:t>|<a:t(?:\s[^>]*)?/>", re.DOTALL)

ESCAPES = (("&", "&amp;"), ("<", "&lt;"), (">", "&gt;"))


class RunText(NamedTuple):
    """One ``<a:t>`` occurrence: what it says, and where the text lives in the part's bytes-as-text."""

    text: str
    start: int
    end: int


class ParagraphText(NamedTuple):
    """A paragraph flattened to what a human would read, plus the runs it was flattened from."""

    index: int
    text: str
    runs: List[RunText]


def escape_xml(text: str) -> str:
    """Escape text for XML character data. Quotes are irrelevant in character data but harmless to escape."""
    escaped = text
    for character, entity in ESCAPES:
        escaped = escaped.replace(character, entity)
    return escaped


def _unescape_xml(text: str) -> str:
    return text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")


def _paragraph_spans(xml: str) -> List[tuple]:
    return [(match.start(1), match.end(1)) for match in PARAGRAPH.finditer(xml)]


def _runs_in(xml: str, start: int, end: int) -> List[RunText]:
    runs: List[RunText] = []
    for match in TEXT.finditer(xml, start, end):
        if match.group(1) is None:
            # A self-closing <a:t/> is an empty run: worth reporting, but there is no interior to replace.
            position = match.start()
            runs.append(RunText("", position, position))
            continue
        runs.append(RunText(_unescape_xml(match.group(1)), match.start(1), match.end(1)))
    return runs


def paragraphs(xml: str) -> List[ParagraphText]:
    """Every paragraph, flattened. This is what a search must use, because runs split words apart."""
    flattened: List[ParagraphText] = []
    for index, (start, end) in enumerate(_paragraph_spans(xml)):
        runs = _runs_in(xml, start, end)
        flattened.append(ParagraphText(index, "".join(run.text for run in runs), runs))
    return flattened


def find_paragraph(xml: str, needle: str) -> Optional[ParagraphText]:
    """First paragraph whose flattened text contains `needle`, or None. Case-sensitive."""
    for paragraph in paragraphs(xml):
        if needle in paragraph.text:
            return paragraph
    return None


def replace_paragraph_text(
    package: OoxmlPackage,
    part_name: str,
    paragraph_index: int,
    text: str,
) -> Optional[str]:
    """Set a paragraph's text, keeping the first run's formatting and emptying the runs after it.

    Returns the paragraph's previous flattened text, or None when the index does not exist.

    Collapsing into the first run rather than deleting runs is deliberate: deleting runs drops the run
    properties (font, size, colour) the look depends on, and PowerPoint then reopens the slide in
    default font. Emptying the trailing runs leaves their properties in place and changes only what is read.
    """
    xml = package.text(part_name)
    spans = _paragraph_spans(xml)
    if paragraph_index >= len(spans):
        return None

    start, end = spans[paragraph_index]
    runs = _runs_in(xml, start, end)
    previous = "".join(run.text for run in runs)

    if not runs:
        # A paragraph with no text run at all: insert one so the text has somewhere to live.
        insertion = '<a:r><a:t>' + escape_xml(text) + "</a:t></a:r>"
        package.replace(part_name, xml[:end] + insertion + xml[end:])
        return previous

    # Rewrite from the end backwards so earlier offsets stay valid.
    updated = xml
    needs_preserve = text != text.strip()
    for run in reversed(runs):
        if run.start == run.end:  # self-closing empty run, nothing to hold text
            continue
        replacement = "" if run is not runs[0] else escape_xml(text)
        if run is runs[0] and needs_preserve:
            replacement = escape_xml(text)
        updated = updated[: run.start] + replacement + updated[run.end:]

    if needs_preserve:
        updated = _ensure_space_preserve(updated, runs[0].start)

    package.replace(part_name, updated)
    return previous


def _ensure_space_preserve(xml: str, text_start: int) -> str:
    """Leading or trailing spaces are stripped by XML consumers unless xml:space="preserve" is set."""
    opening = xml.rfind("<a:t", 0, text_start)
    if opening == -1:
        return xml
    tag_end = xml.find(">", opening)
    if tag_end == -1 or "xml:space" in xml[opening:tag_end]:
        return xml
    return xml[:tag_end] + ' xml:space="preserve"' + xml[tag_end:]


def set_first_matching_paragraph(
    package: OoxmlPackage,
    part_name: str,
    needle: str,
    text: str,
) -> Optional[str]:
    """Find a phrase however it was split across runs, then replace that whole paragraph."""
    xml = package.text(part_name)
    paragraph = find_paragraph(xml, needle)
    if paragraph is None:
        return None
    return replace_paragraph_text(package, part_name, paragraph.index, text)


def all_text(package: OoxmlPackage, part_name: str) -> Iterable[str]:
    """Every paragraph of a part, in document order."""
    for paragraph in paragraphs(package.text(part_name)):
        yield paragraph.text
