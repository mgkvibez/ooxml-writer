"""Shape-level editing: find a shape by name, and change only its text body.

A slide is a list of shapes, and "set the text" means different things for different ones, so this module is
deliberately small and explicit:

* It addresses a shape by the name PowerPoint shows in the selection pane: ``Title 1``,
  ``Content Placeholder 2``, ``TextBox 3``. Editing by index is what the paragraph functions already do; by
  name is what a caller with the slide in front of them can actually state.
* It keeps the shape's own run properties, the same way
  :func:`ooxml_writer.paragraphs.replace_paragraph_text` does, by collapsing the new text into the first run
  and removing the rest.
* A body of several paragraphs becomes one paragraph. Deleting the extra ``<a:p>`` elements is the point:
  replacing three bullets with a sentence and leaving two empty bullets behind is a slide that looks broken,
  and an empty ``<a:p>`` still draws its bullet glyph. Removing them loses only the run properties of
  paragraphs that are not being kept.

Everything is string-surgical, for the reasons given in :mod:`ooxml_writer.paragraphs`: a DOM round-trip
rewrites namespace declarations and collapses self-closing tags, which is the fidelity this project exists to
preserve.
"""

import re
from typing import List, NamedTuple, Optional

from .package import OoxmlPackage
from .paragraphs import _paragraph_spans, _unescape_xml, escape_xml, replace_paragraph_text

SHAPE = re.compile(r"<p:sp\b[^>]*>.*?</p:sp>", re.DOTALL)
CONTROL = re.compile(r"<p:cNvPr\b[^>]*>")
ID_ATTR = re.compile(r'\bid="(\d+)"')
NAME_ATTR = re.compile(r'\bname="([^"]*)"')
TEXT_BODY = re.compile(r"<p:txBody>(.*?)</p:txBody>", re.DOTALL)
PARAGRAPH_ELEMENT = re.compile(r"<a:p\b[^>]*/>|<a:p\b[^>]*>.*?</a:p>", re.DOTALL)
TEXT = re.compile(r"<a:t(?:\s[^>]*)?>(.*?)</a:t>|<a:t(?:\s[^>]*)?/>", re.DOTALL)


class ShapeError(Exception):
    """A shape that cannot be addressed or edited as asked: missing, ambiguous, or without a text body."""


class ShapeText(NamedTuple):
    """One ``<p:sp>``: the name a person uses, the id PowerPoint uses, and the text as a reader sees it."""

    name: str
    shape_id: str
    text: str
    paragraphs: int


def _text_in(xml: str) -> str:
    return _unescape_xml("".join(match.group(1) or "" for match in TEXT.finditer(xml)))


def shapes(xml: str) -> List[ShapeText]:
    """Every shape in a slide part, in document order.

    Shapes inside a group are found too: a ``<p:sp>`` is a ``<p:sp>`` whether or not a group wraps it, and the
    scan runs over the whole part rather than over the children of ``<p:spTree>``.
    """
    found: List[ShapeText] = []
    for match in SHAPE.finditer(xml):
        block = match.group(0)
        control = CONTROL.search(block)
        name = NAME_ATTR.search(control.group(0)) if control else None
        shape_id = ID_ATTR.search(control.group(0)) if control else None
        body = TEXT_BODY.search(block)
        elements = PARAGRAPH_ELEMENT.findall(body.group(1)) if body else []
        found.append(
            ShapeText(
                name=_unescape_xml(name.group(1)) if name else "",
                shape_id=shape_id.group(1) if shape_id else "",
                # newline-joined, so three bullets read as three lines rather than one run-on string
                text="\n".join(_text_in(element) for element in elements),
                paragraphs=len(elements),
            )
        )
    return found


def find_shape(xml: str, name: str) -> Optional[ShapeText]:
    """The shape with this exact name, else the only one that matches it ignoring case.

    PowerPoint keeps names unique within a slide, so an exact match is the normal path. Case-insensitive
    matching is for callers typing a name from memory, and it refuses when two shapes would both answer,
    because picking one of them silently is how a title gets overwritten by its subtitle.
    """
    candidates = shapes(xml)
    exact = [shape for shape in candidates if shape.name == name]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        raise ShapeError(f"{len(exact)} shapes are named {name!r}; names have to be unique to address one")
    folded = [shape for shape in candidates if shape.name.casefold() == name.casefold()]
    if len(folded) == 1:
        return folded[0]
    if len(folded) > 1:
        listed = ", ".join(shape.name for shape in folded)
        raise ShapeError(f"{len(folded)} shapes match {name!r} ignoring case: {listed}")
    return None


def _shape_block(xml: str, name: str) -> "tuple[int, int]":
    target = find_shape(xml, name)
    if target is None:
        listed = ", ".join(shape.name for shape in shapes(xml))
        raise ShapeError(f"no shape named {name!r} in this part: {listed}")
    for match in SHAPE.finditer(xml):
        block = match.group(0)
        control = CONTROL.search(block)
        found = NAME_ATTR.search(control.group(0)) if control else None
        if found and _unescape_xml(found.group(1)) == target.name:
            return match.start(), match.end()
    raise ShapeError(f"shape {name!r} was listed but could not be located")  # pragma: no cover - defensive


def set_shape_text(package: OoxmlPackage, part_name: str, shape_name: str, text: str) -> str:
    """Set the text of a shape's body, keeping the first run's formatting. Returns the text it replaced.

    The body ends up as one paragraph. A shape with no text body at all is an error rather than a silent
    no-op: a caller who asked to set a picture's text asked the wrong question, and saying so is the useful
    answer.
    """
    xml = package.text(part_name)
    start, end = _shape_block(xml, shape_name)
    body = TEXT_BODY.search(xml, start, end)
    if body is None:
        raise ShapeError(f"shape {shape_name!r} has no text body")
    body_start, body_end = body.start(1), body.end(1)

    elements = [
        (match.start(), match.end()) for match in PARAGRAPH_ELEMENT.finditer(xml, body_start, body_end)
    ]
    previous = "\n".join(_text_in(xml[start_of:end_of]) for start_of, end_of in elements)

    if not elements:
        # A body with no paragraph at all: give the text somewhere to live, before </p:txBody>.
        opening = '<a:p><a:r><a:t>' + escape_xml(text) + "</a:t></a:r></a:p>"
        package.replace(part_name, xml[:body_end] + opening + xml[body_end:])
        return previous

    if len(elements) > 1:
        # From the end backwards, so earlier offsets stay valid.
        trimmed = xml
        for element_start, element_end in reversed(elements[1:]):
            trimmed = trimmed[:element_start] + trimmed[element_end:]
        package.replace(part_name, trimmed)

    # The first paragraph's index, counted over the whole part the way replace_paragraph_text counts.
    current = package.text(part_name)
    first_open = PARAGRAPH_ELEMENT.search(current, body_start, body_end)
    assert first_open is not None  # there was at least one element and we removed only the later ones
    index = sum(1 for span in _paragraph_spans(current) if span[1] <= first_open.start())
    replace_paragraph_text(package, part_name, index, text)
    return previous


def shape_text(package: OoxmlPackage, part_name: str, shape_name: str) -> Optional[str]:
    """What a shape says now, or None when there is no shape with that name."""
    target = find_shape(package.text(part_name), shape_name)
    return None if target is None else target.text
