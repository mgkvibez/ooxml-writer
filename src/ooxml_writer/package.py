"""Open an OOXML file, change one part, write it back byte-faithful everywhere else.

Renderers for PPTX and DOCX already exist. None of them lets you *edit* a document and save a file that the
originating application opens unchanged. The reason is the write path: an OOXML file is a zip of parts with
relationships between them, and most editors rebuild the archive from a model, which silently drops everything
the model did not represent.

This module takes the opposite approach. The package is the source of truth; parts stay byte arrays until
something asks to parse one; only the parts that were edited are re-serialised, and everything else is copied
verbatim. The test that matters follows from that: open, save without edits, and every part must be identical.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Union

PathLike = Union[str, Path]

# Parts that must exist in a valid presentation, and the relationship type that points at the main one.
CONTENT_TYPES = "[Content_Types].xml"


class PartMissing(KeyError):
    """A part was requested that the package does not contain."""


@dataclass
class OoxmlPackage:
    """An immutable-ish zip: parts are loaded on demand and written back untouched unless replaced.

Order matters for round-trip fidelity. OOXML readers do not require a particular part order, but tools and
humans
    diffing two archives do, so the original order is preserved and new parts are appended.
    """

    parts: Dict[str, bytes] = field(default_factory=dict)
    order: List[str] = field(default_factory=list)

    @classmethod
    def open(cls, path: PathLike) -> "OoxmlPackage":
        package = cls()
        with zipfile.ZipFile(path, "r") as archive:
            for info in archive.infolist():
                package.parts[info.filename] = archive.read(info.filename)
                package.order.append(info.filename)
        if CONTENT_TYPES not in package.parts:
            raise PartMissing(CONTENT_TYPES)
        return package

    def part(self, name: str) -> bytes:
        try:
            return self.parts[name]
        except KeyError:
            raise PartMissing(name) from None

    def text(self, name: str) -> str:
        return self.part(name).decode("utf-8")

    def replace(self, name: str, data: Union[bytes, str]) -> None:
        payload = data.encode("utf-8") if isinstance(data, str) else data
        if name not in self.parts:
            self.order.append(name)
        self.parts[name] = payload

    def matches(self, pattern: str) -> Iterable[str]:
        """Part names ending with a suffix, in archive order (e.g. 'slide1.xml')."""
        return [name for name in self.order if name.endswith(pattern)]

    def save(self, path: PathLike) -> None:
        """Write the package, keeping every untouched part byte-identical."""
        with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
            for name in self.order:
                archive.writestr(name, self.parts[name])

    def diff(self, other: "OoxmlPackage") -> List[str]:
        """Names of parts that differ, so a round-trip test can name what was lost."""
        changed: List[str] = []
        for name in sorted(set(self.parts) | set(other.parts)):
            if self.parts.get(name) != other.parts.get(name):
                changed.append(name)
        return changed


def replace_text_run(package: OoxmlPackage, part_name: str, index: int, text: str) -> Optional[str]:
    """Replace the contents of the nth ``<a:t>`` run in a slide, leaving every other byte alone.

Deliberately string-surgical rather than DOM-based: a full XML round-trip reorders namespace declarations
and
    rewrites self-closing tags, which is exactly the fidelity loss this project exists to avoid.
    """
    import re

    xml = package.text(part_name)
    matches = list(re.finditer(r"<a:t>(.*?)</a:t>", xml, re.DOTALL))
    if index >= len(matches):
        return None
    match = matches[index]
    updated = xml[: match.start(1)] + text + xml[match.end(1) :]
    package.replace(part_name, updated)
    return match.group(1)
