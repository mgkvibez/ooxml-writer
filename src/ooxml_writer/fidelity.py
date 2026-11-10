"""A fidelity scoreboard: which elements survived an edit, named individually.

The round-trip report answers "did the file come back the same". This answers the question underneath it: when
a part did change, *what* changed. Counting elements by name is coarse on purpose, and the coarseness is the
point: "did the table survive" is answerable this way, and "does it still look right" is not, because that
needs a renderer.

So the scoreboard measures structure preserved, not appearance. Two parts can have an identical census and
still look different (a moved shape, a changed colour); a census difference always means something was added
or removed, and is named.

Reading uses ElementTree, which is legitimate in a way it is not for writing: nothing here produces XML, so
there is no namespace rewrite and no collapsed self-closing tag to worry about. What this module emits is
strings.
"""

import re
import sys
import tempfile
import xml.etree.ElementTree as ElementTree
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

from .package import OoxmlPackage, PathLike

# A parsed tree keeps namespaces as URIs and drops the prefixes, and "{http://schemas...}sp" is unreadable in
# a report. So the prefixes are read from the part's own ``xmlns:`` declarations before parsing: the part says
# what it calls things, and there is no table of namespaces here to go stale.
XMLNS = re.compile(rb'xmlns:([A-Za-z0-9_.-]+)="([^"]+)"')


def prefixes(data: bytes) -> "Dict[str, str]":
    """The namespace URI to prefix map a part declares, in the order it declares them."""
    declared: Dict[str, str] = {}
    for match in XMLNS.finditer(data):
        prefix = match.group(1).decode("utf-8", "replace")
        uri = match.group(2).decode("utf-8", "replace")
        declared.setdefault(uri, prefix)
    return declared


def local_name(tag: str, declared: "Dict[str, str]") -> str:
    """`{uri}sp` as `p:sp`, using the prefix this part declares for that URI.

    An element in a namespace the part never declared keeps its URI: a made-up prefix would be a claim about a
    vocabulary rather than a reading of the file.
    """
    if not tag.startswith("{"):
        return tag
    uri, name = tag[1:].split("}", 1)
    prefix = declared.get(uri)
    return f"{prefix}:{name}" if prefix else f"{{{uri}}}{name}"


def census(data: bytes) -> "Counter[str]":
    """How many of each element a part holds, by prefix-qualified name.

    A part that does not parse returns an empty census rather than raising: the caller is comparing two
    versions of a file, and "this part is not XML at all" is one of the answers, not an error.
    """
    counts: Counter = Counter()
    try:
        root = ElementTree.fromstring(data)
    except ElementTree.ParseError:
        return counts
    declared = prefixes(data)
    for element in root.iter():
        counts[local_name(element.tag, declared)] += 1
    return counts


@dataclass(frozen=True)
class PartFidelity:
    """One part, its census before and after, and the elements that differ."""

    name: str
    before: Counter
    after: Counter

    @property
    def lost(self) -> "Dict[str, int]":
        """Elements that were there and are not, with how many went."""
        return {
            element: self.before[element] - self.after[element]
            for element in self.before
            if self.before[element] > self.after[element]
        }

    @property
    def gained(self) -> "Dict[str, int]":
        return {
            element: self.after[element] - self.before[element]
            for element in self.after
            if self.after[element] > self.before[element]
        }

    @property
    def structural(self) -> bool:
        """True when the elements changed, which is all this can honestly claim to detect."""
        return bool(self.lost or self.gained)

    def describe(self) -> str:
        if not self.structural:
            return f"{self.name:<44} structure kept   {sum(self.before.values()):>5} elements"
        parts: List[str] = []
        for element, count in sorted(self.lost.items()):
            parts.append(f"-{count} {element}")
        for element, count in sorted(self.gained.items()):
            parts.append(f"+{count} {element}")
        return f"{self.name:<44} structure changed {' '.join(parts)}"


@dataclass(frozen=True)
class FidelityReport:
    """Every part, compared element by element."""

    parts: Tuple[PartFidelity, ...]

    @property
    def changed(self) -> List[PartFidelity]:
        return [part for part in self.parts if part.structural]

    @property
    def kept(self) -> List[PartFidelity]:
        return [part for part in self.parts if not part.structural]

    def summarise(self) -> str:
        total = sum(sum(part.before.values()) for part in self.parts)
        lines = [
            f"{len(self.kept)} of {len(self.parts)} parts kept their structure, "
            f"{total} elements counted in the original."
        ]
        for part in self.changed:
            lines.append(part.describe())
        return "\n".join(lines)


def compare(original: OoxmlPackage, saved: OoxmlPackage) -> FidelityReport:
    """Census every part of both packages, in the original's archive order."""
    parts: List[PartFidelity] = []
    for name in original.order:
        before = census(original.part(name))
        try:
            after = census(saved.part(name))
        except KeyError:
            after = Counter()
        parts.append(PartFidelity(name, before, after))
    for name in saved.order:
        if name not in original.order:
            parts.append(PartFidelity(name, Counter(), census(saved.part(name))))
    return FidelityReport(tuple(parts))


def scoreboard(source: PathLike) -> str:
    """Save a deck untouched, reopen it, and report what each part's structure did.

    The untouched save is the baseline worth publishing, because it is the claim the project makes: an editor
    that preserves everything nobody edited. Pairwise comparisons after an edit go through :func:`compare`
    directly, since only the caller knows which part it touched.
    """
    with tempfile.TemporaryDirectory() as directory:
        saved_path = Path(directory) / "saved.pptx"
        original = OoxmlPackage.open(source)
        original.save(saved_path)
        saved = OoxmlPackage.open(saved_path)
    return compare(original, saved).summarise()


def main(argv: Sequence[str]) -> int:
    if len(argv) != 2:
        print("usage: python -m ooxml_writer.fidelity <deck.pptx>", file=sys.stderr)
        return 2
    print(scoreboard(argv[1]))
    return 0


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main(sys.argv))
