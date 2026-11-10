"""Round-trip the package and report, part by part, what survived the save.

The claim this project makes is narrow and checkable: open a deck, write it back, and every part that was not
edited is byte-identical. `OoxmlPackage.diff` answers that as a list of names, which is enough for a unit test
and not enough for a reader deciding whether to trust the tool on their own file. This module is the version
of the same check that says *which* parts came back, what the tool did to each one, and how large each one
was either side of the trip.

Two deliberate choices:

* The comparison is by content, not by name alone. A part that exists in both files but has a different byte
  count is reported as changed with both sizes, because "it is still there" is the reassurance that hides the
  damage.
* Archive order is checked separately. OOXML readers do not need a particular order, so a reordering would
  pass every content check while making two archives that a person comparing them sees as entirely different.
"""

from __future__ import annotations

import hashlib
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional, Sequence, Tuple

from .package import OoxmlPackage, PathLike

IDENTICAL = "identical"
CHANGED = "changed"
ADDED = "added"
DROPPED = "dropped"


@dataclass(frozen=True)
class PartReport:
    """One part, as it was and as it came back."""

    name: str
    status: str
    original_size: Optional[int]
    saved_size: Optional[int]
    original_sha256: Optional[str]
    saved_sha256: Optional[str]

    def describe(self) -> str:
        if self.status == IDENTICAL:
            return f"{self.name:<44} {self.status:<9} {self.original_size:>7} bytes"
        if self.status == CHANGED:
            return (
                f"{self.name:<44} {self.status:<9} {self.original_size:>7} -> {self.saved_size} bytes"
            )
        if self.status == ADDED:
            return f"{self.name:<44} {self.status:<9} {'-':>7} -> {self.saved_size} bytes"
        return f"{self.name:<44} {self.status:<9} {self.original_size:>7} -> - bytes"


@dataclass(frozen=True)
class RoundTripReport:
    """Every part of the archive, and the order it was written in."""

    outcome: str
    parts: Tuple[PartReport, ...]
    original_order: Tuple[str, ...]
    saved_order: Tuple[str, ...]

    @property
    def identical(self) -> List[PartReport]:
        return [part for part in self.parts if part.status == IDENTICAL]

    @property
    def changed(self) -> List[PartReport]:
        return [part for part in self.parts if part.status == CHANGED]

    @property
    def added(self) -> List[PartReport]:
        return [part for part in self.parts if part.status == ADDED]

    @property
    def dropped(self) -> List[PartReport]:
        return [part for part in self.parts if part.status == DROPPED]

    @property
    def order_preserved(self) -> bool:
        return self.original_order == self.saved_order

    def summarise(self) -> str:
        lines = [
            f"parts: {len(self.parts)} "
            f"({len(self.identical)} identical, {len(self.changed)} changed, "
            f"{len(self.added)} added, {len(self.dropped)} dropped)",
            f"archive order: {'preserved' if self.order_preserved else 'CHANGED'}",
            "",
        ]
        lines.extend(part.describe() for part in self.parts)
        return "\n".join(lines)


def _digest(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def compare(original: OoxmlPackage, saved: OoxmlPackage) -> RoundTripReport:
    """Every part of both packages, classified, in the original's order and then the new parts."""
    names = list(original.order) + [name for name in saved.order if name not in original.parts]
    parts: List[PartReport] = []
    for name in names:
        before = original.parts.get(name)
        after = saved.parts.get(name)
        if before is None:
            status = ADDED
        elif after is None:
            status = DROPPED
        elif before == after:
            status = IDENTICAL
        else:
            status = CHANGED
        parts.append(
            PartReport(
                name=name,
                status=status,
                original_size=None if before is None else len(before),
                saved_size=None if after is None else len(after),
                original_sha256=None if before is None else _digest(before),
                saved_sha256=None if after is None else _digest(after),
            )
        )
    clean = all(part.status == IDENTICAL for part in parts)
    outcome = "identical" if clean and original.order == saved.order else "differs"
    return RoundTripReport(
        outcome=outcome,
        parts=tuple(parts),
        original_order=tuple(original.order),
        saved_order=tuple(saved.order),
    )


def roundtrip(source: PathLike, destination: Optional[PathLike] = None) -> RoundTripReport:
    """Save the package and open the saved file again, then report what each part did.

    The destination defaults to a temporary file, because the question being asked is about the write path and
    not about where the bytes land; the temporary is removed before returning, so nothing accumulates.
    """
    package = OoxmlPackage.open(source)
    if destination is not None:
        package.save(destination)
        saved = OoxmlPackage.open(destination)
        return compare(package, saved)

    with tempfile.TemporaryDirectory() as directory:
        target = Path(directory) / "roundtrip.zip"
        package.save(target)
        saved = OoxmlPackage.open(target)
        return compare(package, saved)


def main(argv: Sequence[str]) -> int:
    """`python -m ooxml_writer.roundtrip deck.pptx`: print the report, exit non-zero if anything differs."""
    if len(argv) != 1:
        print("usage: python -m ooxml_writer.roundtrip <file>", file=sys.stderr)
        return 2
    report = roundtrip(argv[0])
    print(report.summarise())
    return 0 if report.outcome == "identical" else 1


if __name__ == "__main__":  # pragma: no cover - exercised through the CLI test
    sys.exit(main(sys.argv[1:]))
