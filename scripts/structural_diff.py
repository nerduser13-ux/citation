#!/usr/bin/env python3
"""Structural diff between two .docx packages (keys, key order, run
structure -- not values).  Usage:

    python scripts/structural_diff.py A.docx B.docx [--only PREFIX,...]

Exit code 0 when no differences remain after --only filtering, 1 otherwise.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from zoteroconv.docx_io import DocxPackage            # noqa: E402
from zoteroconv.structure_diff import (                # noqa: E402
    diff_signatures,
    package_signature,
)


def main(argv: list[str]) -> int:
    if len(argv) < 3:
        print(__doc__)
        return 2
    a, b = DocxPackage.open(argv[1]), DocxPackage.open(argv[2])
    only = None
    if "--only" in argv:
        only = argv[argv.index("--only") + 1].split(",")
    lines = diff_signatures(package_signature(a), package_signature(b),
                            path=Path(argv[1]).name + " vs " + Path(argv[2]).name)
    if only:
        prefix = (Path(argv[1]).name + " vs " + Path(argv[2]).name + ".")
        kept = []
        for ln in lines:
            tail = ln.split(":", 1)[0]
            if tail.startswith(prefix):
                tail = tail[len(prefix):]
            if any(tail == o or tail.startswith(o + ".") or
                   tail.startswith(o + "[") for o in only):
                kept.append(ln)
        lines = kept
    if not lines:
        print("STRUCTURES MATCH (no differences in the compared signature)")
        return 0
    print(f"{len(lines)} structural difference(s):")
    for ln in lines:
        print("  " + ln)
    return 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
