"""
zoteroconv.docx_io -- DocxPackage: load every part of a .docx, expose the
XML trees that matter (document, footnotes, endnotes, custom.xml,
[Content_Types].xml, _rels/.rels, styles, numbering), and save with ALL
modified parts written.

Guarantee (tested in tests/test_docx_io.py): parts that were not marked
modified are written back BYTE-IDENTICAL (the original bytes are copied
verbatim). Parts that were modified are re-serialised from their lxml tree,
so word/document.xml is NOT byte-identical after a modification -- callers
must not claim otherwise.

The input file is never modified: save() refuses to write to the path the
package was loaded from, and writes via a temporary file + atomic replace.
"""

from __future__ import annotations

import os
import zipfile

from lxml import etree

DOCUMENT_PART = "word/document.xml"
FOOTNOTES_PART = "word/footnotes.xml"
ENDNOTES_PART = "word/endnotes.xml"
CUSTOM_PROPS_PART = "docProps/custom.xml"
CONTENT_TYPES_PART = "[Content_Types].xml"
ROOT_RELS_PART = "_rels/.rels"
STYLES_PART = "word/styles.xml"
NUMBERING_PART = "word/numbering.xml"

CUSTOM_PROPS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.custom-properties+xml"
)
CUSTOM_PROPS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
)


def serialize_xml(root) -> bytes:
    """Serialise an lxml tree the way OOXML parts are stored."""
    return etree.tostring(
        root, xml_declaration=True, encoding="UTF-8", standalone=True
    )


class DocxPackage:
    """In-memory representation of a .docx package."""

    def __init__(self, path: str | os.PathLike):
        self.path = os.path.abspath(os.fspath(path))
        if not os.path.isfile(self.path):
            raise FileNotFoundError(self.path)
        self._infos: list[zipfile.ZipInfo] = []
        self._parts: dict[str, bytes] = {}
        self._trees: dict[str, etree._Element] = {}
        self._added_bytes: dict[str, bytes] = {}
        self.modified: set[str] = set()
        with zipfile.ZipFile(self.path, "r") as zf:
            for info in zf.infolist():
                self._infos.append(info)
                self._parts[info.filename] = zf.read(info.filename)

    # ------------------------------------------------------------------ open
    @classmethod
    def open(cls, path: str | os.PathLike) -> "DocxPackage":
        return cls(path)

    # ----------------------------------------------------------------- parts
    def part_names(self) -> list[str]:
        names = [i.filename for i in self._infos]
        names += [n for n in self._added_bytes if n not in names]
        names += [n for n in self._trees if n not in names]
        return names

    def has_part(self, name: str) -> bool:
        return name in self._parts or name in self._added_bytes or name in self._trees

    def read_bytes(self, name: str) -> bytes:
        if name in self._parts:
            return self._parts[name]
        if name in self._added_bytes:
            return self._added_bytes[name]
        if name in self._trees:
            return serialize_xml(self._trees[name])
        raise KeyError(f"part not present in package: {name}")

    def tree(self, name: str):
        """Parsed root element of a part (cached). Raises KeyError if absent."""
        if name not in self._trees:
            self._trees[name] = etree.fromstring(self.read_bytes(name))
        return self._trees[name]

    def tree_or_none(self, name: str):
        return self.tree(name) if self.has_part(name) else None

    # ------------------------------------------------- convenience accessors
    @property
    def document(self):
        return self.tree(DOCUMENT_PART)

    @property
    def footnotes(self):
        return self.tree_or_none(FOOTNOTES_PART)

    @property
    def endnotes(self):
        return self.tree_or_none(ENDNOTES_PART)

    @property
    def custom_props(self):
        return self.tree_or_none(CUSTOM_PROPS_PART)

    @property
    def content_types(self):
        return self.tree(CONTENT_TYPES_PART)

    @property
    def root_rels(self):
        return self.tree(ROOT_RELS_PART)

    @property
    def styles(self):
        return self.tree_or_none(STYLES_PART)

    @property
    def numbering(self):
        return self.tree_or_none(NUMBERING_PART)

    # ---------------------------------------------------------- modification
    def mark_modified(self, name: str) -> None:
        """Flag an existing part as re-serialise-on-save."""
        if not self.has_part(name):
            raise KeyError(f"cannot mark missing part modified: {name}")
        self.modified.add(name)

    def set_tree(self, name: str, root) -> None:
        """Replace the parsed tree of a part and mark it modified."""
        self._trees[name] = root
        self.modified.add(name)

    def add_part_bytes(self, name: str, data: bytes) -> None:
        """Register a brand-new part (not in the original zip) from bytes."""
        if name in self._parts:
            raise ValueError(f"part already exists: {name} (use set_tree/mark_modified)")
        self._added_bytes[name] = data
        self.modified.add(name)

    def add_part_tree(self, name: str, root) -> None:
        """Register a brand-new part from an lxml tree."""
        if name in self._parts:
            raise ValueError(f"part already exists: {name} (use set_tree/mark_modified)")
        self._trees[name] = root
        self.modified.add(name)

    # ------------------------------------------------------------------ save
    def save(self, out_path: str | os.PathLike) -> str:
        """
        Write the package to out_path.

        - unmodified parts  -> original bytes, original ZipInfo metadata
        - modified parts    -> re-serialised tree (or added bytes)
        - added parts       -> appended after the original ones
        Refuses to write to the file this package was loaded from.
        """
        out_abs = os.path.abspath(os.fspath(out_path))
        if out_abs == self.path:
            raise ValueError(
                "refusing to save onto the input file "
                f"({self.path}); choose a different output path"
            )
        out_dir = os.path.dirname(out_abs)
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        tmp = out_abs + ".zoteroconv.tmp"
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
                written: set[str] = set()
                for info in self._infos:
                    name = info.filename
                    if name in self.modified and name in self._trees:
                        data = serialize_xml(self._trees[name])
                    else:
                        # untouched (or bytes-only) part: verbatim original
                        data = self._parts[name]
                    new_info = zipfile.ZipInfo(name, date_time=info.date_time)
                    new_info.compress_type = info.compress_type
                    new_info.external_attr = info.external_attr
                    new_info.internal_attr = info.internal_attr
                    new_info.create_system = info.create_system
                    zout.writestr(new_info, data)
                    written.add(name)
                for name in list(self._added_bytes) + [
                    n for n in self._trees if n not in self._parts
                ]:
                    if name in written:
                        continue
                    data = (
                        serialize_xml(self._trees[name])
                        if name in self._trees
                        else self._added_bytes[name]
                    )
                    new_info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
                    new_info.compress_type = zipfile.ZIP_DEFLATED
                    zout.writestr(new_info, data)
                    written.add(name)
            os.replace(tmp, out_abs)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
        return out_abs
