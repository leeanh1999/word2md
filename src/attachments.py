"""Recover files attached to a Word document as OLE objects.

Word stores an attachment (a PDF, an HTML page, a ZIP…) as two separate parts:
a preview picture - almost always an .emf icon - and the file itself, wrapped
in an OLE container under `word/embeddings/`. Mammoth only ever sees the
picture, so an attachment used to come out of the converter as a meaningless
`.emf` image.

This module unwraps the container, writes the real file next to the Markdown
and hands mammoth a rewritten copy of the document where each object has
become a hyperlink to that file. Rewriting the source rather than
post-processing the HTML keeps the link exactly where the object was, even
inside tables, headers or footnotes.
"""

from __future__ import annotations

import hashlib
import io
import posixpath
import re
import struct
import zipfile
from dataclasses import dataclass, field
from pathlib import Path, PureWindowsPath
from urllib.parse import quote
from xml.sax.saxutils import escape, unescape

OLE10_NATIVE = "\x01Ole10Native"

HYPERLINK_RELATIONSHIP = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)

# Body parts that may hold an object; each has its own relationship file.
_BODY_PART = re.compile(
    r"^word/(document\d*|header\d+|footer\d+|footnotes|endnotes)\.xml$"
)

_OBJECT = re.compile(r"<w:object\b.*?</w:object>", re.S)
_OLE_OBJECT = re.compile(r"<o:OLEObject\b[^>]*?/?>", re.S)
_RELATIONSHIP = re.compile(r"<Relationship\b[^>]*?/?>", re.S)
_ATTRIBUTE = re.compile(r"([\w:.\-]+)\s*=\s*\"([^\"]*)\"|([\w:.\-]+)\s*=\s*'([^']*)'")

# Containers that hold a whole Office document instead of a packaged file.
_STREAM_SUFFIXES = (
    ("WordDocument", ".doc"),
    ("Workbook", ".xls"),
    ("Book", ".xls"),
    ("PowerPoint Document", ".ppt"),
    ("VisioDocument", ".vsd"),
)

# ProgIDs whose payload sits in a plain CONTENTS stream.
_PROGID_SUFFIXES = {
    "acroexch.document": ".pdf",
    "word.document": ".doc",
    "word.openxmldocument": ".docx",
    "excel.sheet": ".xls",
    "powerpoint.show": ".ppt",
}

_SIGNATURES = (
    (b"%PDF-", ".pdf"),
    (b"PK\x03\x04", ".zip"),
    (b"{\\rt", ".rtf"),
    (b"\x89PNG\r\n", ".png"),
    (b"\xff\xd8\xff", ".jpg"),
    (b"GIF8", ".gif"),
    (b"\xd0\xcf\x11\xe0", ".bin"),
)

_ILLEGAL_NAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


@dataclass
class PreparedDocument:
    """The outcome of scanning a .docx for attachments.

    `data` holds the rewritten document, or None when the file has no object
    worth touching and mammoth can read the original straight from disk.
    """

    data: bytes | None = None
    files: list[Path] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


# ------------------------------------------------------- Ole10Native decoding


class _Reader:
    """Little-endian cursor that raises rather than returning junk."""

    def __init__(self, data: bytes):
        self.data = data
        self.pos = 0

    def uint16(self) -> int:
        value = struct.unpack_from("<H", self.data, self.pos)[0]
        self.pos += 2
        return value

    def uint32(self) -> int:
        value = struct.unpack_from("<I", self.data, self.pos)[0]
        self.pos += 4
        return value

    def take(self, size: int) -> bytes:
        if size < 0 or self.pos + size > len(self.data):
            raise ValueError("Vượt quá độ dài dữ liệu.")
        chunk = self.data[self.pos : self.pos + size]
        self.pos += size
        return chunk

    def asciiz(self) -> str:
        end = self.data.find(b"\0", self.pos)
        if end < 0:
            raise ValueError("Chuỗi ANSI không kết thúc bằng NUL.")
        chunk = self.data[self.pos : end]
        self.pos = end + 1
        return chunk.decode("cp1252", "replace")

    def ansi_sized(self) -> str:
        return self.take(self.uint32()).split(b"\0", 1)[0].decode("cp1252", "replace")

    def utf16_sized(self) -> str:
        return self.take(self.uint32() * 2).decode("utf-16-le", "replace")

    @property
    def remaining(self) -> int:
        return len(self.data) - self.pos


def read_ole10_native(blob: bytes) -> tuple[str, bytes] | None:
    """Decode an `\\x01Ole10Native` stream into (file name, contents).

    Layout, per [MS-OLEDS] and Apache POI: total size, flags, the ANSI label
    and source path, a temp path, then the payload. Word also appends UTF-16
    copies of the strings, and those are the ones to trust: the ANSI label
    mangles anything outside the system code page, such as Vietnamese names.
    """
    try:
        reader = _Reader(blob)
        total = reader.uint32()
        if total + 4 > len(blob):
            return None
        if reader.uint16() != 2:  # unparsed variant: raw bytes, no file name
            return None
        label = reader.asciiz()
        source_path = reader.asciiz()
        reader.uint16()  # flags2
        reader.uint16()  # unknown, always 3 in the wild
        reader.ansi_sized()  # temporary path Word copied the file to
        data = reader.take(reader.uint32())
    except (struct.error, ValueError, IndexError):
        return None

    if not data:
        return None

    unicode_label = unicode_path = ""
    try:
        if reader.remaining > 4:
            reader.utf16_sized()  # temporary path again
            unicode_label = reader.utf16_sized()
            unicode_path = reader.utf16_sized()
    except (struct.error, ValueError):
        unicode_label = unicode_path = ""

    for candidate in (unicode_label, unicode_path, label, source_path):
        name = _clean_name(PureWindowsPath(candidate.strip()).name if candidate else "")
        if name:
            if not posixpath.splitext(name)[1]:
                name += _suffix_from_content(data)
            return name, data

    return f"attachment{_suffix_from_content(data)}", data


# ------------------------------------------------------------ OLE containers


def _open_ole(blob: bytes):
    import olefile

    if not olefile.isOleFile(io.BytesIO(blob)):
        return None
    return olefile.OleFileIO(io.BytesIO(blob))


def _read_container(blob: bytes, prog_id: str) -> tuple[str, bytes] | None:
    """Pull the original file out of a `word/embeddings/*.bin` container."""
    ole = _open_ole(blob)
    if ole is None:
        return None
    try:
        if ole.exists(OLE10_NATIVE):
            return read_ole10_native(ole.openstream(OLE10_NATIVE).read())

        for stream, suffix in _STREAM_SUFFIXES:
            if ole.exists(stream):
                return f"attachment{suffix}", blob

        if ole.exists("CONTENTS"):
            data = ole.openstream("CONTENTS").read()
            suffix = _suffix_from_prog_id(prog_id) or _suffix_from_content(data)
            return f"attachment{suffix}", data
    except Exception:  # noqa: BLE001 - an unreadable object keeps its icon
        return None
    finally:
        ole.close()

    return None


def _suffix_from_prog_id(prog_id: str) -> str:
    """`AcroExch.Document.DC` -> `.pdf`, ignoring the version suffix."""
    parts = prog_id.lower().split(".")
    while parts:
        suffix = _PROGID_SUFFIXES.get(".".join(parts))
        if suffix:
            return suffix
        parts.pop()
    return ""


def _suffix_from_content(data: bytes) -> str:
    head = data[:8]
    for signature, suffix in _SIGNATURES:
        if head.startswith(signature):
            return suffix
    lowered = data[:512].lstrip().lower()
    if lowered.startswith(b"<!doctype html") or lowered.startswith(b"<html"):
        return ".html"
    return ".dat"


def _clean_name(name: str) -> str:
    name = _ILLEGAL_NAME_CHARS.sub("_", name).strip(" .")
    return name[:120]


# --------------------------------------------------------------- output side


class _AttachmentStore:
    """Writes attachments to disk, once each, under a unique name."""

    def __init__(self, directory: Path | None):
        self.directory = directory
        self.written: list[Path] = []
        self._by_digest: dict[str, str] = {}
        self._taken: set[str] = set()

    def add(self, name: str, data: bytes) -> str:
        """Save `data` and return the file name to link to."""
        digest = hashlib.sha1(data).hexdigest()
        existing = self._by_digest.get(digest)
        if existing is not None:
            return existing

        name = self._unique(name)
        if self.directory is not None:
            self.directory.mkdir(parents=True, exist_ok=True)
            path = self.directory / name
            path.write_bytes(data)
            self.written.append(path)
        self._by_digest[digest] = name
        return name

    def _unique(self, name: str) -> str:
        stem, dot, suffix = name.rpartition(".")
        if not dot:
            stem, suffix = name, ""
        candidate = name
        counter = 1
        while candidate.lower() in self._taken or (
            self.directory is not None and (self.directory / candidate).exists()
        ):
            candidate = f"{stem} ({counter}){dot}{suffix}"
            counter += 1
        self._taken.add(candidate.lower())
        return candidate


# ------------------------------------------------------------- xml rewriting


def _attributes(tag: str) -> dict[str, str]:
    found = {}
    for match in _ATTRIBUTE.finditer(tag):
        key = match.group(1) or match.group(3)
        value = match.group(2) if match.group(1) else match.group(4)
        found[key] = unescape(value or "")
    return found


def _rels_name(part: str) -> str:
    folder, _, name = part.rpartition("/")
    return f"{folder}/_rels/{name}.rels"


def _read_relationships(archive: zipfile.ZipFile, part: str) -> dict[str, dict]:
    try:
        xml = archive.read(_rels_name(part)).decode("utf-8")
    except KeyError:
        return {}
    relationships = {}
    for match in _RELATIONSHIP.finditer(xml):
        attributes = _attributes(match.group(0))
        identifier = attributes.get("Id")
        if identifier:
            relationships[identifier] = attributes
    return relationships


def _add_relationships(xml: str, added: list[tuple[str, str]]) -> str:
    entries = "".join(
        f'<Relationship Id="{identifier}" Type="{HYPERLINK_RELATIONSHIP}"'
        f' Target="{escape(target, {chr(34): "&quot;"})}" TargetMode="External"/>'
        for identifier, target in added
    )
    return xml.replace("</Relationships>", f"{entries}</Relationships>", 1)


def _link_xml(relationship_id: str | None, label: str) -> str:
    run = f'<w:r><w:t xml:space="preserve">{escape(label)}</w:t></w:r>'
    if relationship_id is None:
        return run
    return f'<w:hyperlink r:id="{relationship_id}">{run}</w:hyperlink>'


class _Rewriter:
    """Replaces the OLE objects of one body part with links."""

    def __init__(self, archive: zipfile.ZipFile, store: _AttachmentStore, folder: str):
        self.archive = archive
        self.store = store
        self.folder = folder
        self.warnings: list[str] = []
        self._counter = 0

    def rewrite(
        self, xml: str, relationships: dict[str, dict]
    ) -> tuple[str, list[tuple[str, str]]]:
        added: list[tuple[str, str]] = []
        taken = set(relationships)

        def next_id() -> str:
            while True:
                self._counter += 1
                identifier = f"rIdAttachment{self._counter}"
                if identifier not in taken:
                    return identifier

        def replace(match: re.Match) -> str:
            block = match.group(0)
            resolved = self._resolve(block, relationships)
            if resolved is None:
                return block
            label, href, icon = resolved
            if href is None:
                link = _link_xml(None, label)
            else:
                identifier = next_id()
                added.append((identifier, href))
                link = _link_xml(identifier, label)
            return link if icon else block + link

        return _OBJECT.sub(replace, xml), added

    def _resolve(
        self, block: str, relationships: dict[str, dict]
    ) -> tuple[str, str | None, bool] | None:
        """Return (label, href, replaces_icon) for one <w:object> block."""
        tag = _OLE_OBJECT.search(block)
        if tag is None:
            return None
        attributes = _attributes(tag.group(0))
        relationship = relationships.get(attributes.get("r:id", ""))
        if relationship is None:
            return None

        target = relationship.get("Target", "")
        icon = attributes.get("DrawAspect", "").lower() == "icon"
        prog_id = attributes.get("ProgID", "")

        # A linked object keeps its file outside the document: point at it.
        if relationship.get("TargetMode", "").lower() == "external":
            if not target:
                return None
            return PureWindowsPath(target).name or target, target, icon

        entry = posixpath.normpath(posixpath.join("word", target))
        try:
            blob = self.archive.read(entry)
        except KeyError:
            self.warnings.append(f"Không tìm thấy dữ liệu đính kèm: {target}")
            return None

        suffix = posixpath.splitext(entry)[1].lower()
        if suffix and suffix != ".bin":
            # Embedded Office files are stored as-is, already correctly named.
            name = _clean_name(posixpath.basename(entry))
            unwrapped = (name, blob)
        else:
            unwrapped = _read_container(blob, prog_id)

        if unwrapped is None:
            return None

        name, data = unwrapped
        stored = self.store.add(name, data)
        href = quote(posixpath.join(self.folder, stored), safe="/")
        return stored, (href if self.store.directory is not None else None), icon


def prepare_attachments(
    source: Path, directory: Path | None = None
) -> PreparedDocument:
    """Extract the attachments of a .docx and rewrite it to link to them.

    With `directory` set, every attachment is written there and the objects
    become hyperlinks. Without one - the outline preview, which must not touch
    the disk - the objects become their plain file name, so the text of the
    document stays the same either way.
    """
    source = Path(source)
    prepared = PreparedDocument()
    store = _AttachmentStore(directory)

    try:
        with zipfile.ZipFile(source) as archive:
            parts = [name for name in archive.namelist() if _BODY_PART.match(name)]
            if not parts:
                return prepared
            folder = directory.name if directory is not None else ""
            rewriter = _Rewriter(archive, store, folder)
            edits: dict[str, str] = {}

            for part in parts:
                xml = archive.read(part).decode("utf-8")
                if "<w:object" not in xml:
                    continue
                relationships = _read_relationships(archive, part)
                updated, added = rewriter.rewrite(xml, relationships)
                if updated == xml:
                    continue
                edits[part] = updated
                if added:
                    rels = _rels_name(part)
                    edits[rels] = _add_relationships(
                        archive.read(rels).decode("utf-8"), added
                    )

            prepared.warnings.extend(rewriter.warnings)
            if edits:
                prepared.data = _repack(archive, edits)
    except zipfile.BadZipFile:
        return prepared
    except Exception as exc:  # noqa: BLE001 - never lose the document over this
        prepared.warnings.append(f"Không xử lý được file đính kèm: {exc}")
        return PreparedDocument(warnings=prepared.warnings)

    prepared.files = store.written
    return prepared


def _repack(archive: zipfile.ZipFile, edits: dict[str, str]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as target:
        for item in archive.infolist():
            replacement = edits.get(item.filename)
            if replacement is None:
                target.writestr(item, archive.read(item.filename))
            else:
                target.writestr(item.filename, replacement.encode("utf-8"))
    return buffer.getvalue()


__all__ = [
    "PreparedDocument",
    "prepare_attachments",
    "read_ole10_native",
]
