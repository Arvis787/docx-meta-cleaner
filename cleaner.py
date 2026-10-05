"""Remove Office package properties without rewriting document content."""

from __future__ import annotations

import os
import posixpath
import shutil
import tempfile
import zipfile
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET

WORD_EXTENSIONS = frozenset({'.docx', '.docm', '.dotx', '.dotm'})
POWERPOINT_EXTENSIONS = frozenset({'.pptx', '.pptm', '.potx', '.potm', '.ppsx', '.ppsm'})
EXCEL_EXTENSIONS = frozenset({'.xlsx', '.xlsm', '.xltx', '.xltm'})
OFFICE_EXTENSIONS = WORD_EXTENSIONS | POWERPOINT_EXTENSIONS | EXCEL_EXTENSIONS
FIXED_ZIP_DATE = (1980, 1, 1, 0, 0, 0)
CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
PROPERTY_TYPES = frozenset({
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
    'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/extended-properties',
    'http://purl.oclc.org/ooxml/officeDocument/relationships/custom-properties',
})


@dataclass(frozen=True)
class CleanResult:
    path: Path
    removed_items: int
    backup_path: Path | None


def is_office_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in OFFICE_EXTENSIONS and not path.name.startswith('~$')


def iter_office_files(folder: Path, recursive: bool):
    for path in folder.rglob('*') if recursive else folder.iterdir():
        if is_office_file(path):
            yield path


def package_name(value: str) -> str:
    return posixpath.normpath(unquote(value).replace('\\', '/').lstrip('/'))


def metadata_parts(archive: zipfile.ZipFile) -> set[str]:
    parts = {name for name in archive.namelist() if package_name(name).startswith('docProps/')}
    root = ET.fromstring(archive.read('_rels/.rels'))
    for rel in root:
        if rel.get('Type') in PROPERTY_TYPES and rel.get('TargetMode') != 'External':
            parts.add(package_name(rel.get('Target', '')))
    # Include relationships belonging to removed parts, even for nonstandard paths.
    for part in tuple(parts):
        folder, name = posixpath.split(part)
        parts.add(posixpath.join(folder, '_rels', name + '.rels'))
    return {package_name(name) for name in parts}


def clean_content_types_xml(data: bytes, parts: set[str]) -> tuple[bytes, int]:
    ET.register_namespace('', CONTENT_TYPES_NS)
    root = ET.fromstring(data)
    removed = 0
    for child in list(root):
        if child.tag == f'{{{CONTENT_TYPES_NS}}}Override' and package_name(child.get('PartName', '')) in parts:
            root.remove(child)
            removed += 1
    return (ET.tostring(root, encoding='utf-8', xml_declaration=True) if removed else data), removed


def clean_relationships_xml(data: bytes, source: str, parts: set[str]) -> tuple[bytes, int]:
    ET.register_namespace('', RELS_NS)
    root = ET.fromstring(data)
    removed = 0
    for child in list(root):
        target = child.get('Target', '').replace('\\', '/')
        target_part = package_name(target if target.startswith('/') else posixpath.join(source, target))
        if child.tag == f'{{{RELS_NS}}}Relationship' and (
            child.get('Type') in PROPERTY_TYPES
            or (child.get('TargetMode') != 'External' and target_part in parts)
        ):
            root.remove(child)
            removed += 1
    return (ET.tostring(root, encoding='utf-8', xml_declaration=True) if removed else data), removed


def make_backup_copy(path: Path) -> Path:
    for number in range(10000):
        backup = path.with_name(path.name + ('.bak' if number == 0 else f'.bak.{number}'))
        try:
            output = backup.open('xb')
        except FileExistsError:
            continue
        try:
            with output, path.open('rb') as original:
                shutil.copyfileobj(original, output)
            shutil.copystat(path, backup)
        except BaseException:
            backup.unlink(missing_ok=True)
            raise
        return backup
    raise OSError('Не удалось создать резервную копию.')


def remove_office_metadata(path: Path, make_backup: bool = True) -> CleanResult:
    path = path.resolve()
    if not is_office_file(path):
        raise ValueError('Неподдерживаемый формат. Выберите файл Word, PowerPoint или Excel OpenXML.')
    if not zipfile.is_zipfile(path):
        raise ValueError('Файл повреждён или зашифрован. Сначала снимите пароль в Office.')
    fd, name = tempfile.mkstemp(prefix=f'.{path.stem}_clean_', suffix=path.suffix, dir=path.parent)
    os.close(fd)
    temporary = Path(name)
    try:
        removed_items = 0
        with zipfile.ZipFile(path) as original:
            names = original.namelist()
            if len(names) != len(set(names)):
                raise ValueError('В архиве обнаружены дублирующиеся части.')
            if '[Content_Types].xml' not in names or '_rels/.rels' not in names:
                raise ValueError('Это не пакет Office OpenXML.')
            parts = metadata_parts(original)
            with zipfile.ZipFile(temporary, 'w', compression=zipfile.ZIP_DEFLATED) as cleaned:
                for entry in original.infolist():
                    filename = package_name(entry.filename)
                    if filename in parts:
                        removed_items += 1
                        continue
                    data = original.read(entry)
                    if filename == '[Content_Types].xml':
                        data, count = clean_content_types_xml(data, parts)
                        removed_items += count
                    elif filename.endswith('.rels') and (filename == '_rels/.rels' or '/_rels/' in filename):
                        source = '' if filename == '_rels/.rels' else filename.rsplit('/_rels/', 1)[0]
                        data, count = clean_relationships_xml(data, source, parts)
                        removed_items += count
                    info = zipfile.ZipInfo(entry.filename, FIXED_ZIP_DATE)
                    info.compress_type = zipfile.ZIP_DEFLATED
                    info.external_attr = entry.external_attr
                    cleaned.writestr(info, data)
        shutil.copymode(path, temporary)
        backup = make_backup_copy(path) if make_backup else None
        os.replace(temporary, path)
        return CleanResult(path, removed_items, backup)
    finally:
        temporary.unlink(missing_ok=True)
