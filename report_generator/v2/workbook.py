"""Generic OOXML banner adaptation and reversible caller-defined metadata.

Schema validation and banner wording belong to the caller. The metadata envelope
and fingerprint keep existing formats readable without introducing a new format.
"""

import base64
import hashlib
import io
import math
import re
import unicodedata
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import cast
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import column_index_from_string, get_column_letter

_S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
_P = 'http://schemas.openxmlformats.org/package/2006/relationships'
_C = 'http://schemas.openxmlformats.org/package/2006/content-types'
_COORD = re.compile(r'(?P<col>\$?[A-Za-z]{1,3})?(?P<row>\$?\d+)$')
_REFERENCE = re.compile(
    r"(?:(?P<sheet>'(?:[^']|'')+'|[^!]+)!)?(?P<coords>\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?\d+:\$?\d+|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3})$"
)


def _xml(root: ET.Element) -> bytes:
    return cast('bytes', ET.tostring(root, encoding='utf-8', xml_declaration=True))


def _parse(content: bytes) -> ET.Element:
    # Generated OOXML is UTF-8. Reject DTDs/entities and alternate encodings
    # before using the standard parser; no dependency source change is needed.
    if b'\x00' in content or b'<!DOCTYPE' in content.upper() or b'<!ENTITY' in content.upper():
        raise ValueError('Workbook XML must not contain DTDs or entities')
    return ET.fromstring(content)  # noqa: S314 - declarations rejected above


@dataclass(frozen=True)
class MetadataFormat:
    """Existing custom XML envelope supplied by its schema owner."""

    path: str
    namespace: str
    root_name: str
    version: str
    relationship_id: str

    def __post_init__(self) -> None:
        values = (self.path, self.namespace, self.root_name, self.version, self.relationship_id)
        if any(not isinstance(value, str) or not value for value in values):
            raise ValueError('Metadata format values must be nonempty strings')
        if self.path.startswith('/') or '..' in self.path.split('/'):
            raise ValueError('Metadata part path must remain inside the workbook package')


MetadataValidator = Callable[[ET.Element], object]


def package_fingerprint(parts: Mapping[str, bytes], format: MetadataFormat) -> str:
    """Bind package payloads and metadata, excluding only the digest itself."""
    digest = hashlib.sha256()
    for name in sorted(parts):
        content = parts[name]
        if name == format.path:
            metadata = _parse(content)
            metadata.attrib.pop('package_sha256', None)
            content = _xml(metadata)
        digest.update(name.encode('utf-8') + b'\x00')
        digest.update(hashlib.sha256(content).digest())
    return digest.hexdigest()


def read_metadata(
    content: bytes, format: MetadataFormat, validate: MetadataValidator | None = None
) -> ET.Element | None:
    """Read and authenticate caller metadata; validation precedes fingerprint comparison."""
    with ZipFile(io.BytesIO(content)) as archive:
        if format.path not in archive.namelist():
            return None
        root = _parse(archive.read(format.path))
        parts = {name: archive.read(name) for name in archive.namelist()}
    if (
        root.tag != f'{{{format.namespace}}}{format.root_name}'
        or root.get('version') != format.version
    ):
        raise ValueError('Unsupported workbook model identity metadata')
    expected = root.get('package_sha256')
    if expected is None or re.fullmatch(r'[0-9a-f]{64}', expected) is None:
        raise ValueError('Workbook identity metadata has no valid package fingerprint')
    if validate is not None:
        validate(root)
    if expected != package_fingerprint(parts, format):
        raise ValueError(
            'Workbook changed after model identity annotation; regenerate the workbook '
            'before reading its metrics or identities'
        )
    return root


def original_workbook(
    content: bytes, format: MetadataFormat, validate: MetadataValidator | None = None
) -> bytes:
    """Recover the byte-exact original, or return an unannotated package unchanged."""
    root = read_metadata(content, format, validate)
    if root is None:
        return content
    original = root.findtext(f'{{{format.namespace}}}original')
    if original is None:
        raise ValueError('Workbook identity metadata has no original content')
    return base64.b64decode(original, validate=True)


def _coordinate(value: str, shift: int) -> str:
    match = _COORD.fullmatch(value)
    if match is None or match.group('row') is None:
        return value
    row = match.group('row')
    column = match.group('col') or ''
    absolute = '$' if row.startswith('$') else ''
    return f'{column}{absolute}{int(row.lstrip("$")) + shift}'


def _range(value: str, shift: int) -> str:
    return ' '.join(
        ':'.join(_coordinate(endpoint, shift) for endpoint in span.split(':'))
        for span in value.split()
    )


def _formula(value: str, shift: int, sheet_names: set[str], current: str | None) -> str:
    """Shift A1 references, including absolute and cross-worksheet references."""
    prefix = '' if value.startswith('=') else '='
    tokens = Tokenizer(prefix + value).items
    result: list[str] = []
    for token in tokens:
        text: str = token.value
        if token.type == 'OPERAND' and token.subtype == 'RANGE':
            match = _REFERENCE.fullmatch(text)
            if match is not None:
                sheet = match.group('sheet')
                name = (
                    (sheet[1:-1].replace("''", "'") if sheet.startswith("'") else sheet)
                    if sheet
                    else current
                )
                if name in sheet_names:
                    text = (sheet + '!' if sheet else '') + _range(match.group('coords'), shift)
        result.append(text)
    return ('=' if not prefix else '') + ''.join(result)


def _banner_style(entries: dict[str, bytes]) -> int:
    path = 'xl/styles.xml'
    root = _parse(entries[path])
    fonts = root.find(f'{{{_S}}}fonts')
    xfs = root.find(f'{{{_S}}}cellXfs')
    if fonts is None or xfs is None:
        raise ValueError('Generated workbook is missing required cell styles')
    font_id = len(fonts)
    font = ET.SubElement(fonts, f'{{{_S}}}font')
    ET.SubElement(font, f'{{{_S}}}b')
    ET.SubElement(font, f'{{{_S}}}sz', {'val': '11'})
    ET.SubElement(font, f'{{{_S}}}name', {'val': 'Calibri'})
    fonts.set('count', str(len(fonts)))
    style_id = len(xfs)
    xf = ET.SubElement(
        xfs,
        f'{{{_S}}}xf',
        {
            'numFmtId': '0',
            'fontId': str(font_id),
            'fillId': '0',
            'borderId': '0',
            'xfId': '0',
            'applyFont': '1',
            'applyAlignment': '1',
        },
    )
    ET.SubElement(
        xf, f'{{{_S}}}alignment', {'horizontal': 'left', 'vertical': 'center', 'wrapText': '1'}
    )
    xfs.set('count', str(len(xfs)))
    entries[path] = _xml(root)
    return style_id


def _shift_sheet(root: ET.Element, name: str, names: set[str], shift: int) -> None:
    for element in root.iter():
        tag = element.tag.rsplit('}', 1)[-1]
        for attribute in ('ref', 'sqref', 'activeCell', 'topLeftCell'):
            if attribute in element.attrib:
                element.set(attribute, _range(element.attrib[attribute], shift))
        if tag in ('row', 'c') and 'r' in element.attrib:
            value = element.attrib['r']
            element.set('r', str(int(value) + shift) if tag == 'row' else _coordinate(value, shift))
        if tag == 'hyperlink' and element.get('location'):
            element.set('location', _formula(element.attrib['location'], shift, names, name))
        if tag in ('f', 'formula', 'formula1', 'formula2') and element.text:
            element.text = _formula(element.text, shift, names, name)
        if tag == 'pane' and element.get('state') in ('frozen', 'frozenSplit'):
            element.set('ySplit', str(float(element.get('ySplit', '0')) + shift))
    breaks = root.find(f'{{{_S}}}rowBreaks')
    if breaks is not None:
        for element in breaks:
            element.set('id', str(int(element.attrib['id']) + shift))


def _banner_width(root: ET.Element, end_column: str) -> float:
    count = column_index_from_string(end_column)
    widths = [13.0] * count
    cols = root.find(f'{{{_S}}}cols')
    if cols is not None:
        for col in cols:
            for index in range(int(col.get('min', '1')) - 1, min(count, int(col.get('max', '1')))):
                widths[index] = 0.0 if col.get('hidden') == '1' else float(col.get('width', '13'))
    return sum(widths)


def _print_width(root: ET.Element) -> None:
    properties = root.find(f'{{{_S}}}sheetPr')
    if properties is None:
        properties = ET.Element(f'{{{_S}}}sheetPr')
        root.insert(0, properties)
    setup_properties = properties.find(f'{{{_S}}}pageSetUpPr')
    if setup_properties is None:
        setup_properties = ET.SubElement(properties, f'{{{_S}}}pageSetUpPr')
    setup_properties.set('fitToPage', '1')
    setup = root.find(f'{{{_S}}}pageSetup')
    if setup is None:
        setup = ET.Element(f'{{{_S}}}pageSetup')
        after = {
            'sheetCalcPr',
            'sheetProtection',
            'protectedRanges',
            'scenarios',
            'autoFilter',
            'sortState',
            'dataConsolidate',
            'customSheetViews',
            'phoneticPr',
            'hyperlinks',
            'printOptions',
            'pageMargins',
            'dataValidations',
            'conditionalFormatting',
            'mergeCells',
            'sheetData',
        }
        insertion = max(
            (
                index + 1
                for index, element in enumerate(root)
                if element.tag.rsplit('}', 1)[-1] in after
            ),
            default=0,
        )
        root.insert(insertion, setup)
        # A newly introduced setup must keep vertical pages unconstrained.
        setup.set('fitToHeight', '0')
    setup.set('fitToWidth', '1')


def _sheet(root: ET.Element, name: str, names: set[str], lines: list[str], style: int) -> None:
    shift = len(lines)
    data = root.find(f'{{{_S}}}sheetData')
    if data is None:
        raise ValueError('Generated worksheet is missing sheetData')
    dimension = root.find(f'{{{_S}}}dimension')
    end_column = 'D'
    if dimension is not None:
        match = _COORD.fullmatch(dimension.get('ref', 'A1').split(':')[-1])
        if match is not None:
            end_column = match.group('col') or 'D'
    end_column = get_column_letter(max(8, column_index_from_string(end_column)))
    _shift_sheet(root, name, names, shift)
    _print_width(root)
    merges = root.find(f'{{{_S}}}mergeCells')
    if merges is None:
        merges = ET.Element(f'{{{_S}}}mergeCells')
        # mergeCells follows sheetData, protection, scenarios and autoFilter,
        # and precedes conditional formatting/data validations/page settings.
        after = {
            'sheetData',
            'sheetCalcPr',
            'sheetProtection',
            'protectedRanges',
            'scenarios',
            'autoFilter',
            'sortState',
            'dataConsolidate',
            'customSheetViews',
        }
        insertion = max(
            (
                index + 1
                for index, element in enumerate(root)
                if element.tag.rsplit('}', 1)[-1] in after
            ),
            default=0,
        )
        root.insert(insertion, merges)
    width = _banner_width(root, end_column)
    for index, line in enumerate(lines, start=1):
        characters = sum(
            2 if unicodedata.east_asian_width(char) in ('W', 'F') else 1 for char in line
        )
        height = max(30, 16 * math.ceil(characters / max(1, width - 2)))
        row = ET.Element(f'{{{_S}}}row', {'r': str(index), 'ht': str(height), 'customHeight': '1'})
        cell = ET.SubElement(
            row, f'{{{_S}}}c', {'r': f'A{index}', 't': 'inlineStr', 's': str(style)}
        )
        inline = ET.SubElement(cell, f'{{{_S}}}is')
        ET.SubElement(
            inline, f'{{{_S}}}t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}
        ).text = line
        data.insert(index - 1, row)
        if end_column != 'A':
            ET.SubElement(merges, f'{{{_S}}}mergeCell', {'ref': f'A{index}:{end_column}{index}'})
    merges.set('count', str(len(merges)))
    if dimension is not None:
        last = dimension.attrib['ref'].split(':')[-1]
        match = _COORD.fullmatch(last)
        last_row = match.group('row') if match is not None else str(shift + 1)
        dimension.set('ref', f'A1:{end_column}{last_row}')


def _print_area_end(match: re.Match[str]) -> str:
    column = get_column_letter(max(8, column_index_from_string(match.group('column'))))
    return f':{match.group("absolute")}{column}{match.group("row")}'


def _defined_names(
    workbook: ET.Element, sheets_element: ET.Element, sheets: dict[str, str], shift: int
) -> None:
    defined = workbook.find(f'{{{_S}}}definedNames')
    if defined is None:
        defined = ET.Element(f'{{{_S}}}definedNames')
        position = list(workbook).index(sheets_element) + 1
        workbook.insert(position, defined)
    titles: dict[int, ET.Element] = {}
    for element in defined:
        sheet_index = int(element.get('localSheetId', '-1'))
        current = list(sheets)[sheet_index] if 0 <= sheet_index < len(sheets) else None
        if element.text:
            element.text = _formula(element.text, shift, set(sheets), current)
        if element.get('name') == '_xlnm.Print_Titles':
            titles[sheet_index] = element
        if element.get('name') == '_xlnm.Print_Area' and element.text:
            # Keep the right/bottom limits but include the new repeated banner.
            element.text = re.sub(r'(!\$?[A-Z]+\$?)\d+(?=:)', r'\g<1>1', element.text)
            element.text = re.sub(
                r':(?P<absolute>\$?)(?P<column>[A-Z]+)(?P<row>\$?\d+)',
                _print_area_end,
                element.text,
            )
    _print_titles(defined, titles, sheets, shift)


def _print_titles(
    defined: ET.Element, titles: dict[int, ET.Element], sheets: dict[str, str], shift: int
) -> None:
    for index, name in enumerate(sheets):
        quoted = "'" + name.replace("'", "''") + "'!"
        row_end = shift
        title = titles.get(index)
        prior = title.text if title is not None else None
        columns: list[str] = []
        if prior:
            for token in Tokenizer('=' + prior).items:
                if token.type != 'OPERAND' or token.subtype != 'RANGE':
                    continue
                span = token.value
                coords = span.rsplit('!', 1)[-1]
                if re.fullmatch(r'\$?\d+:\$?\d+', coords):
                    row_end = max(row_end, int(coords.split(':')[-1].lstrip('$')))
                else:
                    columns.append(span)
        if title is None:
            title = ET.SubElement(
                defined,
                f'{{{_S}}}definedName',
                {'name': '_xlnm.Print_Titles', 'localSheetId': str(index)},
            )
        title.text = ','.join([quoted + f'$1:${row_end}', *columns])


def _shift_table(root: ET.Element, sheets: dict[str, str], shift: int) -> None:
    for element in root.iter():
        if 'ref' in element.attrib:
            element.set('ref', _range(element.attrib['ref'], shift))
        if (
            element.tag.rsplit('}', 1)[-1] in ('calculatedColumnFormula', 'totalsRowFormula')
            and element.text
        ):
            element.text = _formula(element.text, shift, set(sheets), next(iter(sheets)))


def _shift_related_parts(entries: dict[str, bytes], sheets: dict[str, str], shift: int) -> None:
    for filename in list(entries):
        if not filename.endswith('.xml') or not filename.startswith(
            ('xl/tables/', 'xl/drawings/', 'xl/charts/')
        ):
            continue
        root = _parse(entries[filename])
        if filename.startswith('xl/tables/'):
            _shift_table(root, sheets, shift)
        else:
            for element in root.iter():
                tag = element.tag.rsplit('}', 1)[-1]
                if filename.startswith('xl/drawings/') and tag == 'row' and element.text:
                    element.text = str(int(element.text) + shift)
                elif tag == 'f' and element.text:
                    element.text = _formula(element.text, shift, set(sheets), None)
        entries[filename] = _xml(root)


def _add_metadata(
    entries: dict[str, bytes],
    metadata_texts: Mapping[str, str],
    original: bytes,
    format: MetadataFormat,
) -> None:
    metadata = ET.Element(f'{{{format.namespace}}}{format.root_name}', {'version': format.version})
    for name, text in metadata_texts.items():
        ET.SubElement(metadata, f'{{{format.namespace}}}{name}').text = text
    ET.SubElement(metadata, f'{{{format.namespace}}}original').text = base64.b64encode(
        original
    ).decode('ascii')
    entries[format.path] = _xml(metadata)
    types = _parse(entries['[Content_Types].xml'])
    ET.SubElement(
        types,
        f'{{{_C}}}Override',
        {'PartName': '/' + format.path, 'ContentType': 'application/xml'},
    )
    entries['[Content_Types].xml'] = _xml(types)
    root_rels = _parse(entries['_rels/.rels'])
    used = {element.attrib['Id'] for element in root_rels}
    identity_id = format.relationship_id
    while identity_id in used:
        identity_id += '_'
    ET.SubElement(
        root_rels,
        f'{{{_P}}}Relationship',
        {'Id': identity_id, 'Type': _R + '/customXml', 'Target': format.path},
    )
    entries['_rels/.rels'] = _xml(root_rels)
    metadata.set('package_sha256', package_fingerprint(entries, format))
    entries[format.path] = _xml(metadata)


def annotate_workbook(
    path: str | Path,
    *,
    banner_lines: Sequence[str],
    metadata_texts: Mapping[str, str],
    format: MetadataFormat,
    validate: MetadataValidator | None = None,
) -> None:
    """Prepend literal banners while preserving OOXML coordinates and original bytes.

    The caller chooses banner text and metadata schema. The embedded original,
    part names, XML envelope and fingerprint algorithm remain caller-defined.
    Reannotation starts from authenticated original bytes, making it idempotent.
    """
    if not banner_lines or any(not isinstance(line, str) for line in banner_lines):
        raise ValueError('Workbook annotation requires literal banner lines')
    if 'original' in metadata_texts:
        raise ValueError('Original workbook metadata is owned by the package writer')
    if any(
        not isinstance(name, str) or not isinstance(text, str)
        for name, text in metadata_texts.items()
    ):
        raise ValueError('Metadata fields must map text names to text values')
    destination = Path(path)
    original = original_workbook(destination.read_bytes(), format, validate)
    with ZipFile(io.BytesIO(original)) as archive:
        entries = {info.filename: archive.read(info.filename) for info in archive.infolist()}
    workbook = _parse(entries['xl/workbook.xml'])
    relationships = _parse(entries['xl/_rels/workbook.xml.rels'])
    targets = {element.attrib['Id']: element.attrib['Target'] for element in relationships}
    sheets_element = workbook.find(f'{{{_S}}}sheets')
    if sheets_element is None:
        raise ValueError('Workbook contains no worksheets')
    sheets: dict[str, str] = {}
    for element in sheets_element:
        target = targets[element.attrib[f'{{{_R}}}id']]
        resolved = target.lstrip('/') if target.startswith('/') else 'xl/' + target
        if '/worksheets/' in resolved:
            sheets[element.attrib['name']] = resolved
    lines = list(banner_lines)
    shift = len(lines)
    style = _banner_style(entries)
    for name, target in sheets.items():
        root = _parse(entries[target])
        _sheet(root, name, set(sheets), lines, style)
        entries[target] = _xml(root)
    _defined_names(workbook, sheets_element, sheets, shift)
    entries['xl/workbook.xml'] = _xml(workbook)
    _shift_related_parts(entries, sheets, shift)
    _add_metadata(entries, metadata_texts, original, format)
    output = io.BytesIO()
    with ZipFile(output, 'w', compression=ZIP_DEFLATED) as archive:
        for filename, content in entries.items():
            archive.writestr(filename, content)
    destination.write_bytes(output.getvalue())
