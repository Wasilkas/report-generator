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
import weakref
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, cast
from urllib.parse import urlsplit
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr
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


_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
_MC_PREFIX_ATTRIBUTES = frozenset(
    f'{{{_MC}}}{name}'
    for name in (
        'Ignorable',
        'ProcessContent',
        'MustUnderstand',
        'PreserveElements',
        'PreserveAttributes',
    )
)
# ElementTree renames prefixes and declares only namespaces used by element and
# attribute names, but markup-compatibility values (written by Excel, e.g.
# mc:Ignorable="x14ac xr") name prefixes as text. _parse records the prefixes
# those values name for each parsed root; _xml declares the missing ones on the
# serialized root. Adding declarations to ElementTree's output instead of
# serializing with original prefixes leaves every part without such markup,
# including openpyxl output and the metadata part, byte-identical.
# A prefix that names two namespaces in one part maps to None: such a part can
# still be read, and is rejected only when it has to be written back.
_COMPAT_PREFIXES: 'weakref.WeakKeyDictionary[ET.Element, dict[str, str | None]]' = (
    weakref.WeakKeyDictionary()
)


def _parse_compat(content: bytes) -> tuple[ET.Element, dict[str, 'str | None']]:
    """Parse a part once, mapping prefixes named by compatibility values to in-scope URIs."""
    named: dict[str, str | None] = {}
    scopes: list[dict[str, str]] = [{}]
    pending: dict[str, str] = {}
    root: ET.Element | None = None
    events = ('start-ns', 'start', 'end')
    for event, item in ET.iterparse(io.BytesIO(content), events=events):  # noqa: S314
        if event == 'start-ns':
            prefix, uri = item
            pending[prefix] = uri
        elif event == 'end':
            scopes.pop()
        else:
            if root is None:
                root = item
            scope = {**scopes[-1], **pending}
            pending = {}
            scopes.append(scope)
            for key, value in item.attrib.items():
                if key not in _MC_PREFIX_ATTRIBUTES and (
                    key != 'Requires' or item.tag != f'{{{_MC}}}Choice'
                ):
                    continue
                for token in value.split():
                    prefix = token.partition(':')[0]
                    uri = scope.get(prefix)
                    if not prefix or prefix == 'xml' or uri is None:
                        continue
                    if named.setdefault(prefix, uri) != uri:
                        named[prefix] = None
    if root is None:
        raise ValueError('Workbook XML has no root element')
    return root, named


def _declare_prefixes(content: bytes, prefixes: Mapping[str, 'str | None']) -> bytes:
    """Declare named prefixes on the serialized root; ElementTree declares all there."""
    emitted: dict[str, str] = {}
    for event, item in ET.iterparse(io.BytesIO(content), events=('start-ns', 'start')):  # noqa: S314
        if event == 'start':
            break
        emitted[item[0]] = item[1]
    declarations = []
    for prefix, uri in sorted(prefixes.items()):
        if uri is None:
            raise ValueError(
                f'Markup-compatibility prefix {prefix!r} names two namespaces in one workbook part'
            )
        if prefix not in emitted:
            declarations.append(f' xmlns:{prefix}={quoteattr(uri)}')
        elif emitted[prefix] != uri:
            raise ValueError(
                f'Markup-compatibility prefix {prefix!r} conflicts with a serialized prefix'
            )
    if not declarations:
        return content
    root = re.match(rb'<\?xml[^>]*\?>\s*<[^\s/>]+', content)
    if root is None:
        raise ValueError('Serialized workbook XML has no root element')
    return content[: root.end()] + ''.join(declarations).encode('utf-8') + content[root.end() :]


def _xml(root: ET.Element) -> bytes:
    content = cast('bytes', ET.tostring(root, encoding='utf-8', xml_declaration=True))
    prefixes = _COMPAT_PREFIXES.get(root)
    return _declare_prefixes(content, prefixes) if prefixes else content


def _parse(content: bytes) -> ET.Element:
    # Generated OOXML is UTF-8. Reject DTDs/entities and alternate encodings
    # before using the standard parser; no dependency source change is needed.
    if b'\x00' in content or b'<!DOCTYPE' in content.upper() or b'<!ENTITY' in content.upper():
        raise ValueError('Workbook XML must not contain DTDs or entities')
    if _MC.encode('utf-8') not in content:
        return ET.fromstring(content)  # noqa: S314 - declarations rejected above
    root, prefixes = _parse_compat(content)
    if prefixes:
        _COMPAT_PREFIXES[root] = prefixes
    return root


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


def _banner_row(index: int, line: str, style: int, width: float) -> ET.Element:
    characters = sum(2 if unicodedata.east_asian_width(char) in ('W', 'F') else 1 for char in line)
    height = max(30, 16 * math.ceil(characters / max(1, width - 2)))
    row = ET.Element(f'{{{_S}}}row', {'r': str(index), 'ht': str(height), 'customHeight': '1'})
    cell = ET.SubElement(row, f'{{{_S}}}c', {'r': f'A{index}', 't': 'inlineStr', 's': str(style)})
    inline = ET.SubElement(cell, f'{{{_S}}}is')
    ET.SubElement(
        inline, f'{{{_S}}}t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}
    ).text = line
    return row


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
        data.insert(index - 1, _banner_row(index, line, style, width))
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


_XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
_HYPERLINK = _R + '/hyperlink'
_EMU_PER_POINT = 12700
_FOOTER_LIMIT = 255
# CT_Worksheet child order; new elements are inserted after their predecessors.
_WORKSHEET_ORDER = (
    'sheetPr',
    'dimension',
    'sheetViews',
    'sheetFormatPr',
    'cols',
    'sheetData',
    'sheetCalcPr',
    'sheetProtection',
    'protectedRanges',
    'scenarios',
    'autoFilter',
    'sortState',
    'dataConsolidate',
    'customSheetViews',
    'mergeCells',
    'phoneticPr',
    'conditionalFormatting',
    'dataValidations',
    'hyperlinks',
    'printOptions',
    'pageMargins',
    'pageSetup',
    'headerFooter',
    'rowBreaks',
    'colBreaks',
    'customProperties',
    'cellWatches',
    'ignoredErrors',
    'smartTags',
    'drawing',
    'legacyDrawing',
    'legacyDrawingHF',
    'drawingHF',
    'picture',
    'oleObjects',
    'controls',
    'webPublishItems',
    'tableParts',
    'extLst',
)


def _child(root: ET.Element, tag: str) -> ET.Element:
    """Return a worksheet child, creating it at its schema position."""
    existing = root.find(f'{{{_S}}}{tag}')
    if existing is not None:
        return existing
    before = set(_WORKSHEET_ORDER[: _WORKSHEET_ORDER.index(tag)])
    insertion = max(
        (
            index + 1
            for index, element in enumerate(root)
            if element.tag.rsplit('}', 1)[-1] in before
        ),
        default=0,
    )
    element = ET.Element(f'{{{_S}}}{tag}')
    root.insert(insertion, element)
    return element


def _part_rels_path(part: str) -> str:
    directory, _, name = part.rpartition('/')
    return f'{directory}/_rels/{name}.rels'


def _resolve_target(source: str, target: str) -> str:
    if target.startswith('/'):
        return target.lstrip('/')
    parts = source.rpartition('/')[0].split('/')
    for segment in target.split('/'):
        if segment == '..':
            parts.pop()
        elif segment not in ('', '.'):
            parts.append(segment)
    return '/'.join(parts)


def _anchor_end_row(anchor: ET.Element, default_height_points: float) -> int:
    """Return the last 1-based row covered by one drawing anchor."""
    tag = anchor.tag.rsplit('}', 1)[-1]
    if tag == 'twoCellAnchor':
        return int(anchor.findtext(f'{{{_XDR}}}to/{{{_XDR}}}row', '0')) + 1
    row_height = max(1.0, default_height_points) * _EMU_PER_POINT
    extent = anchor.find(f'{{{_XDR}}}ext')
    height = int(extent.get('cy', '0')) if extent is not None else 0
    if tag == 'oneCellAnchor':
        start = int(anchor.findtext(f'{{{_XDR}}}from/{{{_XDR}}}row', '0'))
        offset = int(anchor.findtext(f'{{{_XDR}}}from/{{{_XDR}}}rowOff', '0'))
        return start + math.ceil((offset + height) / row_height)
    position = anchor.find(f'{{{_XDR}}}pos')
    top = int(position.get('y', '0')) if position is not None else 0
    return math.ceil((top + height) / row_height)


def _last_used_row(root: ET.Element, target: str, entries: Mapping[str, bytes]) -> int:
    """Last row holding cells, merged ranges or drawings, so a footer covers nothing."""
    last = 0
    data = root.find(f'{{{_S}}}sheetData')
    for row in data if data is not None else ():
        if len(row):
            last = max(last, int(row.get('r', '0')))
    merges = root.find(f'{{{_S}}}mergeCells')
    for merge in merges if merges is not None else ():
        match = _COORD.fullmatch(merge.get('ref', '').split(':')[-1])
        if match is not None:
            last = max(last, int(match.group('row').lstrip('$')))
    drawing = root.find(f'{{{_S}}}drawing')
    rels_path = _part_rels_path(target)
    if drawing is None or rels_path not in entries:
        return last
    relationship = drawing.get(f'{{{_R}}}id')
    format_properties = root.find(f'{{{_S}}}sheetFormatPr')
    default_height = float(
        format_properties.get('defaultRowHeight', '15') if format_properties is not None else '15'
    )
    for element in _parse(entries[rels_path]):
        if element.get('Id') != relationship:
            continue
        part = _resolve_target(target, element.attrib['Target'])
        if part in entries:
            for anchor in _parse(entries[part]):
                last = max(last, _anchor_end_row(anchor, default_height))
    return last


def _validated_links(lines: Sequence[str], links: Sequence[str | None] | None) -> list[str | None]:
    if links is None:
        return [None] * len(lines)
    if len(links) != len(lines):
        raise ValueError('Banner links must match banner lines one to one')
    for link in links:
        if link is None:
            continue
        parts = urlsplit(link) if isinstance(link, str) else None
        if parts is None or parts.scheme not in ('http', 'https') or not parts.netloc:
            raise ValueError('Banner links must be absolute http or https URLs')
    return list(links)


def _footer_text(text: str) -> str:
    """Escape page-footer codes and keep Excel's 255-character limit."""
    prefix = '&L'
    escaped = text.replace('&', '&&')
    if len(prefix) + len(escaped) <= _FOOTER_LIMIT:
        return prefix + escaped
    kept = escaped[: _FOOTER_LIMIT - len(prefix) - 1]
    if (len(kept) - len(kept.rstrip('&'))) % 2:
        kept = kept[:-1]
    return prefix + kept + '…'


def _add_hyperlinks(
    root: ET.Element,
    target: str,
    entries: dict[str, bytes],
    cells: Sequence[tuple[str, str]],
) -> None:
    rels_path = _part_rels_path(target)
    rels = (
        _parse(entries[rels_path]) if rels_path in entries else ET.Element(f'{{{_P}}}Relationships')
    )
    used = {element.get('Id') for element in rels}
    hyperlinks = _child(root, 'hyperlinks')
    counter = 1
    for reference, url in cells:
        while f'rIdLink{counter}' in used:
            counter += 1
        identifier = f'rIdLink{counter}'
        used.add(identifier)
        ET.SubElement(
            rels,
            f'{{{_P}}}Relationship',
            {'Id': identifier, 'Type': _HYPERLINK, 'Target': url, 'TargetMode': 'External'},
        )
        ET.SubElement(
            hyperlinks, f'{{{_S}}}hyperlink', {'ref': reference, f'{{{_R}}}id': identifier}
        )
    entries[rels_path] = _xml(rels)


def _footer_sheet(
    root: ET.Element,
    target: str,
    entries: dict[str, bytes],
    lines: Sequence[str],
    links: Sequence[str | None],
    page_footer: str | None,
    style: int,
) -> tuple[int, str]:
    """Append literal lines after all content; return the footer's last row and column."""
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
    last = _last_used_row(root, target, entries)
    if lines or page_footer is not None:
        # Printed identity stays readable: fit the merged footer width to one page.
        _print_width(root)
    start = last + 2
    if lines:
        merges = _child(root, 'mergeCells')
        width = _banner_width(root, end_column)
        linked: list[tuple[str, str]] = []
        for offset, (line, link) in enumerate(zip(lines, links, strict=True)):
            index = start + offset
            data.append(_banner_row(index, line, style, width))
            ET.SubElement(merges, f'{{{_S}}}mergeCell', {'ref': f'A{index}:{end_column}{index}'})
            if link is not None:
                linked.append((f'A{index}', link))
        merges.set('count', str(len(merges)))
        if linked:
            _add_hyperlinks(root, target, entries, linked)
        if dimension is not None:
            dimension.set('ref', f'A1:{end_column}{start + len(lines) - 1}')
    if page_footer is not None:
        header_footer = _child(root, 'headerFooter')
        odd = header_footer.find(f'{{{_S}}}oddFooter')
        if odd is None:
            # oddHeader precedes oddFooter in CT_HeaderFooter.
            odd = ET.Element(f'{{{_S}}}oddFooter')
            odd_header = header_footer.find(f'{{{_S}}}oddHeader')
            header_footer.insert(0 if odd_header is None else 1, odd)
        odd.text = _footer_text(page_footer)
    return (start + len(lines) - 1, end_column) if lines else (last, 'A')


_AREA_END = re.compile(r':(?P<absolute>\$?)(?P<column>[A-Z]+)\$?(?P<row>\d+)$')


def _extend_print_areas(
    workbook: ET.Element, sheets: Sequence[str], footers: Sequence[tuple[int, str]]
) -> None:
    """Grow only the final print area so the merged footer prints uncut."""
    defined = workbook.find(f'{{{_S}}}definedNames')
    if defined is None:
        return
    for element in defined:
        if element.get('name') != '_xlnm.Print_Area' or not element.text:
            continue
        index = int(element.get('localSheetId', '-1'))
        if not 0 <= index < len(sheets):
            continue
        last_row, last_column = footers[index]
        areas = element.text.split(',')
        match = _AREA_END.search(areas[-1])
        if match is None:
            continue
        column = get_column_letter(
            max(
                column_index_from_string(match.group('column')),
                column_index_from_string(last_column),
            )
        )
        row = max(int(match.group('row')), last_row)
        areas[-1] = areas[-1][: match.start()] + f':{match.group("absolute")}{column}${row}'
        element.text = ','.join(areas)


def annotate_workbook(
    path: str | Path,
    *,
    banner_lines: Sequence[str],
    metadata_texts: Mapping[str, str],
    format: MetadataFormat,
    validate: MetadataValidator | None = None,
    placement: Literal['top', 'bottom'] = 'top',
    banner_links: Sequence[str | None] | None = None,
    page_footer: str | None = None,
) -> None:
    """Prepend literal banners while preserving OOXML coordinates and original bytes.

    The caller chooses banner text and metadata schema. The embedded original,
    part names, XML envelope and fingerprint algorithm remain caller-defined.
    Reannotation starts from authenticated original bytes, making it idempotent.

    ``placement='bottom'`` leaves every original coordinate unchanged: lines start
    one empty row after the last cell, merged range or drawing of each worksheet,
    optional ``banner_links`` become external hyperlinks on their lines and
    ``page_footer`` becomes the printed left footer. Bottom placement accepts no
    lines for a metadata-only envelope and never adds print titles.
    """
    if placement not in ('top', 'bottom'):
        raise ValueError("Banner placement must be 'top' or 'bottom'")
    if any(not isinstance(line, str) for line in banner_lines):
        raise ValueError('Workbook annotation requires literal banner lines')
    if placement == 'top':
        if not banner_lines:
            raise ValueError('Workbook annotation requires literal banner lines')
        if banner_links is not None or page_footer is not None:
            raise ValueError('Banner links and page footers require bottom placement')
    if page_footer is not None and not isinstance(page_footer, str):
        raise ValueError('Page footer must be text')
    links = _validated_links(banner_lines, banner_links)
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
    if placement == 'bottom':
        style = _banner_style(entries) if lines else 0
        footers: list[tuple[int, str]] = []
        for target in sheets.values():
            root = _parse(entries[target])
            footers.append(_footer_sheet(root, target, entries, lines, links, page_footer, style))
            entries[target] = _xml(root)
        _extend_print_areas(workbook, list(sheets), footers)
        entries['xl/workbook.xml'] = _xml(workbook)
    else:
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
