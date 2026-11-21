"""Excel-saved markup-compatibility prefixes stay declared after workbook annotation."""

import io
import re
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.table import Table

from report_generator.v2 import MetadataFormat, annotate_workbook, original_workbook, read_metadata

FORMAT = MetadataFormat(
    'customXml/example-identity.xml', 'urn:example:identity', 'identity', '1', 'rIdIdentity'
)
MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
URIS = {
    'mc': MC,
    'x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
    'x15ac': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac',
    'xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
    'xr2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/revision2',
    'xr3': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision3',
    'xr6': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision6',
    'xr10': 'http://schemas.microsoft.com/office/spreadsheetml/2016/revision10',
    'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
    'x16r2': 'http://schemas.microsoft.com/office/spreadsheetml/2015/02/main',
}
COMPAT_ATTRIBUTES = {
    f'{{{MC}}}{name}'
    for name in (
        'Ignorable',
        'ProcessContent',
        'MustUnderstand',
        'PreserveElements',
        'PreserveAttributes',
    )
}
UID = '{00000000-0001-0000-0000-000000000000}'


def declarations(*prefixes):
    return ' '.join(f'xmlns:{prefix}="{URIS[prefix]}"' for prefix in prefixes)


def openpyxl_parts():
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = 'Main'
    sheet.append(['Class', 'Value'])
    sheet.append(['cat', 1])
    sheet['B3'] = '=SUM(B2:B2)'
    sheet.add_table(Table(displayName='Results', ref='A1:B2'))
    workbook.create_sheet('Other').append(['=Main!B2+1'])
    output = io.BytesIO()
    workbook.save(output)
    with ZipFile(output) as archive:
        return {info.filename: archive.read(info.filename) for info in archive.infolist()}


def insert_after_tag(content, tag, text):
    """Add attributes right after the first start-tag name, as Excel orders its roots."""
    updated, count = re.subn(rf'<{tag}\b'.encode(), f'<{tag} {text}'.encode(), content, count=1)
    assert count == 1, tag
    return updated


def excel_style(parts, workbook_ignorable='x15 xr xr6 xr10 xr2', sheet_extra=''):
    """Give an openpyxl package the roots and compatibility markup Excel writes."""
    parts = dict(parts)
    workbook = insert_after_tag(
        parts['xl/workbook.xml'],
        'workbook',
        f'xmlns:mc="{MC}" mc:Ignorable="{workbook_ignorable}" '
        + declarations('x15', 'xr', 'xr6', 'xr10', 'xr2'),
    )
    # The absPath block declares mc again and x15ac only below the root, as Excel does.
    alternate = (
        f'<mc:AlternateContent xmlns:mc="{MC}"><mc:Choice Requires="x15">'
        f'<x15ac:absPath xmlns:x15ac="{URIS["x15ac"]}" mc:Ignorable="x15ac" url="C:\\Reports\\"/>'
        '</mc:Choice></mc:AlternateContent>'
        '<xr:revisionPtr revIDLastSave="0" documentId="8_{00000000-0000-0000-0000-000000000000}"'
        ' xr6:coauthVersionLast="47" xr6:coauthVersionMax="47"'
        ' xr10:uidLastSave="{00000000-0000-0000-0000-000000000000}"/>'
    )
    workbook, count = re.subn(
        rb'<workbookPr\b[^>]*>', lambda match: match.group() + alternate.encode(), workbook
    )
    assert count == 1
    parts['xl/workbook.xml'] = workbook
    for name in ('xl/worksheets/sheet1.xml', 'xl/worksheets/sheet2.xml'):
        sheet = insert_after_tag(
            parts[name],
            'worksheet',
            f'xmlns:mc="{MC}" mc:Ignorable="x14ac xr xr2 xr3" '
            + declarations('x14ac', 'xr', 'xr2', 'xr3')
            + f' xr:uid="{UID}" {sheet_extra}',
        )
        parts[name] = insert_after_tag(sheet, 'sheetFormatPr', 'x14ac:dyDescent="0.25"')
    styles = insert_after_tag(
        parts['xl/styles.xml'],
        'styleSheet',
        f'xmlns:mc="{MC}" mc:Ignorable="x14ac x16r2 xr" ' + declarations('x14ac', 'x16r2', 'xr'),
    )
    parts['xl/styles.xml'] = insert_after_tag(styles, 'fonts', 'x14ac:knownFonts="1"')
    parts['xl/tables/table1.xml'] = insert_after_tag(
        parts['xl/tables/table1.xml'],
        'table',
        f'xmlns:mc="{MC}" mc:Ignorable="xr xr3" ' + declarations('xr', 'xr3') + f' xr:uid="{UID}"',
    )
    return parts


def package(parts):
    output = io.BytesIO()
    with ZipFile(output, 'w', compression=ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return output.getvalue()


def compat_bindings(content):
    """Each prefix named by a compatibility value, with its URI in scope at that element."""
    scopes = [{}]
    pending = {}
    found = []
    for event, item in ET.iterparse(io.BytesIO(content), events=('start-ns', 'start', 'end')):
        if event == 'start-ns':
            pending[item[0]] = item[1]
        elif event == 'end':
            scopes.pop()
        else:
            scope = {**scopes[-1], **pending}
            pending = {}
            scopes.append(scope)
            for key, value in item.attrib.items():
                if key in COMPAT_ATTRIBUTES or (
                    key == 'Requires' and item.tag == f'{{{MC}}}Choice'
                ):
                    for token in value.split():
                        prefix = token.partition(':')[0]
                        found.append((item.tag, key, prefix, scope.get(prefix)))
    return found


def annotated_parts(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def annotate(path, placement):
    if placement == 'bottom':
        annotate_workbook(
            path,
            banner_lines=['Candidate: example', 'Training task: open'],
            metadata_texts={'note': 'markup compatibility'},
            format=FORMAT,
            placement='bottom',
            banner_links=[None, 'https://app.example/tasks/1'],
            page_footer='Candidate: example',
        )
    else:
        annotate_workbook(
            path,
            banner_lines=['Candidate: example'],
            metadata_texts={'note': 'markup compatibility'},
            format=FORMAT,
        )


COMPAT_PARTS = (
    'xl/workbook.xml',
    'xl/worksheets/sheet1.xml',
    'xl/worksheets/sheet2.xml',
    'xl/styles.xml',
    'xl/tables/table1.xml',
)


@pytest.mark.parametrize('placement', ['bottom', 'top'])
def test_excel_compatibility_prefixes_stay_bound_after_annotation(tmp_path, placement):
    source = excel_style(openpyxl_parts())
    original = package(source)
    path = tmp_path / 'excel.xlsx'
    path.write_bytes(original)
    annotate(path, placement)
    parts = annotated_parts(path)
    for name, content in parts.items():
        if name.endswith(('.xml', '.rels')):
            ET.fromstring(content)  # every part stays well-formed
    for name in COMPAT_PARTS:
        expected = compat_bindings(source[name])
        assert expected and all(uri == URIS.get(prefix) for _, _, prefix, uri in expected), name
        assert compat_bindings(parts[name]) == expected, name
    workbook_bindings = compat_bindings(parts['xl/workbook.xml'])
    assert (f'{{{MC}}}Choice', 'Requires', 'x15', URIS['x15']) in workbook_bindings
    assert ('{%s}absPath' % URIS['x15ac'], f'{{{MC}}}Ignorable', 'x15ac', URIS['x15ac']) in (
        workbook_bindings
    )
    assert read_metadata(path.read_bytes(), FORMAT) is not None
    assert original_workbook(path.read_bytes(), FORMAT) == original
    workbook = load_workbook(path)
    main = workbook['Main']
    if placement == 'bottom':
        assert main['A5'].value == 'Candidate: example'
        assert main['A6'].hyperlink.target == 'https://app.example/tasks/1'
    else:
        assert main['A1'].value == 'Candidate: example'
        assert main.tables['Results'].ref == 'A2:B3'
    workbook.close()


def test_named_prefix_already_serialized_with_the_same_uri_is_declared_once(tmp_path):
    source = excel_style(openpyxl_parts(), workbook_ignorable='x15 xr xr6 xr10 xr2 r')
    path = tmp_path / 'excel.xlsx'
    path.write_bytes(package(source))
    annotate(path, 'bottom')
    content = annotated_parts(path)['xl/workbook.xml']
    root_tag = re.search(rb'<[^?][^>]*>', content).group()
    assert root_tag.count(b'xmlns:r=') == 1
    assert compat_bindings(content) == compat_bindings(source['xl/workbook.xml'])
    assert ('{%s}workbook' % MAIN, f'{{{MC}}}Ignorable', 'r', REL) in compat_bindings(content)


@pytest.mark.parametrize(
    'sheet_extra',
    [
        # A source prefix spelled like a generated one but bound elsewhere.
        ' '.join(f'xmlns:ns{index}="urn:example:conflict"' for index in range(10))
        + ' mc:MustUnderstand="'
        + ' '.join(f'ns{index}' for index in range(10))
        + '"',
    ],
)
def test_named_prefix_conflicting_with_a_serialized_prefix_is_rejected(tmp_path, sheet_extra):
    original = package(excel_style(openpyxl_parts(), sheet_extra=sheet_extra))
    path = tmp_path / 'excel.xlsx'
    path.write_bytes(original)
    with pytest.raises(ValueError, match='prefix'):
        annotate(path, 'bottom')
    assert path.read_bytes() == original


def test_named_prefix_bound_to_two_uris_in_one_part_is_rejected(tmp_path):
    parts = excel_style(openpyxl_parts())
    parts['xl/styles.xml'] = insert_after_tag(
        parts['xl/styles.xml'], 'fonts', f'xmlns:xr="{URIS["xr2"]}" mc:Ignorable="xr"'
    )
    original = package(parts)
    path = tmp_path / 'excel.xlsx'
    path.write_bytes(original)
    with pytest.raises(ValueError, match='prefix'):
        annotate(path, 'bottom')
    assert path.read_bytes() == original
