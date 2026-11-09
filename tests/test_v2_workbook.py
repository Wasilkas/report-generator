"""Generic OOXML preserves characterized payloads, original bytes and validation order."""

import base64
import io
import json
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

import pytest
from openpyxl import load_workbook

from report_generator.v2 import (
    MetadataFormat,
    annotate_workbook,
    original_workbook,
    package_fingerprint,
    read_metadata,
)
from report_generator.v2.workbook import _parse, _xml

FORMAT = MetadataFormat(
    'customXml/clearml-yolo-identity.xml',
    'https://clearml-yolo/model-identity/v1',
    'identity',
    '1',
    'clearmlYoloIdentity',
)
FIXTURE = Path(__file__).parent / 'fixtures/ooxml-characterization.json'


def characterized():
    value = json.loads(FIXTURE.read_text())
    return (
        base64.b64decode(value['original']),
        {key: base64.b64decode(content) for key, content in value['parts'].items()},
        value,
    )


def package(parts):
    output = io.BytesIO()
    with ZipFile(output, 'w', compression=ZIP_DEFLATED) as archive:
        for name, content in parts.items():
            archive.writestr(name, content)
    return output.getvalue()


def test_characterized_v1_payloads_byte_exact_recovery_and_reannotation(tmp_path):
    original, expected, fixture = characterized()
    path = tmp_path / 'book.xlsx'
    path.write_bytes(original)
    annotate_workbook(
        path,
        banner_lines=fixture['banner_lines'],
        metadata_texts=fixture['metadata_texts'],
        format=FORMAT,
    )
    with ZipFile(path) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    assert parts == expected
    assert original_workbook(path.read_bytes(), FORMAT) == original
    assert original_workbook(original, FORMAT) == original
    assert read_metadata(original, FORMAT) is None
    first = parts.copy()
    annotate_workbook(
        path,
        banner_lines=fixture['banner_lines'],
        metadata_texts=fixture['metadata_texts'],
        format=FORMAT,
    )
    with ZipFile(path) as archive:
        assert {name: archive.read(name) for name in archive.namelist()} == first
    workbook = load_workbook(path)
    shift = len(fixture['banner_lines'])
    main = workbook['Main sheet']
    other = workbook['Other']
    assert main.cell(3 + shift, 2).value == f'=SUM($B${2 + shift}:B{2 + shift})'
    assert other.cell(1 + shift, 1).value == f"='Main sheet'!$B${2 + shift}+1"
    assert main.tables['Results'].ref == f'A{1 + shift}:B{2 + shift}'
    assert main.auto_filter.ref == f'A{1 + shift}:B{2 + shift}'
    assert main.data_validations.dataValidation[0].formula1 == f'B{2 + shift}'
    assert main.freeze_panes == f'B{2 + shift}'
    assert main.row_breaks.brk[0].id == 2 + shift
    assert main.print_title_cols == '$A:$A'
    for sheet in workbook:
        assert sheet.page_setup.fitToWidth == 1
        assert sheet.print_title_rows.startswith('$1:')
        assert all(sheet.cell(row, 1).data_type == 's' for row in range(1, shift + 1))
    workbook.close()


@pytest.mark.parametrize('part', ['xl/worksheets/sheet1.xml', FORMAT.path, 'xl/styles.xml'])
def test_fingerprint_rejects_edited_annotated_parts(part):
    _, parts, _ = characterized()
    if part == FORMAT.path:
        root = ET.fromstring(parts[part])
        root.find(f'{{{FORMAT.namespace}}}models').text = '{}'
        parts[part] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    else:
        parts[part] = parts[part].replace(b'>', b'> ', 1)
    with pytest.raises(ValueError, match='changed after'):
        original_workbook(package(parts), FORMAT)


def test_schema_owner_validator_runs_before_fingerprint_comparison():
    _, parts, _ = characterized()
    root = ET.fromstring(parts[FORMAT.path])
    root.find(f'{{{FORMAT.namespace}}}models').text = 'invalid'
    parts[FORMAT.path] = ET.tostring(root, encoding='utf-8', xml_declaration=True)

    def validator(metadata):
        raise ValueError('schema owner rejected data')

    with pytest.raises(ValueError, match='schema owner rejected'):
        original_workbook(package(parts), FORMAT, validator)


@pytest.mark.parametrize(
    'change,error',
    [('version', 'Unsupported'), ('fingerprint', 'fingerprint'), ('original', 'original content')],
)
def test_metadata_required_envelope_fields(change, error):
    _, parts, _ = characterized()
    root = ET.fromstring(parts[FORMAT.path])
    if change == 'version':
        root.set('version', '2')
    elif change == 'fingerprint':
        root.attrib.pop('package_sha256')
    else:
        root.remove(root.find(f'{{{FORMAT.namespace}}}original'))
    parts[FORMAT.path] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    if change == 'original':
        root.set('package_sha256', package_fingerprint(parts, FORMAT))
        parts[FORMAT.path] = ET.tostring(root, encoding='utf-8', xml_declaration=True)
    with pytest.raises(ValueError, match=error):
        original_workbook(package(parts), FORMAT)


@pytest.mark.parametrize('content', [b'<!DOCTYPE x><x/>', b'<!ENTITY x "y"><x/>', b'\x00<x/>'])
def test_unsafe_xml_rejected(content):
    with pytest.raises(ValueError, match='DTDs or entities'):
        _parse(content)


def test_nonapplication_metadata_envelope_and_literal_banner(tmp_path):
    original, _, _ = characterized()
    path = tmp_path / 'book.xlsx'
    path.write_bytes(original)
    format = MetadataFormat(
        'customXml/acme.xml', 'urn:acme:metadata', 'report', '2026', 'acmeReport'
    )
    annotate_workbook(
        path,
        banner_lines=['=literal', '  Unicode 模型  '],
        metadata_texts={'owner': 'Acme', 'run': 'full-id'},
        format=format,
    )
    metadata = read_metadata(path.read_bytes(), format)
    assert metadata.findtext('{urn:acme:metadata}owner') == 'Acme'
    assert original_workbook(path.read_bytes(), format) == original
    workbook = load_workbook(path)
    for sheet in workbook:
        assert sheet['A1'].value == '=literal' and sheet['A1'].data_type == 's'
        assert sheet['A2'].value == '  Unicode 模型  '
    workbook.close()


def _annotate_bottom(path, lines, **kwargs):
    annotate_workbook(
        path,
        banner_lines=lines,
        metadata_texts={'models': '{}'},
        format=FORMAT,
        placement='bottom',
        **kwargs,
    )


def _sheet_parts(path):
    with ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def test_bottom_placement_keeps_coordinates_and_follows_drawings(tmp_path):
    original, _, _ = characterized()
    path = tmp_path / 'book.xlsx'
    path.write_bytes(original)
    lines = ['Candidate: =model 模型', 'Training task: open in ClearML', 'plain']
    links = ['https://app.example/projects/*/models/m1', 'https://app.example/t/1', None]
    _annotate_bottom(
        path, lines, banner_links=links, page_footer='Candidate: A&B · Training source: recorded'
    )
    workbook = load_workbook(path)
    main, other = workbook['Main sheet'], workbook['Other']
    # Original coordinates, formulas and dependent ranges are unchanged.
    assert main['A1'].value == 'Class' and main['B3'].value == '=SUM($B$2:B2)'
    assert other['A1'].value == "='Main sheet'!$B$2+1"
    assert main.tables['Results'].ref == 'A1:B2'
    assert main.auto_filter.ref == 'A1:B2'
    assert main.data_validations.dataValidation[0].formula1 == 'B2'
    assert main.freeze_panes == 'B2'
    assert main.print_title_rows == '$1:$1'
    assert other.print_title_rows == '$1:$2'
    # The chart starts at row 2 and spans ~2700000 EMU (15 rows of 15pt); the footer follows it.
    start = 2 + 15 + 1
    assert main.cell(start - 1, 1).value is None
    assert [main.cell(start + offset, 1).value for offset in range(3)] == lines
    assert all(main.cell(start + offset, 1).data_type == 's' for offset in range(3))
    assert main.cell(start, 1).hyperlink.target == links[0]
    assert main.cell(start + 1, 1).hyperlink.target == links[1]
    assert main.cell(start + 2, 1).hyperlink is None
    assert main.print_area == f"'Main sheet'!$A$1:$E${start + 2}"
    # Excel footer codes start with &; a literal ampersand is stored doubled.
    assert main.oddFooter.left.text == 'Candidate: A&&B · Training source: recorded'
    # Sheets without drawings start one empty row after their last cell.
    source = load_workbook(io.BytesIO(original))
    other_last = source['Other'].max_row
    source.close()
    assert other.cell(other_last + 1, 1).value is None
    assert other.cell(other_last + 2, 1).value == lines[0]
    workbook.close()
    assert original_workbook(path.read_bytes(), FORMAT) == original
    first = _sheet_parts(path)
    _annotate_bottom(
        path, lines, banner_links=links, page_footer='Candidate: A&B · Training source: recorded'
    )
    assert _sheet_parts(path) == first


def test_bottom_placement_without_lines_is_metadata_only(tmp_path):
    original, _, _ = characterized()
    path = tmp_path / 'book.xlsx'
    path.write_bytes(original)
    _annotate_bottom(path, [])
    with ZipFile(io.BytesIO(original)) as archive:
        sheet = archive.read('xl/worksheets/sheet1.xml')
    assert _sheet_parts(path)['xl/worksheets/sheet1.xml'] == _xml(_parse(sheet))
    assert read_metadata(path.read_bytes(), FORMAT) is not None
    assert original_workbook(path.read_bytes(), FORMAT) == original


def test_bottom_placement_on_empty_sheet_starts_at_second_row(tmp_path):
    from openpyxl import Workbook

    path = tmp_path / 'empty.xlsx'
    Workbook().save(path)
    _annotate_bottom(path, ['Model: name'])
    workbook = load_workbook(path)
    assert workbook.active['A1'].value is None
    assert workbook.active['A2'].value == 'Model: name'
    workbook.close()


def test_page_footer_escapes_codes_and_truncates_to_excel_limit(tmp_path):
    original, _, _ = characterized()
    path = tmp_path / 'book.xlsx'
    path.write_bytes(original)
    text = 'Role: ' + '&' * 300
    _annotate_bottom(path, ['line'], page_footer=text)
    sheet = _parse(_sheet_parts(path)['xl/worksheets/sheet1.xml'])
    footer = sheet.find(
        '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}headerFooter/'
        '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}oddFooter'
    ).text
    assert len(footer) <= 255
    assert footer.startswith('&LRole: ') and footer.endswith('…')
    assert (len(footer[len('&LRole: ') : -1])) % 2 == 0


@pytest.mark.parametrize(
    'kwargs,error',
    [
        ({'placement': 'top', 'banner_lines': ['x'], 'page_footer': 'f'}, 'bottom placement'),
        ({'placement': 'top', 'banner_lines': ['x'], 'banner_links': [None]}, 'bottom placement'),
        ({'placement': 'top', 'banner_lines': []}, 'literal banner lines'),
        ({'placement': 'side', 'banner_lines': ['x']}, 'placement'),
        ({'placement': 'bottom', 'banner_lines': ['x'], 'banner_links': []}, 'one to one'),
        (
            {'placement': 'bottom', 'banner_lines': ['x'], 'banner_links': ['javascript:alert(1)']},
            'http',
        ),
        (
            {'placement': 'bottom', 'banner_lines': ['x'], 'banner_links': ['file:///etc/passwd']},
            'http',
        ),
    ],
)
def test_invalid_placement_arguments_are_rejected(tmp_path, kwargs, error):
    original, _, _ = characterized()
    path = tmp_path / 'book.xlsx'
    path.write_bytes(original)
    with pytest.raises(ValueError, match=error):
        annotate_workbook(path, metadata_texts={}, format=FORMAT, **kwargs)
    assert path.read_bytes() == original
