"""Physical roundtrips using only the actually installed product's backend/runtime."""
from pathlib import Path
import hashlib, io, json, os, sys
from unittest.mock import patch

resources = Path(sys.argv[1]).resolve()
out = Path(sys.argv[2]).resolve()
out.mkdir(parents=True, exist_ok=True)
os.environ.update(HIBRO_RESOURCE_ROOT=str(resources / 'workspace'), MOA_DATA_DIR=str(out / 'isolated-data'), HIBRO_HWP_BINARY=str(resources / 'hwp' / ('hwp.exe' if os.name=='nt' else 'hwp')))
sys.dont_write_bytecode = True
sys.path.insert(0, str(resources / 'workspace'))
import server
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image
from openpyxl import load_workbook

manifest = json.loads((Path(__file__).parent / 'resource-manifest.json').read_text())
for name, digest in manifest['files'].items():
    assert hashlib.sha256((resources / 'workspace' / name).read_bytes()).hexdigest() == digest, name

def pack(book):
    stream = io.BytesIO()
    book.save(stream)
    return stream.getvalue()

def order(paragraph):
    result = []
    for node in paragraph._p.xpath('./w:r/w:t | ./w:r/w:tab | ./w:r/w:br | .//a:blip'):
        if node.tag == qn('a:blip'):
            blob = paragraph.part.related_parts[node.get(qn('r:embed'))].blob
            extent = node.xpath('ancestor::wp:inline/wp:extent')[0]
            result.append([hashlib.sha256(blob).hexdigest(), int(extent.get('cx')), int(extent.get('cy'))])
        else:
            result.append(node.text if node.tag == qn('w:t') else node.tag)
    return result

images = []
for color in ['red', 'blue']:
    stream = io.BytesIO()
    Image.new('RGB', (12, 8), color).save(stream, 'PNG')
    images.append(stream.getvalue())
doc = Document()
for paragraph in [doc.add_paragraph(), doc.add_table(1, 1).cell(0, 0).paragraphs[0]]:
    run = paragraph.add_run('Before')
    run.bold = True
    run.add_picture(io.BytesIO(images[0]), width=Pt(27))
    run.add_text('Middle')
    run.add_tab()
    run.add_break()
    run.add_picture(io.BytesIO(images[1]), width=Pt(36))
    run.add_text('After')
raw = pack(doc)
expected = order(doc.paragraphs[0])
previous = 'Middle'
for cycle in range(1, 4):
    project = server.import_project('synthetic.docx', raw, persist=False)
    assert project['content']['wordPropertiesVersion'] == 3
    replacement = 'Edited' + str(cycle)
    project['content']['html'] = project['content']['html'].replace(previous, replacement)
    raw, _ = server.export_project(project, 'docx')
    file = out / ('word-' + str(cycle) + '.docx')
    file.write_bytes(raw)
    reopened = Document(file)
    expected[2] = replacement
    for paragraph in [reopened.paragraphs[0], reopened.tables[0].cell(0, 0).paragraphs[0]]:
        assert order(paragraph) == expected
        assert all(run.bold for run in paragraph.runs if run.text.strip())
    previous = replacement

formulas = {'B1':'=SUM("5",15,TRUE)', 'B2':'=COUNT(5,"5",TRUE)', 'B3':'=COUNTA(A4)',
            'B4':'=COUNTA(A5)', 'B5':'=COUNT(A5)', 'B6':'=SUM(A1:A4)', 'B7':'=-2^2',
            'B8':'=2^3^2', 'B9':'=IFERROR(1/0,"#N/A")', 'B10':'=SUM(INDEX(A1:A3,0,1))',
            'B11':'=SUBTOTAL(3,A1:A6)', 'B12':'=SUM(A5)', 'B13':'=B7*A1'}
cells = {'A1':{'raw':10}, 'A2':{'raw':'5','valueType':'text'}, 'A3':{'raw':True},
         'A4':{'raw':'=""'}, 'A5':{'raw':'#DIV/0!','valueType':'error'},
         **{ref:{'raw':value} for ref,value in formulas.items()}}
project = server.validate_project({'kind':'sheet','title':'Synthetic core results',
                                  'content':{'sheets':[{'name':'Results','rows':40,'cols':12,'cells':cells}]}})
for cycle in range(1, 4):
    project['content']['sheets'][0]['cells']['A1']['raw'] = cycle * 10
    with patch.object(server, 'soffice_path', return_value=None):
        raw, _ = server.export_project(project, 'xlsx')
    file = out / ('sheet-' + str(cycle) + '.xlsx')
    file.write_bytes(raw)
    saved = load_workbook(file, data_only=False).active
    values = load_workbook(file, data_only=True).active
    expected_values = {'B1':21,'B2':3,'B3':1,'B4':1,'B5':0,'B6':cycle*10,'B7':4,'B8':64,
                       'B9':'#N/A','B10':cycle*10,'B11':5,'B12':'#DIV/0!','B13':cycle*40}
    for ref,value in expected_values.items():
        assert values[ref].value == value, (ref, values[ref].value, value)
        assert saved[ref].value == formulas[ref], ref
    assert values['B9'].data_type == 's' and values['B12'].data_type == 'e'
    project = server.import_project(file.name, file.read_bytes(), persist=False)
    assert project['content']['sheets'][0]['cells']['A2']['valueType'] == 'text'

# Values observed in actual Microsoft Excel, rather than this application's evaluator.
fixtures = Path(__file__).parent / 'fixtures'
golden = json.loads((fixtures / 'excel-functions-observed.json').read_text())
raw = (fixtures / 'Function-Comparison.xlsx').read_bytes()
for cycle in range(1, 4):
    project = server.import_project('Excel-functions.xlsx', raw, persist=False)
    project['content']['sheets'][0]['cells']['A1']['raw'] = 'Installed cycle ' + str(cycle)
    with patch.object(server, 'soffice_path', return_value=None):
        raw, _ = server.export_project(project, 'xlsx')
    file = out / ('excel-functions-' + str(cycle) + '.xlsx')
    file.write_bytes(raw)
    saved = load_workbook(file, data_only=False).active
    values = load_workbook(file, data_only=True).active
    for row, (label, formula, expected) in enumerate(golden['formulasAndValues'], 2):
        cell = values['B' + str(row)]
        assert saved[cell.coordinate].value == formula, (label, 'formula changed')
        assert cell.value == expected, (label, cell.value, expected, cycle)
        if isinstance(expected, bool):
            assert cell.data_type == 'b', (label, cell.data_type)
        elif isinstance(expected, str) and expected.startswith('#'):
            assert cell.data_type == 'e', (label, cell.data_type)
    assert values['A1'].value == 'Installed cycle ' + str(cycle)

# Extra acceptance checks run against this installed runtime/source only.
import unittest
sys.path.insert(0,str(Path(__file__).parent/'checks'))
loader=unittest.TestLoader();batch=unittest.TestSuite()
for name in ['test_document_review','test_document_compare','test_document_notes',
             'test_document_shapes','test_document_section_bands','test_sheet_large',
             'test_hangul_image_wrap','test_slide_richtext']:
    batch.addTests(loader.loadTestsFromName(name))
acceptance=unittest.TextTestRunner(stream=sys.stderr,verbosity=2).run(batch)
assert acceptance.wasSuccessful(), 'Installed batch acceptance failed'
print(json.dumps({'passed':True, 'execution':'installed native Python and shipped resources; no UI automation',
                  'resourceHashesMatched':len(manifest['files']), 'batchAcceptanceTests':acceptance.testsRun, 'batchAcceptancePassed':True, 'wordCycles':3, 'sheetCycles':3,
                  'actualExcelGoldenCases':len(golden['formulasAndValues']), 'actualExcelGoldenCycles':3,
                  'wordTextImageOrderAndImageBytes':True, 'formulaStringsAndCachedTypes':True,
                  'externalRecalculationRequired':False, 'finalValues':expected_values}))
