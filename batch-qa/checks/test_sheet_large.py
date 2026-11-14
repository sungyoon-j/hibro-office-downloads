import io
import unittest
from unittest.mock import patch
from openpyxl import Workbook,load_workbook
import server

class LargeSheetTest(unittest.TestCase):
    def test_oversized_native_merge_is_rejected_before_parser_allocates_merged_cells(self):
        import zipfile
        book=Workbook();book.active['A1']='keep';stream=io.BytesIO();book.save(stream);output=io.BytesIO()
        with zipfile.ZipFile(io.BytesIO(stream.getvalue())) as source,zipfile.ZipFile(output,'w',zipfile.ZIP_DEFLATED) as target:
            for name in source.namelist():
                data=source.read(name)
                if name=='xl/worksheets/sheet1.xml':data=data.replace(b'</sheetData>',b'</sheetData><mergeCells count="1"><mergeCell ref="A1:A200001"/></mergeCells>')
                target.writestr(name,data)
        with patch.object(server,'load_workbook',side_effect=AssertionError('native parser should not run')):
            with self.assertRaisesRegex(ValueError,'200,000'):server.import_sheet(output.getvalue())
    def test_sparse_excel_boundary_and_more_than_thousand_rows_roundtrip(self):
        book=Workbook();sheet=book.active;sheet['A1']='sparse';sheet['CW1001']=42;sheet['XFD1048576']=99
        stream=io.BytesIO();book.save(stream)
        content=server.import_sheet(stream.getvalue());source=content['sheets'][0]
        self.assertEqual((source['rows'],source['cols']),(1048576,16384));self.assertEqual(source['cells']['CW1001']['raw'],'42')
        self.assertLess(len(source['cells']),10)
        project=server.validate_project({'kind':'sheet','title':'large','content':content})
        for _ in range(2):
            with patch.object(server,'soffice_path',return_value=None):raw,_=server.export_project(project,'xlsx')
            reopened=load_workbook(io.BytesIO(raw));self.assertEqual(reopened.active['XFD1048576'].value,99);self.assertEqual(reopened.active['CW1001'].value,42)
            project['content']=server.import_sheet(raw)

    def test_larger_sheet_calculates_and_caches_full_column_totals_sparsely(self):
        source={'name':'data','rows':10001,'cols':200,'cells':{'A1001':{'raw':3},'A10001':{'raw':4},'B1':{'raw':'=SUM(A:A)'},'B2':{'raw':'=COUNTIF(A:A,"")'},'B3':{'raw':'=AND(TRUE,FALSE)'},'B4':{'raw':'=NETWORKDAYS(DATE(2026,11,2),DATE(2026,11,6))'}}}
        project=server.validate_project({'kind':'sheet','title':'large-calculation','content':{'sheets':[source]}})
        with patch.object(server,'soffice_path',return_value=None):raw,_=server.export_project(project,'xlsx')
        sheet=load_workbook(io.BytesIO(raw),data_only=True).active
        self.assertEqual(sheet['B1'].value,7);self.assertEqual(sheet['B2'].value,1048574);self.assertIs(sheet['B3'].value,False);self.assertEqual(sheet['B4'].value,5)
