import io,tempfile,unittest
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
import server
from document_compare import compare_documents
from office_safety import inspect_editable

def content(text):
 d=Document();p=d.add_paragraph();p.add_run(text).bold=True;p.add_run(' / 끝').italic=True
 cell=d.add_table(rows=1,cols=1).cell(0,0);cell.text='표 원본 100';stream=io.BytesIO();d.save(stream);return server.import_doc(stream.getvalue())

class DocumentCompareTests(unittest.TestCase):
 def test_independent_comparison_retains_styles_tables_and_native_revision_roundtrip(self):
  original=content('기존 문구');revised=content('새 문구');revised['html']=revised['html'].replace('표 원본 100','표 수정 200');before,after=deepcopy(original),deepcopy(revised)
  result=compare_documents(original,revised,'합성 비교');self.assertEqual(result['summary']['paragraphs'],2);self.assertEqual(original,before);self.assertEqual(revised,after)
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'compare.docx';server.export_doc({'kind':'document','title':'비교','content':result['content']},p);raw=p.read_bytes();inspect_editable(raw,'docx');doc=Document(p)
   self.assertEqual(len(doc.tables),1);self.assertEqual(doc.element.xpath('.//w:del//w:delText/text()'),['기존','원본','100']);self.assertEqual(doc.element.xpath('.//w:ins//w:t/text()'),['새','수정','200'])
   self.assertTrue(doc.element.xpath('.//w:ins//w:rPr/w:b'));self.assertTrue(doc.element.xpath('.//w:rPr/w:i'))
   next_content=server.import_doc(raw);server.export_doc({'kind':'document','title':'재열기','content':next_content},p);self.assertEqual(len(Document(p).element.xpath('.//w:ins')),3)
 def test_unchanged_document_has_no_synthetic_revisions(self):
  c=content('변경 없음');r=compare_documents(c,c);self.assertEqual(r['summary'],{'insertions':0,'deletions':0,'paragraphs':0});self.assertNotIn('data-word-revision=',r['content']['html'])
 def test_structure_and_existing_review_records_are_refused_without_mutation(self):
  a=content('원본');b=content('수정');b['html']+='<p>새 문단</p>'
  with self.assertRaisesRegex(ValueError,'구성이 다른'):compare_documents(a,b)
  b=content('수정');b['comments']=[{'id':'1','text':'검토'}]
  with self.assertRaisesRegex(ValueError,'기존 변경 기록'):compare_documents(a,b)

if __name__=='__main__':unittest.main()
