import io
import tempfile
import unittest
from pathlib import Path
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import server
from office_safety import inspect_editable
from document_review import revision_metadata

def packed(doc):
    stream=io.BytesIO();doc.save(stream);return stream.getvalue()

def revision(p,kind,text,identifier,author='검토자 A'):
    node=OxmlElement('w:'+kind)
    for key,value in [('id',identifier),('author',author),('date','2026-11-14T10:00:00Z')]:node.set(qn('w:'+key),value)
    run=p.add_run(text);run.bold=True;run.font.name='Arial'
    if kind=='del':
        for t in run._r.iter(qn('w:t')):t.tag=qn('w:delText')
    node.append(run._r);p._p.append(node);return node

class DocumentReviewTests(unittest.TestCase):
    def test_text_revisions_and_metadata_survive_edit_and_three_physical_docx_cycles(self):
        doc=Document();p=doc.add_paragraph('원래 본문 / ');revision(p,'ins','추가한 문구','7');p.add_run(' / ');revision(p,'del','삭제한 문구','8','검토자 B')
        p.add_run(' / 끝\t마침');cell=doc.add_table(rows=1,cols=1).cell(0,0);revision(cell.paragraphs[0],'ins','표 안의 추가','9')
        raw=packed(doc);inspect_editable(raw,'docx')
        with tempfile.TemporaryDirectory() as temp:
            for cycle in range(3):
                content=server.import_doc(raw)
                self.assertEqual(content['html'].count('data-word-revision='),3)
                self.assertIn('삭제한 문구',content['html']);self.assertIn('끝\t마침',content['html'])
                if cycle==0:content['html']=content['html'].replace('원래 본문','수정한 본문')
                destination=Path(temp)/f'cycle-{cycle}.docx';server.export_doc({'kind':'document','title':'합성 검토','content':content},destination);raw=destination.read_bytes()
                reopened=Document(io.BytesIO(raw));tree=reopened.element
                self.assertIn('수정한 본문',''.join(tree.xpath('.//w:t/text()')))
                self.assertEqual([n.get(qn('w:id')) for n in tree.xpath('.//w:ins')],['7','9'])
                self.assertEqual(tree.xpath('.//w:del/@w:author'),['검토자 B'])
                self.assertEqual(tree.xpath('.//w:del//w:delText/text()'),['삭제한 문구'])
                self.assertEqual(tree.xpath('.//w:del//w:t'),[])
                self.assertEqual(tree.xpath('.//w:ins//w:rPr/w:b/@w:val'),[])
                self.assertEqual(len(tree.xpath('.//w:ins//w:rPr/w:b')),2)
                inspect_editable(raw,'docx')

    def test_inserted_text_can_be_edited_without_erasing_other_revision(self):
        doc=Document();p=doc.add_paragraph('앞 ');revision(p,'ins','삽입','1');revision(p,'del','삭제','2');content=server.import_doc(packed(doc));content['html']=content['html'].replace('삽입','교체')
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'edited.docx';server.export_doc({'kind':'document','title':'검토','content':content},path);reopened=Document(path)
            self.assertEqual(reopened.element.xpath('.//w:ins//w:t/text()'),['교체']);self.assertEqual(reopened.element.xpath('.//w:del//w:delText/text()'),['삭제'])

    def test_structural_and_formatting_revisions_are_refused_before_loss(self):
        for tag in ['rPrChange','moveFrom','ins']:
            doc=Document();p=doc.add_paragraph('유지');node=OxmlElement('w:'+tag);node.set(qn('w:id'),'1');p._p.get_or_add_pPr().append(node)
            with self.assertRaisesRegex(ValueError,'변경 추적'):inspect_editable(packed(doc),'docx')

    def test_revision_metadata_is_typed_and_xml_escaped(self):
        for args in [('insert','-1'),('insert','2147483648'),('other','1'),('delete','1','a\nb'),('insert','1','','yesterday')]:
            with self.assertRaises(ValueError):revision_metadata(*args)
        doc=Document();revision(doc.add_paragraph(),'ins','<유지>','1','A & "B"');content=server.import_doc(packed(doc));self.assertIn('A &amp; &quot;B&quot;',content['html']);self.assertIn('&lt;유지&gt;',content['html'])

    def test_existing_comment_and_range_survive_three_saves_and_comment_edit(self):
        doc=Document();p=doc.add_paragraph('앞 ');r=p.add_run('메모 대상');p.add_run(' 뒤');c=doc.add_comment(r,'첫 줄\n두 번째 줄',author='검토자',initials='AB');identifier=str(c.comment_id)
        raw=packed(doc)
        with tempfile.TemporaryDirectory() as temp:
            for i in range(3):
                inspect_editable(raw,'docx');content=server.import_doc(raw);self.assertEqual(len(content['comments']),1);self.assertEqual(content['html'].count('data-word-comment='),3)
                if i==0:content['comments'][0]['text']='고친 메모\n줄바꿈 유지'
                path=Path(temp)/f'comment-{i}.docx';server.export_doc({'kind':'document','title':'검토','content':content},path);raw=path.read_bytes();reopened=Document(path)
                self.assertEqual(reopened.comments.get(int(identifier)).text,'고친 메모\n줄바꿈 유지')
                self.assertEqual(reopened.element.xpath('.//w:commentReference/@w:id'),[identifier]);self.assertEqual(reopened.element.xpath('.//w:commentRangeStart/@w:id'),[identifier]);self.assertEqual(reopened.element.xpath('.//w:commentRangeEnd/@w:id'),[identifier])

    def test_new_comment_is_native_word_comment_and_orphan_anchor_is_refused(self):
        content={'html':'<p>앞 <span data-word-comment="start" data-word-comment-id="4"></span>선택<span data-word-comment="end" data-word-comment-id="4"></span><span data-word-comment="reference" data-word-comment-id="4"></span> 뒤</p>','comments':[{'id':'4','text':'새 메모','author':'작성자','date':'2026-11-14T10:00:00Z'}]}
        with tempfile.TemporaryDirectory() as temp:
            path=Path(temp)/'new.docx';server.export_doc({'kind':'document','title':'검토','content':content},path);doc=Document(path);self.assertEqual(doc.comments.get(4).text,'새 메모');self.assertEqual(doc.comments.get(4).author,'작성자')
            content['html']=content['html'].replace('data-word-comment="start"','data-word-comment="end"')
            with self.assertRaisesRegex(ValueError,'메모 범위'):server.export_doc({'kind':'document','title':'검토','content':content},path)

if __name__=='__main__':unittest.main()
