import io,unittest,zipfile
from copy import deepcopy
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import parse_xml
import server
from office_safety import inspect_editable

def export(c):
 out=io.BytesIO();server.export_doc({'kind':'document','title':'각주 검사','content':c},out);return out.getvalue()
def sample():
 return {'html':'<p>본문<span data-word-note="footnote" data-word-note-id="7"></span> · 뒤<span data-word-note="endnote" data-word-note-id="2"></span></p>','notes':[{'kind':'footnote','id':'7','html':'<p><strong>첫 각주</strong> · 내용</p>'},{'kind':'endnote','id':'2','html':'<p>미주 <em>설명</em></p>'}]}
class NotesTests(unittest.TestCase):
 def test_insert_edit_remove_and_three_native_roundtrips_preserve_rich_notes(self):
  raw=export(sample());first=raw
  for i in range(3):
   inspect_editable(raw,'docx');c=server.import_doc(raw);self.assertEqual(len(c['notes']),2);self.assertEqual(c['html'].count('data-word-note='),2)
   c['html']=c['html'].replace('본문','수정 본문');raw=export(c)
   with zipfile.ZipFile(io.BytesIO(first)) as a,zipfile.ZipFile(io.BytesIO(raw)) as b:
    self.assertEqual(a.read('word/footnotes.xml'),b.read('word/footnotes.xml'));self.assertEqual(a.read('word/endnotes.xml'),b.read('word/endnotes.xml'))
  c=server.import_doc(raw);c['notes'][0]['html']=c['notes'][0]['html'].replace('첫 각주','고친 각주');raw=export(c)
  self.assertIn('고친 각주',server.import_doc(raw)['notes'][0]['html']);self.assertIn('<strong>',server.import_doc(raw)['notes'][0]['html'])
  c=server.import_doc(raw);c['notes']=[n for n in c['notes'] if n['kind']!='footnote'];import re;c['html']=re.sub('<span data-word-note="footnote".*?</span>','',c['html']);raw=export(c);self.assertEqual(len(server.import_doc(raw)['notes']),1)
 def test_duplicate_or_orphan_reference_and_nested_note_are_refused(self):
  for change in ['duplicate','orphan','nested']:
   c=sample()
   if change=='duplicate':c['html']=c['html'].replace('본문','본문<span data-word-note="footnote" data-word-note-id="7"></span>')
   if change=='orphan':c['notes']=[]
   if change=='nested':c['notes'][0]['html']='<p><span data-word-note="footnote" data-word-note-id="7"></span></p>'
   with self.assertRaises(ValueError):export(c)
 def test_malicious_original_notes_cannot_suppress_a_real_edit(self):
  c=server.import_doc(export(sample()));c['notes'][0]['html']='<p>실제 수정</p>';c['originalNotes']=deepcopy(c['notes']);self.assertIn('실제 수정',server.import_doc(export(c))['notes'][0]['html'])
if __name__=='__main__':unittest.main()
