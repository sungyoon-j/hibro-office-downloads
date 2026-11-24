import io,tempfile,unittest,zipfile,hashlib
from copy import deepcopy
from pathlib import Path
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt,Mm
from PIL import Image
import server
from office_safety import inspect_editable

def packed(doc):
 out=io.BytesIO();doc.save(out);return out.getvalue()
def fixture():
 d=Document();d.settings.odd_and_even_pages_header_footer=True
 image=io.BytesIO();Image.new('RGB',(60,30),'#238260').save(image,'PNG');image.seek(0)
 for i in range(3):
  s=d.sections[0] if i==0 else d.add_section(WD_SECTION.NEW_PAGE);d.add_paragraph('구역 본문 '+str(i+1));s.different_first_page_header_footer=i!=1
  if i==1:continue
  for side in ('header','footer'):
   for flavor,attribute in [('default',side),('first','first_page_'+side),('even','even_page_'+side)]:
    target=getattr(s,attribute);target.is_linked_to_previous=False;p=target.paragraphs[0];r=p.add_run(f'{i+1}-{side}-{flavor}');r.font.name='Arial';r.font.size=Pt(10);r.bold=True
  image.seek(0);s.header.add_paragraph().add_run().add_picture(image,width=Mm(15),height=Mm(7.5))
  for n in s._sectPr.xpath('./w:pgNumType'):s._sectPr.remove(n)
  n=OxmlElement('w:pgNumType');n.set(qn('w:start'),str(3 if i==0 else 7));s._sectPr.append(n)
 return packed(d)
def export(content):
 out=io.BytesIO();server.export_doc({'kind':'document','title':'구역 머리글','content':content},out);return out.getvalue()
def band_hashes(raw):
 with zipfile.ZipFile(io.BytesIO(raw)) as z:return {n:hashlib.sha256(z.read(n)).hexdigest() for n in z.namelist() if n.startswith(('word/header','word/footer')) and n.endswith('.xml')}

class SectionBandTests(unittest.TestCase):
 def test_first_even_odd_linked_bands_and_logos_survive_body_edit_and_three_saves(self):
  raw=fixture();expected=band_hashes(raw)
  for cycle in range(3):
   inspect_editable(raw,'docx');c=server.import_doc(raw);b=c['pageBands'];self.assertTrue(b['evenOdd']);self.assertEqual(len(b['sectionBands']),3);entries=b['sectionBands']
   self.assertTrue(entries[0]['firstDifferent']);self.assertFalse(entries[1]['firstDifferent']);self.assertTrue(entries[1]['header']['default']['linked']);self.assertEqual(entries[2]['header']['first']['text'],'3-header-first');self.assertEqual(entries[0]['header']['default']['logo']['widthMm'],15)
   c['html']=c['html'].replace('구역 본문','수정 본문');raw=export(c);self.assertEqual(band_hashes(raw),expected)
   d=Document(io.BytesIO(raw));self.assertEqual(d.sections[1].header.paragraphs[0].text,'1-header-default');self.assertEqual(d.sections[2].header.paragraphs[0].text,'3-header-default');self.assertEqual(d.sections[2]._sectPr.xpath('./w:pgNumType/@w:start'),['7'])
 def test_edit_one_section_detaches_definition_and_preserves_other_sections_and_logo(self):
  c=server.import_doc(fixture());c['pageBands']['sectionBands'][2]['header']['default']['text']='독립 머리글 수정';d=Document(io.BytesIO(export(c)))
  self.assertEqual(d.sections[0].header.paragraphs[0].text,'1-header-default');self.assertEqual(d.sections[1].header.paragraphs[0].text,'1-header-default');self.assertEqual(d.sections[2].header.paragraphs[0].text,'독립 머리글 수정');self.assertEqual(len(d.sections[2].header.paragraphs[1]._p.xpath('.//a:blip')),1)
  c=server.import_doc(packed(d));c['pageBands']['sectionBands'][2]['header']['default']['linked']=True;d=Document(io.BytesIO(export(c)));self.assertTrue(d.sections[2].header.is_linked_to_previous);self.assertEqual(d.sections[2].header.paragraphs[0].text,'1-header-default')
 def test_same_direct_relationship_in_two_sections_does_not_share_a_later_edit(self):
  d=Document(io.BytesIO(fixture()));source=d.sections[0]._sectPr.xpath('./w:headerReference[@w:type="default"]')[0];other=d.sections[2]._sectPr
  for n in other.xpath('./w:headerReference[@w:type="default"]'):other.remove(n)
  other.insert(0,deepcopy(source));c=server.import_doc(packed(d));c['pageBands']['sectionBands'][2]['header']['default']['text']='세 번째만';d=Document(io.BytesIO(export(c)))
  self.assertEqual(d.sections[0].header.paragraphs[0].text,'1-header-default');self.assertEqual(d.sections[2].header.paragraphs[0].text,'세 번째만')
 def test_inline_logo_stays_between_its_text_runs_when_header_is_edited(self):
  d=Document();d.add_paragraph('본문');p=d.sections[0].header.paragraphs[0];p.add_run('Company ');image=io.BytesIO();Image.new('RGB',(40,20),'green').save(image,'PNG');image.seek(0);p.add_run().add_picture(image,width=Mm(10));p.add_run('end')
  c=server.import_doc(packed(d));b=c['pageBands']['sectionBands'][0]['header']['default'];self.assertEqual(b['logo']['position'],'inline');self.assertEqual(b['logo']['offset'],8);self.assertFalse(b.get('advanced',False))
  b['text']='Company LTD end';b['logo']['offset']=12;d=Document(io.BytesIO(export(c)));p=d.sections[0].header.paragraphs[0];self.assertEqual(p.runs[0].text,'Company LTD ');self.assertTrue(p.runs[1]._r.xpath('.//a:blip'));self.assertEqual(p.runs[2].text,'end')
 def test_complex_native_header_is_preserved_and_refuses_lossy_rewrite(self):
  d=Document();d.add_paragraph('본문');d.sections[0].header.add_table(rows=1,cols=1,width=Mm(50)).cell(0,0).text='복합 머리글 표';raw=packed(d);c=server.import_doc(raw);self.assertTrue(c['pageBands']['sectionBands'][0]['header']['default']['advanced']);self.assertEqual(band_hashes(export(c)),band_hashes(raw))
  c['pageBands']['sectionBands'][0]['header']['default']['text']='잘못 덮기'
  with self.assertRaisesRegex(ValueError,'복합 서식'):export(c)

if __name__=='__main__':unittest.main()
