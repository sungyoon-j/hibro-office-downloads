import io,unittest
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
import server
from office_safety import inspect_editable

def shape(i,t):return {'id':'shape'+str(i),'type':t,'text':'상자 내용\n둘째 줄','widthMm':60,'heightMm':25,'fill':'#edf3fa','line':'#345678','textColor':'#123456','font':'Arial','fontSize':12,'align':'center'}
def export(c):
 out=io.BytesIO();server.export_doc({'kind':'document','title':'도형 검사','content':c},out);return out.getvalue()
class ShapeTests(unittest.TestCase):
 def test_shapes_remain_native_editable_objects_through_three_cycles_and_one_edit(self):
  values=[shape(i,t) for i,t in enumerate(['textbox','rect','roundRect','ellipse'],1)];c={'html':''.join('<p>앞<span data-word-shape="'+s['id']+'"></span>뒤</p>' for s in values),'shapes':values}
  for i in range(3):
   raw=export(c);inspect_editable(raw,'docx');d=Document(io.BytesIO(raw));self.assertEqual(len(d.element.xpath('.//w:pict')),4);self.assertEqual(len(d.element.xpath('.//w:txbxContent')),4);c=server.import_doc(raw);self.assertEqual(len(c['shapes']),4);self.assertEqual([v['text'] for v in c['shapes']],['상자 내용\n둘째 줄']*4);self.assertAlmostEqual(c['shapes'][0]['widthMm'],60,places=3)
  c['shapes'][1]['text']='둘째 도형만 수정';reopened=server.import_doc(export(c));self.assertEqual(reopened['shapes'][0]['text'],'상자 내용\n둘째 줄');self.assertEqual(reopened['shapes'][1]['text'],'둘째 도형만 수정')
 def test_unchanged_source_shape_xml_survives_body_edit_and_three_saves(self):
  c={'html':'<p>본문<span data-word-shape="shape1"></span></p>','shapes':[shape(1,'roundRect')]};raw=export(c);expected=etree.tostring(Document(io.BytesIO(raw)).element.xpath('.//w:pict')[0])
  for _ in range(3):
   c=server.import_doc(raw);c['html']=c['html'].replace('본문','바꾼 본문');raw=export(c);self.assertEqual(etree.tostring(Document(io.BytesIO(raw)).element.xpath('.//w:pict')[0]),expected)
 def test_unmodeled_drawingml_textbox_cannot_silently_disappear(self):
  from docx.oxml import OxmlElement
  d=Document();p=d.add_paragraph();r=p.add_run();drawing=OxmlElement('w:drawing');box=OxmlElement('w:txbxContent');box.append(OxmlElement('w:p'));drawing.append(box);r._r.append(drawing);out=io.BytesIO();d.save(out)
  with self.assertRaisesRegex(ValueError,'글상자'):inspect_editable(out.getvalue(),'docx')
 def test_orphan_duplicate_invalid_dimension_or_xml_are_refused(self):
  for mode in ['orphan','duplicate','dimension','xml']:
   value=shape(1,'rect');c={'html':'<p><span data-word-shape="shape1"></span></p>','shapes':[value]}
   if mode=='orphan':c['shapes']=[]
   if mode=='duplicate':c['html']*=2
   if mode=='dimension':value['widthMm']=float('nan')
   if mode=='xml':value['id']='"><x>'
   with self.assertRaises(ValueError):export(c)
if __name__=='__main__':unittest.main()
