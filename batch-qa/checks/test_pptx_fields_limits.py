import unittest,io,uuid,zipfile,copy,json
from pathlib import Path
from pptx import Presentation
from pptx.util import Pt
from pptx.oxml.xmlchemy import OxmlElement
from pptx.dml.color import RGBColor
from lxml import etree
import server,slide_text,office_limits
A=slide_text.A

def raw(p):
 b=io.BytesIO();p.save(b);return b.getvalue()
def fixture():
 p=Presentation();p._element.set('firstSlideNum','7')
 for index in range(2):
  s=p.slides.add_slide(p.slide_layouts[6]);box=s.shapes.add_textbox(Pt(20),Pt(30),Pt(300),Pt(80));para=box.text_frame.paragraphs[0];para.add_run().text='번호: '
  for kind,text in [('slidenum',str(7+index)),('datetime1','11/24/2026')]:
   fld=OxmlElement('a:fld');fld.set('id','{'+str(uuid.uuid4()).upper()+'}');fld.set('type',kind);pr=OxmlElement('a:rPr');pr.set('b','1');fld.append(pr);t=OxmlElement('a:t');t.text=text;fld.append(t);para._p.append(fld)
  para.add_run().text=' 끝'
 return raw(p)
def model(data):return server.import_project('fields.pptx',data,persist=False)
def export(m):return server.export_project(m,'pptx')[0]
def fields(data):
 with zipfile.ZipFile(io.BytesIO(data)) as z:return {n:[etree.tostring(e,method='c14n') for e in etree.fromstring(z.read(n)).iter(A+'fld')] for n in z.namelist() if n.startswith('ppt/slides/slide') and n.endswith('.xml')}
class FieldLimits(unittest.TestCase):
 def test_field_keeps_native_identity_when_other_text_changes_through_three_cycles(self):
  source=fixture();m=model(source)
  for i in range(3):
   item=m['content']['slides'][0]['items'][0];item['text']=item['text'].replace('번호:' if i==0 else f'수정{i-1}:',f'수정{i}:');out=export(m);self.assertEqual(fields(source),fields(out));m=model(out);self.assertIn(f'수정{i}:',m['content']['slides'][0]['items'][0]['text'])
 def test_no_edit_zip_parts_are_identical(self):
  data=fixture();out=export(model(data))
  with zipfile.ZipFile(io.BytesIO(data)) as a,zipfile.ZipFile(io.BytesIO(out)) as b:self.assertEqual({n:a.read(n) for n in a.namelist()},{n:b.read(n) for n in b.namelist()})
 def test_whole_field_style_preserves_field_and_editing_field_turns_only_it_into_regular_text(self):
  data=fixture();m=model(data);item=m['content']['slides'][0]['items'][0];r=next(r for r in item['richText'][0]['runs'] if r.get('field',{}).get('type')=='slidenum');r['color']='#ff0000';out=export(m);again=model(out);self.assertEqual(again['content']['slides'][0]['items'][0]['richText'][0]['runs'][1]['field'],r['field'])
  item['text']=item['text'].replace('7','직접입력',1);out=export(m);self.assertEqual(sum(len(v) for v in fields(out).values()),3);self.assertIn('직접입력',model(out)['content']['slides'][0]['items'][0]['text'])
 def test_reordering_updates_display_number_and_native_field_identity(self):
  m=model(fixture());source=m['content']['slides'][0]['items'][0]['richText'][0]['runs'][1]['field'];m['content']['slides'].reverse();server.validate_project(m);self.assertTrue(m['content']['slides'][1]['items'][0]['text'].startswith('번호: 8'));out=export(m);again=model(out);r=again['content']['slides'][1]['items'][0]['richText'][0]['runs'][1];self.assertEqual(r['field'],source);self.assertEqual(r['text'],'8')
 def test_first_slide_number_change_is_saved_in_native_presentation(self):
  m=model(fixture());m['content']['firstSlideNum']=30;server.validate_project(m);out=export(m);again=model(out);self.assertEqual(again['content']['firstSlideNum'],30);self.assertTrue(again['content']['slides'][0]['items'][0]['text'].startswith('번호: 30'))
 def test_grouped_nonrect_shape_can_edit_nested_text_and_save(self):
  from pptx.enum.shapes import MSO_SHAPE
  p=Presentation();s=p.slides.add_slide(p.slide_layouts[6]);g=s.shapes.add_group_shape();n=g.shapes.add_group_shape();box=n.shapes.add_shape(MSO_SHAPE.OVAL,Pt(30),Pt(20),Pt(150),Pt(75));box.text='그룹 원본';data=raw(p);m=model(data);item=m['content']['slides'][0]['items'][0]['children'][0]['children'][0];item['text']='그룹 수정';out=export(m);self.assertEqual(model(out)['content']['slides'][0]['items'][0]['children'][0]['children'][0]['text'],'그룹 수정')
 def test_unknown_or_invalid_field_stays_guarded(self):
  p=Presentation(io.BytesIO(fixture()));p.slides[0].shapes[0]._element.xpath('.//a:fld')[0].set('type','externalExecutable')
  with self.assertRaisesRegex(ValueError,'자동 글자 필드'):model(raw(p))
  m=model(fixture());m['content']['slides'][0]['items'][0]['richText'][0]['runs'][1]['field']='invalid'
  with self.assertRaises(ValueError):server.validate_project(m)
 def test_pptx_admission_is_one_hundred_mib_other_office_stays_twenty(self):
  for name in ['a.pptx','A.PPTX','a.pdf']:self.assertEqual(office_limits.file_limit(name),100*1024**2)
  for name in ['a.docx','a.xlsx','a.hwp']:self.assertEqual(office_limits.file_limit(name),20*1024**2)
  self.assertTrue(office_limits.accepts_large_request('/api/import',{'filename':'a.pptx'}));self.assertTrue(office_limits.accepts_large_request('/api/export',{'project':{'kind':'slides'}}));self.assertFalse(office_limits.accepts_large_request('/api/import',{'filename':'a.docx'}));self.assertFalse(office_limits.accepts_large_request('/api/projects',{'kind':'document'}))
 def test_http_import_save_and_export_large_pptx(self):
  import tempfile,threading,urllib.request,urllib.error,base64
  from http.server import ThreadingHTTPServer
  source=fixture();output=io.BytesIO()
  with zipfile.ZipFile(io.BytesIO(source)) as a,zipfile.ZipFile(output,'w',compression=zipfile.ZIP_STORED) as z:
   for n in a.namelist():z.writestr(n,a.read(n))
   z.writestr('padding.dat',b'\0'*(23*1024**2))
  data=output.getvalue();old=server.DATA
  with tempfile.TemporaryDirectory() as directory:
   server.DATA=Path(directory);http=ThreadingHTTPServer(('127.0.0.1',0),server.Handler);thread=threading.Thread(target=http.serve_forever,daemon=True);thread.start()
   def post(path,value):
    request=urllib.request.Request(f'http://127.0.0.1:{http.server_port}'+path,data=json.dumps(value).encode(),headers={'Content-Type':'application/json'})
    return urllib.request.urlopen(request).read()
   try:
    imported=json.loads(post('/api/import',{'filename':'large.pptx','data':base64.b64encode(data).decode()}));imported['title']='큰 PPT 저장 검사';saved=json.loads(post('/api/projects',imported));result=post('/api/export',{'project':saved,'format':'pptx'});self.assertTrue(model(result)['content']['nativePptx'])
    with self.assertRaises(urllib.error.HTTPError) as caught:post('/api/import',{'filename':'wrong.docx','data':base64.b64encode(data).decode()})
    self.assertEqual(caught.exception.code,413)
   finally:http.shutdown();http.server_close();thread.join();server.DATA=old
 def test_large_pptx_can_save_without_weakening_expanded_zip_limits(self):
  source=fixture();out=io.BytesIO()
  with zipfile.ZipFile(io.BytesIO(source)) as a,zipfile.ZipFile(out,'w',compression=zipfile.ZIP_STORED) as z:
   for n in a.namelist():z.writestr(n,a.read(n))
   z.writestr('padding.dat',b'\0'*(21*1024**2))
  source=out.getvalue();m=model(source);again=export(m);self.assertGreater(len(source),20*1024**2);self.assertTrue(model(again)['content']['nativePptx'])
  with self.assertRaisesRegex(ValueError,'100MB'):office_limits.check_file('large.pptx',bytearray(100*1024**2+1))
if __name__=='__main__':unittest.main()
