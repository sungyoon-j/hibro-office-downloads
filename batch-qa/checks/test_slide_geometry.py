from pathlib import Path
from copy import deepcopy
import io,json,tempfile,unittest,uuid,zipfile,threading,urllib.request
from http.server import ThreadingHTTPServer
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.util import Pt
from pptx.oxml.xmlchemy import OxmlElement
import server,slide_geometry as geometry
from office_safety import inspect_editable

A=geometry.A;P=geometry.P
def bytes_of(p):
 b=io.BytesIO();p.save(b);return b.getvalue()
def fixture():
 p=Presentation();s=p.slides.add_slide(p.slide_layouts[6])
 one=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Pt(25),Pt(30),Pt(160),Pt(90));one.text='원본 도형';two=s.shapes.add_shape(MSO_SHAPE.OVAL,Pt(240),Pt(130),Pt(140),Pt(80));two.text='타원'
 c=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Pt(185),Pt(75),Pt(240),Pt(170));c.line.width=Pt(1.5);ln=c._element.find(P+'spPr/'+A+'ln');head=OxmlElement('a:tailEnd');head.set('type','triangle');ln.append(head)
 return p
def canonical(raw):
 with zipfile.ZipFile(io.BytesIO(raw)) as z:return {n:z.read(n) for n in z.namelist()}
def nodes(raw,tag):
 with zipfile.ZipFile(io.BytesIO(raw)) as z:
  root=etree.fromstring(z.read('ppt/slides/slide1.xml'));return [etree.tostring(n,method='c14n') for n in root.iter(tag)]
class GeometryTests(unittest.TestCase):
 def export(self,content):
  with tempfile.TemporaryDirectory() as d:
   f=Path(d)/'roundtrip.pptx';server.export_slides({'kind':'slides','title':'검사','content':content},f);return f.read_bytes()
 def test_all_standard_shapes_import_and_no_edit_save_preserves_package(self):
  p=Presentation()
  for i,name in enumerate(geometry.presets()):
   if i%16==0:s=p.slides.add_slide(p.slide_layouts[6])
   shape=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,Pt(10+(i%4)*170),Pt(10+((i%16)//4)*110),Pt(150),Pt(90));shape._element.find(P+'spPr/'+A+'prstGeom').set('prst',name)
  raw=bytes_of(p);inspect_editable(raw,'pptx');model=server.import_slides(raw);server.validate_project({'kind':'slides','content':model})
  self.assertEqual(sum(bool(item.get('geometry')) for s in model['slides'] for item in s['items']),186)
  self.assertEqual(canonical(raw),canonical(self.export(model)))
 def test_opc_entry_order_is_preserved_but_changed_targets_are_not_ignored(self):
  from slide_preservation import equal_part
  raw=bytes_of(fixture());out=io.BytesIO()
  with zipfile.ZipFile(io.BytesIO(raw)) as source,zipfile.ZipFile(out,'w') as target:
   for info in source.infolist():
    data=source.read(info.filename)
    if info.filename.endswith('.rels') or info.filename=='[Content_Types].xml':
     root=etree.fromstring(data);children=list(root)
     for n in children:root.remove(n)
     root.extend(reversed(children));data=etree.tostring(root)
    target.writestr(info,data)
  raw=out.getvalue();self.assertEqual(canonical(raw),canonical(self.export(server.import_slides(raw))))
  a=b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Target="a.xml" Type="test"/></Relationships>'
  self.assertFalse(equal_part(a,a.replace(b'a.xml',b'b.xml')))
 def test_dense_slide_uses_the_existing_bounded_thousand_shape_limit(self):
  p=Presentation();s=p.slides.add_slide(p.slide_layouts[6])
  for i in range(120):s.shapes.add_shape(MSO_SHAPE.OVAL,Pt(i%20*30),Pt(i//20*40),Pt(25),Pt(25))
  raw=bytes_of(p);m=server.import_project('many-shapes.pptx',raw,persist=False);self.assertEqual(len(m['content']['slides'][0]['items']),120);self.assertEqual(canonical(raw),canonical(self.export(m['content'])))
  item=m['content']['slides'][0]['items'][0];m['content']['slides'][0]['items']=[deepcopy(item) for _ in range(1001)]
  with self.assertRaises(ValueError):server.validate_project(m)
 def test_connections_and_geometry_survive_text_move_save_and_reopen(self):
  raw=bytes_of(fixture());inspect_editable(raw,'pptx');model=server.import_slides(raw);preset=model['slides'][0]['items'][2]['geometry']['preset']
  for cycle in range(3):
   one=model['slides'][0]['items'][0];one['text']='편집 검사';one['richText'][0]['runs'][0]['text']='편집 검사';one['x']+=5
   out=self.export(model);self.assertEqual(nodes(raw,A+'prstGeom'),nodes(out,A+'prstGeom'));self.assertEqual(nodes(raw,P+'nvCxnSpPr'),nodes(out,P+'nvCxnSpPr'))
   model=server.import_slides(out);self.assertEqual(model['slides'][0]['items'][0]['text'],'편집 검사');self.assertEqual(model['slides'][0]['items'][2]['geometry']['preset'],preset)
 def test_adjustments_are_kept_and_recipe_rewrites_are_refused(self):
  p=fixture();shape=p.slides[0].shapes[0];shape.adjustments[0]=.3;raw=bytes_of(p);m=server.import_slides(raw);self.assertEqual(m['slides'][0]['items'][0]['geometry']['adjust'][0][1],['val','30000'])
  m['slides'][0]['items'][0]['geometry']['adjust'][0][1]=['val','20000']
  with self.assertRaisesRegex(ValueError,'경로'):self.export(m)
 def test_enabling_same_color_fill_and_line_survives_save(self):
  p=fixture();shape=p.slides[0].shapes[0];shape.fill.background();shape.line.fill.background();m=server.import_slides(bytes_of(p));item=m['slides'][0]['items'][0];item.pop('transparentFill');item['strokeWidth']=2
  reopened=Presentation(io.BytesIO(self.export(m)));shape=reopened.slides[0].shapes[0];self.assertEqual(int(shape.fill.type),1);self.assertEqual(int(shape.line.fill.type),1);self.assertEqual(shape.line.width.pt,2)
 def test_connector_stroke_edit_preserves_arrow_and_has_no_fill_control_data(self):
  raw=bytes_of(fixture());m=server.import_slides(raw);item=m['slides'][0]['items'][2];item['stroke']='#abcdef';item['strokeWidth']=3
  out=self.export(m);shape=Presentation(io.BytesIO(out)).slides[0].shapes[2];self.assertEqual(str(shape.line.color.rgb),'ABCDEF');self.assertEqual(shape.line.width.pt,3);self.assertEqual(nodes(raw,A+'tailEnd'),nodes(out,A+'tailEnd'));self.assertTrue(server.import_slides(out)['slides'][0]['items'][2]['geometry']['connector'])
  item['fill']='#abcdef'
  with self.assertRaisesRegex(ValueError,'연결선'):self.export(m)
 def test_native_fill_opacity_is_visible_and_retained_on_recolor(self):
  from pptx.dml.color import RGBColor
  p=fixture();shape=p.slides[0].shapes[0];shape.fill.solid();shape.fill.fore_color.rgb=RGBColor(18,52,86);solid=shape._element.find(P+'spPr/'+A+'solidFill');alpha=OxmlElement('a:alpha');alpha.set('val','35000');next(iter(solid)).append(alpha)
  m=server.import_slides(bytes_of(p));item=m['slides'][0]['items'][0];self.assertEqual(item['geometry']['fillOpacity'],.35);item['fill']='#abcdef';out=self.export(m);again=server.import_slides(out);self.assertEqual(again['slides'][0]['items'][0]['geometry']['fillOpacity'],.35);self.assertEqual(again['slides'][0]['items'][0]['fill'].lower(),'#abcdef')
 def test_unsafe_geometry_and_unsupported_media_remain_guarded(self):
  raw=bytes_of(fixture());m=server.import_slides(raw);g=m['slides'][0]['items'][0]['geometry'];g['paths'][0]['commands'][0][1]='<script>'
  with self.assertRaises(ValueError):server.validate_project({'kind':'slides','content':m})
  p=fixture();timing=OxmlElement('p:timing');p.slides[0]._element.append(timing)
  with self.assertRaisesRegex(ValueError,'애니메이션'):inspect_editable(bytes_of(p),'pptx')
 def test_group_duplicate_updates_connector_local_shape_references(self):
  p=Presentation();s=p.slides.add_slide(p.slide_layouts[6]);g=s.shapes.add_group_shape();a=g.shapes.add_shape(MSO_SHAPE.OVAL,Pt(10),Pt(10),Pt(50),Pt(50));b=g.shapes.add_shape(MSO_SHAPE.DIAMOND,Pt(120),Pt(20),Pt(50),Pt(50));c=g.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Pt(60),Pt(35),Pt(120),Pt(45))
  props=c._element.find(P+'nvCxnSpPr/'+P+'cNvCxnSpPr')
  for tag,shape in [('stCxn',a),('endCxn',b)]:n=OxmlElement('a:'+tag);n.set('id',str(shape.shape_id));n.set('idx','0');props.append(n)
  m=server.import_slides(bytes_of(p));copy=deepcopy(m['slides'][0]['items'][0]);copy['id']=str(uuid.uuid4());m['slides'][0]['items'].append(copy);out=self.export(m);opened=Presentation(io.BytesIO(out));groups=list(opened.slides[0].shapes);self.assertEqual(len(groups),2)
  for group in groups:
   ids={str(sh.shape_id) for sh in group.shapes}
   for n in group._element.xpath('.//a:stCxn|.//a:endCxn'):self.assertIn(n.get('id'),ids)
 def test_legacy_display_upgrade_keeps_user_text_and_adds_geometry(self):
  raw=bytes_of(fixture());m=server.import_slides(raw);m['slideDisplayVersion']=3;one=m['slides'][0]['items'][0];one.pop('geometry');one['text']='사용자 편집';one['richText'][0]['runs'][0]['text']='사용자 편집'
  from slide_preservation import upgrade_display
  upgraded=upgrade_display(m,server.import_slides);self.assertEqual(upgraded['slideDisplayVersion'],5);self.assertEqual(upgraded['slides'][0]['items'][0]['text'],'사용자 편집');self.assertEqual(upgraded['slides'][0]['items'][0]['geometry']['preset'],'roundRect')
 def test_http_cache_marker_keeps_current_version_and_does_not_rewrite_legacy_copy(self):
  old=server.DATA
  with tempfile.TemporaryDirectory() as d:
   server.DATA=Path(d);http=ThreadingHTTPServer(('127.0.0.1',0),server.Handler);thread=threading.Thread(target=http.serve_forever,daemon=True);thread.start()
   try:
    for version in [4,5]:
     p=server.import_project('cache.pptx',bytes_of(fixture()),persist=False);p['content']['slideDisplayVersion']=version;file=server.project_path(p['id']);file.write_text(json.dumps(p));before=file.read_bytes()
     with urllib.request.urlopen(f'http://127.0.0.1:{http.server_port}/api/projects/'+p['id']) as response:content=json.load(response)['content']
     self.assertEqual(content.get('slideDisplayVersion'),5 if version==5 else None);self.assertEqual(file.read_bytes(),before)
   finally:http.shutdown();http.server_close();thread.join();server.DATA=old
if __name__=='__main__':unittest.main()
