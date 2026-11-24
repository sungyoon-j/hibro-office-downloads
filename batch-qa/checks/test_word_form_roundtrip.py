"""Synthetic first-page forms only; no user file is uploaded to hosted QA."""
from copy import deepcopy
import hashlib,html,io,json,unittest,zipfile
from lxml import etree
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from PIL import Image
import server

def mark(p,size=18,bold=False):
    r=OxmlElement('w:rPr');fonts=OxmlElement('w:rFonts')
    for key in ['asciiTheme','eastAsiaTheme','hAnsiTheme']:fonts.set(qn('w:'+key),'minorEastAsia')
    r.append(fonts)
    if bold:r.append(OxmlElement('w:b'))
    n=OxmlElement('w:sz');n.set(qn('w:val'),str(size));r.append(n);p._p.get_or_add_pPr().append(r)
def pack(doc):
    out=io.BytesIO();doc.save(out);return out.getvalue()
def seed():
    d=Document();s=d.sections[0];s.different_first_page_header_footer=True
    s.footer.paragraphs[0].text='Default footer';s.first_page_footer.paragraphs[0].text='First footer'
    image=io.BytesIO();Image.new('RGB',(8,8),'green').save(image,'PNG');s.header.paragraphs[0].add_run().add_picture(io.BytesIO(image.getvalue()),width=Pt(9))
    p=d.add_paragraph();mark(p,36,True);r=p.add_run('Form title');r.bold=True;r.font.size=Pt(18)
    d.add_paragraph();mark(d.paragraphs[-1],18)
    p=d.add_paragraph('Body text');mark(p,18)
    table=d.add_table(rows=1,cols=2);table.cell(0,0).text='Label';mark(table.cell(0,1).paragraphs[0],18)
    return pack(d)
def native_parts(raw):
    with zipfile.ZipFile(io.BytesIO(raw)) as z:
        parts={}
        for name in z.namelist():
            if not name.startswith(('word/header','word/footer','word/media/','word/_rels/header','word/_rels/footer')):continue
            value=z.read(name)
            if name.endswith(('.xml','.rels')):value=etree.tostring(etree.fromstring(value),method='c14n')
            parts[name]=hashlib.sha256(value).hexdigest()
        root=etree.fromstring(z.read('word/document.xml'));ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        fonts=[]
        for p in root.xpath('//w:body//w:p',namespaces=ns):
            mark=p.xpath('./w:pPr/w:rPr',namespaces=ns);fonts.append(etree.tostring(mark[0],method='c14n',exclusive=True) if mark else None)
        return parts,fonts
def tree(content):return server.HtmlTree(content['html']).root
def find(node,tag):
    if not isinstance(node,str):
        if node['tag']==tag:return node
        for child in node['children']:
            found=find(child,tag)
            if found is not None:return found
def findall(node,tags):
    if isinstance(node,str):return []
    return ([node] if node['tag'] in tags else [])+[found for child in node['children'] for found in findall(child,tags)]
def serial(node):
    if isinstance(node,str):return html.escape(node)
    inner=''.join(serial(c) for c in node['children'])
    if node['tag']=='root':return inner
    attrs=''.join(' '+k+'="'+html.escape(v,quote=True)+'"' for k,v in node['attrs'].items())
    return '<'+node['tag']+attrs+'>'+inner+('' if node['tag'] in {'br','img','wbr','col'} else '</'+node['tag']+'>')
def export(p):return server.export_project(p,'docx')[0]

class NativeFormRoundtrip(unittest.TestCase):
    def test_blank_cell_edit_keeps_first_footer_logo_marks_and_other_body_over_three_cycles(self):
        raw=seed();expected=native_parts(raw)
        for _ in range(3):
            p=server.import_project('synthetic-form.docx',raw,persist=False);self.assertEqual(p['content']['wordPropertiesVersion'],5)
            t=tree(p['content']);table=find(t,'table');cells=findall(table,{'td','th'});self.assertEqual(len(cells),2)
            find(cells[1],'p')['children']=['Saved value'];p['content']['html']=serial(t);raw=export(p);d=Document(io.BytesIO(raw))
            self.assertEqual(native_parts(raw),expected);self.assertTrue(d.sections[0].different_first_page_header_footer)
            self.assertEqual(d.tables[0].cell(0,1).text,'Saved value');self.assertEqual(d.tables[0].cell(0,0).text,'Label');self.assertEqual(d.paragraphs[2].text,'Body text')
    def test_complete_replacement_uses_native_theme_font_and_keeps_bold_size(self):
        p=server.import_project('synthetic-form.docx',seed(),persist=False);t=tree(p['content']);paragraph=find(t,'p')
        paragraph['children']=[{'tag':'strong','attrs':{},'children':['Replacement']}];p['content']['html']=serial(t)
        raw=export(p);r=Document(io.BytesIO(raw)).paragraphs[0].runs[0]
        self.assertEqual(r._r.rPr.rFonts.get(qn('w:eastAsiaTheme')),'minorEastAsia');self.assertTrue(r.bold);self.assertEqual(r.font.size.pt,18)
        for _ in range(3):p=server.import_project('roundtrip.docx',raw,persist=False);raw=export(p);self.assertEqual(Document(io.BytesIO(raw)).paragraphs[0].text,'Replacement')
    def test_explicit_new_font_overrides_mark_without_changing_the_mark(self):
        initial=seed();p=server.import_project('synthetic-form.docx',initial,persist=False);t=tree(p['content']);find(t,'p')['children']=[{'tag':'span','attrs':{'style':'font-family:Arial'},'children':['Replacement']}];p['content']['html']=serial(t);raw=export(p)
        self.assertEqual(Document(io.BytesIO(raw)).paragraphs[0].runs[0].font.name,'Arial');self.assertEqual(native_parts(raw),native_parts(initial))
    def test_untrusted_mark_is_rejected_atomically(self):
        p=server.import_project('synthetic-form.docx',seed(),persist=False);t=tree(p['content']);find(t,'p')['attrs']['data-word-mark-format']=json.dumps([{'tag':'object','attrs':{},'children':[]}]);p['content']['html']=serial(t);before=deepcopy(p)
        with self.assertRaises(ValueError):export(p)
        self.assertEqual(p,before)
