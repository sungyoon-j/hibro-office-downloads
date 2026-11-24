import base64
from copy import deepcopy
import io
from pathlib import Path
import tempfile
import unittest
import zipfile
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Pt
from pptx.oxml.xmlchemy import OxmlElement
import server
from slide_text import plain, sync_text


def fixture():
    deck = Presentation(); deck.slide_width = Pt(960); deck.slide_height = Pt(540)
    deck.core_properties.title = '원본 내부 제목 유지'
    slide = deck.slides.add_slide(deck.slide_layouts[6])
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(70), Pt(90), Pt(700), Pt(170))
    shape.fill.solid(); shape.fill.fore_color.rgb = RGBColor.from_string('D7EE82')
    shape.line.color.rgb = RGBColor.from_string('173E38'); shape.line.width = Pt(2)
    shape.text_frame.margin_left = Pt(12); shape.text_frame.margin_top = Pt(8)
    p = shape.text_frame.paragraphs[0]; p.line_spacing = 1.5; p.space_after = Pt(7)
    for text, font, size, bold, color in [('브라더 ', 'Arial', 22, False, '173E38'), ('중요한 제목', 'NanumMyeongjo', 32, True, 'B42318'), (' 안내', 'Arial', 18, False, '173E38')]:
        r = p.add_run(); r.text = text; r.font.name = font; r.font.size = Pt(size); r.font.bold = bold; r.font.color.rgb = RGBColor.from_string(color)
    p.runs[1].hyperlink.address = 'https://example.com/reference'
    r = OxmlElement('a:extLst'); ext = OxmlElement('a:ext'); ext.set('uri', 'rich-test-retained'); r.append(ext); p.runs[1]._r.append(r)
    p = shape.text_frame.add_paragraph(); p.alignment = 3; p.level = 1
    r = p.add_run(); r.text = '둘째 문단'; r.font.name = 'Arial'; r.font.size = Pt(16); r.font.italic = True; r.font.underline = True
    p.add_line_break(); r = p.add_run(); r.text = '같은 문단의 다음 줄'; r.font.size = Pt(14)
    box = slide.shapes.add_textbox(Pt(70), Pt(330), Pt(700), Pt(60)); box.text = '수정하지 않는 글상자'
    slide.notes_slide.notes_text_frame.text = '발표자 메모 유지'
    slide = deck.slides.add_slide(deck.slide_layouts[6]); slide.shapes.add_textbox(Pt(80), Pt(100), Pt(500), Pt(80)).text = '다른 슬라이드 유지'
    data = io.BytesIO(); deck.save(data); return data.getvalue()


def project(raw): return server.import_project('혼합 서식.pptx', raw, persist=False)
def exported(model): return server.export_project(model, 'pptx')[0]


class RichSlideTest(unittest.TestCase):
    def test_original_shape_runs_and_every_package_part_survive_unedited_save(self):
        raw = fixture(); model = project(raw); item = model['content']['slides'][0]['items'][0]
        self.assertTrue(item['textShape']); self.assertEqual(item['fill'], '#D7EE82')
        self.assertEqual([r['fontSize'] for r in item['richText'][0]['runs']], [22, 32, 18])
        output = exported(model)
        with zipfile.ZipFile(io.BytesIO(raw)) as a, zipfile.ZipFile(io.BytesIO(output)) as b:
            self.assertEqual(a.namelist(), b.namelist())
            self.assertTrue(all(a.read(n) == b.read(n) for n in a.namelist()))

    def test_changed_length_text_keeps_other_runs_shape_and_unrelated_package_bytes_three_cycles(self):
        raw = fixture(); model = project(raw); item = model['content']['slides'][0]['items'][0]
        item['text'] = item['text'].replace('중요한 제목', '수정한 긴 제목입니다')
        for cycle in range(3):
            output = exported(model); actual = Presentation(io.BytesIO(output)); shape = actual.slides[0].shapes[0]
            self.assertEqual(shape.text, item['text']); self.assertEqual(shape.auto_shape_type, MSO_SHAPE.RECTANGLE)
            self.assertEqual(shape.fill.fore_color.rgb, RGBColor.from_string('D7EE82')); self.assertEqual(shape.line.width.pt, 2)
            self.assertEqual(shape.text_frame.margin_left.pt, 12); self.assertEqual(shape.text_frame.paragraphs[0].space_after.pt, 7)
            self.assertEqual(shape.text_frame.paragraphs[1].alignment, 3); self.assertEqual(shape.text_frame.paragraphs[1].level, 1)
            self.assertTrue(shape.text_frame.paragraphs[1].runs[0].font.italic); self.assertTrue(shape.text_frame.paragraphs[1].runs[0].font.underline)
            for run in shape.text_frame.paragraphs[0].runs:
                if run.text.strip() and run.text not in {'브라더 ', ' 안내'}:
                    self.assertTrue(run.font.bold); self.assertEqual(run.font.size.pt, 32)
                    self.assertEqual(run.hyperlink.address, 'https://example.com/reference')
            with zipfile.ZipFile(io.BytesIO(raw)) as a, zipfile.ZipFile(io.BytesIO(output)) as b:
                changed = [n for n in a.namelist() if a.read(n) != b.read(n)]
                self.assertEqual(changed, ['ppt/slides/slide1.xml'])
                source = etree.fromstring(a.read(changed[0])); target = etree.fromstring(b.read(changed[0])); ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
                self.assertEqual(etree.tostring(source.findall('.//p:sp', ns)[1]), etree.tostring(target.findall('.//p:sp', ns)[1]))
            model = project(output); item = model['content']['slides'][0]['items'][0]

    def test_run_only_format_change_does_not_flatten_surrounding_styles(self):
        model = project(fixture()); item = model['content']['slides'][0]['items'][0]
        item['richText'][0]['runs'][1]['color'] = '#0055CC'
        item['richText'][0]['runs'][1]['italic'] = True
        shape = Presentation(io.BytesIO(exported(model))).slides[0].shapes[0]
        a, b, c = shape.text_frame.paragraphs[0].runs
        self.assertEqual(a.font.size.pt, 22); self.assertEqual(c.font.size.pt, 18)
        self.assertEqual(b.font.color.rgb, RGBColor.from_string('0055CC')); self.assertTrue(b.font.italic); self.assertTrue(b.font.bold)
        self.assertEqual(b.hyperlink.address, 'https://example.com/reference')

    def test_object_move_and_duplicate_slide_keep_independent_text_and_notes(self):
        model = project(fixture()); original = model['content']['slides'][0]
        duplicate = deepcopy(original); duplicate['id'] = 'copied-slide'; duplicate['items'][0]['text'] = '복제 문구\n다음 문단'; duplicate['items'][0]['x'] += 25
        duplicate['notes'] = '사본 메모'; model['content']['slides'].insert(0, duplicate)
        result = Presentation(io.BytesIO(exported(model)))
        self.assertEqual(result.slides[0].shapes[0].text, duplicate['items'][0]['text']); self.assertEqual(result.slides[0].shapes[0].left.pt, 95)
        self.assertEqual(result.slides[1].shapes[0].text, original['items'][0]['text'])
        self.assertEqual(result.slides[0].notes_slide.notes_text_frame.text, '사본 메모'); self.assertEqual(result.slides[1].notes_slide.notes_text_frame.text, '발표자 메모 유지')

    def test_added_text_and_removed_object_dont_drop_source_slide_parts(self):
        model = project(fixture()); slide = model['content']['slides'][0]; slide['items'].pop(1)
        slide['items'].append({'type': 'text', 'x': 10, 'y': 10, 'w': 200, 'h': 40, 'text': '새 글상자', 'fontSize': 20})
        result = project(exported(model)); self.assertEqual(len(result['content']['slides'][0]['items']), 2)
        self.assertEqual(result['content']['slides'][0]['items'][1]['text'], '새 글상자')

    def test_new_paragraph_and_empty_text_are_exportable_without_losing_shape(self):
        for text in ['', '바꾼 첫 문단\n새 문단\n']:
            model = project(fixture()); model['content']['slides'][0]['items'][0]['text'] = text
            shape = Presentation(io.BytesIO(exported(model))).slides[0].shapes[0]
            self.assertEqual(shape.text, text); self.assertEqual(shape.fill.fore_color.rgb, RGBColor.from_string('D7EE82'))
            self.assertEqual(project(exported(model))['content']['slides'][0]['items'][0]['type'], 'text')

    def test_inherited_run_style_and_theme_paint_do_not_copy_first_runs_emphasis(self):
        deck = Presentation(); slide = deck.slides.add_slide(deck.slide_layouts[6])
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Pt(10), Pt(10), Pt(400), Pt(100))
        p = shape.text_frame.paragraphs[0]
        first = p.add_run(); first.text = '강조'; first.font.bold = True; first.font.size = Pt(32)
        second = p.add_run(); second.text = ' 기본 글자'
        data = io.BytesIO(); deck.save(data); raw = data.getvalue()
        model = project(raw); item = model['content']['slides'][0]['items'][0]
        self.assertTrue(item['richText'][0]['runs'][0]['bold'])
        self.assertFalse(item['richText'][0]['runs'][1]['bold'])
        self.assertNotEqual(item['richText'][0]['runs'][1]['fontSize'], 32)
        self.assertNotEqual(item['fill'].lower(), '#ffffff')
        with zipfile.ZipFile(io.BytesIO(raw)) as a, zipfile.ZipFile(io.BytesIO(exported(model))) as b:
            self.assertTrue(all(a.read(n) == b.read(n) for n in a.namelist()))


if __name__ == '__main__': unittest.main()
