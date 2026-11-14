"""Real HWP wrapping, edits, mode conversion and atomic refusal boundaries."""
from copy import deepcopy
from pathlib import Path
import base64,io,struct,tempfile,unittest
from PIL import Image
import hangul_backend as native
import native_page
from hangul_commands import apply_command,envelope
from hangul_object_edit import shaped
from hangul_images import pictures,summary
from hangul_paragraph_patch import without_cache
from hangul_image_wrap import wrap_caches


class NativePictureWrapping(unittest.TestCase):
    def setUp(self):
        self.tmp=tempfile.TemporaryDirectory();self.data=Path(self.tmp.name)
        self.seed=native.create_native('둘레 시험',False,self.data,markdown='제목은 유지합니다.\n\n'+('그림 옆의 본문이 이어지고 아래에서는 전체 너비로 돌아옵니다. '*10)+'\n\n마지막 문단도 보존합니다.')
        folder=self.data/'shape';folder.mkdir()
        file,_=shaped(envelope(self.seed,self.data)['document'],folder,native.native_path(self.seed['content'],self.data))
        self.seed=native.import_native('source.hwp',file.read_bytes(),self.data)
        out=io.BytesIO();Image.new('RGB',(120,80),'#327863').save(out,format='PNG')
        self.op={'type':'hangul_insert_image','after_key':'0:0','placement':'square','width_mm':40,'height_mm':25,'x_mm':10,'y_mm':0,'data_base64':base64.b64encode(out.getvalue()).decode()}

    def tearDown(self):self.tmp.cleanup()

    def insert(self,mode):
        return apply_command(self.seed,{**self.op,'placement':mode},self.data)['project']

    def test_both_wraps_keep_text_styles_streams_and_three_physical_reopens(self):
        original=native.native_path(self.seed['content'],self.data).read_bytes();before=envelope(self.seed,self.data)['document']
        for mode,flag in [('square',0),('topBottom',1)]:
            with self.subTest(mode=mode):
                result=self.insert(mode);after=envelope(result,self.data)['document']
                self.assertEqual(summary(after)[0]['placement'],mode)
                self.assertEqual((struct.unpack_from('<I',bytes.fromhex(pictures(after)[0]['picture']['common_data']))[0]>>21)&7,flag)
                remaining=deepcopy(after['sections'][0]['paragraphs']);remaining.pop(1)
                self.assertEqual(without_cache(remaining),without_cache(before['sections'][0]['paragraphs']))
                self.assertEqual(before['header']['fonts'],after['header']['fonts'])
                self.assertEqual(before['header']['char_shapes'],after['header']['char_shapes'])
                self.assertEqual(native_page.render(result,self.data)['totalPages'],1)
                for _ in range(3):
                    result=native.import_native('copy.hwp',native.export_native(result,'hwp',self.data)[0],self.data)
                    self.assertEqual(envelope(result,self.data)['document'],after)
        self.assertEqual(native.native_path(self.seed['content'],self.data).read_bytes(),original)

    def test_edit_resize_move_and_unwrap_preserve_body(self):
        result=self.insert('square')
        text='수정한 본문이 그림 옆에서 이어집니다. '*12
        result=native_page.edit(result,[{'key':'0:2','parts':[{'text':text}]}],self.data)
        for patch in [{'width_mm':45,'height_mm':28},{'x_mm':60,'y_mm':5},{'placement':'topBottom'},{'placement':'inline','x_mm':0,'y_mm':0},{'placement':'front','x_mm':10,'y_mm':0},{'placement':'square'}]:
            result=apply_command(result,{'type':'hangul_set_image_layout','image_index':0,**patch},self.data)['project']
            result=native.import_native('copy.hwp',native.export_native(result,'hwp',self.data)[0],self.data)
            self.assertIn(text,[native.paragraph_text(p) for p in envelope(result,self.data)['document']['sections'][0]['paragraphs']])
        self.assertEqual(summary(envelope(result,self.data)['document'])[0]['placement'],'square')

    def test_unsupported_page_geometry_and_hwpx_are_atomic(self):
        result=self.insert('square');old=deepcopy(result);raw=native.native_path(result['content'],self.data).read_bytes()
        for patch in [{'x_mm':190},{'y_mm':200,'width_mm':40,'height_mm':200}]:
            with self.assertRaises(ValueError):apply_command(result,{'type':'hangul_set_image_layout','image_index':0,**patch},self.data)
        other=native.import_native('source.hwpx',native.export_native(self.seed,'hwpx',self.data)[0],self.data)
        with self.assertRaisesRegex(ValueError,'HWPX'):apply_command(other,self.op,self.data)
        self.assertEqual(result,old);self.assertEqual(native.native_path(result['content'],self.data).read_bytes(),raw)

    def test_native_multiple_columns_are_refused_before_shaping(self):
        result=self.insert('square');before=envelope(result,self.data)['document'];desired=deepcopy(before)
        column=next(c['Generic']['column_def'] for c in desired['sections'][0]['paragraphs'][0]['controls'] if 'column_def' in c.get('Generic',{}))
        column['count']=2;pic=pictures(desired)[0]['picture']
        with self.assertRaisesRegex(ValueError,'여러 단'):wrap_caches(before,desired,0,[1],pictures(before)[0]['picture'],pic)


if __name__=='__main__':unittest.main()
