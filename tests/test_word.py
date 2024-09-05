import unittest

from poword.api.word import *


class TestPoword(unittest.TestCase):
    def test_docx2pdf(self):
        docx2pdf(path=r'./docx/word2pdf/aaa - 副本.docx', output_path=r'./docx/res')

    def test_doc2docx(self):
        doc2docx(input_path=r'./docx/res/aaa - 副本.doc', output_path=r'./out/resx', output_name='abc.docx')

    def test_docx2doc(self):
        docx2doc(input_path=r'./docx', output_path=r'./docx/ress')

    def test_merge4docx(self):
        merge4docx(input_path=r'./docx/res/', output_path=r'./docx/res/out', new_word_name='合并的Word.docx')
