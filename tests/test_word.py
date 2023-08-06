import unittest

from poword.api.word import *


class TestPoword(unittest.TestCase):
    def test_docx2pdf(self):
        docx2pdf(path=r'./docx', output_path=r'./docx/res')

    def test_doc2docx(self):
        doc2docx(input_path=r'./docx/res/aaa - 副本.doc', output_path=r'./docx/resx')

    def test_docx2doc(self):
        docx2doc(input_path=r'./docx', output_path=r'./docx/ress')
