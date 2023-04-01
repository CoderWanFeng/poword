import unittest

from poword.api.word import docx2pdf


class TestPoword(unittest.TestCase):
    def test_docx2pdf(self):
        docx2pdf(path=r'./docx',output_path=r'./docx/res')
