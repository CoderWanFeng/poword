import unittest

from poword.api.word import docx2pdf


class TestWechat(unittest.TestCase):
    def test_send_file(self):
        docx2pdf(path=r'D:\workplace\code\github\python-office\tests\test_files\pdf\add_img.docx')
