import importlib.util
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, patch


def load_word_type_module():
    fake_docx = types.ModuleType("docx")
    fake_docx.Document = Mock()
    fake_docx.ImagePart = type("ImagePart", (), {})

    fake_pofile = types.ModuleType("pofile")
    fake_pofile.get_files = Mock()
    fake_pofile.mkdir = Mock()

    fake_progress = types.ModuleType("poprogress")
    fake_progress.simple_progress = lambda items: items

    fake_client = types.ModuleType("win32com.client")
    fake_client.constants = types.SimpleNamespace(wdExportFormatPDF=17)
    fake_client.gencache = types.SimpleNamespace(EnsureDispatch=Mock())

    fake_win32com = types.ModuleType("win32com")
    fake_win32com.client = fake_client

    dependencies = {
        "docx": fake_docx,
        "pofile": fake_pofile,
        "poprogress": fake_progress,
        "win32com": fake_win32com,
        "win32com.client": fake_client,
    }
    module_path = Path(__file__).parents[1] / "poword" / "core" / "WordType.py"
    spec = importlib.util.spec_from_file_location("word_type_under_test", module_path)
    module = importlib.util.module_from_spec(spec)

    with patch.dict(sys.modules, dependencies):
        spec.loader.exec_module(module)

    return module


class TestCreatePdfResourceCleanup(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.word_type = load_word_type_module()

    def setUp(self):
        self.document = Mock()
        self.word_app = Mock()
        self.word_app.Documents.Open.return_value = self.document
        self.word_type.gencache.EnsureDispatch = Mock(return_value=self.word_app)
        self.main_word = self.word_type.MainWord()

    def test_closes_document_after_successful_export(self):
        self.main_word.createpdf("source.docx", "output.pdf")

        self.document.ExportAsFixedFormat.assert_called_once_with("output.pdf", 17)
        self.document.Close.assert_called_once_with(False)

    def test_closes_document_when_export_fails(self):
        self.document.ExportAsFixedFormat.side_effect = RuntimeError("export failed")

        with self.assertRaisesRegex(RuntimeError, "export failed"):
            self.main_word.createpdf("source.docx", "output.pdf")

        self.document.Close.assert_called_once_with(False)


if __name__ == "__main__":
    unittest.main()
