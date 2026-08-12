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
    fake_progress.simple_progress = Mock(side_effect=lambda items: items)

    fake_client = types.ModuleType("win32com.client")
    fake_client.constants = types.SimpleNamespace()
    fake_client.gencache = types.SimpleNamespace()

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


class TestDoc2DocxProgress(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.word_type = load_word_type_module()

    def setUp(self):
        self.word_type.get_files = Mock(return_value=[])
        self.word_type.mkdir = Mock(return_value=(False, Path("output")))
        self.word_type.simple_progress = Mock(side_effect=lambda items: items)
        self.main_word = self.word_type.MainWord()

    def test_shows_progress_by_default(self):
        self.main_word.doc2docx("source.doc", "output")

        self.word_type.simple_progress.assert_called_once_with([])

    def test_can_hide_progress(self):
        self.main_word.doc2docx("source.doc", "output", show_progress=False)

        self.word_type.simple_progress.assert_not_called()


if __name__ == "__main__":
    unittest.main()
