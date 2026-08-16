import base64
import importlib
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import Mock, call, patch

from docx import Document


fake_client = types.ModuleType("win32com.client")
fake_client.constants = types.SimpleNamespace(wdExportFormatPDF=17)
fake_client.gencache = types.SimpleNamespace(EnsureDispatch=Mock())
fake_client.Dispatch = Mock()
fake_win32com = types.ModuleType("win32com")
fake_win32com.client = fake_client
sys.modules["win32com"] = fake_win32com
sys.modules["win32com.client"] = fake_client

import poword


word_type = importlib.import_module("poword.core.WordType")


def make_dir(path):
    path = Path(path)
    existed = path.exists()
    path.mkdir(parents=True, exist_ok=True)
    return existed, path


class TestPowordPublicApi(unittest.TestCase):
    def test_doc2docx_uses_doc_inputs_and_distinct_batch_names(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "output"
            input_files = [Path(temp_dir) / "a.doc", Path(temp_dir) / "b.doc"]
            documents = [Mock(), Mock()]
            word_app = Mock()
            word_app.Documents.Open.side_effect = documents

            with (
                patch.object(word_type, "get_files", return_value=input_files) as get_files,
                patch.object(word_type, "mkdir", side_effect=make_dir),
                patch.object(word_type, "simple_progress") as progress,
                patch.object(word_type.gencache, "EnsureDispatch", return_value=word_app),
            ):
                poword.doc2docx(temp_dir, output_path, "same.docx", show_progress=False)

            get_files.assert_called_once_with(Path(temp_dir).absolute(), suffix=".doc")
            progress.assert_not_called()
            documents[0].SaveAs.assert_called_once_with(os.path.join(output_path, "a") + ".docx", 16)
            documents[1].SaveAs.assert_called_once_with(os.path.join(output_path, "b") + ".docx", 16)
            for document in documents:
                document.Close.assert_called_once_with(False)
            word_app.Quit.assert_not_called()

    def test_docx2doc_keeps_single_output_name(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output_path = Path(temp_dir) / "output"
            input_file = Path(temp_dir) / "source.docx"
            document = Mock()
            word_app = Mock()
            word_app.Documents.Open.return_value = document

            with (
                patch.object(word_type, "get_files", return_value=[input_file]) as get_files,
                patch.object(word_type, "mkdir", side_effect=make_dir),
                patch.object(word_type, "simple_progress", side_effect=lambda items: items),
                patch.object(word_type.gencache, "EnsureDispatch", return_value=word_app),
            ):
                poword.docx2doc(temp_dir, output_path, "renamed.doc")

            get_files.assert_called_once_with(Path(temp_dir).absolute(), suffix=".docx")
            document.SaveAs.assert_called_once_with(os.path.join(output_path, "renamed") + ".doc", 0)
            document.Close.assert_called_once_with(False)
            word_app.Quit.assert_not_called()

    def test_conversion_failure_preserves_original_error(self):
        document = Mock()
        document.SaveAs.side_effect = RuntimeError("save failed")
        document.Close.side_effect = RuntimeError("close failed")
        word_app = Mock()
        word_app.Documents.Open.return_value = document

        with (
            patch.object(word_type, "get_files", return_value=[Path("source.doc")]),
            patch.object(word_type, "mkdir", return_value=(False, Path("output"))),
            patch.object(word_type, "simple_progress", side_effect=lambda items: items),
            patch.object(word_type.gencache, "EnsureDispatch", return_value=word_app),
        ):
            with self.assertRaisesRegex(RuntimeError, "save failed"):
                poword.doc2docx("input", "output", "result.docx")

        document.Close.assert_called_once_with(False)
        word_app.Quit.assert_not_called()

    def test_merge_filters_sorts_and_cleans_up(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = Path(temp_dir)
            first = input_path / "a.DOCX"
            second = input_path / "b.docx"
            target = input_path / "merged.docx"
            for path in (first, second, target, input_path / "notes.txt", input_path / "~$lock.docx"):
                path.touch()
            (input_path / "subdir").mkdir()

            output_document = Mock()
            word_app = Mock()
            word_app.Documents.Add.return_value = output_document
            failed_document = Mock()
            failed_document.Application.Selection.InsertFile.side_effect = RuntimeError("insert failed")
            failed_app = Mock()
            failed_app.Documents.Add.return_value = failed_document

            empty_input = input_path / "empty"
            empty_input.mkdir()
            with (
                patch.object(word_type, "mkdir", side_effect=make_dir),
                patch.object(word_type.client, "Dispatch", side_effect=[word_app, failed_app]) as dispatch,
            ):
                poword.merge4docx(input_path, input_path, target.name)
                with self.assertRaisesRegex(RuntimeError, "insert failed"):
                    poword.merge4docx(input_path, input_path, target.name)
                poword.merge4docx(empty_input, input_path / "empty-output", target.name)

            self.assertEqual(
                output_document.Application.Selection.InsertFile.call_args_list,
                [call(str(first)), call(str(second))],
            )
            output_document.SaveAs.assert_called_once_with(str(target))
            output_document.Close.assert_called_once_with(False)
            word_app.Quit.assert_not_called()
            failed_document.Close.assert_called_once_with(False)
            failed_app.Quit.assert_not_called()
            self.assertEqual(dispatch.call_count, 2)

    def test_docx2pdf_reuses_application_and_cleans_up_on_failure(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            input_files = [Path(temp_dir) / "a.docx", Path(temp_dir) / "b.docx"]
            documents = [Mock(), Mock()]
            documents[1].ExportAsFixedFormat.side_effect = RuntimeError("export failed")
            word_app = Mock()
            word_app.Documents.Open.side_effect = documents

            with (
                patch.object(word_type, "get_files", return_value=input_files),
                patch.object(word_type, "mkdir", side_effect=make_dir),
                patch.object(word_type, "simple_progress", side_effect=lambda items: items),
                patch.object(word_type.gencache, "EnsureDispatch", return_value=word_app) as dispatch,
            ):
                with self.assertRaisesRegex(RuntimeError, "export failed"):
                    poword.docx2pdf(temp_dir, Path(temp_dir) / "pdf")

            dispatch.assert_called_once_with("Word.Application")
            for document in documents:
                document.Close.assert_called_once_with(False)
            word_app.Quit.assert_not_called()

    def test_docx4imgs_preserves_images_from_same_stem_documents(self):
        red = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1sAAAAASUVORK5CYII="
        )
        blue = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output_path = root / "images"
            sources = []
            for folder, blob in (("a", red), ("b", blue)):
                source_dir = root / folder
                source_dir.mkdir()
                image_path = source_dir / "image.png"
                image_path.write_bytes(blob)
                document_path = source_dir / "report.docx"
                document = Document()
                document.add_picture(str(image_path))
                document.save(document_path)
                sources.append(document_path)

            poword.docx4imgs(sources[0], output_path)
            poword.docx4imgs(sources[1], output_path)
            poword.docx4imgs(sources[1], output_path)

            extracted = list(output_path.rglob("image1.png"))
            self.assertEqual(len(extracted), 2)
            self.assertEqual({path.read_bytes() for path in extracted}, {red, blue})


if __name__ == "__main__":
    unittest.main()
