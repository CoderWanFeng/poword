import hashlib
import os
import sys
from os.path import basename, join
from pathlib import Path

from docx import ImagePart, Document
from pofile import get_files, mkdir
from poprogress import simple_progress
from win32com import client
from win32com.client import constants, gencache


class MainWord():

    def __init__(self):
        self.app = 'Word.Application'

    def docx2pdf(self, path, output_path, docxSuffix=".docx", pdfSuffix='.pdf'):
        """
        将指定路径下的所有docx文件转换为pdf文件。

        参数:
        - path: str, docx文件所在的目录路径。
        - output_path: str, 转换后的pdf文件保存的输出目录路径。
        - docxSuffix: str, docx文件的后缀，默认为".docx"。
        - pdfSuffix: str, 生成的pdf文件的后缀，默认为".pdf"。

        返回:
        无
        """
        # 获取需要转换的所有docx文件列表
        waiting_covert_docx_files = get_files(path, suffix=docxSuffix)

        # 检查是否有需要转换的docx文件
        if waiting_covert_docx_files:
            print(f'一共有{len(waiting_covert_docx_files)}个docx文件')

            word_app = gencache.EnsureDispatch(self.app)
            word_app.Visible = False
            # 遍历每个docx文件并转换为pdf
            for i, docx_file in simple_progress(enumerate(waiting_covert_docx_files)):
                # 获取输出目录的绝对路径
                abs_output_path = Path(output_path).absolute()

                # 如果输出目录不存在，则创建
                mkdir(abs_output_path)
                if not abs_output_path.exists():
                    abs_output_path.mkdir()

                # 获取当前docx文件的绝对路径
                abs_single_docx_path = Path(docx_file).absolute()
                print(f'正在转换的是第 {str(i + 1)} 个，文档名字是： {abs_single_docx_path}')

                # 构建对应的pdf文件路径
                abs_pdf_path = abs_output_path / (abs_single_docx_path.stem + pdfSuffix)

                # 调用方法将docx文件转换为pdf文件
                self._createpdf(word_app, str(abs_single_docx_path), str(abs_pdf_path))

    def createpdf(self, wordPath, pdfPath):
        """
        将Word文档转换为PDF格式文件。

        参数:
        wordPath (str): Word文档的路径。
        pdfPath (str): 生成的PDF文件的路径。

        说明:
        本函数通过调用Word应用程序接口将Word文档转换为PDF文件，适用于需要将Word格式文档转换为不可编辑的PDF格式的场景。
        """
        word_app = gencache.EnsureDispatch(self.app)
        word_app.Visible = False
        self._createpdf(word_app, wordPath, pdfPath)

    def _createpdf(self, word_app, word_path, pdf_path):
        doc = word_app.Documents.Open(word_path, ReadOnly=1)
        try:
            doc.ExportAsFixedFormat(pdf_path, constants.wdExportFormatPDF)
        finally:
            self._close_document(doc)

    @staticmethod
    def _close_document(doc):
        has_original_error = sys.exc_info()[0] is not None
        try:
            doc.Close(False)
        except Exception:
            if not has_original_error:
                raise

    def merge4docx(self, input_path, output_path, new_word_name):
        """
        合并同一目录下的所有docx文件到一个新文档中。

        参数:
        input_path: str，输入文件夹的路径。
        output_path: str，输出文件夹的路径。
        new_word_name: str，合并后新文档的文件名。

        该函数通过将同一目录下的所有docx文件合并到一个新文档中，实现了批量文档合并的功能。
        """
        # 将输入和输出路径从相对路径转换为绝对路径
        abs_input_path = Path(input_path).absolute()
        abs_output_path = Path(output_path).absolute()

        # 创建输出目录
        mkdir(abs_output_path)

        # 拼接保存路径和文件名，得到合并后的文档完整路径
        save_path = abs_output_path / new_word_name

        # 打印合并开始的标志信息
        print('-' * 10 + '开始合并!' + '-' * 10)

        # 获取输入目录下的所有docx文件路径
        folder = Path(abs_input_path)
        save_path_key = os.path.normcase(str(save_path.resolve()))
        waiting_files = sorted(
            (
                path for path in folder.iterdir()
                if path.is_file()
                and path.suffix.lower() == '.docx'
                and not path.name.startswith('~$')
                and os.path.normcase(str(path.resolve())) != save_path_key
            ),
            key=lambda path: os.path.normcase(str(path.resolve())),
        )

        if not waiting_files:
            print('-' * 10 + '合并完成!' + '-' * 10)
            return

        word_app = client.Dispatch(self.app)
        word_app.Visible = False
        # 新建一个Word文档，用于保存合并后的文件
        output_file = word_app.Documents.Add()
        try:
            # 遍历所有文件，逐个将文件内容插入到新建的Word文档中
            for single_file in waiting_files:
                output_file.Application.Selection.InsertFile(str(single_file))

            # 保存合并后的文档
            output_file.SaveAs(str(save_path))
        finally:
            self._close_document(output_file)

        # 打印合并完成的标志信息
        print('-' * 10 + '合并完成!' + '-' * 10)

    def doc2docx(self, input_path, output_path, output_name=None, docSuffix='.doc', type_id=16,
                 show_progress=True):
        """
        将doc文件转换为docx文件
        :param input_path: 输入文件的路径
        :param output_path: 输出文件的路径
        :param output_name: (可选) 输出文件的名称
        :param docSuffix: (可选) 输入文件的后缀，默认为'.doc'
        :param type_id: (可选) 文件类型ID，用于识别文件格式，默认为16，代表存为.docx文件
        :param show_progress: (可选) 是否显示转换进度条，默认为True
        :return: 无返回值
        """
        # 调用convert4word方法进行文件格式转换
        self._convert4word(type_id, input_path, output_path, docSuffix, output_name, show_progress)

    def docx2doc(self, input_path, output_path='./', output_name=None, docSuffix='.docx', type_id=0):
        """
        将docx格式文件转换为doc格式文件。

        :param input_path: 输入文件的路径
        :param output_path: 输出文件的保存路径，默认为当前路径
        :param output_name: 输出文件的名称，默认为None
        :param docSuffix: 文档的后缀名，默认为'.docx'
        :param type_id: 转换类型标识，默认为0，代表存为.doc文件
        :return: 无返回值
        """
        # 调用convert4word方法实现docx到doc的格式转换
        self._convert4word(type_id, input_path, output_path, docSuffix, output_name)

    def _convert4word(self, type_id, input_path, output_path, docSuffix, output_name,
                      show_progress=True):
        """

        :param type_id: 16-docx,0-doc
        :param show_progress: 是否显示转换进度条
        :return:
        """
        abs_input_path = Path(input_path).absolute()
        exsit, abs_output_path = mkdir(output_path)
        word_file_list = list(get_files(abs_input_path, suffix=docSuffix))
        out_suffix = '.doc' if type_id == 0 else '.docx'
        files_to_convert = simple_progress(word_file_list) if show_progress else word_file_list
        if not word_file_list:
            return

        word_app = gencache.EnsureDispatch(self.app)
        word_app.Visible = False
        for word_file in files_to_convert:
            doc = word_app.Documents.Open(str(word_file), ReadOnly=1)
            try:
                output_file_name = (
                    Path(output_name).stem
                    if output_name and len(word_file_list) == 1
                    else Path(word_file).stem
                )
                output_word_name = os.path.join(abs_output_path, output_file_name) + out_suffix
                doc.SaveAs(output_word_name, type_id)
            finally:
                self._close_document(doc)

    def docx4imgs(self, word_path, img_path):
        """
        从docx文件中提取并保存图片。

        该方法旨在解析指定的docx文件，提取其中包含的所有图片，并将它们保存到指定的目录中。
        这对于需要从大量文档中自动抽取图片资源以进行进一步处理或归档的场景特别有用。

        :author Wang Peng

        :param word_path: str，要处理的docx文件的路径。此参数确定了输入文档的位置。
        :param img_path: str，保存提取图片的目标目录路径。此参数指定了图片的保存位置。
        :return: 该方法没有返回值，但会在指定的目录下生成提取的图片文件。
        """
        doc_obj = Document(word_path)
        images = []
        for rel in doc_obj.part.rels.values():
            if "image" in rel.reltype:
                img_part = rel.target_part
                if isinstance(img_part, ImagePart):
                    images.append((basename(img_part.partname), img_part.blob))

        if not images:
            return

        source_path = Path(word_path).resolve()
        output_root = Path(img_path)
        preferred_dir = output_root / source_path.stem
        source_key = os.path.normcase(str(source_path)).encode('utf-8')
        conflict_dir = output_root / f'{source_path.stem}-{hashlib.sha256(source_key).hexdigest()[:12]}'

        if conflict_dir.exists():
            output_dir = conflict_dir
        elif preferred_dir.exists() and any(
            (preferred_dir / name).exists() and (preferred_dir / name).read_bytes() != blob
            for name, blob in images
        ):
            output_dir = conflict_dir
        else:
            output_dir = preferred_dir

        mkdir(output_dir)
        for name, blob in images:
            save_path = join(output_dir, name)
            with open(save_path, "wb") as img_file:
                img_file.write(blob)
