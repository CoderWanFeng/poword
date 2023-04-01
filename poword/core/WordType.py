import os

from pofile import get_files
from poprogress import simple_progress
from win32com.client import constants, gencache
# import win32com.client as win32
from pathlib import Path


class MainWord():

    def docx2pdf(self, path, output_path, docxSuffix=".docx", pdfSuffix='.pdf'):
        waiting_covert_docx_files = get_files(path, name=docxSuffix)
        if waiting_covert_docx_files:
            print(f'一共有{len(waiting_covert_docx_files)}个docx文件')
            for i, docx_file in simple_progress(enumerate(waiting_covert_docx_files)):
                abs_output_path = Path(output_path).absolute()
                if not abs_output_path.exists():
                    abs_output_path.mkdir()
                abs_single_docx_path = Path(docx_file).absolute()
                print(f'正在转换的是第 {str(i + 1)} 个，文档名字是： {abs_single_docx_path}')
                abs_pdf_path = abs_output_path / (abs_single_docx_path.stem + pdfSuffix)
                self.createpdf(str(abs_single_docx_path), str(abs_pdf_path))

    def createpdf(self, wordPath, pdfPath):
        word_app = gencache.EnsureDispatch('Word.Application')
        word_app.Visible = False  # 是否可视化
        doc = word_app.Documents.Open(wordPath, ReadOnly=1)
        # 转换方法
        doc.ExportAsFixedFormat(pdfPath, constants.wdExportFormatPDF)
        # word_app.Quit() #不注释，不能批量转换，必须注释

    def merge4docx(self, input_path, output_path, new_word_name):
        abs_input_path = Path(input_path).absolute()  # 相对路径→绝对路径
        abs_output_path = Path(output_path).absolute()  # 相对路径→绝对路径
        save_path = abs_output_path / new_word_name
        print('-' * 10 + '开始合并!' + '-' * 10)
        word_app = gencache.EnsureDispatch('Word.Application')  # 打开word程序
        word_app.Visible = False  # 是否可视化
        folder = Path(abs_input_path)
        waiting_files = [path for path in folder.iterdir()]
        output_file = word_app.Documents.Add()  # 新建合并后的文档
        for single_file in waiting_files:
            output_file.Application.Selection.InsertFile(single_file)  # 拼接文档
        output_file.SaveAs(str(save_path))  # 保存
        output_file.Close()
        print('-' * 10 + '合并完成!' + '-' * 10)
