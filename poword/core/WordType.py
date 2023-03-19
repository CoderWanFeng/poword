import os
from win32com.client import constants, gencache
# import win32com.client as win32
from pathlib import Path

class MainWord():

    def docx2pdf(self, path, docxSuffix=".docx"):
        wordFiles = []
        # 如果不存在，则不做处理
        if not os.path.exists(path):
            print("path does not exist path = " + path)
            return
        # 判断是否是文件
        elif os.path.isfile(path):
            print("path file type is file " + path)
            wordFiles.append(path)
        # 如果是目录，则遍历目录下面的文件
        elif os.path.isdir(path):
            print(os.listdir(path))
            # 填充路径，补充完整路径
            if not path.endswith("/") or not path.endswith("\\"):
                path = path + "/"
            for file in os.listdir(path):
                if file.endswith(docxSuffix):
                    wordFiles.append(path + file)
        print(wordFiles)
        for file in wordFiles:
            filepath = os.path.abspath(file)
            index = filepath.rindex('.')
            pdfPath = os.path.abspath(filepath[:index] + '.pdf')
            print(pdfPath)
            self.createpdf(filepath, pdfPath)

    def createpdf(self, wordPath, pdfPath):
        word = gencache.EnsureDispatch('Word.Application')
        doc = word.Documents.Open(wordPath, ReadOnly=1)
        # 转换方法
        doc.ExportAsFixedFormat(pdfPath, constants.wdExportFormatPDF)
        word.Quit()

    def merge4docx(self, input_path, output_path, new_word_name):
        abs_input_path = Path(input_path).absolute()  # 相对路径→绝对路径
        abs_output_path = Path(output_path).absolute()  # 相对路径→绝对路径
        save_path = abs_output_path / new_word_name
        print('-' * 10 + '开始合并!' + '-' * 10)
        word_app = gencache.EnsureDispatch('Word.Application')   # 打开word程序
        word_app.Visible = False  # 是否可视化
        folder = Path(abs_input_path)
        waiting_files = [path for path in folder.iterdir()]
        output_file = word_app.Documents.Add()  # 新建合并后的文档
        for single_file in waiting_files:
            output_file.Application.Selection.InsertFile(single_file)  # 拼接文档
        output_file.SaveAs(str(save_path))  # 保存
        output_file.Close()
        print('-' * 10 + '合并完成!' + '-' * 10)
