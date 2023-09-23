# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/B1V6KeXc7IOEB8DgXLWv3g
@个人网站 ：www.python-office.com
@Date    ：2023/4/2 1:39 
@Description     ：
'''
import os
import shutil
from pathlib import Path

from win32com import client
from win32com.client import gencache


def doc2docx(input_path, output_path):
    """
    doc转为docx（win）
    :param path_original:
    :param path_final:
    :return:
    """
    abs_output_path = str(Path(output_path).absolute())
    # 获取文件的格式后缀
    file_suffix = os.path.splitext(input_path)[1]
    if file_suffix == ".doc":
        word_app = gencache.EnsureDispatch('Word.Application')  # 打开word程序
        word_app.Visible = False  # 是否可视化
        # 源文件
        doc = word_app.Documents.Open(input_path, ReadOnly=1)
        # 生成的新文件
        doc.SaveAs(abs_output_path, 16)
        doc.Close()
        # word.Quit()
    elif file_suffix == ".docx":
        shutil.copy(abs_output_path, output_path)


"""
mac\linux
import os

source = "./doc/"
dest = "./docx/"
g = os.walk(source)

# 遍历文件夹
for root, dirs, files in g:
    for file in files:
        # 源文件完整路径
        file_input_path = os.path.join(root, file)
        print(file_input_path)

        os.system("soffice --headless --convert-to docx {} --outdir {}".format(file_input_path, dest))
"""
import popip
import poword

if __name__ == '__main__':
    # popip.pip_times('python-office')
    # popip.pip_times('poprogress')
    doc2docx(input_path=r'C:\Users\Lenovo\Desktop\temp\test\aa.doc',output_path=r'./fdadasf')
    # poword.docx2pdf(path=r'C:\Users\Lenovo\Desktop\temp\test\aa.doc')
