# -*- coding: UTF-8 -*-
'''
@Author  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@WeChat     ：CoderWanFeng
@Blog      ：www.python-office.com
@Date    ：2023/4/2 1:39 
@Description     ：
'''
import os
import shutil

from win32com import client

def doc_to_docx_in_win(path_raw, path_output):
    """
    doc转为docx（win）
    :param path_original:
    :param path_final:
    :return:
    """
    # 获取文件的格式后缀
    file_suffix = os.path.splitext(path_raw)[1]
    if file_suffix == ".doc":
        word = client.Dispatch('Word.Application')
        # 源文件
        doc = word.Documents.Open(path_raw)
        # 生成的新文件
        doc.SaveAs(path_output, 16)
        doc.Close()
        word.Quit()
    elif file_suffix == ".docx":
        shutil.copy(path_raw, path_output)


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
        file_path_raw = os.path.join(root, file)
        print(file_path_raw)

        os.system("soffice --headless --convert-to docx {} --outdir {}".format(file_path_raw, dest))
"""