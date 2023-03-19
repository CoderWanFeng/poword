#!/usr/bin/env python
# -*- coding:utf-8 -*-

#############################################
# File Name: word.py
# 公众号/B站/小红书/抖音: 程序员晚枫
# Mail: 1957875073@qq.com
# Created Time:  2022-4-25 10:17:34
# Description: 有关word的自动化操作
#############################################

# 创建对象
from poword.core.WordType import MainWord

mainWord = MainWord()


# 1、文件的批量转换
# 自己指定路径，
# 为了适配wps不能转换doc的问题，这里限定：只能转换docx
# @except_dec()
def docx2pdf(path):
    mainWord.docx2pdf(path)
def merge4docx(input_path, output_path, new_word_name):
    mainWord.merge4docx(input_path, output_path, new_word_name)
