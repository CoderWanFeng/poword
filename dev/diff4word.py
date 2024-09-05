# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/8x7c9qiAneTsDJq9JnWLgA
@个人网站 ：www.python-office.com
@Date    ：2023/4/2 1:41 
@Description     ：https://juejin.cn/post/6899035340095520775
'''

from spire.doc import *
from spire.doc.common import *

# 加载Word文档
document = Document()
document.LoadFromFile("fdadd.docx")

# 遍历所有页面
for i in range(document.GetPageCount()):
    # 转换指定页面为图片流
    imageStream = document.SaveImageToStreams(i, ImageType.Bitmap)
    # 保存为.png图片（也可以保存为jpg或bmp等图片格式）
    with open("图片\\图-{0}.png".format(i), 'wb') as imageFile:
        imageFile.write(imageStream.ToArray())

# 关闭文档
document.Close()
