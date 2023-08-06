# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/yFcocJbfS9Hs375NhE8Gbw
@个人网站 ：www.python-office.com
@Date    ：2023/4/2 1:38 
@Description     ：
'''

from docxcompose.composer import Composer

def compose_files(self, files, output_file_path):
    """
    合并多个word文件到一个文件中
    :param files:待合并文件的列表
    :param output_file_path 新的文件路径
    :return:
    """
    composer = Composer(Document())
    for file in files:
        composer.append(Document(file))

    # 保存到新的文件中
    composer.save(output_file_path)

# 作者：AirPython
# 链接：https://juejin.cn/post/6899035340095520775
# 来源：稀土掘金
# 著作权归作者所有。商业转载请联系作者获得授权，非商业转载请注明出处。