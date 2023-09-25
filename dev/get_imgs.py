# -*- coding: UTF-8 -*-
'''
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信     ：CoderWanFeng : https://mp.weixin.qq.com/s/HYOWV7ImvTXImyYWtwADog
@个人网站      ：www.python-office.com
@代码日期    ：2023/9/25 22:26 
@本段代码的视频说明     ：
'''

from os.path import basename, dirname, join
from docx import Document, ImagePart
def extract_image(document):
    for rel in document.part.rels.values():  # 遍历文档中的所有关联对象
        if "image" in rel.reltype:  # 找到关联类型为图片的对象
            part = rel.target_part
            if isinstance(part, ImagePart):  # 如果是图片对象
                save_dir = dirname(__file__)  # 提取路径部分，丢掉文件名
                save_path = join(save_dir, basename(part.partname))  # 默认文件名image1.img
                with open(save_path, "wb") as f:
                    f.write(part.blob)
import poword
if __name__ == '__main__':
    # doc = Document("./fdadasf.docx")
    # extract_image(doc)
    # print("提取图片成功")
    poword.docx4imgs(word_path=r'./fdadasf.docx', img_path=r'./out')