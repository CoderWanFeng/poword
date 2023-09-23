# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/B1V6KeXc7IOEB8DgXLWv3g
@个人网站 ：www.python-office.com
@Date    ：2023/4/2 1:41 
@Description     ：https://juejin.cn/post/6899035340095520775
'''

# 两个 Word 文档的对比也是工作中比较常见的需求了
#
# 首先，遍历文档中所有段落，过滤掉空行，获取所有文本内容

# 分别获取段落内容
content1 = ''
content2 = ''
file1 = ''
file2 = ''
for paragraph in file1.paragraphs:
    if "" == paragraph.text.strip():
        continue
    content1 += paragraph.text + '\n'

for paragraph in file2.paragraphs:
    if "" == paragraph.text.strip():
        continue
    content2 += paragraph.text + '\n'

# 如果参数 keepends 为 False，不包含换行符，如果为 True，则保留换行符。
print("第二个文档数据如下：\n", content1.splitlines(keepends=False))
print("第一个文档数据如下：\n", content1.splitlines(keepends=False))

# 接着，使用 Python 中的标准依赖库 difflib 对比文字间的差异，最后生成 HTML 差异报告

import codecs
from difflib import HtmlDiff

# 差异内容
diff_html = HtmlDiff(wrapcolumn=100).make_file(content1.split("\n"), content2.split("\n"))

# 写入到文件中
with codecs.open('./diff_result.html', 'w', encoding='utf-8') as f:
     f.write(diff_html)