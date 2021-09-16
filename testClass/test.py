# -*- coding: utf-8 -*-
# Author    ： ly
# Datetime  ： 2021/9/15 11:44
import os

CHECK_TYPE = {
    1: u'检查',
    11: u'胃镜',
    111: u'电子胃镜',
    112: u'染色放大',
    113: u'超声胃镜',
    12: u'肠镜',
    121: u'电子肠镜',
    122: u'染色放大',
    123: u'超声肠镜',
    2: u"治疗",
    21: u"ESD",
    211: u"胃ESD",
    212: u"肠ESD",
    22: u"静脉曲张",
    23: u"ERCP",
    231: u"ERCP",
    24: u"息肉",
    241: u"胃息肉治疗",
    242: u"肠息肉治疗",

}
d = {'folder': 'C:/Users/Administrator/Desktop/葵花宝典/Python书籍',
     'files': 'C:/Users/Administrator/Desktop/葵花宝典/Python书籍\\Python爬虫开发与项目实战.pdf\nC:/Users/Administrator/Desktop/葵花宝典/Python书籍\\Python知识手册-V3.0.pdf\nC:/Users/Administrator/Desktop/葵花宝典/Python书籍\\Python编程快速上手—让繁琐工作自动化.pdf\nC:/Users/Administrator/Desktop/葵花宝典/Python书籍\\《Python+Cookbook》第三版中文v2.0.0.pdf\nC:/Users/Administrator/Desktop/葵花宝典/Python书籍\\《Python深度学习》2018中文版.pdf\nC:/Users/Administrator/Desktop/葵花宝典/Python书籍\\《Python深度学习》2018英文版.pdf\nC:/Users/Administrator/Desktop/葵花宝典/Python书籍\\像计算机科学家一样思考python.pdf\n'}

if __name__ == '__main__':
    # 从前端获取的value值 作为key查字典得到对应的value值
    # check_type = 111
    # print(CHECK_TYPE[check_type])
    folder = d.get('files', '').replace('\n', ',').split(',')
    # print(folder)
    flag = None
    for f in folder:
        ext = os.path.splitext(f)[1]
        if ext == '.pdf':
            flag = True
            break
    print(flag)
