# -*- coding:utf-8 -*-
import re

if __name__ == '__main__':
    words = '"//misc.360buyimg.com","//img10.360buyimg.com"/,"//img12.360buyimg.com"' # 字符串太长,不好粘贴 先不粘贴了
    pattern = re.compile(r'(?://img).*?"')
    print(pattern.findall(words))

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
