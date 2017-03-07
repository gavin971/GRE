import re
import xlwt
data = open('C:/Users/watersir zhangga/Documents/0英语/单词.txt', 'r')
i = 0
# 如果开头是单词，那么该行为单词及其释义、词源
# 开头为单词的定义：一行以空格分割后第一个list中全为字母，不含符号数字
# 开头不是单词，则是上一行的词源或解释。
word = []
ciyuan = {}
meaning = {}
i = 0
for l in data:
    
    if re.match('^[a-zA-Z]+$', l.split()[0]):
        tw = l.split()[0] # this word
        word += [tw]
        i += 1
        ciyuan[tw] = ''
        meaning[tw] = ''
        for x in range(len(l.split())):
            if x != 0:
                meaning[tw] += l.split()[x]
    elif re.match('^[0-9]+$', l.split()[0]):
        print(l)
    else:
        ciyuan[word[i-1]] += l    
dic = open('dic.txt', 'w')
f = xlwt.Workbook()
sheet = f.add_sheet(u'sheet')
for i in range(len(word)):
    sheet.write(i, 0, word[i])
    sheet.write(i, 1, meaning[word[i]])
    sheet.write(i, 2, ciyuan[word[i]])
f.save('test.xls')