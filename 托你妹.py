import xlrd
import xlwt
import datetime
from xlutils.copy import copy
import xlutils
import os
import requests
import re
import pygame
filepath = r'C:\Users\watersir zhangga\Desktop\托福词汇'+'\\'
def ciyuan(w):
    excel = filepath + '词源.xls'
    workbook = xlrd.open_workbook(excel)
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    i = 0
    while True:
        if i >= nrows:
            break
        else:
            word_byte_utf8 = sheet.cell(i,0).value.encode('utf-8')
            word = word_byte_utf8.decode()
            if word==w:
                ciyuan_byte_utf8 = sheet.cell(i,2).value.encode('utf-8')
                print('词源：', ciyuan_byte_utf8.decode())
                break
            else:
                i += 1
def vocal(word):

    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:39.0) Gecko/20100101 Firefox/39.0',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate',
        'Connection': 'keep-alive'}
    url = 'http://www.iciba.com/' + word
    html=requests.get(url,headers=headers).text
    m = html.split('\n')
    v = []
    for l in m:
        t = re.findall(r'(<i class="new-speak-step" ms-on-mouseover="sound..)(.*)(...></i>)', l)
        if len(t) != 0:
            v.append(t[0][1])
    r = requests.get(v[1]) # 美音
    path = 'E:temp/'
    name = word + '.mp3'
    if not os.path.exists(path + name):
        with open(path + name, 'wb') as f:
            f.write(r.content)
        f.close()
    file = path + name
    pygame.mixer.init()
    track = pygame.mixer.music.load(file)
    pygame.mixer.music.play()
    #os.system(path + name)

def load():
    doc=open(filepath+'doc.txt','r')
    last = 0
    list = 0
    unit = 0
    for l in doc:
        l = l.split(':')
        if l[0]=='记至':
            last = int(l[1])
        if l[0]=='list':
            list = int(l[1])
        if l[0]=='unit':
            unit = int(l[1])
    return (last,list,unit)
def getkeys(keys):
    while(True):
        k = input()
        flag = 1
        for i in range(len(keys)):
            if k== keys[i]:
                flag = 0
                break
        if flag == 0:
            break
    return k
def archive(last,list,unit,num,time):
    doc=open(filepath+'doc.txt','a')
    doc.write('\n时间:'+str(datetime.datetime.now())+'\n')
    doc.write('list:'+str(list)+'\n')
    doc.write('unit:'+str(unit)+'\n')
    doc.write('记至:'+str(last)+'\n')
    doc.write('记忆个数:'+str(num)+'\n')
    doc.write('记忆时长:'+str(time)+'\n')
def review(list,unit):#小单元复习
    print('>>>复习list',list+1,',单元',unit)
    excel = filepath+'小3000.xls'
    workbook = xlrd.open_workbook(excel)
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    wb = copy(workbook)
    shit = wb.get_sheet(0)
    n = 0#本次记忆个数
    last = list*100+unit*10+1
    while True and n<10:
        word = sheet.cell(last,2).value.encode('utf-8')
        print(last,':',word.decode('utf-8'))
        paraphrase_eng = sheet.cell(last,4).value.encode('utf-8')
        print('音标:',paraphrase_eng.decode('utf-8'))
        whole = sheet.cell(last,0).value
        forget = sheet.cell(last,1).value
        choose = getkeys(['1','2','3','4','1 note','2 note','3 note'])
        key = int(choose.split()[0])
        note_flag = ''
        if len(choose.split())==2:
            note_flag = choose.split()[1]
        if key==1:
            whole+=1
            shit.write(last,0,whole)
        if key==2:
            whole+=1
            forget+=1
            shit.write(last,0,whole)
            shit.write(last,1,forget)
        if key==3:
            whole+=1
            shit.write(last,0,whole)
            shit.write(last,4,'easy')
        if key==4:
            break
        paraphrase_chn = sheet.cell(last,3).value.encode('utf-8')
        print('中文释义:',paraphrase_chn.decode('utf-8'))

        if note_flag == 'note':
            note = input('输入笔记：')
            shit.write(last,5,note)
        n += 1
        last += 1
        wb.save(excel)
    print('>>>复习结束！list', list + 1, ',单元', unit + 1)
def getnote():
    print('请输入笔记:')
    note = []
    while True:
        t = input()
        if t != 'end':
            note.append(t)
        else:
            break
    return note
if __name__ == '__main__':
    #获得当前时间
    begin = datetime.datetime.now()
    (last,list,unit) = load()
    print('上次list:', list + 1, '上次unit:', unit + 1, ',上次记到:', last)
    choose = int(input('继续请按1，新选择list按2：'))
    if choose == 2:
        list = int(input('请输入list:')) - 1
        last = list*100 + 1
    print('开始记单词，list:', list + 1)
    excel = filepath + '托你妹.xls'
    workbook = xlrd.open_workbook(excel)
    sheet = workbook.sheet_by_index(0)
    nrows = sheet.nrows
    ncols = sheet.ncols
    wb = copy(workbook)
    shit = wb.get_sheet(0)
    num = 0#本次记忆个数
    thisnum = 0
    while True:    
        word = sheet.cell(last,2).value.encode('utf-8')
        print(last,':',word.decode('utf-8'))
        paraphrase_eng = sheet.cell(last,4).value.encode('utf-8')
        print('音标:',paraphrase_eng.decode('utf-8'))
        vocal(str(word.decode('utf-8'))) # 发音
        
        whole = sheet.cell(last,0).value
        forget = sheet.cell(last,1).value
        choose = getkeys(['1','2','3','4','1 note','2 note','3 note','re'])
        key = choose.split()[0]
        note_flag = ''
        if len(choose.split())==2:
            note_flag = choose.split()[1]
        if key == 're':
            this_forget = sheet.cell(last - 1, 1).value
            shit.write(last - 1, 1, this_forget + 1)
            print('已撤销！\n')
            choose = getkeys(['1','2','3','4','1 note','2 note','3 note','re'])
            key = choose.split()[0]
        if key == '1':
            whole += 1
            shit.write(last, 0, whole)
        if key == '2':
            whole += 1
            forget += 1
            shit.write(last, 0, whole)
            shit.write(last, 1, forget)
        if key == '3':
            whole += 1
            shit.write(last, 0, whole)
            shit.write(last, 4, 'easy')
        if key == '4':
            break
        chn = sheet.cell(last,3).value.encode('utf-8')
        print('中文释义:', chn.decode('utf-8'))
        paraphrase_note = sheet.cell(last,5).value.encode('utf-8')
        if len(paraphrase_note) != 0:
            print('笔记:',paraphrase_note.decode('utf-8'))
        ciyuan = sheet.cell(last, 6).value.encode('utf-8')
        if len(ciyuan) != 0:
            print('词源:', ciyuan.decode('utf-8')) 
        if note_flag == 'note':
            note = getnote()
            shit.write(last, 5, note)
        num += 1
        thisnum += 1
        last += 1
        wb.save(excel)
        # IfReview = input('是否复习？')
        # if IfReview == 1:
        # if thisnum>=10:
            # review(list,unit)
            # unit+=1
            # thisnum=0
        # if unit==5:
            # for i in range(6):
                # review(list,i)
        # if unit==11:
            # for i in range(12):
                # review(list,i)
        if int(last/100)>list:
            print('>>>list', list + 1, '完成！开始list',list+2)
            list += 1
    end = datetime.datetime.now()
    time = end - begin
    list = int(last/100)
    archive(last,list,unit,num,time)
    print('记至：',last,'\n记忆个数',num,'\n记忆时长',time)
    
            
            
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
