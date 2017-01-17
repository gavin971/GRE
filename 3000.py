import xlrd,xlwt
import datetime
from xlutils.copy import copy
import xlutils
import sys,os
last = 0
list = 0
unit = 0
#获取脚本文件的当前路径
def cur_file_dir():
     #获取脚本路径
     path = sys.path[0]
     #判断为脚本文件还是py2exe编译后的文件，如果是脚本文件，则返回的是脚本的目录，如果是py2exe编译后的文件，则返回的是编译后的文件路径
     if os.path.isdir(path):
         return path
     elif os.path.isfile(path):
         return os.path.dirname(path)
filepath = cur_file_dir()
def load():
	doc=open(filepath+'\doc.txt','r')
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
	doc=open(filepath+'\doc.txt','a')
	doc.write('\n时间:'+str(datetime.datetime.now())+'\n')
	doc.write('list:'+str(list)+'\n')
	doc.write('unit:'+str(unit)+'\n')
	doc.write('记至:'+str(last)+'\n')
	doc.write('记忆个数:'+str(num)+'\n')
	doc.write('记忆时长:'+str(time)+'\n')
def review(list,unit):#小单元复习
	print('>>>复习list',list+1,',单元',unit)
	excel = filepath+'\小3000.xls'
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
		paraphrase_eng = sheet.cell(last,4).value.encode('utf-8')
		print('英文释义:',paraphrase_eng.decode('utf-8'))
		if note_flag == 'note':
			note = input('输入笔记：')
			shit.write(last,5,note)
		n += 1
		last += 1
		wb.save(excel)
	print('>>>复习结束！list',list+1,',单元',unit+1)

#获得当前时间
begin = datetime.datetime.now()
(last,list,unit) = load()
last = list*100+unit*10+1
print('上次list:',list+1,'上次unit:',unit+1,',上次记到:',last)
choose = int(input('继续请按1，新选择list按2：'))

if choose == 2:
	list = int(input('请输入list:'))-1
	last = list*100 + 1
print('开始记单词，list:',list)
excel = filepath+'\小3000.xls'
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
	paraphrase_eng = sheet.cell(last,4).value.encode('utf-8')
	print('英文释义:',paraphrase_eng.decode('utf-8'))
	if note_flag == 'note':
		note = input('输入笔记：')
		shit.write(last,5,note)
	num += 1
	thisnum += 1
	last += 1
	wb.save(excel)
	if thisnum>=10:
		review(list,unit)
		unit+=1
		thisnum=0
	if unit==5:
		for i in range(6):
			review(list,i)
	if unit==11:
		for i in range(12):
			review(list,i)
	if int(last/100)>list:
		print('>>>list',list+1,'完成！开始list',list+2)
		list += 1
end = datetime.datetime.now()
time = end - begin
list = int(last/100)
archive(last,list,unit,num,time)
print('记至：',last,'\n记忆个数',num,'\n记忆时长',time)

			
			
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
