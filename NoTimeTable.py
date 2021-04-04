# -*- coding: utf-8 -*-
# @author: Alan Harris-腰间盘同学
# @datetime: 2021/4/4 12:25
# @email: alanharrisno.1@outlook.com
# @github: github.com/AlanHarrisNo1

import os
import re
import xlrd
import xlwt

zkb = []
#总课表列表
nm = []
#名单
for i in range(42):
	zkb.append('')

#获取目录下的所有xls课表
def getfiles():
	dirs = os.listdir(os.getcwd())
	xlslist = []
	for i in dirs:
		if os.path.splitext(i)[1] == '.xls':
			xlslist.append(i)
	return xlslist

#每个单元格1-16周是否有课判断
def jag(ifmt):
	st = []
	for i in range(1,17):
		st.append(i)
	for line in ifmt.splitlines():
		if line.find('周]')!=-1:
			week = re.split(r'[,\[\]]',line)
			n=0
			m=-1
			if week[-2] == '周':
				n=1
			elif week[-2] == '单周':
				m=0
			else:
				m=1
			del week[-2:]
			for items in week:
				if items.find('-')!=-1:
					item = re.split(r'-',items)
					for i in range(int(item[0]),int(item[1])+1):
						if n == 1:
							if i in st:
								st.remove(i)
						if m == 0:
							if (i in st) & (i%2 == 1):
								st.remove(i)
						if m == 1:
							if (i in st) & (i%2 == 0):
								st.remove(i)
				else:
					if int(items) in st:
						st.remove(int(items))

	return st

#每个课表处理
def echfl(xls):
	wb = xlrd.open_workbook(xls)
	sh1 = wb.sheet_by_index(0)
	title = sh1.cell_value(0, 0)
	title = title.split( )
	nm.append(title[1])
	n = 0
	for i in range(3, 9):						#修改行
		for j in range(1, 8):					#修改列
			kb = str(jag(sh1.cell_value(i,j)))
			if kb != '[]' and len(kb) != 55:
				zkb[n] += title[1] + kb + '\n'
			if len(kb) == 55:
				zkb[n] += title[1] + '\n'
			n+=1

#数据处理
def lalala():	
	xlslist = getfiles()
	for xls in xlslist:
		echfl(xls)
	print(nm,'\n','共',len(nm),'人')

#制作课表
def hahaha():
	xxx = xlwt.Workbook()
	sh1 = xxx.add_sheet('无课表')
	week = ['一','二','三','四','五','六','日']
	i = 1 
	for days in week:
		sh1.write(0, i, '星期'+days)
		i+=1
	j = 102
	for i in range(1, 7):
		sh1.write(i, 0, (str(j)+'节').zfill(5))
		j+=202
	n = 0
	i = j = 1
	for i in range(1, 7):
		for j in range(1, 8):
			sh1.write(i, j, zkb[n])
			n+=1
	xxx.save('无课表.xls')
	print('无课表制作完成，请将其移出该目录，防止下次运行出错！')

if __name__ == '__main__':
	lalala()
	hahaha()


