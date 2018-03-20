#!/usr/local/bin/python
# -*- coding: utf-8 -*-


########################################################################

#功能：将propertise文件转换成excel，支持传文件夹和文件的绝对路径,
#      输入文件夹路径将转换文件夹内所有.properties文件到一个excel文件。
#Author：leixin

########################################################################

import os
import xlwt

#解析propertise 文件，追加到excel中
def propertise_to_excel(fileName,path):
	global ws
	global row
	
	#校验文件后缀
	if(os.path.splitext(path)[1] != '.properties'):
		return
	file = open(path,'r')
	
	while 1:
		lines = file.readlines(10000)
		if not lines:
			break
		for line in lines:
			line = line.decode("utf-8").strip()
			if(line == ''):
				continue
			#print "line --- " + line.decode("utf-8")
			strs = line.split('=',1)
			if(len(strs) != 2):
				continue
			ws.write(row, 0, fileName)
			ws.write(row, 1, strs[0].strip())
			ws.write(row, 2, strs[1].strip())
			row+=1

#递归遍历path下及其所有层级子目录下的文件
def transfer_all_propertise(path):
	if not os.path.isdir(path):
		propertise_to_excel(path.split("\\")[-1],path)
	else:
		files = os.listdir(path)
		for file in files:
			if  os.path.isdir(path+'\\'+file):
				transfer_all_propertise(path+'\\'+file)
			else:
				propertise_to_excel(file,path+'\\'+file)
			
			
if __name__ == '__main__':
	#获得文件目录/文件
	while(1):
		root_path = raw_input("please input path [input '-1' to exit] : ")
		if(root_path == '' or root_path == '-1'):
			exit()
		fileName = root_path.split('\\')[-1]
		print fileName
		wb = xlwt.Workbook();
		ws = wb.add_sheet(fileName,'utf-8');
		row = 2	
		transfer_all_propertise(root_path)
		wb.save(fileName+'.xlsx')
	
	
