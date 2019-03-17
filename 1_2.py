#!/usr/bin/python
# -*- coding: UTF-8 -*-
import pickle
import xlrd 
import xlwt
import json
from functools import reduce
import math
from openpyxl import Workbook

dicttf_df_idf = {}
with open('temp/list', 'rb') as data_file:
	listAll = pickle.load(data_file)
with open("temp/dicttf_df_idf", "rb") as f:
	dicttf_df_idf = pickle.load(f)

tfidfDictList = dicttf_df_idf.keys()


list1, list2, list3, list4, list5, list6 = [], [], [], [],[],[]

n1 = 7735
n2 = 232
n3 = 2342
n4 = 1444
n5 = 25974
n6 = 4329
n_all = n1 + n2 + n3 + n4 + n5 + n6
#workbook = xlwt.Workbook(encoding='utf-8')


for i in range(len(listAll)):
	if '銀行' in listAll[i]:
		list1.append(listAll[i])

	elif ('信用卡' in listAll[i]) or ('用卡' in listAll[i]) or ('信用' in listAll[i]):
		list2.append(listAll[i])

	elif '台積電' in listAll[i]:
		list3.append(listAll[i])

	elif '匯率' in listAll[i]:
		list4.append(listAll[i])

	elif '台灣' in listAll[i]:
		list5.append(listAll[i])

	elif '日本' in listAll[i]:
		list6.append(listAll[i])


#98484,7735,232,2342,1444,25974,4329
def wirteintopickle():
	with open("temp/dict1", "wb") as f:
		pickle.dump(dict1, f)

	with open("temp/dict2", "wb") as f:
		pickle.dump(dict2, f)

	with open("temp/dict3", "wb") as f:
		pickle.dump(dict3, f)

	with open("temp/dict4", "wb") as f:
		pickle.dump(dict4, f)

	with open("temp/dict5", "wb") as f:
		pickle.dump(dict5, f)

	with open("temp/dict6", "wb") as f:
		pickle.dump(dict6, f)

	print (len(listAll))
	print (len(list1))
	print (len(list2))
	print (len(list3))
	print (len(list4))
	print (len(list5))
	print (len(list6))




#print ("dere", len(dicttf_df_idf))
# [word] : [tf, df, idf, dfInTopic, dfOfTopic, dfAll]
def calulate(nthistopic, list):
	dictNew = {}
	global dicttf_df_idf
	for i in range(len(list)):
		set_temp = set(list[i])
		#print (len(list[i]), len(set_temp)) 
		for word in set_temp:
			if word in dicttf_df_idf.keys():
				if word in dictNew.keys():
					temp = dictNew[word][3]
					temp += 1
					dictNew[word][3] = temp
				else:
					dictNew[word] = [dicttf_df_idf[word][0], dicttf_df_idf[word][1], dicttf_df_idf[word][2], 1, nthistopic, n_all]
	return dictNew


dict1 = calulate(n1, list1)
dict2 = calulate(n2, list2)
dict3 = calulate(n3, list3)
dict4 = calulate(n4, list4)
dict5 = calulate(n5, list5)
dict6 = calulate(n6, list6)

f = xlwt.Workbook()
def writeintoExcel(name, dict):	
	sheet1 = f.add_sheet(name,cell_overwrite_ok=True)
	sheet1.write(0, 0, "name")
	sheet1.write(0, 1, "tf")
	sheet1.write(0, 2, "df")
	sheet1.write(0, 3, "idf")
	sheet1.write(0, 4, "df In Topic")
	sheet1.write(0, 5, "num of Topic")
	sheet1.write(0, 6, "all document num of this six catalog")
	i = 1
	for word in dict.keys():
		sheet1.write(i, 0, word)
		sheet1.write(i, 1, dict[word][0])
		sheet1.write(i, 2, dict[word][1])
		sheet1.write(i, 3, dict[word][2])
		sheet1.write(i, 4, dict[word][3])
		sheet1.write(i, 5, dict[word][4])
		sheet1.write(i, 6, dict[word][5])
		i += 1

writeintoExcel(u'銀行', dict1)
writeintoExcel(u'信用卡', dict2)
writeintoExcel(u'台積電', dict3)
writeintoExcel(u'匯率', dict4)
writeintoExcel(u'台灣', dict5)
writeintoExcel(u'日本', dict6)


f.save('tf_df_idf_all_catalog.xlsx')

wirteintopickle()






















