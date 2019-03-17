#!/usr/bin/python
# -*- coding: UTF-8 -*-
import pickle
import xlrd 
import xlwt
import jieba
from functools import reduce
import math
from jieba.analyse import extract_tags as et
from openpyxl import Workbook

localName = "data/hw1_text.xlsx"
jieba.set_dictionary('data/dict.txt')

dicttf_df_idf = {}
totalnum = 0
list1 = []

stop_words = set(line.strip() for line in open('data/chinese'))


def isAllChinese(s):
	if not(u'\u4e00' <= s <= u'\u9fa5'):
		return False
	return True

wb = xlrd.open_workbook(localName) 
for sheetNum in range(0, 3):
	count = 0
	sheet = wb.sheet_by_index(sheetNum) 
	#print sheet.cell_value(1, 3), "  ", sheet.cell_value(1, 4)
	for i in range(1, sheet.nrows):
		count += 1
		temp = sheet.cell_value(i, 3) + sheet.cell_value(i, 4)
		#print type(temp)
		words = jieba.cut(temp, cut_all = False)
		temp2 = []

		for word in words:
			#print (str)
			if isAllChinese(word):  # not digit
				#print "here"
				if len(word) > 1:   # not single character				 
					if word not in stop_words:
						temp2.append(word)

		list1.append(temp2)
	print ("sheet num:", sheetNum, " ", count)
	totalnum += count

# tf-df-idf
#print (len(list1))

for i in range(len(list1)):
	set_temp = set(list1[i])
	for word in set_temp:
		if word in dicttf_df_idf.keys():
			temp = dicttf_df_idf[word][1] # df
			#print (word, temp)
			temp += 1
			dicttf_df_idf[word][1] = temp
			#print (word, dicttf_df_idf[word][1])
			temp2 = dicttf_df_idf[word][0] # tf
			temp2 += list1[i].count(word)
			dicttf_df_idf[word][0] = temp2
		else:
			dicttf_df_idf[word] = [list1[i].count(word), 1 ]
print ("total number of text: ", totalnum)

d = {}
for key in dicttf_df_idf.keys():
	if dicttf_df_idf[key][0] >= 1000 and dicttf_df_idf[key][1] >= 500:
		d[key] = [dicttf_df_idf[key][0], dicttf_df_idf[key][1], dicttf_df_idf[key][0] * math.log((totalnum/dicttf_df_idf[key][1]))]


print ("length of dict: ", len(d))




def write(name, dict):	
	f = xlwt.Workbook()
	sheet1 = f.add_sheet(name,cell_overwrite_ok=True)
	count = 0
	sheet1.write(count, 0, "word")
	sheet1.write(count, 1, "tf")
	sheet1.write(count, 2, "df")
	sheet1.write(count, 3, "idf")
	count += 1
	for word in dict.keys():
		#print (len(dict[word]))
		sheet1.write(count, 0, word)
		sheet1.write(count, 1, dict[word][0])
		sheet1.write(count, 2, dict[word][1])
		sheet1.write(count, 3, dict[word][2])
		count += 1
	f.save('tf-df-idf.xlsx')

write(u'tf-df-idf', d)



with open("temp/list", "wb") as f:
	pickle.dump(list1, f)

with open("temp/dicttf_df_idf", "wb") as f:
	pickle.dump(d, f)









