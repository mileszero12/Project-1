#!/usr/bin/python
# -*- coding: UTF-8 -*-
import pickle
import xlrd 
import xlwt
import json
import math
import numpy as np
from decimal import *
from scipy.stats import chi2_contingency
from openpyxl import Workbook

dict1, dict2, dict3, dict4, dict5, dict6 = {}, {}, {}, {}, {}, {}


with open("temp/dict1", "rb") as f:
	dict1 = pickle.load(f)

with open("temp/dict2", "rb") as f:
	dict2 = pickle.load(f)

with open("temp/dict3", "rb") as f:
	dict3 = pickle.load(f)

with open("temp/dict4", "rb") as f:
	dict4 = pickle.load(f)
	#print (len(dict4))

with open("temp/dict5", "rb") as f:
	dict5 = pickle.load(f)

with open("temp/dict6", "rb") as f:
	dict6 = pickle.load(f)

def tf_idf(dict):
	#print (len(dict))
	list1 = sorted(dict.items(), key=lambda item:item[1][2], reverse = True)
	dict11 = {}
	#print (len(list1))
	for i in range(0, 200):
		dict11[list1[i][0]] = list1[i][1]
	return dict11


#print (type(dict1))
dict11 = tf_idf(dict1)
dict22 = tf_idf(dict2)
dict33 = tf_idf(dict3)
dict44 = tf_idf(dict4)
dict55 = tf_idf(dict5)
dict66 = tf_idf(dict6)



# [6] = mi
def featureSelection(dict1):
	for k in dict1.keys():
		a = dict1[k][3]
		b = dict1[k][4] - a
		c = dict1[k][1] - a
		d = dict1[k][5] - dict1[k][4] - c
		obs1 = np.array([a, b, c, d])
		obs2 = obs1.reshape((2, 2))
		g, p, dof, ex = chi2_contingency(obs2, lambda_ = "log-likelihood")
		mi2 = float(0.5 * g / sum(obs1)) / math.log(2)
		#print type(dict1)
		dict1[k].append(mi2)
	list1 = sorted(dict1.items(), key=lambda item:item[1][3], reverse = True)
	#print list1
	for i in range(0, 200):
		dict1[list1[i][0]].append(i)

featureSelection(dict11)
featureSelection(dict22)
featureSelection(dict33)
featureSelection(dict44)
featureSelection(dict55)
featureSelection(dict66)

def chi(array, docNum, n1):
    a = 0
    b = 0
    for i in range(0, 2):
        a += array[i][0]
        b += array[i][1]

 
    a = float(a * docNum) / n1
    b = float(b * (n1 - docNum)) / n1
    aa = 0
    bb = 0
    for i in range(0, 2):
        aa += float((array[i][0] - a)*(array[i][0] - a)) / a
        bb += float((array[i][1] - b)*(array[i][1] - b)) / b
    return aa + bb

#[8]chi
def ChiSquareFeatureSelection(dict1):
	for k in dict1.keys():
		a = dict1[k][3]
		b = dict1[k][4] - a
		c = dict1[k][1] - a
		d = dict1[k][5] - dict1[k][4] - c
		obs1 = np.array([a, b, c, d])
		obs2 = obs1.reshape((2, 2))
		chi2 = chi(obs2, dict1[k][4], dict1[k][5])
		dict1[k].append(chi2)
	list1 = sorted(dict1.items(), key=lambda item:item[1][5], reverse = True)
	#print list1
	for i in range(0, 200):
		dict1[list1[i][0]].append(i)


ChiSquareFeatureSelection(dict11)
ChiSquareFeatureSelection(dict22)
ChiSquareFeatureSelection(dict33)
ChiSquareFeatureSelection(dict44)
ChiSquareFeatureSelection(dict55)
ChiSquareFeatureSelection(dict66)



# 0:tf 1:df 2:idf 3:dfIntopic 4:docNumOfTopic 5:allDocNum 6:mi 7:miRank 8:chi 9:chiRank 10:totalRank
def total(dict1):
	for k in dict1.keys():
		dict1[k].append(dict1[k][6] + dict1[k][8])
	a = []
	temp = 0
	list1 = sorted(dict1.items(), key=lambda item:item[1][10])
	for i in range(0, 200):
		if temp <= 100:
			a.append(list1[i])
			temp += 1
		else:
			break
	return a

#print (dict11)
a1 = total(dict11)
a2 = total(dict22)
a3 = total(dict33)
a4 = total(dict44)
a5 = total(dict55)
a6 = total(dict66)
#print (a1[1][1])

f = xlwt.Workbook()
def writeintoExcel(name, list2):	
	sheet1 = f.add_sheet(name,cell_overwrite_ok=True)
	sheet1.write(0, 0, "Term")
	sheet1.write(0, 1, "tf")
	sheet1.write(0, 2, "df")
	sheet1.write(0, 3, "idf")
	sheet1.write(0, 4, "Doc Num In Topic")
	sheet1.write(0, 5, "Doc Num of Topic")
	sheet1.write(0, 6, "all Doc Num")
	sheet1.write(0, 7, "MI Value")
	sheet1.write(0, 8, "MI Rank")
	sheet1.write(0, 9, "Chi Value")
	sheet1.write(0, 10, "Chi Rank")
	sheet1.write(0, 11, "Total Rank")
	list1 = list2[::-1]
	n = len(list1) - 1
	for i in range(0, n):
		sheet1.write(i + 1, 0, list1[i][0])
		for j in range(0,11):
			sheet1.write(i + 1, j + 1, list1[i][1][j])
		


writeintoExcel(u'銀行', a1)
writeintoExcel(u'信用卡', a2)
writeintoExcel(u'台積電', a3)
writeintoExcel(u'匯率', a4)
writeintoExcel(u'台灣', a5)
writeintoExcel(u'日本', a6)
f.save('100.xlsx')














