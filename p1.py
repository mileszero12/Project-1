import xlrd
import math
import xlwt
import pickle
import numpy as np
from scipy.stats import chi2_contingency
from xlwt import *
N = 42258

stop_words_Dict = 'data/chinese'

stop_words = set(line.strip() for line in open(stop_words_Dict))
file = Workbook(encoding = 'utf-8')
def model(sheetNum, listname, saveSheetName):
	data = xlrd.open_workbook('d.xls')
	tai = data.sheet_by_index(sheetNum)
	#找到台積電的sheet
	cn1 = tai.col_values(0)
	cn2 = tai.col_values(1)
	n = []
	b = []
	for item in cn2:
		word = item[0:150]
		#前150字只要有部分不一樣的就算不一樣的文章
		if word not in n:
			n.append(word)
			b.append(cn2.index(item))
		else:
			continue

	del n
		

	with open(listname, 'rb') as data_file:
		x = pickle.load(data_file)
	word_lst = [] 
	word_dict = {}
	w_dict = {}
	for item in range(len(x)):
		a = []
		if item in b:
			for item2 in x[item]:
				word_dict[item2] = word_dict.get(item2, 0) + 1
				# tf
				if item2 not in a:
					w_dict[item2] = w_dict.get(item2, 0) + 1
					a.append(item2)
				# df

	finale_list = []
	for key in word_dict:
		if key not in stop_words:
			if word_dict[key] >= len(b) * 0.12 and w_dict[key] >= len(b) * 0.06:
				finale_list.append([key, word_dict[key], w_dict[key]])

	del word_dict
	del w_dict
	del x

	data = xlrd.open_workbook('data/hw1_table.xlsx')
	gram2 = data.sheet_by_index(0)
	gram3 = data.sheet_by_index(1)

	cn2 = gram2.col_values(1)
	cn3 = gram3.col_values(1)

	ct2 = gram2.col_values(2)
	ct3 = gram3.col_values(2)

	cd2 = gram2.col_values(3)
	cd3 = gram3.col_values(3)
	nval = 0


	doc = len(b)
	dicttemp = {}
	#找到finale_list中各詞對應在全部2gram 和全部3gram的TF和DF值
	for n in finale_list:
		try :
			nval = cn2.index(n[0])
			n.append(ct2[nval])
			n.append(cd2[nval])
		except ValueError:
			try:
				nval = cn3.index(n[0])
				n.append(ct3[nval])
				n.append(cd3[nval])
			except ValueError:
				continue
# 0tf, 1df, 2totaltf, 3totaldf, 4tfidf, 5mi, 6chi, 7tfidfrank, 8mirank, 9chirank, 10rank
		if n[2] > n[4]:
			n[2] = n[4]
		dicttemp[n[0]] = [n[1], n[2], n[3], n[4]]
		tfidf = (1 + math.log(n[1])) * math.log(doc/n[2])

		mi = abs(math.log(n[2] * N/(n[4] * doc)))
		dicttemp[n[0]].append(tfidf)#4
		dicttemp[n[0]].append(mi)#5

		a = dicttemp[n[0]][1]
		b = doc - a
		c = dicttemp[n[0]][3] - a
		d = N - doc - c
		obs1 = np.array([a, b, c, d])
		obs2 = obs1.reshape((2, 2))
		chi2, p, dof, ex = chi2_contingency(obs2, lambda_ = "log-likelihood")
		dicttemp[n[0]].append(chi2)#6

	list1 = sorted(dicttemp.items(), key=lambda item:item[1][4], reverse = True)
	list2 = sorted(dicttemp.items(), key=lambda item:item[1][5], reverse = True)
	list3 = sorted(dicttemp.items(), key=lambda item:item[1][6], reverse = True)
	for i in range(len(list1)):
		dicttemp[list1[i][0]].append(i)
	#print (len(dicttemp['共同']))
	for i in range(len(list2)):
		dicttemp[list2[i][0]].append(i)
	#print (list2)
	#print (len(dicttemp['共同']))
	for i in range(len(list3)):
		dicttemp[list3[i][0]].append(i)
	#print (len(dicttemp['共同']))
	for k in dicttemp.keys():
		dicttemp[k].append(dicttemp[k][7] + dicttemp[k][8] + dicttemp[k][9])

	res = sorted(dicttemp.items(), key=lambda item:item[1][10], reverse = False)
	x = [0,1,2,3,4,5,6,7,8,9,10]

	#生成xls檔

	
	table = file.add_sheet(saveSheetName)
	table.write(0, 0, 'word')
	table.write(0, 1, 'tf in topic')
	table.write(0, 2, 'df in topic')
	table.write(0, 3, 'tf of total')
	table.write(0, 4, 'df of total')
	table.write(0, 5, 'tf-idf')
	table.write(0, 6, 'MI')
	table.write(0, 7, 'chi')
	table.write(0, 8, 'Tf-idf_Rank')
	table.write(0, 9, 'MI_Rank')
	table.write(0, 10, 'Chi_Rank')
	table.write(0, 11, 'Total_Rank')
	

	for i in range(100):
		table.write(i+1, 0, res[i][0])
		#print (res[i][0])
		for j in range(len(res[i][1])):
			table.write(i+1, j + 1,res[i][1][j] )
	



model(0, 'temp/list1', '銀行')
model(1, 'temp/list2', '信用卡')
model(2, 'temp/list3', '台積電')
model(3, 'temp/list4', '匯率')
model(4, 'temp/list5', '台灣')
model(5, 'temp/list6', '日本')


file.save('res.xls')














