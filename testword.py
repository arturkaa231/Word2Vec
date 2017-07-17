#!/usr/bin/env python
# -*- coding: utf-8 -*-
from gensim.models import Word2Vec
from gensim.models.word2vec import LineSentence
#библиотека для считывания
import xlrd
#библиотека для записи
import xlwt
import os.path
from xlutils.copy import copy as xlcopy
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap
from sklearn import neighbors, datasets
import pylab
#очистка словаря от знаков препинания
def clean(text):
	
	for i in [',','.',':',';','!','?']: 
		text=text.replace(i, '')
	
	return text
#считывания словаря с xls и запись в 	
def ReadXls(xls):
	wb=xlrd.open_workbook(os.path.join(xls))
	wb.sheet_names()
	sh=wb.sheet_by_index(0)
	WL=''
	i=0
	fi=open('words.txt','w')

	while i<sh.nrows:
		Load=sh.cell(i,0).value
		WL+=str(Load.encode('utf-8'))+"  "
		i+=1

	fi.write(WL+'\n')	
	fi.close()

#Запись в xls
def WriteXls(xls):
	#rb=xlrd.open_workbook(a)
	wb=xlwt.Workbook()
	ws=wb.add_sheet('result')
	#sheet=book.sheet_by_index(0)
	#wb=xlcopy(rb)

	ws=wb.get_sheet(0)
	k=0
	#запись отдельных слов в первый столбец
	for i in model.wv.vocab.keys() :
		ws.write(k,0,i)
		k+=1
	k=0
	#запись векторов во сторой столбец
	for i in model.wv.vocab.keys() :
		ws.write(k,1,str(model.wv[i]))
		k+=1
	wb.save(xls)
	

def WordMap():
	vectors=[]
	for i in model.wv.vocab.keys():
		vectors.append(model.wv[i])
	vectors=np.array(vectors)
	n_neighbors=15

	X = vectors

	h = .02  # step size in the mesh

	for weights in ['uniform', 'distance']:
	    # we create an instance of Neighbours Classifier and fit the data.
	    
	    x_min, x_max = X[:, 0].min() - 1, X[:, 0].max() + 1
	    y_min, y_max = X[:, 1].min() - 1, X[:, 1].max() + 1
	    xx, yy = np.meshgrid(np.arange(x_min, x_max, h),
		                 np.arange(y_min, y_max, h))
	  
	    
	    plt.xlim(-0.01, 0.01)
	    plt.ylim(-0.01, 0.01)
	    plt.title('words map')
	    
	    

	for i,text in enumerate(model.wv.vocab.keys()):
		plt.annotate(text,(X[:, 0][i],X[:, 1][i]))
	plt.show()

#ввод параметров модели	
a=input('enter the file address ')
size=input("Enter the vector size: ")
win=input("Enter the window size: ")
minc=input("Enter the minimal count of words: ")

ReadXls(a)


W=[]
#очищаем от знаков припенания и создаем список-словарь
for i in list(LineSentence('words.txt'))[0]:
	i=clean(i)
	W.append(unicode(i))
	
W3=[]
for i in W:
	i=i.split('  ')
	W3.append(i)
	
#Создаем модель
model=Word2Vec(W3,size=size, window=win,min_count=minc)
#print model.most_similar(u"sentence")

#запись векторов и словаря в xls
WriteXls(a)

#построение карты слов
WordMap()

