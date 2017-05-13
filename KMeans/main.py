#! /usr/bin/env python
#coding=utf-8

import kMeans
from numpy import *
import pylab
from numpy import *
import xlwt

def showFigure(dataMat,k,clusterAssment):

    tag=['go','or','yo','ko', 'bo', 'mo']
    for i in range(k):
        datalist = dataMat[nonzero(clusterAssment[:,0].A==i)[0]]
        pylab.plot(datalist[:,0],datalist[:,1],tag[i])
    pylab.show()

    
    row = 0
    for i in range(k):
		datalist = dataMat[nonzero(clusterAssment[:,0].A==i)[0]]
		for j in range(len(datalist)):
			sheet1.write(row, 0, datalist[j,0])
			sheet1.write(row, 1, datalist[j,1])
			sheet1.write(row, 2, tag[i])
			row += 1
	

if __name__ == '__main__':
	outputfilename='D:\\code\\team\\res.xls'
	outputfile=xlwt.Workbook()
	sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
	k=6
	dataMat = mat(kMeans.loadDataSet('D:\\code\\team\\data.txt'))
	myCentroids,clusterAssment=kMeans.kMeans(dataMat,k)
	showFigure(dataMat,k,clusterAssment)
	outputfile.save(outputfilename)




