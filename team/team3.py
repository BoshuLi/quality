import os
import math
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

path='D:\\code\\team\\tmp2'
files=os.listdir(path)

outputfilename='D:\\code\\team\\res2\\teamRes2.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)

count = 0
for f in files:
	tmpresult = 0
	numPro = 0
	print f
	bk=xlrd.open_workbook('D:\\code\\team\\tmp2\\'+f)
	#get the sheets number
	shxrange=range(bk.nsheets)
	print shxrange
	#get the sheets name
	for x in shxrange:
		p=bk.sheets()[x].name.encode('utf-8')
    	print "Sheets Number(%s): %s" %(x,p.decode('utf-8'))
	sh=bk.sheets()[0]
	nrows=sh.nrows
	ncols=sh.ncols
	# return the lines and col number
	print "line:%d  col:%d" %(nrows,ncols)
	i = 0
	while i < nrows:
		valN = 1
		val =[]
		val.append(float(sh.cell_value(i, 6)))
		conNow = sh.cell_value(i, 1)
		if float(sh.cell_value(i, 6)) == 1.0000000:
			while i!=nrows-1 and conNow == sh.cell_value(i+1, 1):
				i = i+1
				valN = valN + 1
			valN = max(valN,2)
			x=float(1.0000000/float(valN))
			var = ((1-x)*(1-x) + (valN-1)*x*x)/float(valN)
			#print var
			tmpresult = tmpresult + var
			numPro = numPro + 1
		else:
			while i!=nrows-1 and conNow == sh.cell_value(i+1, 1):
				i = i+1
				val.append(float(sh.cell_value(i, 6)))
				valN = valN + 1
			
			valN = max(valN,2)
			x=float(1.0000000/float(valN))	
			var = 0		
			for v in val:
				var = var+(v-x)*(v-x)
			var = float(var)/float(valN)
			#print var
			tmpresult = tmpresult + var
			numPro = numPro + 1
		i = i+1	
	result = tmpresult/numPro
	print result

	sheet1.write(count, 0, f.strip('.xls'))
	sheet1.write(count, 1, result)
	count+=1
outputfile.save(outputfilename)	
