import os
import pickle
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

#open the excel file
techfile="D:\\code\\tool\\newtech.txt"
tech=open(techfile,'rb')
#tech dictionary
techData=pickle.load(tech)

outputfilename='D:\\code\\tool\\res\\res.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)

path='D:\\code\\tool\\onlycontest'
files=os.listdir(path)

row = 0
for f in files:
	techN = 0
	typeN = 0

	bk=xlrd.open_workbook('D:\\code\\tool\\onlycontest\\'+f)
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
	columnnum=0
	dep = []
	dep2 = []
	for i in range(nrows):
		pro = str(sh.cell_value(i, 2))
		typ = str(sh.cell_value(i, 3))
		for x in techData[pro]:
			if x not in dep:
				techN += 1
				dep.append(x)
		if typ not in dep2:
			typeN += 1
			dep2.append(typ)

	sheet1.write(row, 0, f.strip('.xls'))
	sheet1.write(row, 1, techN)
	sheet1.write(row, 2, typeN)
	row += 1
outputfile.save(outputfilename)


