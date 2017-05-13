import os
import pickle
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

path='D:\\code\\ltex\\restmp'
files=os.listdir(path)

outputfilename='D:\\code\\ltex\\res\\res.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)

count = 0
for f in files:

	Filename="D:\\code\\ltex\\restmp\\"+f
	bk=xlrd.open_workbook(Filename)
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

	name = f.strip('.xls')

	total = 0
	for i in range(nrows):
		if str(sh.cell_value(i, 2)) != 'Not Specified':
			total += int(sh.cell_value(i, 3))
	sheet1.write(count, 0, name)
	sheet1.write(count, 1, total)
	count +=1

outputfile.save(outputfilename)


