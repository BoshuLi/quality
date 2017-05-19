import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')


outputfilename='D:\\code\\quality\\size\\Res.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
path='D:\\code\\quality\\contest1'
files=os.listdir(path)
#open the total excel file
count = 0
for f in files:
	print f
	bk=xlrd.open_workbook(path+'\\'+f)
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
	prizenum = 6
	diffnum = 33
	typenum = 3
	prize = 0
	Diff = 0
	Ty = 0
	depT = []
	for i in range(nrows):
		p = sh.cell_value(i, prizenum)
		prize += int(p)
		d = sh.cell_value(i, diffnum)
		Diff += float(d)
		t = str(sh.cell_value(i, typenum))
		if t not in depT:
			depT.append(t)
			Ty += 1
	prize = float(prize)/10000.00000
	Diff = Diff/nrows
	Ty = float(Ty)/9.00000
	sheet1.write(count, 0, str(sh.cell_value(i, columnnum)))
	sheet1.write(count, 1, prize)
	sheet1.write(count, 2, Diff)
	sheet1.write(count, 3, Ty)
	sheet1.write(count, 4, nrows)
	count = count + 1
	
outputfile.save(outputfilename)	