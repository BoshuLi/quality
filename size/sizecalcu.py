import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')


outputfilename='D:\\code\\quality\\size\\sres.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
bk=xlrd.open_workbook('D:\\code\\quality\\size\\sizeres.xlsx')
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
count = 0
for i in range(nrows):
	invest = float(sh.cell_value(i, 3))
	avgDif = float(sh.cell_value(i, 4))
	completion = float(sh.cell_value(i, 2))
	numPro = float(sh.cell_value(i, 1))
	size = -24.822 + 9.331 * invest + 8.469 * avgDif + 5.452 * completion + 0.062 * numPro
	sheet1.write(count, 0, str(sh.cell_value(i, columnnum)))
	sheet1.write(count, 1, size)

	count = count + 1
outputfile.save(outputfilename)	