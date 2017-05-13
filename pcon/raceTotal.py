import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

keywordfile = open('D:\\code\\prec\\prectotal.txt')
outputfilename='D:\\code\\pron\\res\\precRes.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
path='D:\\code\\prec\\res'
files=os.listdir(path)
#open the total excel file
count = 0
for f in files:
	print f
	bk=xlrd.open_workbook('D:\\code\\prec\\res\\'+f)
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
	titlenum=21

	zongji = keywordfile.readline().strip('\n')
	total=float(zongji)

	sheet1.write(count, 0, f)
	sheet1.write(count, 1, total/float(nrows-1))
	count = count + 1
	
outputfile.save(outputfilename)	
