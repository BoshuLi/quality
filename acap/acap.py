import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

doucument = ['Conceptualization', 'Architecture', 'Specification', 'Design', 'Test Scenarios', 'Test Suites', 'Content Creation']

outputfilename='D:\\code\\acap\\res\\acapRes.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
path='D:\\code\\acap\\handle'
files=os.listdir(path)
#open the total excel file

count = 0
for f in files:
	countDoc = 0
	countRating = 0
	print f
	bk=xlrd.open_workbook('D:\\code\\acap\\handle\\'+f)
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
	topicnum=3
	ratingnum=22
	title = sh.cell_value(0, columnnum)
	for i in range(nrows):
		cell_value=sh.cell_value(i, topicnum)
		if str(cell_value) in doucument:
			countDoc = countDoc+1
			r = (int(sh.cell_value(i, ratingnum)) if str(sh.cell_value(i, ratingnum)) != "" else 0)
			countRating = countRating + r

	sheet1.write(count, 0, title)
	sheet1.write(count, 1, countDoc)
	sheet1.write(count, 2, countRating)
	sheet1.write(count, 3, float(countRating+1)/float(countDoc+1))
	
	count = count + 1
	
outputfile.save(outputfilename)	
