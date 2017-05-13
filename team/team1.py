import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

doucument = ['Conceptualization', 'Architecture', 'Specification']

outputfilename='D:\\code\\team\\res\\teamRes.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
path='D:\\code\\team\\handle'
files=os.listdir(path)

pairs = []
memlist = []
#open the total excel file

count = 0
for f in files:
	print f
	bk=xlrd.open_workbook('D:\\code\\team\\handle\\'+f)
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
	memlist.append([])
	pairs.append([])
	for i in range(nrows):
		cell_value = sh.cell_value(i, 0)
		memlist[count].append(cell_value)
		for j in range(nrows):
			cell_value2 = sh.cell_value(j, 0)
			if cell_value != cell_value2:
				pairs[count].append([cell_value,cell_value2])
	count = count + 1

groupWork = []

for i in range(len(pairs)):
	print len(pairs[i])
	groupWork.append(0)
	for j in range(len(pairs[i])):
		for k in range(len(pairs)):
			if i != k and pairs[i][j] in pairs[k]:
				groupWork[i] = groupWork[i] + 1
	sheet1.write(i, 0, files[i])			
	sheet1.write(i, 1, groupWork[i])
outputfile.save(outputfilename)	


'''
	columnnum=0
	topicnum=3
	prizenum=6
	title = sh.cell_value(0, columnnum)
	for i in range(nrows):
		cell_value=sh.cell_value(i, topicnum)
		prize_value=sh.cell_value(i, prizenum)
		if str(cell_value) in doucument:
			countDoc = countDoc + int(prize_value)
		countAll = countAll + int(prize_value)
	sheet1.write(count, 0, title)
	sheet1.write(count, 1, countDoc)
	sheet1.write(count, 2, countAll)
	sheet1.write(count, 3, float(countDoc+1)/float(countAll+1))
	
	count = count + 1
	
outputfile.save(outputfilename)	
'''