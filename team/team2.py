import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')


path='D:\\code\\team\\handle2'
files=os.listdir(path)

projectlist = []
#open the total excel file

count = 0
for f in files:
	print f
	bk=xlrd.open_workbook('D:\\code\\team\\handle2\\'+f)
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

	projectlist.append([])
	for i in range(nrows):
		cell_value = sh.cell_value(i, 2)
		projectlist[count].append(cell_value)
	count = count + 1


bkT=xlrd.open_workbook('D:\\code\\team\\wp.xls')
#get the sheets number
shxrangeT=range(bkT.nsheets)
print shxrangeT
#get the sheets name
for x in shxrangeT:
	pT=bkT.sheets()[x].name.encode('utf-8')
	print "Sheets Number(%s): %s" %(x,pT.decode('utf-8'))
shT=bkT.sheets()[0]
nrowsT=shT.nrows
ncolsT=shT.ncols
# return the lines and col number
print "line:%d  col:%d" %(nrowsT,ncolsT)

ran = 0
for f in files:
	cf = 0
	outputfilename='D:\\code\\team\\tmp2\\'+f
	outputfile=xlwt.Workbook()
	sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
	for t in projectlist[ran]:
		for i in range(nrowsT):
			if t == shT.cell_value(i, 2):
				sheet1.write(cf, 0, shT.cell_value(i, 1))
				sheet1.write(cf, 1, shT.cell_value(i, 2))
				sheet1.write(cf, 2, shT.cell_value(i, 0))
				sheet1.write(cf, 3, shT.cell_value(i, 3))
				sheet1.write(cf, 4, shT.cell_value(i, 4))
				sheet1.write(cf, 5, shT.cell_value(i, 5))
				sheet1.write(cf, 6, shT.cell_value(i, 6))
				cf = cf + 1
	ran = ran + 1
	outputfile.save(outputfilename)	



'''
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