import os
import pickle
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

#open the total excel file
Filename="D:\\code\\ltex\\alltest(2).xlsx"
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
columnnum=0
titlenum=21

def findFamiliar(project,testin):	
	
	count = []
	for i in range(nrows):
		cell_value=sh.cell_value(i, columnnum)
		title_value=sh.cell_value(i, titlenum)		
		if project != cell_value and testin == title_value:	
			count.append(str(sh.cell_value(i, 2)))
	return count	

def doForall(FilenameS):
	#open the search file
	#FilenameS='aegle-mapa'
	bkS=xlrd.open_workbook('D:\\code\\ltex\\constant\\'+FilenameS+'.xls')
	#get the sheets number
	shxrangeS=range(bkS.nsheets)
	print shxrangeS
 
	#get the sheets name
	for x in shxrangeS:
		pS=bkS.sheets()[x].name.encode('utf-8')
    	print "Sheets Number(%s): %s" %(x,pS.decode('utf-8'))
	shS=bkS.sheets()[0]

	nrowsS=shS.nrows
	ncolsS=shS.ncols
	# return the lines and col number
	print "line:%d  col:%d" %(nrowsS,ncolsS)

	columnnumS=0
	nameS=21
	# input the searching string and column

		#find the rows which you want to select and write to a txt file
	#outputfilename='D:\\code\\ltex\\res\\'+FilenameS+'constant.xls'
	#outputfile=xlwt.Workbook()
	#sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
	
	#dep=[]
	#j=0
	dic = {}
	for i in range(nrowsS):
		project=shS.cell_value(i, columnnumS)
		testin=shS.cell_value(i, nameS)
		projectNum = str(shS.cell_value(i, 2))
		#if testin not in dep:
			#dep.append(testin)
		result=findFamiliar(project,testin)
			#sheet1.write(j, 0 , testin)
			#sheet1.write(j, 1 , result)
		dic[(str(testin),projectNum)] = result
		#j = j + 1
	return dic
	#outputfile.save(outputfilename)


keywordfile = open('D:\\code\\projectKey.txt')
FilenameS = keywordfile.readline().strip('\n')

while(FilenameS):
	dicN = doForall(FilenameS)
	out = open('D:\\code\\ltex\\tmp\\'+FilenameS+'.txt', 'wb')
	print len(dicN)
	pickle.dump(dicN, out)
	out.close()
	FilenameS = keywordfile.readline().strip('\n')


