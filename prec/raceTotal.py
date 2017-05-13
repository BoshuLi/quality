import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

keywordfile = open('D:\\code\\prec\\prectotal.txt')
outputfilename='D:\\code\\prec\\res\\precRes.xls'
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
''' 

#open the search file
#FilenameS='aegle-mapa'
bkS=xlrd.open_workbook('D:\\code\\prec\\'+FilenameS+'.xls')

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
outputfilename='D:\\code\\prec\\res\\'+FilenameS+'constant.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)

def findFamiliar(project,testin):	
	
	count = 0
	for i in range(nrows):
		cell_value=sh.cell_value(i, columnnum)
		title_value=sh.cell_value(i, titlenum)		
		if project != cell_value and testin == title_value:	
			count = count+1
	return count		

dep=[]
j=0
for i in range(nrowsS):
	project=shS.cell_value(i,columnnumS)
	testin=shS.cell_value(i,nameS)
	if testin not in dep:
		dep.append(testin)
		result=findFamiliar(project,testin)
		sheet1.write(j, 0 , testin)
		sheet1.write(j, 1 , result)
		j = j + 1
outputfile.save(outputfilename)
'''