import os
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

#open the excel file
Filename="D:\\code\\quality\\alltest1.0.xlsx"
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
titlenum=2
# input the searching string and column

	#find the rows which you want to select and write to a txt file

def buildNew(testin):	
	outputfilename='D:\\code\\quality\\contant1.0\\'+testin + '.xls'
	outputfile=xlwt.Workbook()
	sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)
	count = 0
	dep = {}
	flag = 0
	for i in range(nrows):
		cell_value=sh.cell_value(i, columnnum)
		title_value=sh.cell_value(i, titlenum)
		if testin == str(cell_value):
			if dep.has_key(title_value):
				if dep[title_value] < 2:
					dep[title_value] += 1
					for j in range(ncols):
						sheet1.write(count, j , sh.cell_value(i,j))
					count = count+1
			else :
				dep[title_value] = 1
				for j in range(ncols):
					sheet1.write(count, j , sh.cell_value(i,j))
				count = count+1
	outputfile.save(outputfilename)


checkfile=open('D:\\code\\quality\\projectKey.txt','r')
testin=checkfile.readline().strip('\n')
while(testin):
	buildNew(testin)
	testin=checkfile.readline().strip('\n')
