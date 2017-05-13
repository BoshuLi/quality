import os
import pickle
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

#open the excel file
Filename="D:\\code\\ltex\\alltest.xlsx"
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

columnnum=2
tech = open('tech.txt','rb')
data = pickle.load(tech)
dic = {}

for i in range(nrows):
	num = str(sh.cell_value(i, columnnum))
	dic[num] = data[num]

print len(dic)
newtech = open('newtech.txt','wb')
pickle.dump(dic, newtech)
tech.close()
newtech.close()
