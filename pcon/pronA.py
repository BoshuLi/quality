import os
import pickle
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

path='D:\\code\\pcon\\constant'
files=os.listdir(path)

#outputfilename='D:\\code\\pcon\\tmp\\res.xls'
#outputfile=xlwt.Workbook()
#sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)

#count = 0
for f in files:
	dic = {}
	Filename="D:\\code\\pcon\\constant\\"+f
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
	name = f.strip('.xls')

	total = 0
	for i in range(nrows):
		pro = str(sh.cell_value(i, 2))
		user = str(sh.cell_value(i, 21))
		if dic.has_key(pro):
			dic[pro].append(user)
		else :
			dic[pro] = [user]
	#sheet1.write(count, 0, name)
	#sheet1.write(count, 1, total)
	#count +=1
	out = open('D:\\code\\pcon\\tmp\\'+name+'.txt', 'wb')
	pickle.dump(dic, out)
	out.close()
print dic
#outputfile.save(outputfilename)


