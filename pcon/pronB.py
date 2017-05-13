import os
import pickle
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

path='D:\\code\\pcon\\constant'
files=os.listdir(path)

outputfilename='D:\\code\\pcon\\res\\res.xls'
outputfile=xlwt.Workbook()
sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)

row = 0

for f in files:
	count = 0
	numclass = []
	person = []
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

	FilenameS="D:\\code\\prec\\tmp\\"+f
	bkS=xlrd.open_workbook(FilenameS)
	#get the sheets number
	shxrangeS=range(bkS.nsheets)

	for x in shxrangeS:
		pS=bkS.sheets()[x].name.encode('utf-8')
	shS=bkS.sheets()[0]
	nrowsS=shS.nrows
	ncolsS=shS.ncols


	total = 0
	pronum = 0
	pro = ''
	for i in range(nrows):
		if pro != str(sh.cell_value(i, 2)):
			pro = str(sh.cell_value(i, 2))
			pronum += 1
		cla = str(sh.cell_value(i, 3))
		user = str(sh.cell_value(i, 21))
		if (pro,cla) not in numclass:
			numclass.append((pro, cla))
			person.append([user])
		else :
			p=numclass.index((pro,cla))
			person[p].append(user)
	for value in numclass:
		ind=numclass.index(value)
		if dic.has_key(value[1]):	
			for y in person[ind]:
				if y not in dic[value[1]]:
					dic[value[1]].append(y)
			for x in dic[value[1]]:
				if x not in person[ind]:
					dic[value[1]].remove(x)
					count += 1
		else:
			dic[value[1]] = person[ind]

		#if dic.has_key(pro):
			#dic[pro].append(user)
		#lse :
			#dic[pro] = [user]
	#sheet1.write(count, 0, name)
	#sheet1.write(count, 1, total)
	#count +=1
	out = open('D:\\code\\pcon\\tmp\\'+name+'.txt', 'wb')
	pickle.dump(dic, out)
	out.close()
	sheet1.write(row, 0, name)
	sheet1.write(row, 1, count)
	sheet1.write(row, 2, nrowsS)
	sheet1.write(row, 3, pronum)
	row += 1


print numclass
print person
outputfile.save(outputfilename)


