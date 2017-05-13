import os
import pickle
import xlrd,xlwt,sys
reload(sys)
sys.setdefaultencoding('utf-8')

#open the excel file
techfile="D:\\code\\ltex\\newtech.txt"
tech=open(techfile,'rb')
#tech dictionary
techData=pickle.load(tech)

path='D:\\code\\ltex\\tmp'
files=os.listdir(path)

for f in files:
	nowtech=open('D:\\code\\ltex\\tmp\\' + f,'rb')
	now=pickle.load(nowtech)
	Filename = f.strip('.txt')

	outputfilename='D:\\code\\ltex\\restmp\\'+Filename+'.xls'
	outputfile=xlwt.Workbook()
	sheet1 = outputfile.add_sheet('sheet1', cell_overwrite_ok=True)

	row = 0
	for key,value in now.items():
		maintech = techData[key[1]]
		for t in maintech:
			count = 0
			for pro in value:
				if t in techData[pro]:
					count += 1
			sheet1.write(row, 0, key[1])
			sheet1.write(row, 1, key[0])
			sheet1.write(row, 2, t)
			sheet1.write(row, 3, count)
			sheet1.write(row, 4, len(value))
			row += 1
	outputfile.save(outputfilename)


