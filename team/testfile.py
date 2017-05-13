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

pairs = [1,2]
memlist = [[1,3],[2,1]]
if pairs in memlist:
	 print pairs
