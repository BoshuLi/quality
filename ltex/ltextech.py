import os
import xlrd,xlwt,sys
import pickle
reload(sys)
sys.setdefaultencoding('utf-8')

path='D:\\code\\ltex\\tech'
files=os.listdir(path)
#open the total excel file
savefile='D:\\code\\ltex\\tech.txt'
dic={}
count = 0
for f in files:
	ntech = []
	name = f.strip('.txt')
	Nfile = open(path+'\\'+f,'r+')
	line = Nfile.readline().strip('\n')
	while(line):
		ntech.append(line)
		line = Nfile.readline().strip('\n')
	dic[name] = ntech
	Nfile.close()
output = open(savefile,'wb')
pickle.dump(dic,output)
output.close()

