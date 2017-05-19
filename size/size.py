import numpy as np  
  
def loadDataSet(fileName):  
    dataMat=[]  
    labelMat=[]  
    fr=open(fileName)  
    for line in fr.readlines():  
        curLine=line.strip().split('\t')  
        dataMat.append(list(map(float,curLine[:-1])))  
        labelMat.append(float(curLine[-1]))  
    return dataMat,labelMat  
  
def standRegress(xArr,yArr):  
    xMat=np.mat(xArr)  
    yMat=np.mat(yArr).T  
    xTx=xMat.T*xMat  
    if np.linalg.det(xTx)==0.0:  
          
        return  
    ws=xTx.I*(xMat.T*yMat)  
    return ws  
  
xArr,yArr=loadDataSet("ex1.txt")  
ws=standRegress(xArr,yArr)  
print ws