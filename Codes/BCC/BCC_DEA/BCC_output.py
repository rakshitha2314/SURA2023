from pulp import *
import xlrd
import xlwt
from xlwt import Workbook
import numpy as np
m=int(input()) #This is the number of inputs we want
s=int(input()) #This is the number of outputs we want
n=int(input()) #This is the total number of DMUs
X=[]           #This is the matrix containing the inputs
Y=[]           #This is the matrix containing the outputs 
loc = ("BCC_DEA.xls") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
wr=Workbook()
sheet1=wr.add_sheet('Sheet 1')
for i in range(n):
    sheet1.write(i+1,0,sheet.cell_value(i+1,0))
sheet1.write(0,1,'Efficiency')
for i in range(m):
    D=[]
    for j in range(n):
        D.append(sheet.cell_value(j+1,i+1))
    X.append(D)       
for i in range(s):
    C=[]
    for j in range(n):
        C.append(sheet.cell_value(j+1,i+m+1))
    Y.append(C)
for k in range(n):
    o=k+1 #The dmu which you want to evaluate       
    lamda=[0]*n
    prob=LpProblem("Efficiency Ranking",LpMaximize)
    for i in range(n):
        lamda[i]=LpVariable("lamda"+str(i),0)
    t=LpVariable("Additive",0)
    lamda=np.array(lamda)
    X=np.array(X)
    Y=np.array(Y)
    prob+= t ,"objective"
    lamda.transpose()
    lamdaX=np.matmul(X,lamda)
    lamdaY=np.matmul(Y,lamda)
    for i in range(m):
        prob+= lamdaX[i]<= X[i][o-1],str(i)+"Ayush"      
    for i in range(s):
        prob+= lamdaY[i]>= t*Y[i][o-1],str(i)+"Gupta" 
    dummy1=0
    for i in range(n):
        dummy1+=lamda[i]
    prob+= dummy1 == 1 , "constraint"
    prob.solve()
    sheet1.write(k+1,1,prob.objective.value())
wr.save('BCC_Efficiency.xls')


        





